#!/usr/bin/env node
// 제안서 자동화 MCP 서버: RFP 요구사항마다 표준/커스터마이징/범위 외를 판단 → 사람 승인 → 초안 → 사람 승인 → PPT
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import ExcelJS from "exceljs";
import { join } from "node:path";
import { z } from "zod";
import { ROOT, STATE, WORKSPACE } from "./config.ts";
import { renderPptx } from "./render.ts";
import { FOLDERS, listFiles, loadLibrary, normalize, readRfp } from "./sources.ts";
import { coverage, DECISION_LABEL, hash, isApproved, load, save, traceabilityMarkdown, type State, type Target } from "./state.ts";

const lib = await loadLibrary(ROOT);
const server = new McpServer({ name: "proposal-mcp", version: "0.1.0" });

const text = (t: string) => ({ content: [{ type: "text" as const, text: t }] });
const fail = (t: string) => ({ content: [{ type: "text" as const, text: t }], isError: true });
const json = (v: unknown) => JSON.stringify(v, null, 2);
async function state(): Promise<State> {
  const s = await load(STATE);
  if (!s) throw new Error("먼저 load_rfp로 RFP를 불러오세요");
  return s;
}

server.registerTool("list_sources", {
  description: "참고 폴더 3개(기존 제안서, 이번 고객 RFP, 솔루션 표준 기능 정의서)의 파일 목록과 현재 진행 상태를 보여준다. 작업 시작 시 먼저 호출한다.",
}, async () => {
  const s = await load(STATE);
  return text(json({
    기존제안서: await listFiles(ROOT, FOLDERS.past, ".pptx"),
    RFP: await listFiles(ROOT, FOLDERS.rfp, ".xlsx"),
    표준기능정의서: await listFiles(ROOT, FOLDERS.features, ".md"),
    표준기능소제목: lib.sections.map((x) => x.title),
    기존제안서슬라이드수: lib.slides.length,
    진행상태: s ? { rfp: s.rfpFile, 요구사항: s.requirements.length, 판단완료: Object.keys(s.judgments).length, 추적표승인: isApproved(s, "traceability"), 초안: !!s.draft, 초안승인: isApproved(s, "draft") } : "시작 전",
  }));
});

server.registerTool("load_rfp", {
  description: "RFP 엑셀에서 요구사항 목록을 읽어 새 작업을 시작한다. 기존 진행 상태는 초기화된다.",
  inputSchema: { file: z.string().optional().describe("rfp 폴더의 파일명. 하나뿐이면 생략") },
}, async ({ file }) => {
  const files = await listFiles(ROOT, FOLDERS.rfp, ".xlsx");
  const f = file ?? (files.length === 1 ? files[0] : undefined);
  if (!f || !files.includes(f)) return fail(`RFP 파일을 지정하세요: ${files.join(", ")}`);
  const requirements = await readRfp(join(ROOT, FOLDERS.rfp, f));
  const ids = requirements.map((r) => r.id);
  const dup = ids.filter((id, i) => ids.indexOf(id) !== i);
  if (dup.length) return fail(`요구사항 ID가 중복됩니다: ${dup.join(", ")}`);
  await save(STATE, { rfpFile: f, requirements, judgments: {}, approvals: {} });
  return text(`${f}에서 요구사항 ${requirements.length}건을 읽었습니다. next_requirement로 하나씩 판단하세요.\n` + requirements.map((r) => `- ${r.id} [${r.category}/${r.priority}] ${r.title}`).join("\n"));
});

server.registerTool("read_standard_features", {
  description: "솔루션 표준 기능 정의서(사람이 자연어로 쓴 Markdown) 전체를 읽는다. 요구사항 판단 전에 반드시 한 번 읽고, 판단할 때 이 문서의 ## 소제목과 원문 문장을 근거로 쓴다.",
}, async () => text(lib.docs.map((d) => `<!-- 파일: ${d.file} -->\n${d.markdown}`).join("\n\n")));

server.registerTool("next_requirement", {
  description: "아직 판단하지 않은 다음 요구사항과 기존 제안서의 비슷한 슬라이드(키워드 검색 결과라 관련 없을 수 있음)를 준다. 판단은 read_standard_features로 읽은 표준 기능 정의서를 기준으로 한다. 다 끝나면 done=true."
}, async () => {
  const s = await state();
  const r = s.requirements.find((x) => !s.judgments[x.id]);
  if (!r) return text(json({ done: true, next: "get_traceability로 추적표를 확인하고 request_approval(traceability)로 사람 승인을 받으세요" }));
  const q = `${r.title} ${r.detail}`;
  return text(json({
    남은요구사항: s.requirements.filter((x) => !s.judgments[x.id]).length,
    요구사항: r,
    기존제안서후보: lib.searchSlides(q, 3).map(({ id, title, text, score }) => ({ id, title, text, score })),
    판단기준: {
      기준문서: "read_standard_features로 읽은 표준 기능 정의서. 정의서에 '별도 개발', '연동 필요', '표준에 없음'처럼 적힌 제약을 반드시 반영",
      standard: "표준 기능만으로 요구사항 전체를 충족 (설정·옵션 모듈 포함)",
      custom: "표준 기능을 바탕으로 추가 개발·연동이 필요 (customScope에 추가 개발 범위를 적는다)",
      unsupported: "솔루션 범위 밖이라 이번 제안에서 제외하거나 별도 협의가 필요",
    },
  }));
});

server.registerTool("record_judgment", {
  description: "요구사항 하나의 판단을 기록한다. 같은 ID로 다시 호출하면 판단을 고친다(이 경우 추적표 승인은 다시 받아야 한다).",
  inputSchema: {
    requirementId: z.string(),
    decision: z.enum(["standard", "custom", "unsupported"]),
    sections: z.array(z.string()).describe("근거가 되는 표준 기능 정의서의 ## 소제목 (그대로). standard·custom은 1개 이상"),
    quotes: z.array(z.string()).describe("판단 근거가 된 정의서 원문 문장을 그대로 옮긴 것. 위 소제목 안에 실제로 있어야 한다. standard·custom은 1개 이상"),
    rationale: z.string().min(10).describe("요구사항 문구와 기능 설명을 비교한 판단 이유"),
    customScope: z.string().optional().describe("custom일 때 추가 개발·연동 범위"),
    evidence: z.array(z.string()).optional().describe("참고한 기존 제안서 슬라이드 id (예: 파일명#5)"),
  },
}, async ({ requirementId, decision, sections, quotes, rationale, customScope, evidence }) => {
  const s = await state();
  if (!s.requirements.some((r) => r.id === requirementId)) return fail(`없는 요구사항 ID: ${requirementId}`);
  const cited = lib.sections.filter((x) => sections.includes(x.title));
  const unknown = sections.filter((t) => !cited.some((x) => x.title === t));
  if (unknown.length) return fail(`표준 기능 정의서에 없는 소제목: ${unknown.join(", ")}. 있는 소제목: ${lib.sections.map((x) => x.title).join(" / ")}`);
  if (decision !== "unsupported" && (!sections.length || !quotes.length)) return fail(`${decision} 판단에는 근거 소제목과 원문 인용이 1개 이상 필요합니다`);
  const notFound = quotes.filter((q) => !cited.some((x) => normalize(x.text).includes(normalize(q))));
  if (notFound.length) return fail(`인용한 문장이 근거 소제목 안에 없습니다 (원문 그대로 옮기세요): ${notFound.join(" / ")}`);
  if (decision === "custom" && !customScope?.trim()) return fail("custom 판단에는 customScope(추가 개발 범위)가 필요합니다");
  const badEvidence = (evidence ?? []).filter((id) => !lib.slides.some((x) => x.id === id));
  if (badEvidence.length) return fail(`없는 슬라이드 id: ${badEvidence.join(", ")}`);
  s.judgments[requirementId] = { decision, sections, quotes, rationale, customScope: decision === "custom" ? customScope : undefined, evidence };
  await save(STATE, s);
  const left = s.requirements.filter((r) => !s.judgments[r.id]).length;
  return text(`${requirementId}: ${DECISION_LABEL[decision]} 기록. 남은 요구사항 ${left}건.${left ? " next_requirement를 호출하세요." : " get_traceability로 넘어가세요."}`);
});

server.registerTool("get_traceability", {
  description: "요구사항 추적표(판단 결과)를 보여주고 output/요구사항추적표.xlsx로 저장한다. 커스터마이징 목록은 사람이 견적·일정을 판단할 때 쓰는 내부용이다.",
}, async () => {
  const s = await state();
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet("추적표");
  ws.addRow(["요구사항 ID", "요구사항", "중요도", "판단", "근거 소제목", "정의서 인용", "커스터마이징 범위", "판단 이유", "참고 제안서"]).font = { bold: true };
  for (const r of s.requirements) {
    const j = s.judgments[r.id];
    ws.addRow([r.id, r.title, r.priority, j ? DECISION_LABEL[j.decision] : "미판단", j?.sections.join(", "), j?.quotes.join("\n"), j?.customScope, j?.rationale, j?.evidence?.join(", ")]);
  }
  const out = join(WORKSPACE, "output", "요구사항추적표.xlsx");
  await (await import("node:fs/promises")).mkdir(join(WORKSPACE, "output"), { recursive: true });
  await wb.xlsx.writeFile(out);
  return text(`${traceabilityMarkdown(s)}\n\n저장: ${out}\n추적표 승인: ${isApproved(s, "traceability") ? "승인됨" : "필요"}`);
});

async function askHuman(target: Target, summary: string): Promise<{ ok: boolean; comment?: string; via: string } | null> {
  if (!server.server.getClientCapabilities()?.elicitation) return null;
  const res = await server.server.elicitInput({
    mode: "form",
    message: `${target === "traceability" ? "요구사항 추적표" : "제안서 초안"} 승인 요청\n\n${summary}`,
    requestedSchema: {
      type: "object",
      properties: {
        decision: { type: "string", title: "결정", enum: ["승인", "수정 요청"] },
        comment: { type: "string", title: "수정 요청 내용 (선택)" },
      },
      required: ["decision"],
    },
  });
  if (res.action !== "accept") return { ok: false, comment: "사용자가 승인 창을 닫았습니다", via: "elicitation" };
  return { ok: res.content?.decision === "승인", comment: res.content?.comment as string | undefined, via: "elicitation" };
}

server.registerTool("request_approval", {
  description: "사람에게 승인을 요청한다. traceability: 판단 결과(특히 커스터마이징·범위 외 목록), draft: 제안서 초안. LLM은 승인을 대신할 수 없고, 수정 요청이 오면 반영한 뒤 다시 요청한다.",
  inputSchema: { target: z.enum(["traceability", "draft"]) },
}, async ({ target }) => {
  const s = await state();
  if (target === "traceability" && Object.keys(s.judgments).length < s.requirements.length) return fail("모든 요구사항을 판단한 뒤에 승인을 요청하세요");
  if (target === "draft" && !s.draft) return fail("save_draft로 초안을 먼저 저장하세요");
  if (isApproved(s, target)) return text("이미 현재 내용으로 승인되어 있습니다");
  const summary = target === "traceability" ? traceabilityMarkdown(s) : `초안 ${s.draft!.split(/^---\s*$/m).length}장 · ${join(WORKSPACE, "draft.md")}`;
  const answer = await askHuman(target, summary);
  if (!answer) return text(`이 클라이언트는 승인 창(elicitation)을 지원하지 않습니다. 사람이 터미널에서 직접 승인해야 합니다:\n  npm run approve -- ${target}\n승인 후 list_sources로 상태를 확인하세요.`);
  if (!answer.ok) return text(`승인되지 않았습니다. 수정 요청: ${answer.comment ?? "(없음)"}\n${target === "traceability" ? "record_judgment로 고친 뒤" : "save_draft로 고친 뒤"} 다시 승인을 요청하세요.`);
  s.approvals[target] = { hash: hash(target === "traceability" ? s.judgments : s.draft), at: new Date().toISOString(), via: answer.via, comment: answer.comment };
  await save(STATE, s);
  return text(`승인되었습니다.${answer.comment ? ` 코멘트: ${answer.comment}` : ""} ${target === "traceability" ? "이제 제안서 초안을 작성하세요." : "render_pptx로 PPT를 만드세요."}`);
});

server.registerTool("save_draft", {
  description: "제안서 초안(Markdown)을 저장한다. 슬라이드는 '---'로 나누고 '# 제목'과 '- 항목'으로 쓴다. 각 슬라이드에 <!-- req: SFR-001, SFR-002 --> 처럼 다루는 요구사항 ID를 표시해야 한다(PPT에서는 지워짐). 제안 범위 외가 아닌 요구사항이 하나라도 빠지면 저장을 거절한다. 커스터마이징 여부는 제안서 본문에 따로 표기하지 않는다.",
  inputSchema: { markdown: z.string().min(50) },
}, async ({ markdown }) => {
  const s = await state();
  if (!isApproved(s, "traceability")) return fail("추적표가 승인되지 않았습니다. request_approval(traceability)를 먼저 받으세요");
  const c = coverage(s, markdown);
  if (c.missing.length || c.unknown.length || c.outOfScope.length)
    return fail(json({ 저장거절: "요구사항 반영 검사 실패", 빠진요구사항: c.missing, 없는ID: c.unknown, 범위외인데포함: c.outOfScope }));
  s.draft = markdown;
  await save(STATE, s);
  await (await import("node:fs/promises")).writeFile(join(WORKSPACE, "draft.md"), markdown);
  return text(`초안 저장 (${markdown.split(/^---\s*$/m).length}장, 요구사항 반영 검사 통과). ${join(WORKSPACE, "draft.md")}\nrequest_approval(draft)로 사람 승인을 받으세요.`);
});

server.registerTool("render_pptx", {
  description: "승인된 초안을 PPTX로 만든다. 초안이 승인 이후 바뀌었으면 거절한다.",
}, async () => {
  const s = await state();
  if (!s.draft) return fail("저장된 초안이 없습니다");
  if (!isApproved(s, "traceability")) return fail("초안 작성 이후 판단이 바뀌어 추적표 승인이 풀렸습니다. 추적표 승인 → 초안 수정·승인을 다시 진행하세요");
  if (!isApproved(s, "draft")) return fail("초안이 승인되지 않았거나 승인 이후 바뀌었습니다. request_approval(draft)를 받으세요");
  const out = join(WORKSPACE, "output", `${s.rfpFile.replace(/\.xlsx$/i, "")}_제안서.pptx`);
  const n = await renderPptx(s.draft, out);
  return text(`PPT 생성: ${out} (${n}장)`);
});

server.registerPrompt("create_proposal", {
  description: "RFP 한 건으로 제안서를 만드는 전체 순서",
}, () => ({
  messages: [{ role: "user", content: { type: "text", text: `proposal-mcp 도구로 제안서를 만들어 주세요. 순서를 지키고, 단계마다 무엇을 했는지 짧게 알려 주세요.

1. list_sources로 폴더와 진행 상태를 확인하고, read_standard_features로 표준 기능 정의서를 끝까지 읽은 뒤 load_rfp로 요구사항을 불러온다.
2. next_requirement → record_judgment를 done이 될 때까지 반복한다.
   - 요구사항 문구를 기준으로 판단한다. 정의서 내용이 요구사항의 일부만 충족하면 custom이다.
   - 정의서에 적힌 제약(예: "결재는 포털 안에서만", "배치로만", "표준에 없습니다")을 반드시 확인하고, 그 문장을 quotes에 원문 그대로 남긴다.
   - 기존 제안서에 비슷한 구축 사례가 있으면 evidence로 남기고 customScope 작성에 참고한다.
3. get_traceability로 추적표를 보여주고 request_approval(traceability)로 승인받는다. 수정 요청이 오면 record_judgment로 고친 뒤 다시 요청한다.
4. 승인된 판단과 기존 제안서의 표현을 참고해 초안을 쓰고 save_draft로 저장한다.
   - 구성: 사업 이해 → 제안 솔루션 개요 → 요구사항별 적용 방안 → 추진 일정 → 지원 체계
   - 커스터마이징 여부는 본문에 표기하지 않되, 표준 기능에 없는 내용을 표준인 것처럼 쓰지 않는다.
   - 제안 범위 외 요구사항은 초안에서 다루지 않는다.
   - 모든 문장은 정의서, RFP, 기존 제안서, 승인된 판단 중 하나에 근거해야 한다. 근거가 없는 약속(지원 인력, 일정, 교육 등)은 지어내지 말고 "[확인 필요: 무엇을 정해야 하는지]"로 남겨 사람이 채우게 한다.
5. request_approval(draft)로 승인받고, 승인되면 render_pptx로 PPT를 만든다.` } }],
}));

await server.connect(new StdioServerTransport());
