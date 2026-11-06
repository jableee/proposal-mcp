// 전체 흐름 확인: 판단 → (CLI) 추적표 승인 → 초안 → (CLI) 초안 승인 → PPT. 임시 작업 폴더에서만 실행
import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { StdioClientTransport } from "@modelcontextprotocol/sdk/client/stdio.js";
import assert from "node:assert/strict";
import { execFileSync } from "node:child_process";
import { mkdtempSync, readFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { readPptx } from "../src/sources.ts";

const pkg = new URL("..", import.meta.url).pathname;
const ws = mkdtempSync(join(tmpdir(), "proposal-e2e-"));
const env = { ...process.env, PROPOSAL_WORKSPACE: ws };
const client = new Client({ name: "e2e", version: "0" });
await client.connect(new StdioClientTransport({ command: "node", args: [join(pkg, "src/server.ts")], env }));
const call = async (name: string, args: Record<string, unknown> = {}) => {
  const r = await client.callTool({ name, arguments: args });
  return { error: !!r.isError, text: (r.content as { text: string }[])[0].text };
};
const approve = (t: string) => execFileSync("node", [join(pkg, "src/approve.ts"), t, "테스트 승인"], { env, input: "y\n" }).toString();

const expected = JSON.parse(readFileSync(join(pkg, "samples/expected_judgments.json"), "utf8"));
const ev: Record<string, [string, string]> = {
  "SFR-001": ["메타데이터 수집", "정해 둔 주기(보통 매일 새벽)에 자동으로 수집합니다."],
  "SFR-002": ["통합검색", "테이블, 컬럼, 표준용어, 게시글을 검색창 하나에서 찾을 수 있습니다."],
  "SFR-003": ["AI 검색 (옵션 모듈)", "인터넷이 안 되는 폐쇄망에서도 동작합니다."],
  "SFR-004": ["표준용어와 표준코드 관리", "바뀐 내용은 이력으로 남습니다."],
  "SFR-005": ["데이터 이용 신청과 승인", "결재는 포털 안에서만 이뤄집니다."],
  "SFR-006": ["알림", "사내 메신저 알림은 표준에 들어 있지 않아서"],
  "SFR-007": ["데이터 분석과 리포트", "결과는 엑셀로 내려받을 수 있습니다."],
  "SFR-008": ["권한과 로그인", "컬럼 값 단위로 가리는 마스킹 기능은 표준에 없습니다."],
  "SFR-009": ["데이터 품질 진단", "분 단위의 실시간 감시는 하지 않습니다."],
  "SFR-010": ["권한과 로그인", "SAML 2.0이나 OIDC를 지원하는 사내 SSO와 연동할 수 있습니다."],
  "SFR-011": ["사용 이력과 감사 로그", "보관 기간은 설정으로 정하며 기본값은 1년입니다."],
  "PER-001": ["통합검색", "보통 1초 안팎에 결과가 나옵니다."],
  "AVL-001": ["설치와 이중화", "한 대가 멈춰도 서비스는 계속됩니다."],
  "SFR-012": ["화면과 접속 환경", "모바일 전용 앱은 없고"],
};
await call("load_rfp");
for (;;) {
  const n = JSON.parse((await call("next_requirement")).text);
  if (n.done) break;
  const id = n.요구사항.id, d = expected[id];
  const r = await call("record_judgment", { requirementId: id, decision: d, sections: [ev[id][0]], quotes: [ev[id][1]], rationale: "테스트용 정답 판단을 기록합니다", customScope: d === "custom" ? "연동 모듈 추가 개발" : undefined });
  assert.ok(!r.error, r.text);
}
assert.match((await call("get_traceability")).text, /표준 충족 9 · 커스터마이징 3 · 제안 범위 외 2/);
assert.match((await call("request_approval", { target: "traceability" })).text, /npm run approve/, "elicitation 미지원이면 CLI 안내");
approve("traceability");

const covered = Object.keys(expected).filter((k) => k !== "_note" && expected[k] !== "unsupported");
const half = covered.slice(0, 6), rest = covered.slice(6);
const bad = await call("save_draft", { markdown: `# 사업 이해\n<!-- req: ${half.join(", ")} -->\n- 테스트 초안 본문입니다` });
assert.ok(bad.error && /빠진요구사항/.test(bad.text), "요구사항 누락 시 저장 거절");
const draft = `# 사업 이해\n<!-- req: ${half.join(", ")} -->\n- 테스트 초안 본문입니다\n---\n# 요구사항별 적용 방안\n<!-- req: ${rest.join(", ")} -->\n- 항목\n  - 하위 항목`;
assert.ok(!(await call("save_draft", { markdown: draft })).error);
assert.ok((await call("render_pptx")).error, "초안 승인 전 PPT 거절");
approve("draft");
await call("record_judgment", { requirementId: "SFR-012", decision: "unsupported", sections: [], quotes: [], rationale: "승인 후 판단을 고치면 추적표 승인이 풀린다" });
assert.ok((await call("render_pptx")).error, "승인 후 판단이 바뀌면 PPT 거절");
await call("record_judgment", { requirementId: "SFR-012", decision: "unsupported", sections: [ev["SFR-012"][0]], quotes: [ev["SFR-012"][1]], rationale: "테스트용 정답 판단을 기록합니다" });
const out = await call("render_pptx");
assert.ok(!out.error, out.text);
const pptx = out.text.match(/PPT 생성: (.*\.pptx)/)![1];
const slides = await readPptx(pptx, "out");
assert.equal(slides.length, 2);
assert.ok(!slides.some((s) => s.text.includes("req:")), "요구사항 표시 주석은 PPT에서 지워져야 함");
await client.close();
console.log("e2e ok →", pptx);
