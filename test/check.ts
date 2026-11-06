// 최소 확인: 실제 MCP 클라이언트로 서버를 띄워 승인 강제·검증 규칙이 동작하는지 본다
import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { StdioClientTransport } from "@modelcontextprotocol/sdk/client/stdio.js";
import assert from "node:assert/strict";
import { mkdtempSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { tokenize } from "../src/sources.ts";

assert.deepEqual(tokenize("데이터를 검색"), ["데이", "이터", "터를", "검색"]);

const ws = mkdtempSync(join(tmpdir(), "proposal-mcp-"));
const client = new Client({ name: "check", version: "0" }); // elicitation 미지원 클라이언트
await client.connect(new StdioClientTransport({ command: "node", args: [new URL("../src/server.ts", import.meta.url).pathname], env: { ...process.env, PROPOSAL_WORKSPACE: ws } }));
const call = async (name: string, args: Record<string, unknown> = {}) => {
  const r = await client.callTool({ name, arguments: args });
  return { error: !!r.isError, text: (r.content as { text: string }[])[0].text };
};

assert.match((await call("load_rfp")).text, /요구사항 14건/);
const first = JSON.parse((await call("next_requirement")).text);
assert.equal(first.요구사항.id, "SFR-001");
assert.match((await call("read_standard_features")).text, /## 메타데이터 수집/);

const j = (o: Record<string, unknown>) => call("record_judgment", { requirementId: "SFR-001", rationale: "검증 규칙 확인용 판단입니다", ...o });
assert.ok((await j({ decision: "standard", sections: ["없는 소제목"], quotes: ["x"] })).error, "없는 소제목 거절");
assert.ok((await j({ decision: "standard", sections: ["메타데이터 수집"], quotes: [] })).error, "standard는 인용 필수");
assert.ok((await j({ decision: "standard", sections: ["메타데이터 수집"], quotes: ["수집 실패 시 담당자에게 문자를 보냅니다"] })).error, "정의서에 없는 문장 인용 거절");
assert.ok((await j({ decision: "standard", sections: ["통합검색"], quotes: ["정해 둔 주기(보통 매일 새벽)에 자동으로 수집합니다"] })).error, "다른 소제목의 문장 인용 거절");
assert.ok((await j({ decision: "custom", sections: ["메타데이터 수집"], quotes: ["정해 둔 주기(보통 매일 새벽)에 자동으로 수집합니다."] })).error, "custom은 범위 필수");
assert.ok(!(await j({ decision: "standard", sections: ["메타데이터 수집"], quotes: ["정해 둔 주기(보통   매일 새벽)에\n자동으로 수집합니다."] })).error, "공백 차이는 허용");
assert.ok((await call("save_draft", { markdown: "# 제목\n- 승인 전에 초안을 저장하려는 시도입니다 ......" })).error, "추적표 승인 전 초안 거절");
assert.ok((await call("render_pptx")).error, "초안 없으면 PPT 거절");
assert.ok((await call("request_approval", { target: "traceability" })).error, "판단이 끝나기 전 승인 요청 거절");
await client.close();
console.log("check ok");
