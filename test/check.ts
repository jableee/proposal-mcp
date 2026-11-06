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
assert.ok(first.표준기능후보.some((f: { id: string }) => f.id === "F-COL-01"), "수집 요구사항에 수집 기능이 후보로 나와야 함");

assert.ok((await call("record_judgment", { requirementId: "SFR-001", decision: "standard", featureIds: ["F-XXX"], rationale: "없는 기능 ID로 판단" })).error, "없는 기능 ID 거절");
assert.ok((await call("record_judgment", { requirementId: "SFR-001", decision: "custom", featureIds: ["F-COL-01"], rationale: "범위 없이 custom 판단" })).error, "custom은 범위 필수");
assert.ok((await call("save_draft", { markdown: "# 제목\n- 승인 전에 초안을 저장하려는 시도입니다 ......" })).error, "추적표 승인 전 초안 거절");
assert.ok((await call("render_pptx")).error, "초안 없으면 PPT 거절");
assert.ok((await call("request_approval", { target: "traceability" })).error, "판단이 끝나기 전 승인 요청 거절");
await client.close();
console.log("check ok");
