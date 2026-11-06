// elicitation을 지원하지 않는 클라이언트용: 사람이 터미널에서 직접 승인한다
import { createInterface } from "node:readline/promises";
import { join } from "node:path";
import { STATE, WORKSPACE } from "./config.ts";
import { hash, isApproved, load, save, traceabilityMarkdown, type Target } from "./state.ts";

const target = process.argv[2] as Target;
const comment = process.argv[3]; // 선택: npm run approve -- draft "코멘트"
if (target !== "traceability" && target !== "draft") { console.error('사용법: npm run approve -- traceability|draft ["코멘트"]'); process.exit(1); }
const s = await load(STATE);
if (!s) { console.error("진행 중인 작업이 없습니다"); process.exit(1); }
if (target === "draft" && !s.draft) { console.error("저장된 초안이 없습니다"); process.exit(1); }
if (isApproved(s, target)) { console.log("이미 승인되어 있습니다"); process.exit(0); }
console.log(target === "traceability" ? traceabilityMarkdown(s) : `초안: ${join(WORKSPACE, "draft.md")}\n\n${s.draft}`);
const rl = createInterface({ input: process.stdin, output: process.stdout });
const ans = (await rl.question("\n승인할까요? (y/N) ")).trim().toLowerCase();
rl.close();
if (ans !== "y") { console.log("승인하지 않았습니다"); process.exit(0); }
s.approvals[target] = { hash: hash(target === "traceability" ? s.judgments : s.draft), at: new Date().toISOString(), via: "cli", comment };
await save(STATE, s);
console.log("승인했습니다");
