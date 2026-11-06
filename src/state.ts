// 진행 상태는 JSON 파일 하나에 둔다. 승인은 "승인 당시 내용의 해시"로 기록해, 내용이 바뀌면 승인이 자동으로 무효가 된다
import { createHash } from "node:crypto";
import { mkdir, readFile, writeFile } from "node:fs/promises";
import { dirname } from "node:path";
import type { Requirement } from "./sources.ts";

export type Decision = "standard" | "custom" | "unsupported";
export type Judgment = { decision: Decision; featureIds: string[]; rationale: string; customScope?: string; evidence?: string[] };
export type Approval = { hash: string; at: string; via: string; comment?: string };
export type Target = "traceability" | "draft";
export type State = {
  rfpFile: string;
  requirements: Requirement[];
  judgments: Record<string, Judgment>;
  draft?: string;
  approvals: Partial<Record<Target, Approval>>;
};

export const hash = (v: unknown) => createHash("sha256").update(JSON.stringify(v)).digest("hex").slice(0, 12);
export const contentOf = (s: State, t: Target) => (t === "traceability" ? s.judgments : s.draft);
export const isApproved = (s: State, t: Target) => s.approvals[t]?.hash === hash(contentOf(s, t));

export async function load(path: string): Promise<State | null> {
  try { return JSON.parse(await readFile(path, "utf8")); } catch { return null; }
}
export async function save(path: string, s: State) {
  await mkdir(dirname(path), { recursive: true });
  await writeFile(path, JSON.stringify(s, null, 2));
}

export const DECISION_LABEL: Record<Decision, string> = { standard: "표준 충족", custom: "커스터마이징", unsupported: "제안 범위 외" };

export function traceabilityMarkdown(s: State) {
  const rows = s.requirements.map((r) => {
    const j = s.judgments[r.id];
    return `| ${r.id} | ${r.title} | ${r.priority} | ${j ? DECISION_LABEL[j.decision] : "미판단"} | ${j?.featureIds.join(", ") ?? ""} | ${j?.customScope ?? ""} | ${j?.rationale ?? ""} |`;
  });
  const count = (d: Decision) => Object.values(s.judgments).filter((j) => j.decision === d).length;
  return [
    `요구사항 ${s.requirements.length}건 · 표준 충족 ${count("standard")} · 커스터마이징 ${count("custom")} · 제안 범위 외 ${count("unsupported")} · 미판단 ${s.requirements.length - Object.keys(s.judgments).length}`,
    "",
    "| ID | 요구사항 | 중요도 | 판단 | 근거 기능 | 커스터마이징 범위 | 판단 이유 |",
    "|---|---|---|---|---|---|---|",
    ...rows,
  ].join("\n");
}

// 초안의 각 슬라이드는 <!-- req: SFR-001, SFR-002 --> 로 다루는 요구사항을 표시한다 (PPT에서는 지워짐)
export function coverage(s: State, markdown: string) {
  const referenced = new Set([...markdown.matchAll(/<!--\s*req:([^>]*)-->/g)].flatMap((m) => m[1].split(",").map((x) => x.trim()).filter(Boolean)));
  const ids = new Set(s.requirements.map((r) => r.id));
  const mustCover = s.requirements.filter((r) => s.judgments[r.id]?.decision !== "unsupported").map((r) => r.id);
  return {
    missing: mustCover.filter((id) => !referenced.has(id)),
    unknown: [...referenced].filter((id) => !ids.has(id)),
    outOfScope: [...referenced].filter((id) => s.judgments[id]?.decision === "unsupported"),
  };
}
