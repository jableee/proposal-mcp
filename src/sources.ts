// 세 폴더의 문서를 읽어 검색 가능한 형태로 만든다
import ExcelJS from "exceljs";
import JSZip from "jszip";
import MiniSearch from "minisearch";
import { readdir, readFile } from "node:fs/promises";
import { join } from "node:path";

export type Requirement = { id: string; category: string; title: string; detail: string; priority: string };
export type Section = { file: string; title: string; text: string }; // 표준 기능 정의서의 ## 소제목 단위
export type Slide = { id: string; file: string; slide: number; title: string; text: string };

export const FOLDERS = { past: "past_proposals", rfp: "rfp", features: "standard_features" } as const;

// 한국어는 조사가 붙어("데이터를") 단어 단위 매칭이 약해서 2글자 단위로 자른다
export function tokenize(s: string): string[] {
  const out: string[] = [];
  for (const w of s.toLowerCase().split(/[^\p{L}\p{N}]+/u)) {
    if (!w) continue;
    if (/\p{Script=Hangul}/u.test(w) && w.length > 2) for (let i = 0; i < w.length - 1; i++) out.push(w.slice(i, i + 2));
    else out.push(w);
  }
  return out;
}

export const listFiles = async (root: string, folder: string, ext: string) =>
  (await readdir(join(root, folder))).filter((f) => f.toLowerCase().endsWith(ext) && !f.startsWith("~$")).sort();

// 첫 시트의 머리글을 보고 열을 찾는다. 고객사마다 열 이름이 조금씩 달라서 포함 여부로 맞춘다
async function readSheet(path: string, cols: Record<string, string[]>) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(path);
  const ws = wb.worksheets[0];
  const header = Array.from(ws.getRow(1).values as unknown[], (v) => String(v ?? "")); // values는 1부터 시작하는 희소 배열
  const idx: Record<string, number> = {};
  for (const [key, names] of Object.entries(cols)) {
    const i = header.findIndex((h) => names.some((n) => h.includes(n)));
    if (i < 0) throw new Error(`${path}: '${names.join("/")}' 열을 찾지 못했습니다. 머리글: ${header.filter(Boolean).join(", ")}`);
    idx[key] = i;
  }
  const rows: Record<string, string>[] = [];
  ws.eachRow((row, n) => {
    if (n === 1) return;
    const r = Object.fromEntries(Object.entries(idx).map(([k, i]) => [k, String(row.getCell(i).text ?? "").trim()]));
    if (Object.values(r).some(Boolean)) rows.push(r);
  });
  return rows;
}

export const readRfp = (path: string) =>
  readSheet(path, {
    id: ["고유번호", "요구사항 ID", "요구사항ID", "번호"],
    category: ["분류", "구분"],
    title: ["명칭", "요구사항명"],
    detail: ["상세"],
    priority: ["중요도", "우선순위"],
  }) as Promise<Requirement[]>;

// 표준 기능 정의서는 사람이 자연어로 쓴 Markdown. ## 소제목을 기능 단위로 본다
export function splitSections(md: string, file: string): Section[] {
  return md.split(/^## /m).slice(1).map((block) => {
    const [title, ...rest] = block.split("\n");
    return { file, title: title.trim(), text: rest.join("\n").trim() };
  });
}

// 인용문 검사용: 공백·줄바꿈 차이는 무시한다
export const normalize = (s: string) => s.replace(/\s+/g, " ").trim();

// pptx는 zip 안의 slideN.xml에서 <a:t> 텍스트만 뽑는다. 첫 텍스트를 제목으로 본다
export async function readPptx(path: string, file: string): Promise<Slide[]> {
  const zip = await JSZip.loadAsync(await readFile(path));
  const names = Object.keys(zip.files).filter((n) => /^ppt\/slides\/slide\d+\.xml$/.test(n));
  names.sort((a, b) => Number(a.match(/\d+/g)!.pop()) - Number(b.match(/\d+/g)!.pop()));
  const slides: Slide[] = [];
  for (const [i, n] of names.entries()) {
    const xml = await zip.file(n)!.async("string");
    const texts = [...xml.matchAll(/<a:t>([^<]*)<\/a:t>/g)].map((m) => decodeXml(m[1]));
    slides.push({ id: `${file}#${i + 1}`, file, slide: i + 1, title: texts[0] ?? "", text: texts.slice(1).join("\n") });
  }
  return slides;
}
const decodeXml = (s: string) => s.replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, "&");

// 기존 제안서 검색. ponytail: BM25 키워드 검색만, 벡터 검색은 여기서 결과를 RRF로 합치면 된다
function index<T extends { id: string }>(docs: T[], fields: (keyof T & string)[]) {
  const ms = new MiniSearch<T>({ fields, storeFields: Object.keys(docs[0] ?? {}), tokenize, searchOptions: { tokenize, combineWith: "OR" } });
  ms.addAll(docs);
  return (q: string, k: number) => ms.search(q).slice(0, k).map((r) => ({ ...(r as unknown as T), score: Number(r.score.toFixed(2)) }));
}

export async function loadLibrary(root: string) {
  const featureFiles = await listFiles(root, FOLDERS.features, ".md");
  if (!featureFiles.length) throw new Error(`${FOLDERS.features} 폴더에 표준 기능 정의서(.md)가 없습니다`);
  const docs = await Promise.all(featureFiles.map(async (f) => ({ file: f, markdown: await readFile(join(root, FOLDERS.features, f), "utf8") })));
  const sections = docs.flatMap((d) => splitSections(d.markdown, d.file));
  const slides = (await Promise.all((await listFiles(root, FOLDERS.past, ".pptx")).map((f) => readPptx(join(root, FOLDERS.past, f), f)))).flat();
  return { docs, sections, slides, searchSlides: index(slides, ["title", "text"]) };
}
