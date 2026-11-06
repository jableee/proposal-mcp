// 승인된 Markdown 초안 → PPTX. 슬라이드는 '---'로 나누고, '# 제목', '- 항목', '  - 하위 항목'만 해석한다
import pptxgen from "pptxgenjs";

export async function renderPptx(markdown: string, fileName: string) {
  const p = new pptxgen();
  p.layout = "LAYOUT_WIDE";
  const slides = markdown.replace(/<!--[\s\S]*?-->/g, "").split(/^---\s*$/m).map((s) => s.trim()).filter(Boolean);
  for (const block of slides) {
    const lines = block.split("\n");
    const title = (lines.find((l) => /^#\s/.test(l)) ?? "").replace(/^#+\s*/, "");
    const body = lines.filter((l) => l.trim() && !/^#\s/.test(l)).map((l) => {
      const m = l.match(/^(\s*)[-*]\s+(.*)$/);
      const text = (m ? m[2] : l.trim()).replace(/\*\*(.*?)\*\*/g, "$1");
      return { text, options: m ? { bullet: true, indentLevel: m[1].length >= 2 ? 1 : 0 } : {} };
    });
    const s = p.addSlide();
    s.addText(title, { x: 0.5, y: 0.3, w: 12.3, h: 0.8, fontSize: 26, bold: true, color: "1F3864" });
    if (body.length) s.addText(body, { x: 0.7, y: 1.4, w: 11.9, h: 5.6, fontSize: 16, valign: "top", paraSpaceAfter: 6 });
  }
  await p.writeFile({ fileName });
  return slides.length;
}
