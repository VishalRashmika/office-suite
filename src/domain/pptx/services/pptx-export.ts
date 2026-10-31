import { PptxPresentation, PptxSlide } from "../models/slide";

export class PptxExport {
  /**
   * Generates a clean, comprehensive Obsidian Markdown export of the entire presentation.
   */
  static toMarkdown(presentation: PptxPresentation, fileName: string): string {
    const lines: string[] = [];

    lines.push(`---`);
    lines.push(`title: "${presentation.title || fileName}"`);
    lines.push(`source: "${fileName}"`);
    lines.push(`slides_count: ${presentation.slides.length}`);
    lines.push(`type: presentation-notes`);
    lines.push(`date: ${new Date().toISOString().split("T")[0]}`);
    lines.push(`---`);
    lines.push(``);
    lines.push(`# ${presentation.title || fileName}`);
    lines.push(``);
    lines.push(`> [!abstract] Presentation Overview`);
    lines.push(`> - **Total Slides:** ${presentation.slides.length}`);
    lines.push(`> - **Dimensions:** ${presentation.width} × ${presentation.height}px`);
    lines.push(`> - **Source File:** [[${fileName}]]`);
    lines.push(``);

    // Table of contents
    lines.push(`## Table of Contents`);
    for (const slide of presentation.slides) {
      const slideTitle = slide.title || `Slide ${slide.slideNumber}`;
      const anchor = `slide-${slide.slideNumber}`;
      lines.push(`- [${slide.slideNumber}. ${slideTitle}](#${anchor})`);
    }
    lines.push(``);
    lines.push(`---`);
    lines.push(``);

    // Individual slides
    for (const slide of presentation.slides) {
      lines.push(this.slideMarkdown(slide, fileName));
      lines.push(``);
      lines.push(`---`);
      lines.push(``);
    }

    return lines.join("\n");
  }

  /**
   * Generates markdown content for a single slide.
   */
  private static slideMarkdown(slide: PptxSlide, fileName: string): string {
    const lines: string[] = [];
    const title = slide.title || `Slide ${slide.slideNumber}`;

    lines.push(`### Slide ${slide.slideNumber}: ${title} {#slide-${slide.slideNumber}}`);
    lines.push(`*[[${fileName}#slide=${slide.slideNumber}|Open Slide in Presentation View]]*`);
    lines.push(``);

    // Extract text paragraphs
    const paragraphs: string[] = [];
    for (const el of slide.elements) {
      if (el.paragraphs) {
        for (const p of el.paragraphs) {
          const runTexts = p.runs
            .map((r) => {
              let t = r.isSlideNumber ? String(slide.slideNumber) : r.text;
              if (r.bold) t = `**${t}**`;
              if (r.italic) t = `*${t}*`;
              if (r.hyperlink?.url) t = `[${t}](${r.hyperlink.url})`;
              return t;
            })
            .join("");

          const trimmed = runTexts.trim();
          if (trimmed) {
            if (p.bullet && p.bullet.type !== "none") {
              paragraphs.push(`- ${trimmed}`);
            } else {
              paragraphs.push(trimmed);
            }
          }
        }
      }
    }

    if (paragraphs.length > 0) {
      lines.push(paragraphs.join("\n\n"));
    } else {
      lines.push(`*(No text elements)*`);
    }

    // Speaker notes
    if (slide.notes && slide.notes.trim()) {
      lines.push(``);
      lines.push(`> [!note] Speaker Notes`);
      for (const noteLine of slide.notes.trim().split("\n")) {
        lines.push(`> ${noteLine}`);
      }
    }

    return lines.join("\n");
  }
}
