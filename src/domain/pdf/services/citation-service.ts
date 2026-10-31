/**
 * Service to manage Citations and Export of PDF annotations into Obsidian Markdown notes.
 */

import { CitationBuilder, CitationMetadata, CitationStyle } from "../models/citation";
import { PdfAnnotation, rgbToHex } from "../models/pdf-annotation";

export class CitationService {
  /**
   * Generates a citation string for given metadata.
   */
  static generateCitation(style: CitationStyle, metadata: CitationMetadata): string {
    return CitationBuilder.format(style, metadata);
  }

  /**
   * Generates deep link to a specific page or annotation.
   */
  static generateReferenceLink(
    pdfFileName: string,
    pageNumber: number,
    annotationId?: string,
    displayText?: string
  ): string {
    return CitationBuilder.buildDeepLink(
      {
        pdfFileName,
        pageNumber,
        annotationId,
      },
      displayText
    );
  }

  /**
   * Generates a comprehensive Obsidian Markdown document summarizing all annotations in the PDF.
   */
  static exportAnnotationsToMarkdown(
    pdfFileName: string,
    docTitle: string | undefined,
    annotations: PdfAnnotation[],
    totalPages: number
  ): string {
    const title = docTitle || pdfFileName.replace(/\.pdf$/i, "");
    const dateStr = new Date().toISOString().split("T")[0];

    const lines: string[] = [
      `---`,
      `title: "Annotations: ${title}"`,
      `source: "[[${pdfFileName}]]"`,
      `date: ${dateStr}`,
      `total_annotations: ${annotations.length}`,
      `tags:`,
      `  - pdf-annotations`,
      `---`,
      ``,
      `# Annotations for [[${pdfFileName}|${title}]]`,
      ``,
      `*Generated on ${new Date().toLocaleString()} — ${annotations.length} annotation(s) across ${totalPages} page(s).*`,
      ``,
      `---`,
      ``,
    ];

    if (annotations.length === 0) {
      lines.push(`*No annotations found in this document.*`);
      return lines.join("\n");
    }

    // Group annotations by page (1-indexed)
    const pageMap = new Map<number, PdfAnnotation[]>();
    for (const annot of annotations) {
      const pageNum = annot.pageIndex + 1;
      const list = pageMap.get(pageNum) || [];
      list.push(annot);
      pageMap.set(pageNum, list);
    }

    const sortedPages = Array.from(pageMap.keys()).sort((a, b) => a - b);

    if (sortedPages.length > 1) {
      const pageLinks = sortedPages.map((p) => `[[#Page ${p}|Page ${p}]]`).join(" • ");
      lines.push(`**Jump to:** ${pageLinks}`);
      lines.push(``);
    }

    for (const pageNum of sortedPages) {
      const pageAnnots = pageMap.get(pageNum)!;
      lines.push(`## Page ${pageNum}`);
      lines.push(``);

      for (const annot of pageAnnots) {
        const link = this.generateReferenceLink(pdfFileName, pageNum, annot.id, `Page ${pageNum}`);
        const colorHex = rgbToHex(annot.color);

        switch (annot.type) {
          case "highlight": {
            const quote = annot.highlightedText || annot.contents || "";
            lines.push(`> [!quote] Highlight (${link})`);
            if (quote) {
              lines.push(`> "${quote}"`);
            }
            if (annot.contents && annot.contents !== quote) {
              lines.push(`> `);
              lines.push(`> **Note:** ${annot.contents}`);
            }
            if (annot.linkedNote) {
              lines.push(`> `);
              lines.push(`> **Linked Note:** [[${annot.linkedNote}]]`);
            }
            lines.push(``);
            break;
          }

          case "underline": {
            const quote = annot.highlightedText || annot.contents || "";
            lines.push(`> [!quote] Underline (${link})`);
            if (quote) {
              lines.push(`> "${quote}"`);
            }
            if (annot.contents && annot.contents !== quote) {
              lines.push(`> `);
              lines.push(`> **Note:** ${annot.contents}`);
            }
            if (annot.linkedNote) {
              lines.push(`> `);
              lines.push(`> **Linked Note:** [[${annot.linkedNote}]]`);
            }
            lines.push(``);
            break;
          }

          case "text": {
            lines.push(`> [!note] Sticky Note (${link})`);
            lines.push(`> **Author:** ${annot.author || "User"}`);
            if (annot.contents) {
              lines.push(`> ${annot.contents}`);
            }
            if (annot.linkedNote) {
              lines.push(`> `);
              lines.push(`> **Linked Note:** [[${annot.linkedNote}]]`);
            }
            lines.push(``);
            break;
          }

          case "ink": {
            lines.push(`- ✏️ **Drawing / Ink Markup** on ${link}`);
            if (annot.contents) {
              lines.push(`  - Note: ${annot.contents}`);
            }
            if (annot.linkedNote) {
              lines.push(`  - Linked Note: [[${annot.linkedNote}]]`);
            }
            lines.push(``);
            break;
          }

          case "freetext": {
            lines.push(`> [!example] Text Box (${link})`);
            if (annot.contents) {
              lines.push(`> "${annot.contents}"`);
            }
            if (annot.linkedNote) {
              lines.push(`> `);
              lines.push(`> **Linked Note:** [[${annot.linkedNote}]]`);
            }
            lines.push(``);
            break;
          }

          case "square":
          case "circle":
          case "line": {
            lines.push(`- 🔷 **${annot.type.toUpperCase()} Shape** on ${link}`);
            if (annot.contents) {
              lines.push(`  - Note: ${annot.contents}`);
            }
            if (annot.linkedNote) {
              lines.push(`  - Linked Note: [[${annot.linkedNote}]]`);
            }
            lines.push(``);
            break;
          }
        }
      }
    }

    return lines.join("\n");
  }
}
