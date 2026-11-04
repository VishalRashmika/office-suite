/**
 * Service to manage Citations and Export of PDF annotations into Obsidian Markdown notes.
 */

import { CitationBuilder, CitationMetadata, CitationStyle } from "../models/citation";
import { PdfAnnotation, rgbToHex } from "../models/pdf-annotation";
import { PdfColorSchemeProfile, findTagForColor } from "../models/pdf-color-scheme";

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
   * Includes color scheme legend and semantic tags when a color scheme is provided.
   */
  static exportAnnotationsToMarkdown(
    pdfFileName: string,
    docTitle: string | undefined,
    annotations: PdfAnnotation[],
    totalPages: number,
    colorScheme?: PdfColorSchemeProfile
  ): string {
    const title = docTitle || pdfFileName.replace(/\.pdf$/i, "");
    const dateStr = new Date().toISOString().split("T")[0];

    const lines: string[] = [
      `---`,
      `title: "Annotations: ${title}"`,
      `source: "[[${pdfFileName}]]"`,
      `date: ${dateStr}`,
      `total_annotations: ${annotations.length}`,
      `color_scheme: "${colorScheme?.name || "Default"}"`,
      `tags:`,
      `  - pdf-annotations`,
      `---`,
      ``,
      `# Annotations for [[${pdfFileName}|${title}]]`,
      ``,
      `*Generated on ${new Date().toLocaleString()} — ${annotations.length} annotation(s) across ${totalPages} page(s).*`,
      ``,
    ];

    if (colorScheme && colorScheme.colors && colorScheme.colors.length > 0) {
      lines.push(`### Color Scheme: ${colorScheme.name}`);
      lines.push(`| Color | Name | Tag / Meaning |`);
      lines.push(`| :---: | :--- | :--- |`);
      for (const col of colorScheme.colors) {
        lines.push(`| <span style="color:${col.hex}">■</span> | **${col.name}** | \`${col.tag || "Untagged"}\` |`);
      }
      lines.push(``);
    }

    lines.push(`---`);
    lines.push(``);

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
        const taggedColor = colorScheme ? findTagForColor(colorScheme, annot.color) : undefined;
        const tagLabel = taggedColor?.tag ? ` [${taggedColor.tag}]` : "";

        switch (annot.type) {
          case "highlight": {
            const quote = annot.highlightedText || annot.contents || "";
            lines.push(`> [!quote] Highlight${tagLabel} (${link})`);
            if (taggedColor?.tag) {
              lines.push(`> **Tag:** \`${taggedColor.tag}\` (${taggedColor.name})`);
              lines.push(`> `);
            }
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
            lines.push(`> [!quote] Underline${tagLabel} (${link})`);
            if (taggedColor?.tag) {
              lines.push(`> **Tag:** \`${taggedColor.tag}\` (${taggedColor.name})`);
              lines.push(`> `);
            }
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
            lines.push(`> [!note] Sticky Note${tagLabel} (${link})`);
            if (taggedColor?.tag) {
              lines.push(`> **Tag:** \`${taggedColor.tag}\` (${taggedColor.name})`);
            }
            lines.push(`> **Author:** ${annot.author || "User"}`);
            if (annot.contents) {
              lines.push(`> `);
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
            lines.push(`- ✏️ **Drawing${tagLabel}** on ${link}`);
            if (taggedColor?.tag) {
              lines.push(`  - Tag: \`${taggedColor.tag}\``);
            }
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
            lines.push(`> [!example] Text Box${tagLabel} (${link})`);
            if (taggedColor?.tag) {
              lines.push(`> **Tag:** \`${taggedColor.tag}\``);
            }
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
            lines.push(`- 🔷 **${annot.type.toUpperCase()} Shape${tagLabel}** on ${link}`);
            if (taggedColor?.tag) {
              lines.push(`  - Tag: \`${taggedColor.tag}\``);
            }
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

