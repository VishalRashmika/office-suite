/**
 * Citation formats and Deep-linking generator for PDF annotations and pages.
 * Supports Markdown links with page/annot hash, APA, MLA, Chicago, and BibTeX.
 */

export type CitationStyle = "markdown" | "apa" | "mla" | "chicago" | "bibtex";

export interface CitationMetadata {
  title?: string;
  author?: string;
  year?: string | number;
  publisher?: string;
  pdfFileName: string;
  pageNumber: number;
  quotedText?: string;
  annotationId?: string;
}

export class CitationBuilder {
  /**
   * Generates a deep link to a specific page or annotation inside an Obsidian vault.
   * Format: [[Document.pdf#page=3]] or [[Document.pdf#page=3&annot=uuid]]
   */
  static buildDeepLink(metadata: CitationMetadata, customDisplayText?: string): string {
    const file = metadata.pdfFileName;
    const pageParam = `page=${metadata.pageNumber}`;
    const annotParam = metadata.annotationId ? `&annot=${metadata.annotationId}` : "";
    const hash = `#${pageParam}${annotParam}`;
    const display = customDisplayText || `${metadata.title || file}, p. ${metadata.pageNumber}`;
    return `[[${file}${hash}|${display}]]`;
  }

  /**
   * Formats a Markdown quote with deep-linked citation.
   * Example:
   * > "Machine learning models require robust evaluation metrics."
   * > — [[Attention-Is-All-You-Need.pdf#page=4|Attention Is All You Need, p. 4]]
   */
  static buildMarkdownQuote(metadata: CitationMetadata): string {
    const quote = metadata.quotedText ? metadata.quotedText.trim() : "";
    const link = this.buildDeepLink(metadata);
    if (!quote) {
      return link;
    }
    return `> "${quote}"\n> — ${link}`;
  }

  /**
   * APA 7th Edition citation format
   */
  static buildApa(metadata: CitationMetadata): string {
    const author = metadata.author || "Unknown Author";
    const year = metadata.year ? `(${metadata.year})` : "(n.d.)";
    const title = metadata.title ? `*${metadata.title}*` : `*${metadata.pdfFileName}*`;
    const page = `p. ${metadata.pageNumber}`;
    return `${author} ${year}. ${title} (${page}).`;
  }

  /**
   * MLA 9th Edition citation format
   */
  static buildMla(metadata: CitationMetadata): string {
    const author = metadata.author ? `${metadata.author}. ` : "";
    const title = metadata.title ? `"${metadata.title}." ` : `"${metadata.pdfFileName}." `;
    const year = metadata.year ? `${metadata.year}, ` : "";
    const page = `p. ${metadata.pageNumber}.`;
    return `${author}${title}${year}${page}`;
  }

  /**
   * Chicago 17th Edition (Author-Date) format
   */
  static buildChicago(metadata: CitationMetadata): string {
    const author = metadata.author || "Unknown Author";
    const year = metadata.year || "n.d.";
    const title = metadata.title ? `"${metadata.title}"` : `"${metadata.pdfFileName}"`;
    return `${author}. ${year}. ${title}, p. ${metadata.pageNumber}.`;
  }

  /**
   * BibTeX entry format
   */
  static buildBibtex(metadata: CitationMetadata): string {
    const key = (metadata.author ? metadata.author.split(" ")[0].toLowerCase() : "pdf") +
      (metadata.year || new Date().getFullYear()) +
      `_p${metadata.pageNumber}`;
    const author = metadata.author || "Unknown";
    const title = metadata.title || metadata.pdfFileName.replace(/\.pdf$/i, "");
    const year = metadata.year || new Date().getFullYear();

    return `@misc{${key},\n` +
      `  author = {${author}},\n` +
      `  title = {${title}},\n` +
      `  year = {${year}},\n` +
      `  note = {Page ${metadata.pageNumber}}\n` +
      `}`;
  }

  /**
   * Format citation based on chosen style
   */
  static format(style: CitationStyle, metadata: CitationMetadata): string {
    switch (style) {
      case "markdown":
        return this.buildMarkdownQuote(metadata);
      case "apa":
        return this.buildApa(metadata);
      case "mla":
        return this.buildMla(metadata);
      case "chicago":
        return this.buildChicago(metadata);
      case "bibtex":
        return this.buildBibtex(metadata);
      default:
        return this.buildMarkdownQuote(metadata);
    }
  }
}
