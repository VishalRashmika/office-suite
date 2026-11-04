/**
 * Modal to Generate and Copy Academic and Markdown Citations with PDF Deep Links.
 */

import { App, Modal, Notice, MarkdownView } from "obsidian";
import { CitationBuilder, CitationMetadata, CitationStyle } from "../../domain/pdf";

export class PdfCiteModal extends Modal {
  private metadata: CitationMetadata;

  constructor(app: App, metadata: CitationMetadata) {
    super(app);
    this.metadata = metadata;
  }

  onOpen(): void {
    const { contentEl } = this;
    contentEl.empty();
    contentEl.addClass("pdf-cite-modal");

    contentEl.createEl("h2", { text: "Cite & reference PDF" });

    if (this.metadata.quotedText) {
      const quoteBox = contentEl.createDiv({ cls: "pdf-cite-preview-quote" });
      quoteBox.createSpan({ cls: "pdf-cite-preview-label", text: "Selected text:" });
      quoteBox.createEl("p", { text: `“${this.metadata.quotedText}”` });
    }

    const formats: { style: CitationStyle; label: string; desc: string }[] = [
      {
        style: "markdown",
        label: "Obsidian Markdown Quote",
        desc: "Blockquote with deep-linked page reference",
      },
      {
        style: "apa",
        label: "APA (7th Edition)",
        desc: "Author (Year). Title (p. N).",
      },
      {
        style: "mla",
        label: "MLA (9th Edition)",
        desc: "Author. Title. Year, p. N.",
      },
      {
        style: "chicago",
        label: "Chicago (17th Edition)",
        desc: "Author. Year. Title, p. N.",
      },
      {
        style: "bibtex",
        label: "BibTeX",
        desc: "Formatted entry for LaTeX/Zotero reference managers",
      },
    ];

    const listContainer = contentEl.createDiv({ cls: "pdf-cite-list" });

    for (const fmt of formats) {
      const itemEl = listContainer.createDiv({ cls: "pdf-cite-item" });
      const infoEl = itemEl.createDiv({ cls: "pdf-cite-item-info" });
      infoEl.createEl("h4", { text: fmt.label });
      infoEl.createEl("p", { cls: "pdf-cite-item-desc", text: fmt.desc });

      const text = CitationBuilder.format(fmt.style, this.metadata);
      const preEl = itemEl.createEl("pre", { cls: "pdf-cite-text-block", text });

      const btnRow = itemEl.createDiv({ cls: "pdf-cite-btn-row" });

      const copyBtn = btnRow.createEl("button", {
        cls: "mod-primary",
        text: "Copy Citation",
      });
      copyBtn.onclick = async () => {
        await navigator.clipboard.writeText(text);
        new Notice(`Copied ${fmt.label} citation to clipboard!`);
        this.close();
      };

      const insertBtn = btnRow.createEl("button", { text: "Insert into Active Note" });
      insertBtn.onclick = () => {
        this.insertIntoActiveNote(text);
        this.close();
      };
    }

    // Direct Deep Link Section
    const deepLinkBox = contentEl.createDiv({ cls: "pdf-cite-deeplink-box" });
    deepLinkBox.createEl("h4", { text: "Obsidian Deep Link" });
    const directLink = CitationBuilder.buildDeepLink(this.metadata);
    deepLinkBox.createEl("code", { text: directLink });

    const copyDirectBtn = deepLinkBox.createEl("button", {
      cls: "mod-cta",
      text: "Copy Deep Link",
    });
    copyDirectBtn.onclick = async () => {
      await navigator.clipboard.writeText(directLink);
      new Notice("Copied Obsidian deep link!");
      this.close();
    };
  }

  private insertIntoActiveNote(text: string): void {
    const leaves = this.app.workspace.getLeavesOfType("markdown");
    if (leaves.length > 0) {
      const mdView = leaves[0].view as MarkdownView;
      if (mdView && mdView.editor) {
        mdView.editor.replaceSelection(`\n${text}\n`);
        new Notice("Inserted citation into active Markdown note.");
        return;
      }
    }
    new Notice("No active Markdown note found to insert citation.");
  }

  onClose(): void {
    this.contentEl.empty();
  }
}
