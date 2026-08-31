import { FileView, TFile, WorkspaceLeaf, Notice } from "obsidian";
import { EditorState, Transaction } from "prosemirror-state";
import { EditorView } from "prosemirror-view";
import { history, undo, redo } from "prosemirror-history";
import { keymap } from "prosemirror-keymap";
import { baseKeymap, toggleMark } from "prosemirror-commands";
import { splitListItem, sinkListItem, liftListItem } from "prosemirror-schema-list";
import { inputRules, wrappingInputRule, textblockTypeInputRule } from "prosemirror-inputrules";
import { docxSchema as schema, docxToProseMirrorDoc, proseMirrorDocToDocxBody } from "../../domain/docx";
import { DocxDocument } from "../../infrastructure/docx/docx-document";
import { DocxToolbar, DocxToolbarDelegate } from "./docx-toolbar";
import { DocxCommands } from "./docx-commands";

import { DocxImageView } from "./docx-image-view";

export const VIEW_TYPE_DOCX = "docx-editor-view";

const docxInputRules = inputRules({
  rules: [
    wrappingInputRule(/^\s*([-+*])\s$/, schema.nodes.bullet_list),
    wrappingInputRule(
      /^(\d+)\.\s$/,
      schema.nodes.ordered_list,
      (match) => ({ order: +match[1] }),
      (match, node) => node.childCount + (node.attrs.order || 1) == +match[1]
    ),
    wrappingInputRule(/^\s*>\s$/, schema.nodes.blockquote),
    textblockTypeInputRule(/^(#{1,6})\s$/, schema.nodes.heading, (match) => ({
      level: match[1].length,
    })),
  ],
});

export class DocxView extends FileView implements DocxToolbarDelegate {
  allowNoFile = false;
  editorView: EditorView | null = null;
  docxDoc: DocxDocument | null = null;
  toolbarEl: HTMLElement | null = null;
  editorHostEl: HTMLElement | null = null;
  dirty = false;
  paperTheme: "light" | "dark" = "light";

  private toolbar = new DocxToolbar();
  private saveTimer: number | null = null;

  constructor(leaf: WorkspaceLeaf) {
    super(leaf);
  }

  getViewType(): string {
    return VIEW_TYPE_DOCX;
  }

  getIcon(): string {
    return "file-text";
  }

  getDisplayText(): string {
    return this.file?.basename ?? "Docx";
  }

  canAcceptExtension(extension: string): boolean {
    return extension === "docx";
  }

  async onLoadFile(file: TFile): Promise<void> {
    this.contentEl.empty();
    this.contentEl.addClass("docx-editor-root");

    const status = this.contentEl.createDiv({ cls: "docx-editor-status", text: "Loading document..." });

    let bytes: ArrayBuffer;
    try {
      bytes = await this.app.vault.readBinary(file);
      this.docxDoc = await DocxDocument.load(bytes);
    } catch (err) {
      status.setText(`Could not open this .docx file: ${(err as Error).message}`);
      console.error(err);
      return;
    }

    const pmDoc = await docxToProseMirrorDoc(this.docxDoc);

    status.remove();
    this.toolbarEl = this.contentEl.createDiv({ cls: "docx-editor-toolbar" });
    this.editorHostEl = this.contentEl.createDiv({ cls: "docx-editor-host" });

    this.setPaperTheme(this.paperTheme);

    const insertPageBreakCmd = (state: EditorState, dispatch?: (tr: Transaction) => void) => {
      return DocxCommands.insertPageBreak(state, dispatch);
    };

    const state = EditorState.create({
      doc: pmDoc,
      schema,
      plugins: [
        history(),
        keymap({
          Enter: splitListItem(schema.nodes.list_item),
          Tab: sinkListItem(schema.nodes.list_item),
          "Shift-Tab": liftListItem(schema.nodes.list_item),
          "Mod-z": undo,
          "Mod-y": redo,
          "Mod-Shift-z": redo,
          "Mod-b": toggleMark(schema.marks.bold),
          "Mod-i": toggleMark(schema.marks.italic),
          "Mod-u": toggleMark(schema.marks.underline),
          "Mod-Shift-s": toggleMark(schema.marks.strike),
          "Mod-`": toggleMark(schema.marks.code),
          "Mod-Enter": insertPageBreakCmd,
        }),
        docxInputRules,
        keymap(baseKeymap),
      ],
    });

    this.editorView = new EditorView(this.editorHostEl, {
      state,
      nodeViews: {
        image: (node, view, getPos) => new DocxImageView(node, view, getPos as any, this.docxDoc),
      },
      dispatchTransaction: (tr: Transaction) => {
        if (!this.editorView) return;
        const newState = this.editorView.state.apply(tr);
        this.editorView.updateState(newState);
        if (tr.docChanged) this.scheduleSave();
        this.updateToolbarStatus();
      },
      handleDOMEvents: {
        blur: () => {
          if (this.dirty) this.saveNow();
          return false;
        },
        paste: (view, event) => {
          const files = event.clipboardData?.files;
          if (files && files.length > 0) {
            for (let i = 0; i < files.length; i++) {
              const file = files[i];
              if (file.type.startsWith("image/")) {
                event.preventDefault();
                DocxCommands.insertImageFile(view, this.docxDoc, file);
                return true;
              }
            }
          }
          return false;
        },
        drop: (view, event) => {
          const files = event.dataTransfer?.files;
          if (files && files.length > 0) {
            for (let i = 0; i < files.length; i++) {
              const file = files[i];
              if (file.type.startsWith("image/")) {
                event.preventDefault();
                const coords = { left: event.clientX, top: event.clientY };
                const pos = view.posAtCoords(coords);
                if (pos) {
                  const selTr = view.state.tr.setSelection(
                    (view.state.selection.constructor as any).near(view.state.doc.resolve(pos.pos))
                  );
                  view.dispatch(selTr);
                }
                DocxCommands.insertImageFile(view, this.docxDoc, file);
                return true;
              }
            }
          }
          return false;
        },
      },
    });

    window.addEventListener("blur", this.onWindowBlur);
    this.toolbar.build(this.toolbarEl, this.editorView, this);
    this.setPaperTheme(this.paperTheme);

    // Trigger initial render of pagination decorations after DOM paints
    setTimeout(() => {
      if (this.editorView) {
        this.editorView.dispatch(this.editorView.state.tr);
      }
    }, 100);
  }

  private onWindowBlur = () => {
    if (this.dirty) this.saveNow();
  };

  async onUnloadFile(file: TFile): Promise<void> {
    window.removeEventListener("blur", this.onWindowBlur);
    if (this.saveTimer) window.clearTimeout(this.saveTimer);
    if (this.dirty) await this.saveNow();
    this.toolbar.dispose();
    this.editorView?.destroy();
    this.editorView = null;
    this.docxDoc = null;
  }

  setPaperTheme(theme: "light" | "dark"): void {
    this.paperTheme = theme;
    if (theme === "light") {
      this.contentEl.removeClass("docx-theme-dark");
      this.contentEl.addClass("docx-theme-light");
      this.toolbar.lightThemeBtn?.addClass("is-active");
      this.toolbar.darkThemeBtn?.removeClass("is-active");
    } else {
      this.contentEl.removeClass("docx-theme-light");
      this.contentEl.addClass("docx-theme-dark");
      this.toolbar.darkThemeBtn?.addClass("is-active");
      this.toolbar.lightThemeBtn?.removeClass("is-active");
    }
  }

  private scheduleSave(): void {
    this.dirty = true;
    if (this.saveTimer) window.clearTimeout(this.saveTimer);
    this.saveTimer = window.setTimeout(() => this.saveNow(), 800);
  }

  private async saveNow(): Promise<void> {
    if (!this.editorView || !this.docxDoc || !this.file) return;
    try {
      const bodyChildren = proseMirrorDocToDocxBody(this.docxDoc, this.editorView.state.doc);
      const bytes = await this.docxDoc.save(bodyChildren);
      await this.app.vault.modifyBinary(this.file, bytes);
      this.dirty = false;
    } catch (err) {
      console.error("Failed to save .docx:", err);
      new Notice(`Failed to save ${this.file.name}: ${(err as Error).message}`);
    } finally {
      this.updateToolbarStatus();
    }
  }

  /** Force an immediate save, e.g. before the leaf/file closes. */
  async requestSave(): Promise<void> {
    if (this.saveTimer) window.clearTimeout(this.saveTimer);
    if (this.dirty) await this.saveNow();
  }

  private updateToolbarStatus(): void {
    this.toolbar.updateSaveStatus(this.dirty);
  }
}
