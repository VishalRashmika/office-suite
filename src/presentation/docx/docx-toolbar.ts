import { EditorView } from "prosemirror-view";
import { undo, redo } from "prosemirror-history";
import { toggleMark, setBlockType } from "prosemirror-commands";
import { liftListItem } from "prosemirror-schema-list";
import { docxSchema as schema } from "../../domain/docx";
import { ToolbarBuilder, FONT_FAMILIES, FONT_SIZES } from "../common/toolbar-builder";
import { DocxCommands } from "./docx-commands";
import { DocxTableMenu } from "./docx-table-menu";
import { DocxDocument } from "../../infrastructure/docx/docx-document";

export interface DocxToolbarDelegate {
  docxDoc: DocxDocument | null;
  paperTheme: "light" | "dark";
  setPaperTheme(theme: "light" | "dark"): void;
}

export class DocxToolbar {
  private tableMenu = new DocxTableMenu();
  public saveIndicatorEl: HTMLElement | null = null;
  public lightThemeBtn: HTMLElement | null = null;
  public darkThemeBtn: HTMLElement | null = null;

  build(container: HTMLElement, view: EditorView, delegate: DocxToolbarDelegate): void {
    container.empty();

    const addBtn = (label: string, title: string, run: () => void, cls = "") => {
      return ToolbarBuilder.createButton(container, {
        label,
        title,
        className: cls,
        onClick: () => {
          run();
          view.focus();
        },
      });
    };

    const addSep = () => ToolbarBuilder.createSeparator(container);

    // 1. History: Undo / Redo
    addBtn("↶", "Undo (Ctrl+Z)", () => undo(view.state, view.dispatch));
    addBtn("↷", "Redo (Ctrl+Y)", () => redo(view.state, view.dispatch));

    addSep();

    // 2. Style / Heading Selector
    ToolbarBuilder.createSelect(container, {
      options: [
        { label: "Normal text", value: "p" },
        { label: "Heading 1", value: "h1" },
        { label: "Heading 2", value: "h2" },
        { label: "Heading 3", value: "h3" },
        { label: "Heading 4", value: "h4" },
        { label: "Heading 5", value: "h5" },
        { label: "Heading 6", value: "h6" },
        { label: "Quote", value: "quote" },
      ],
      onChange: (val) => {
        if (val === "p") {
          setBlockType(schema.nodes.paragraph)(view.state, view.dispatch);
        } else if (val.startsWith("h")) {
          const level = parseInt(val.slice(1), 10);
          setBlockType(schema.nodes.heading, { level })(view.state, view.dispatch);
        } else if (val === "quote") {
          DocxCommands.toggleBlockquote(view);
        }
        view.focus();
      },
    });

    // 3. Font Family Selector
    ToolbarBuilder.createSelect(container, {
      options: FONT_FAMILIES.map((f) => ({ label: f, value: f })),
      onChange: (font) => {
        DocxCommands.applyMark(view, schema.marks.fontFamily, { font });
        view.focus();
      },
    });

    // 4. Font Size Selector
    ToolbarBuilder.createSelect(container, {
      options: FONT_SIZES.map((sz) => ({ label: `${sz} pt`, value: String(sz), selected: sz === 11 })),
      onChange: (val) => {
        const pt = parseInt(val, 10);
        DocxCommands.applyMark(view, schema.marks.fontSize, { pt });
        view.focus();
      },
    });

    addSep();

    // 5. Basic Formatting
    addBtn("B", "Bold (Ctrl+B)", () => toggleMark(schema.marks.bold)(view.state, view.dispatch), "bold");
    addBtn("I", "Italic (Ctrl+I)", () => toggleMark(schema.marks.italic)(view.state, view.dispatch), "italic");
    addBtn("U", "Underline (Ctrl+U)", () => toggleMark(schema.marks.underline)(view.state, view.dispatch), "underline");
    addBtn("S", "Strikethrough", () => toggleMark(schema.marks.strike)(view.state, view.dispatch), "strike");
    addBtn("x₂", "Subscript", () => toggleMark(schema.marks.subscript)(view.state, view.dispatch));
    addBtn("x²", "Superscript", () => toggleMark(schema.marks.superscript)(view.state, view.dispatch));
    addBtn("</>", "Inline Code", () => toggleMark(schema.marks.code)(view.state, view.dispatch));

    // 6. Text Color Picker
    const textColorPicker = ToolbarBuilder.createColorPicker(container, {
      label: "A",
      title: "Text color",
      initialColor: "#000000",
      buttonColor: "#007acc",
      onChange: (color) => {
        textColorPicker.button.setCssStyles({ color });
        DocxCommands.applyMark(view, schema.marks.color, { color });
        view.focus();
      },
    });
    textColorPicker.button.setCssStyles({ fontWeight: "bold" });

    // 7. Highlight Color Picker
    const hlColorPicker = ToolbarBuilder.createColorPicker(container, {
      label: "HL",
      title: "Highlight color",
      initialColor: "#ffff00",
      buttonBackground: "#fffb8f",
      onChange: (color) => {
        hlColorPicker.button.setCssStyles({ background: color });
        DocxCommands.applyMark(view, schema.marks.highlight, { color });
        view.focus();
      },
    });

    // 8. Clear Formatting
    addBtn("Tx", "Clear formatting", () => DocxCommands.clearFormatting(view));

    addSep();

    // 9. Alignment
    for (const [label, align] of [
      ["Left", "left"],
      ["Center", "center"],
      ["Right", "right"],
      ["Justify", "justify"],
    ] as const) {
      addBtn(label, `Align ${align}`, () => DocxCommands.setAlignment(view, align));
    }

    // 10. Line Spacing
    ToolbarBuilder.createSelect(container, {
      title: "Line spacing",
      options: [
        { label: "1.0", value: "1" },
        { label: "1.15", value: "1.15" },
        { label: "1.5", value: "1.5" },
        { label: "2.0", value: "2" },
      ],
      onChange: (lineSpacing) => {
        DocxCommands.setLineSpacing(view, lineSpacing);
        view.focus();
      },
    });

    // 11. Indent / Outdent
    addBtn("Outdent", "Decrease indent", () => DocxCommands.adjustIndent(view, -1));
    addBtn("Indent", "Increase indent", () => DocxCommands.adjustIndent(view, 1));

    addSep();

    // 12. Lists
    addBtn("Bullet List", "Bulleted list", () => DocxCommands.toggleList(view, schema.nodes.bullet_list));
    addBtn("Numbered List", "Numbered list", () => DocxCommands.toggleList(view, schema.nodes.ordered_list));
    addBtn("Unindent", "Lift out of list", () => {
      liftListItem(schema.nodes.list_item)(view.state, view.dispatch);
      view.focus();
    });

    addSep();

    // 13. Table Menu Dropdown
    const tableMenuWrap = container.createDiv({ cls: "docx-toolbar-color-wrap" });
    const tableMenuBtn = tableMenuWrap.createEl("button", { text: "Table ▾", cls: "docx-toolbar-btn" });
    tableMenuBtn.setAttribute("title", "Table operations");
    tableMenuBtn.onclick = (e) => {
      e.preventDefault();
      e.stopPropagation();
      this.tableMenu.toggle(tableMenuBtn, view);
    };

    // 14. Media & Breaks
    addBtn("Image", "Insert Image from file", () => {
      DocxCommands.promptInsertImage(view, delegate.docxDoc);
    });
    addBtn("Page Break", "Insert Page Break (Ctrl+Enter)", () => {
      DocxCommands.insertPageBreak(view.state, view.dispatch);
    });
    addBtn("Line", "Insert Horizontal Divider", () => DocxCommands.insertHorizontalRule(view));

    // 15. Right Side (Paper Theme Toggle + Save Indicator)
    const rightWrap = container.createDiv({ cls: "docx-toolbar-right" });

    const themeToggle = ToolbarBuilder.createThemeToggle(rightWrap, delegate.paperTheme, (theme) => {
      delegate.setPaperTheme(theme);
    }, {
      lightTitle: "Light paper (Black text on White paper)",
      darkTitle: "Dark paper (White text on Dark paper)",
    });

    this.lightThemeBtn = themeToggle.lightBtn;
    this.darkThemeBtn = themeToggle.darkBtn;

    this.saveIndicatorEl = rightWrap.createSpan({ cls: "docx-save-indicator", text: "Saved" });
  }

  updateSaveStatus(dirty: boolean): void {
    if (this.saveIndicatorEl) {
      this.saveIndicatorEl.setText(dirty ? "Saving…" : "Saved");
    }
  }

  dispose(): void {
    this.tableMenu.close();
  }
}
