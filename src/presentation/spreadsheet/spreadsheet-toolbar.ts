import { CellStyle } from "../../domain/spreadsheet";
import { ToolbarBuilder, FONT_FAMILIES, FONT_SIZES } from "../common/toolbar-builder";

export interface SpreadsheetToolbarDelegate {
  paperTheme: "light" | "dark";
  setTheme(theme: "light" | "dark"): void;
  onUndo(): void;
  onRedo(): void;
  onApplyFormat(format: string): void;
  onAdjustDecimals(delta: number): void;
  onApplyStyle(style: Partial<CellStyle>): void;
  onToggleStyle(prop: keyof CellStyle): void;
  onToggleBorders(): void;
  onInsertRow(): void;
  onInsertColumn(): void;
  onDeleteRow(): void;
  onDeleteColumn(): void;
  onInsertFunction(fn: string): void;
}

export class SpreadsheetToolbar {
  public saveIndicatorEl: HTMLElement | null = null;
  public lightThemeBtn: HTMLElement | null = null;
  public darkThemeBtn: HTMLElement | null = null;

  build(container: HTMLElement, delegate: SpreadsheetToolbarDelegate): void {
    container.empty();

    const addBtn = (label: string, title: string, run: () => void, cls = "") => {
      return ToolbarBuilder.createButton(container, {
        label,
        title,
        className: cls,
        onClick: run,
      });
    };

    const addSep = () => ToolbarBuilder.createSeparator(container);

    // 1. History: Undo / Redo
    addBtn("↶", "Undo (Ctrl+Z)", () => delegate.onUndo());
    addBtn("↷", "Redo (Ctrl+Y)", () => delegate.onRedo());

    addSep();

    // 2. Number Formatting
    addBtn("$", "Format as currency ($)", () => delegate.onApplyFormat("currency"));
    addBtn("%", "Format as percent (%)", () => delegate.onApplyFormat("percent"));
    addBtn(".0", "Decrease decimal places", () => delegate.onAdjustDecimals(-1));
    addBtn(".00", "Increase decimal places", () => delegate.onAdjustDecimals(1));

    addSep();

    // 3. Font Family Selector
    ToolbarBuilder.createSelect(container, {
      options: FONT_FAMILIES.map((f) => ({ label: f, value: f })),
      onChange: (font) => delegate.onApplyStyle({ fontFamily: font }),
    });

    // 4. Font Size Selector
    ToolbarBuilder.createSelect(container, {
      options: FONT_SIZES.map((sz) => ({ label: `${sz}`, value: String(sz), selected: sz === 10 })),
      onChange: (val) => {
        const sz = parseInt(val, 10);
        delegate.onApplyStyle({ fontSize: sz });
      },
    });

    addSep();

    // 5. Bold, Italic, Underline, Strike
    addBtn("B", "Bold (Ctrl+B)", () => delegate.onToggleStyle("bold"), "bold");
    addBtn("I", "Italic (Ctrl+I)", () => delegate.onToggleStyle("italic"), "italic");
    addBtn("U", "Underline (Ctrl+U)", () => delegate.onToggleStyle("underline"), "underline");
    addBtn("S", "Strikethrough", () => delegate.onToggleStyle("strike"), "strike");

    // 6. Text Color Picker
    const textColorPicker = ToolbarBuilder.createColorPicker(container, {
      label: "A",
      title: "Text color",
      initialColor: "#000000",
      buttonColor: "#007acc",
      onChange: (color) => {
        textColorPicker.button.setCssStyles({ color });
        delegate.onApplyStyle({ color });
      },
    });
    textColorPicker.button.setCssStyles({ fontWeight: "bold" });

    // 7. Fill Color Picker
    ToolbarBuilder.createColorPicker(container, {
      label: "Fill",
      title: "Fill color",
      initialColor: "#e6f4ea",
      onChange: (bg) => delegate.onApplyStyle({ background: bg }),
    });

    // 8. Borders Menu
    addBtn("Borders", "Toggle borders", () => delegate.onToggleBorders());

    addSep();

    // 9. Alignment: Left, Center, Right
    for (const [label, align] of [
      ["Left", "left"],
      ["Center", "center"],
      ["Right", "right"],
    ] as const) {
      addBtn(label, `Align ${align}`, () => delegate.onApplyStyle({ align }));
    }

    addSep();

    // 10. Insert / Delete Row / Col
    addBtn("+ Row", "Insert row below", () => delegate.onInsertRow());
    addBtn("+ Col", "Insert column right", () => delegate.onInsertColumn());
    addBtn("Delete Row", "Delete selected row", () => delegate.onDeleteRow());
    addBtn("Delete Col", "Delete selected column", () => delegate.onDeleteColumn());

    addSep();

    // 11. Quick Functions Dropdown
    const fnSelect = ToolbarBuilder.createSelect(container, {
      options: [
        { label: "Functions", value: "" },
        ...["SUM", "AVERAGE", "COUNT", "MAX", "MIN", "IF", "CONCAT", "ROUND"].map((fn) => ({
          label: fn,
          value: fn,
        })),
      ],
      onChange: (fn) => {
        if (!fn) return;
        delegate.onInsertFunction(fn);
        fnSelect.value = "";
      },
    });

    // 12. Right Side: Theme Toggle + Save Indicator
    const rightWrap = container.createDiv({ cls: "docx-toolbar-right" });

    const themeToggle = ToolbarBuilder.createThemeToggle(
      rightWrap,
      delegate.paperTheme,
      (theme) => delegate.setTheme(theme),
      {
        lightTitle: "Light theme",
        darkTitle: "Dark theme",
      }
    );

    this.lightThemeBtn = themeToggle.lightBtn;
    this.darkThemeBtn = themeToggle.darkBtn;

    this.saveIndicatorEl = rightWrap.createSpan({ cls: "docx-save-indicator", text: "Saved" });
  }

  updateSaveStatus(dirty: boolean): void {
    if (this.saveIndicatorEl) {
      this.saveIndicatorEl.setText(dirty ? "Saving…" : "Saved");
    }
  }
}
