export const FONT_FAMILIES = [
  "Calibri",
  "Arial",
  "Times New Roman",
  "Georgia",
  "Courier New",
  "Verdana",
  "Segoe UI",
  "Trebuchet MS",
  "Tahoma",
  "Comic Sans MS",
];

export const FONT_SIZES = [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 32, 36, 48, 72];

export interface ButtonOptions {
  label: string;
  title: string;
  onClick: () => void;
  className?: string;
}

export interface SelectOption {
  label: string;
  value: string;
  selected?: boolean;
}

export interface SelectOptions {
  options: SelectOption[];
  onChange: (value: string) => void;
  title?: string;
  className?: string;
}

export interface ColorPickerOptions {
  label: string;
  title: string;
  initialColor: string;
  onChange: (color: string) => void;
  buttonColor?: string;
  buttonBackground?: string;
  className?: string;
}

export class ToolbarBuilder {
  static createButton(container: HTMLElement, opts: ButtonOptions): HTMLButtonElement {
    const btn = container.createEl("button", {
      text: opts.label,
      cls: `docx-toolbar-btn ${opts.className || ""}`.trim(),
    });
    btn.setAttribute("title", opts.title);
    btn.onclick = (e) => {
      e.preventDefault();
      opts.onClick();
    };
    return btn;
  }

  static createSelect(container: HTMLElement, opts: SelectOptions): HTMLSelectElement {
    const select = container.createEl("select", {
      cls: `docx-toolbar-select ${opts.className || ""}`.trim(),
    });
    if (opts.title) select.setAttribute("title", opts.title);

    for (const opt of opts.options) {
      const el = select.createEl("option", { text: opt.label, value: opt.value });
      if (opt.selected) el.selected = true;
    }

    select.onchange = (e) => {
      const val = (e.target as HTMLSelectElement).value;
      opts.onChange(val);
    };

    return select;
  }

  static createSeparator(container: HTMLElement): HTMLDivElement {
    return container.createDiv({ cls: "docx-toolbar-sep" });
  }

  static createColorPicker(container: HTMLElement, opts: ColorPickerOptions): {
    wrapper: HTMLElement;
    button: HTMLButtonElement;
    input: HTMLInputElement;
  } {
    const wrap = container.createDiv({ cls: "docx-toolbar-color-wrap" });
    const btn = wrap.createEl("button", {
      text: opts.label,
      cls: `docx-toolbar-btn ${opts.className || ""}`.trim(),
    });
    btn.setAttribute("title", opts.title);
    if (opts.buttonColor || opts.buttonBackground) {
      btn.setCssStyles({
        color: opts.buttonColor,
        background: opts.buttonBackground,
      });
    }

    const input = wrap.createEl("input", {
      type: "color",
      cls: "docx-toolbar-color-input",
      value: opts.initialColor,
    });

    input.onchange = (e) => {
      const color = (e.target as HTMLInputElement).value;
      opts.onChange(color);
    };

    return { wrapper: wrap, button: btn, input };
  }

  static createThemeToggle(
    container: HTMLElement,
    currentTheme: "light" | "dark",
    onThemeChange: (theme: "light" | "dark") => void,
    options?: {
      lightTitle?: string;
      darkTitle?: string;
    }
  ): { lightBtn: HTMLButtonElement; darkBtn: HTMLButtonElement } {
    const toggleWrap = container.createDiv({ cls: "docx-view-mode-toggle" });

    const lightBtn = toggleWrap.createEl("button", {
      text: "Light",
      cls: `docx-view-mode-btn ${currentTheme === "light" ? "is-active" : ""}`.trim(),
    });
    lightBtn.setAttribute("title", options?.lightTitle || "Light theme");
    lightBtn.onclick = () => onThemeChange("light");

    const darkBtn = toggleWrap.createEl("button", {
      text: "Dark",
      cls: `docx-view-mode-btn ${currentTheme === "dark" ? "is-active" : ""}`.trim(),
    });
    darkBtn.setAttribute("title", options?.darkTitle || "Dark theme");
    darkBtn.onclick = () => onThemeChange("dark");

    return { lightBtn, darkBtn };
  }
}
