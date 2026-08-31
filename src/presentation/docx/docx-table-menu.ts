import { EditorView } from "prosemirror-view";
import { DocxCommands } from "./docx-commands";

export class DocxTableMenu {
  private activeModal: HTMLElement | null = null;

  toggle(anchorBtn: HTMLElement, view: EditorView): void {
    if (this.activeModal) {
      this.close();
      return;
    }

    const modal = document.body.createDiv({ cls: "docx-table-menu-modal" });
    this.activeModal = modal;
    const rect = anchorBtn.getBoundingClientRect();
    modal.setCssStyles({
      top: `${rect.bottom + 4}px`,
      left: `${rect.left}px`,
    });

    const addItem = (label: string, action: () => void, isDanger = false) => {
      const item = modal.createEl("button", {
        text: label,
        cls: `docx-table-menu-item ${isDanger ? "danger" : ""}`.trim(),
      });
      item.onclick = (ev) => {
        ev.preventDefault();
        ev.stopPropagation();
        this.close();
        action();
        view.focus();
      };
    };

    addItem("Insert 3x3 Table", () => DocxCommands.insertTable(view, 3, 3));
    addItem("Add Row Below", () => DocxCommands.addTableRow(view, true));
    addItem("Add Row Above", () => DocxCommands.addTableRow(view, false));
    addItem("Add Column Right", () => DocxCommands.addTableColumn(view, true));
    addItem("Add Column Left", () => DocxCommands.addTableColumn(view, false));
    addItem("Cell Shading (Blue)", () => DocxCommands.setCellBackground(view, "#EBF3FB"));
    addItem("Cell Shading (Gray)", () => DocxCommands.setCellBackground(view, "#F2F2F2"));
    addItem("Cell Shading (Yellow)", () => DocxCommands.setCellBackground(view, "#FFF8E1"));
    addItem("Cell Shading (Green)", () => DocxCommands.setCellBackground(view, "#E8F5E9"));
    addItem("Cell Shading (Clear)", () => DocxCommands.setCellBackground(view, ""));
    addItem("Delete Row", () => DocxCommands.deleteTableRow(view), true);
    addItem("Delete Column", () => DocxCommands.deleteTableColumn(view), true);
    addItem("Delete Table", () => DocxCommands.deleteTable(view), true);

    const onDocClick = (evt: MouseEvent) => {
      if (this.activeModal && !this.activeModal.contains(evt.target as Node)) {
        this.close();
        document.removeEventListener("click", onDocClick);
      }
    };
    window.setTimeout(() => document.addEventListener("click", onDocClick), 10);
  }

  close(): void {
    if (this.activeModal) {
      this.activeModal.remove();
      this.activeModal = null;
    }
  }
}
