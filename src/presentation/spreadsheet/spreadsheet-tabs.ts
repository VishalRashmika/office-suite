import { WorkbookData } from "../../domain/spreadsheet";

export interface SpreadsheetTabsDelegate {
  onSelectSheet(index: number): void;
  onRenameSheet(index: number, currentName: string): void;
  onAddSheet(): void;
}

export class SpreadsheetTabs {
  render(container: HTMLElement, workbook: WorkbookData, delegate: SpreadsheetTabsDelegate): void {
    container.empty();

    const tabsWrap = container.createDiv({ cls: "sheet-tabs-list" });
    for (let i = 0; i < workbook.sheets.length; i++) {
      const sheet = workbook.sheets[i];
      const isActive = i === (workbook.activeSheetIndex || 0);

      const tab = tabsWrap.createDiv({
        cls: `sheet-tab ${isActive ? "is-active" : ""}`.trim(),
        text: sheet.name,
      });

      tab.onclick = () => {
        delegate.onSelectSheet(i);
      };

      tab.ondblclick = (e) => {
        e.stopPropagation();
        delegate.onRenameSheet(i, sheet.name);
      };
    }

    // Add Sheet Button (+)
    const addSheetBtn = container.createEl("button", {
      text: "+",
      cls: "sheet-add-tab-btn",
    });
    addSheetBtn.setAttribute("title", "Add new sheet");
    addSheetBtn.onclick = () => {
      delegate.onAddSheet();
    };
  }
}
