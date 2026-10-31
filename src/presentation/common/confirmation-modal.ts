import { App, Modal, Setting } from "obsidian";

export class ConfirmationModal extends Modal {
  constructor(
    app: App,
    private titleText: string,
    private messageText: string,
    private confirmButtonText: string,
    private onConfirm: () => void | Promise<void>
  ) {
    super(app);
  }

  onOpen(): void {
    const { contentEl } = this;
    contentEl.empty();
    contentEl.createEl("h3", { text: this.titleText });
    contentEl.createEl("p", { text: this.messageText });

    new Setting(contentEl)
      .addButton((btn) => {
        btn
          .setButtonText(this.confirmButtonText)
          .setWarning()
          .onClick(async () => {
            this.close();
            await this.onConfirm();
          });
      })
      .addButton((btn) => {
        btn.setButtonText("Cancel").onClick(() => {
          this.close();
        });
      });
  }

  onClose(): void {
    const { contentEl } = this;
    contentEl.empty();
  }
}
