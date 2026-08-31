import { App, Modal, Setting } from "obsidian";

export class TextPromptModal extends Modal {
  private value: string;
  private submitted = false;

  constructor(
    app: App,
    private titleText: string,
    private defaultText: string,
    private onConfirm: (val: string) => void
  ) {
    super(app);
    this.value = defaultText;
  }

  onOpen(): void {
    const { contentEl } = this;
    contentEl.empty();
    contentEl.createEl("h3", { text: this.titleText });

    new Setting(contentEl).addText((text) => {
      text.setValue(this.defaultText);
      text.onChange((val) => {
        this.value = val;
      });
      text.inputEl.focus();
      text.inputEl.select();
      text.inputEl.addEventListener("keydown", (e: KeyboardEvent) => {
        if (e.key === "Enter") {
          e.preventDefault();
          this.submitted = true;
          this.close();
          this.onConfirm(this.value);
        }
      });
    });

    new Setting(contentEl)
      .addButton((btn) => {
        btn
          .setButtonText("Confirm")
          .setCta()
          .onClick(() => {
            this.submitted = true;
            this.close();
            this.onConfirm(this.value);
          });
      })
      .addButton((btn) => {
        btn.setButtonText("Cancel").onClick(() => {
          this.close();
        });
      });
  }

  onClose(): void {
    this.contentEl.empty();
  }
}
