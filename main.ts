import {
  App,
  Editor,
  MarkdownView,
  Modal,
  Notice,
  Plugin,
  PluginSettingTab,
  Setting,
  TFile
} from "obsidian";
import {
  DEFAULT_SETTINGS,
  formatWorkIqCliAnswer,
  WorkIqSettings
} from "./src/workiq";
import { askWorkIq } from "./src/workiq-cli";

export default class WorkIqPlugin extends Plugin {
  private workIqSettings: WorkIqSettings = { ...DEFAULT_SETTINGS };

  async onload(): Promise<void> {
    await this.loadSettings();

    this.addCommand({
      id: "ask-work-iq",
      name: "Ask Microsoft Work IQ",
      editorCallback: async (editor) => {
        await this.askAndInsert(editor);
      }
    });

    this.addRibbonIcon("message-circle-question", "Ask Microsoft Work IQ", async () => {
      const activeView = this.app.workspace.getActiveViewOfType(MarkdownView);
      const editor = activeView?.editor;

      if (!editor) {
        new Notice("Open a note before searching WorkIQ.");
        return;
      }

      await this.askAndInsert(editor);
    });

    this.addSettingTab(new WorkIqSettingTab(this.app, this));
  }

  async loadSettings(): Promise<void> {
    const savedSettings = await this.loadData();

    this.workIqSettings = {
      ...DEFAULT_SETTINGS,
      workIqExecutablePath: getString(savedSettings, "workIqExecutablePath") ?? ""
    };
    await this.saveSettings();
  }

  async saveSettings(): Promise<void> {
    await this.saveData(this.workIqSettings);
  }

  getSettings(): WorkIqSettings {
    return { ...this.workIqSettings };
  }

  async updateSettings(settings: Partial<WorkIqSettings>): Promise<void> {
    this.workIqSettings = {
      ...this.workIqSettings,
      ...settings
    };
    await this.saveSettings();
  }

  private async askAndInsert(editor: Editor): Promise<void> {
    const prompt = await this.promptForQuery();

    if (!prompt) {
      return;
    }

    const targetFile = this.app.workspace.getActiveFile();
    const progressNotice = new Notice("Asking Microsoft Work IQ...", 0);

    try {
      await this.writeDiagnostic("request started");
      const answer = await askWorkIq(prompt, this.workIqSettings.workIqExecutablePath);
      await this.writeDiagnostic(`CLI returned ${answer.response.length} characters`);
      const markdown = formatWorkIqCliAnswer(prompt, answer);
      await this.insertAnswer(editor, targetFile, `${markdown}\n`);
      await this.writeDiagnostic("answer saved to note");
      new Notice("Inserted and saved the Work IQ answer.", 10000);
    } catch (error) {
      const message = error instanceof Error ? error.message : "Unknown error";
      await this.writeDiagnostic(`request failed: ${message}`);
      console.error("Work IQ request failed", error);
      new WorkIqErrorModal(this.app, message).open();
    } finally {
      progressNotice.hide();
    }
  }

  private async insertAnswer(originalEditor: Editor, targetFile: TFile | null, markdown: string): Promise<void> {
    const activeView = this.app.workspace.getActiveViewOfType(MarkdownView);
    const activeFile = this.app.workspace.getActiveFile();
    const editor = targetFile && activeFile?.path === targetFile.path ? activeView?.editor ?? originalEditor : originalEditor;
    const contentBefore = editor.getValue();

    editor.replaceSelection(markdown);

    if (editor.getValue().length > contentBefore.length) {
      return;
    }

    if (!targetFile) {
      throw new Error("Work IQ answered, but no target note is open.");
    }

    await this.app.vault.append(targetFile, `\n${markdown}`);
  }

  private async writeDiagnostic(message: string): Promise<void> {
    const pluginDirectory = this.manifest.dir;

    if (!pluginDirectory) {
      return;
    }

    const path = `${pluginDirectory}/workiq.log`;
    const line = `${new Date().toISOString()} ${message}\n`;

    try {
      if (await this.app.vault.adapter.exists(path)) {
        const currentLog = await this.app.vault.adapter.read(path);
        await this.app.vault.adapter.write(path, `${currentLog}${line}`);
      } else {
        await this.app.vault.adapter.write(path, line);
      }
    } catch (error) {
      console.error("Could not write Work IQ diagnostic log", error);
    }
  }

  private promptForQuery(): Promise<string | null> {
    return new Promise((resolve) => {
      new WorkIqSearchModal(this.app, resolve).open();
    });
  }
}

class WorkIqErrorModal extends Modal {
  constructor(app: App, private readonly message: string) {
    super(app);
  }

  onOpen(): void {
    this.contentEl.empty();
    this.contentEl.createEl("h2", { text: "Work IQ request failed" });
    this.contentEl.createEl("p", { text: this.message });
    new Setting(this.contentEl).addButton((button) =>
      button.setButtonText("Close").setCta().onClick(() => this.close())
    );
  }

  onClose(): void {
    this.contentEl.empty();
  }
}

function getString(value: unknown, key: string): string | undefined {
  if (!isRecord(value)) {
    return undefined;
  }

  const property = value[key];
  return typeof property === "string" ? property : undefined;
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
}

class WorkIqSearchModal extends Modal {
  private query = "";
  private resolved = false;

  constructor(app: App, private readonly onSubmit: (query: string | null) => void) {
    super(app);
  }

  onOpen(): void {
    const { contentEl } = this;
    contentEl.empty();
    contentEl.createEl("h2", { text: "Ask Microsoft Work IQ" });

    new Setting(contentEl)
      .setName("Prompt")
      .setDesc("Ask Microsoft 365 Copilot for an answer grounded in your work data.")
      .addText((text) => {
        text
          .setPlaceholder("Summarize the latest decisions for Project Alpha")
          .onChange((value) => {
            this.query = value;
          });
        text.inputEl.addEventListener("keydown", (event) => {
          if (event.key === "Enter") {
            this.finish(this.query);
          }
        });
        window.setTimeout(() => text.inputEl.focus(), 0);
      })
      .addButton((button) =>
        button
          .setButtonText("Ask")
          .setCta()
          .onClick(() => {
            this.finish(this.query);
          })
      );
  }

  onClose(): void {
    this.contentEl.empty();
    if (!this.resolved) {
      this.resolved = true;
      this.onSubmit(null);
    }
  }

  private finish(query: string | null): void {
    if (this.resolved) {
      return;
    }

    this.resolved = true;
    this.onSubmit(query?.trim() || null);
    this.close();
  }
}

class WorkIqSettingTab extends PluginSettingTab {
  constructor(app: App, private readonly plugin: WorkIqPlugin) {
    super(app, plugin);
  }

  display(): void {
    const { containerEl } = this;
    const settings = this.plugin.getSettings();
    containerEl.empty();

    containerEl.createEl("h2", { text: "Work IQ settings" });

    new Setting(containerEl)
      .setName("Work IQ executable")
      .setDesc("Optional path to the official Work IQ executable. Leave blank for the standard global npm location.")
      .addText((text) => {
        text
          .setPlaceholder("C:\\path\\to\\workiq.exe")
          .setValue(settings.workIqExecutablePath)
          .onChange(async (value) => {
            await this.plugin.updateSettings({ workIqExecutablePath: value.trim() });
          });
      });
  }
}
