import {
  App,
  Editor,
  MarkdownView,
  Modal,
  Notice,
  Plugin,
  PluginSettingTab,
  requestUrl,
  Setting
} from "obsidian";
import {
  buildWorkIqSearchRequest,
  DEFAULT_SETTINGS,
  flattenWorkIqHits,
  formatWorkIqHits,
  parseWorkIqSearchResponse,
  WorkIqSearchSettings
} from "./src/workiq";

export default class WorkIqPlugin extends Plugin {
  private workIqSettings: WorkIqSearchSettings = { ...DEFAULT_SETTINGS };

  async onload(): Promise<void> {
    await this.loadSettings();

    this.addCommand({
      id: "search-workiq",
      name: "Search Microsoft WorkIQ",
      editorCallback: async (editor) => {
        await this.searchAndInsert(editor);
      }
    });

    this.addRibbonIcon("search", "Search Microsoft WorkIQ", async () => {
      const activeView = this.app.workspace.getActiveViewOfType(MarkdownView);
      const editor = activeView?.editor;

      if (!editor) {
        new Notice("Open a note before searching WorkIQ.");
        return;
      }

      await this.searchAndInsert(editor);
    });

    this.addSettingTab(new WorkIqSettingTab(this.app, this));
  }

  async loadSettings(): Promise<void> {
    this.workIqSettings = {
      ...DEFAULT_SETTINGS,
      entityTypes: [...DEFAULT_SETTINGS.entityTypes],
      ...(await this.loadData())
    };
  }

  async saveSettings(): Promise<void> {
    await this.saveData(this.workIqSettings);
  }

  getSettings(): WorkIqSearchSettings {
    return {
      ...this.workIqSettings,
      entityTypes: [...this.workIqSettings.entityTypes]
    };
  }

  async updateSettings(settings: Partial<WorkIqSearchSettings>): Promise<void> {
    this.workIqSettings = {
      ...this.workIqSettings,
      ...settings,
      entityTypes: settings.entityTypes ? [...settings.entityTypes] : [...this.workIqSettings.entityTypes]
    };
    await this.saveSettings();
  }

  private async searchAndInsert(editor: Editor): Promise<void> {
    if (!this.workIqSettings.accessToken.trim()) {
      new Notice("Add a Microsoft Graph access token in WorkIQ settings first.");
      return;
    }

    const query = await this.promptForQuery();

    if (!query) {
      return;
    }

    try {
      const response = await requestUrl({
        url: this.workIqSettings.graphSearchEndpoint,
        method: "POST",
        headers: {
          Authorization: ["Bearer", this.workIqSettings.accessToken.trim()].join(" "),
          "Content-Type": "application/json"
        },
        body: JSON.stringify(buildWorkIqSearchRequest(query, this.workIqSettings)),
        throw: false
      });

      const responseBody = readJsonResponse(response);

      if (response.status < 200 || response.status >= 300) {
        throw new Error(getHttpErrorMessage(response.status, responseBody));
      }

      const searchResponse = parseWorkIqSearchResponse(responseBody);
      const markdown = formatWorkIqHits(query, flattenWorkIqHits(searchResponse));
      editor.replaceSelection(`${markdown}\n`);
      new Notice("Inserted WorkIQ search results.");
    } catch (error) {
      const message = error instanceof Error ? error.message : "Unknown error";
      new Notice(`WorkIQ search failed: ${message}`);
    }
  }

  private promptForQuery(): Promise<string | null> {
    return new Promise((resolve) => {
      new WorkIqSearchModal(this.app, resolve).open();
    });
  }
}

function getHttpErrorMessage(status: number, responseBody: unknown): string {
  const statusMessage = `Microsoft Graph returned HTTP ${status}.`;

  if (responseBody === undefined) {
    return statusMessage;
  }

  try {
    parseWorkIqSearchResponse(responseBody);
    return statusMessage;
  } catch (error) {
    const detail = error instanceof Error ? error.message : undefined;
    return detail ? `${statusMessage} ${detail}` : statusMessage;
  }
}

function readJsonResponse(response: { json: unknown }): unknown {
  return response.json;
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
    contentEl.createEl("h2", { text: "Search Microsoft WorkIQ" });

    new Setting(contentEl)
      .setName("Search query")
      .setDesc("Find Microsoft 365 context to insert into the active note.")
      .addText((text) => {
        text
          .setPlaceholder("Project roadmap, recent planning mail, ...")
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
          .setButtonText("Search")
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

    containerEl.createEl("h2", { text: "WorkIQ settings" });

    new Setting(containerEl)
      .setName("Microsoft Graph access token")
      .setDesc("Paste a delegated Microsoft Graph token that can call Microsoft Search.")
      .addText((text) => {
        text
          .setPlaceholder("eyJ...")
          .setValue(settings.accessToken)
          .onChange(async (value) => {
            await this.plugin.updateSettings({ accessToken: value.trim() });
          });
        text.inputEl.type = "password";
      });

    new Setting(containerEl)
      .setName("Microsoft Graph search endpoint")
      .setDesc("Use the default Microsoft Search endpoint unless your environment requires another Graph cloud.")
      .addText((text) =>
        text
          .setValue(settings.graphSearchEndpoint)
          .onChange(async (value) => {
            await this.plugin.updateSettings({
              graphSearchEndpoint: value.trim() || DEFAULT_SETTINGS.graphSearchEndpoint
            });
          })
      );

    new Setting(containerEl)
      .setName("Entity types")
      .setDesc("Comma-separated Microsoft Search entity types, such as driveItem, message, event.")
      .addText((text) =>
        text
          .setValue(settings.entityTypes.join(", "))
          .onChange(async (value) => {
            await this.plugin.updateSettings({
              entityTypes: value
                .split(",")
                .map((entityType) => entityType.trim())
                .filter(Boolean)
            });
          })
      );

    new Setting(containerEl)
      .setName("Maximum results")
      .setDesc("The number of WorkIQ search results to insert into the active note.")
      .addSlider((slider) =>
        slider
          .setLimits(1, 25, 1)
          .setValue(settings.maxResults)
          .setDynamicTooltip()
          .onChange(async (value) => {
            await this.plugin.updateSettings({ maxResults: value });
          })
      );
  }
}
