import {
  App,
  Editor,
  MarkdownView,
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
  WorkIqSearchResponse,
  WorkIqSearchSettings
} from "./src/workiq";

export default class WorkIqPlugin extends Plugin {
  workIqSettings: WorkIqSearchSettings = DEFAULT_SETTINGS;

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
      ...(await this.loadData())
    };
  }

  async saveSettings(): Promise<void> {
    await this.saveData(this.workIqSettings);
  }

  private async searchAndInsert(editor: Editor): Promise<void> {
    if (!this.workIqSettings.accessToken.trim()) {
      new Notice("Add a Microsoft Graph access token in WorkIQ settings first.");
      return;
    }

    const query = window.prompt("Search Microsoft WorkIQ/M365 data");

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
        body: JSON.stringify(buildWorkIqSearchRequest(query, this.workIqSettings))
      });

      const searchResponse = response.json as WorkIqSearchResponse;
      const markdown = formatWorkIqHits(query, flattenWorkIqHits(searchResponse));
      editor.replaceSelection(`${markdown}\n`);
      new Notice("Inserted WorkIQ search results.");
    } catch (error) {
      const message = error instanceof Error ? error.message : "Unknown error";
      new Notice(`WorkIQ search failed: ${message}`);
    }
  }
}

class WorkIqSettingTab extends PluginSettingTab {
  constructor(app: App, private readonly plugin: WorkIqPlugin) {
    super(app, plugin);
  }

  display(): void {
    const { containerEl } = this;
    containerEl.empty();

    containerEl.createEl("h2", { text: "WorkIQ settings" });

    new Setting(containerEl)
      .setName("Microsoft Graph access token")
      .setDesc("Paste a delegated Microsoft Graph token that can call Microsoft Search.")
      .addText((text) => {
        text
          .setPlaceholder("eyJ...")
          .setValue(this.plugin.workIqSettings.accessToken)
          .onChange(async (value) => {
            this.plugin.workIqSettings.accessToken = value.trim();
            await this.plugin.saveSettings();
          });
        text.inputEl.type = "password";
      });

    new Setting(containerEl)
      .setName("Microsoft Graph search endpoint")
      .setDesc("Use the default Microsoft Search endpoint unless your environment requires another Graph cloud.")
      .addText((text) =>
        text
          .setValue(this.plugin.workIqSettings.graphSearchEndpoint)
          .onChange(async (value) => {
            this.plugin.workIqSettings.graphSearchEndpoint = value.trim() || DEFAULT_SETTINGS.graphSearchEndpoint;
            await this.plugin.saveSettings();
          })
      );

    new Setting(containerEl)
      .setName("Entity types")
      .setDesc("Comma-separated Microsoft Search entity types, such as driveItem, message, event.")
      .addText((text) =>
        text
          .setValue(this.plugin.workIqSettings.entityTypes.join(", "))
          .onChange(async (value) => {
            this.plugin.workIqSettings.entityTypes = value
              .split(",")
              .map((entityType) => entityType.trim())
              .filter(Boolean);
            await this.plugin.saveSettings();
          })
      );

    new Setting(containerEl)
      .setName("Maximum results")
      .setDesc("The number of WorkIQ search results to insert into the active note.")
      .addSlider((slider) =>
        slider
          .setLimits(1, 25, 1)
          .setValue(this.plugin.workIqSettings.maxResults)
          .setDynamicTooltip()
          .onChange(async (value) => {
            this.plugin.workIqSettings.maxResults = value;
            await this.plugin.saveSettings();
          })
      );
  }
}
