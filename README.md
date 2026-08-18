# WorkIQ for Obsidian

Search Microsoft WorkIQ/Microsoft 365 data from Obsidian and insert the most relevant context into the active note.

## Features

- Adds a `Search Microsoft WorkIQ` command.
- Calls the Microsoft Graph Search API with a delegated access token.
- Inserts SharePoint/OneDrive, Outlook mail, and calendar results as Markdown context.
- Lets you configure the Graph endpoint, entity types, and maximum result count.

## Development

```bash
npm install
npm test
npm run build
```

## Setup

1. Build the plugin.
2. Copy `main.js`, `manifest.json`, and `styles.css` if present into an Obsidian vault plugin folder such as `.obsidian/plugins/workiq/`.
3. Enable the plugin in Obsidian.
4. Add a Microsoft Graph delegated access token with Microsoft Search permissions in the plugin settings.
