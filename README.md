---
title: Work IQ for Obsidian
description: Ask Microsoft 365 Copilot questions grounded in your work data from an Obsidian note
---

Ask Microsoft 365 Copilot questions from Obsidian and insert answers grounded in
your Microsoft 365 work data into the active note.

## What the plugin does

The plugin opens a prompt inside Obsidian, sends the question to Microsoft's
official Work IQ CLI, and inserts the response at the cursor in the active note.
Work IQ can ground answers in Microsoft 365 content that the signed-in user can
already access, including email, meetings, Teams conversations, and documents.

The plugin:

* Adds an `Ask Microsoft Work IQ` command to the command palette
* Adds a Work IQ action to the left ribbon
* Uses the account and managed authentication held by the official Work IQ CLI
* Inserts the question as a Markdown heading followed by the Work IQ response
* Keeps access tokens out of the Obsidian vault and plugin settings
* Shows a persistent progress notification while a request is running
* Displays CLI or authentication failures in an Obsidian modal
* Records request stages in a local diagnostic log

This is a desktop-only plugin because it starts the locally installed Work IQ
executable. Obsidian Mobile is not supported.

## Prerequisites

You need:

* Obsidian Desktop 1.5.0 or later
* Node.js and npm
* Microsoft Work IQ enabled in your Microsoft 365 tenant
* An account assigned to the tenant's Work IQ usage-based billing plan
* Administrative consent for Microsoft's Work IQ application

You do not need to create an Entra app registration for this plugin. Your tenant
administrator must configure Work IQ once for the organization. See the
[Work IQ CLI documentation](https://learn.microsoft.com/en-us/microsoft-365/copilot/extensibility/work-iq/cli)
for current tenant, billing, consent, and platform requirements.

## Install and initialize Work IQ

Install the official CLI globally:

```powershell
npm install --global @microsoft/workiq
workiq accept-eula
workiq ask -q "Reply with exactly: Work IQ is ready."
```

The final command might open a Microsoft sign-in prompt the first time. Continue
only after it prints `Work IQ is ready.` in the terminal. This proves the CLI,
account, tenant consent, and Work IQ access work before Obsidian is involved.

You can confirm the installed CLI version with:

```powershell
workiq version
```

## Install the plugin

Build the plugin from this repository:

```powershell
npm install
npm test
npm run build
```

Create a `workiq` directory under the vault's community plugin directory, then
copy these files into it:

```text
<vault>/.obsidian/plugins/workiq/
|-- main.js
|-- manifest.json
`-- versions.json
```

In Obsidian:

1. Open **Settings > Community plugins**.
2. Enable **WorkIQ**.
3. Open **Settings > WorkIQ**.
4. Configure the Work IQ executable as described below.
5. Restart Obsidian after replacing plugin files with a newer build.

## Configure the executable

The only plugin setting is **Work IQ executable**.

Leave the setting blank when Work IQ was installed globally with npm in the
standard location. On Windows, the plugin derives the path from `APPDATA` and
the processor architecture:

```text
%APPDATA%\npm\node_modules\@microsoft\workiq\bin\win-x64\workiq.exe
%APPDATA%\npm\node_modules\@microsoft\workiq\bin\win-arm64\workiq.exe
```

For a standard x64 installation, the expanded path resembles:

```text
C:\Users\<username>\AppData\Roaming\npm\node_modules\@microsoft\workiq\bin\win-x64\workiq.exe
```

Enter the full path to the native `workiq.exe` when npm uses a different global
directory or automatic discovery fails. Do not select the `workiq.ps1` wrapper.

Find and verify the native x64 executable with PowerShell:

```powershell
$workIq = Join-Path (npm root --global) '@microsoft\workiq\bin\win-x64\workiq.exe'
$workIq
Test-Path $workIq
& $workIq version
```

`Test-Path` must return `True`. Use `win-arm64` instead of `win-x64` on Windows
ARM64.

On macOS and Linux, a blank setting resolves `workiq` from the environment used
to launch Obsidian. Set the absolute executable path if Obsidian cannot resolve
the command from `PATH`.

## Use the plugin

1. Open the note that should receive the answer.
2. Place the cursor where the answer should be inserted, or select text to
   replace it.
3. Open the command palette with `Ctrl+P` on Windows and Linux or `Cmd+P` on
   macOS.
4. Run **WorkIQ: Ask Microsoft Work IQ**. You can also select the Work IQ icon
   in the left ribbon.
5. Enter a question and press **Enter** or select **Ask**.
6. Keep the note open while Work IQ processes the request.

Requests commonly take several seconds. The progress notification remains open
until Work IQ returns or the request fails. A successful result is inserted in
this form:

```markdown
## Work IQ: Summarize my meetings from today and list the action items

Work IQ response...
```

Example prompts:

```text
Summarize my meetings from today and list the action items.
```

```text
Which unread emails need a response from me?
```

```text
What decisions were made in the latest meeting about Project Alpha?
```

```text
Find recent documents about the customer architecture and summarize them.
```

## How it works

For each question, the plugin starts the official CLI with arguments equivalent
to:

```powershell
workiq ask --json --question "<question>"
```

The process is started without a command shell, and the question is passed as a
separate argument. The plugin reads the structured JSON response, terminates its
CLI child after receiving the complete response, and inserts the answer through
the Obsidian editor. If the editor mutation does not persist, it falls back to
the Obsidian vault API.

Authentication, token caching, token renewal, tenant policies, and Microsoft
365 permissions remain the responsibility of Microsoft's Work IQ client. The
plugin neither receives nor stores a refresh token.

Each prompt is a separate one-shot Work IQ request. The plugin does not preserve
multi-turn conversation state between prompts.

## Troubleshooting

### Nothing is inserted

1. Wait for the progress notification to close. Work IQ queries can take tens
   of seconds.
2. Confirm that the target note remains open and editable.
3. Run the same question directly with `workiq ask -q "Your question"`.
4. Confirm the configured executable exists with `Test-Path`.
5. Restart Obsidian after updating the plugin bundle.
6. Inspect the diagnostic log described below.

### Executable not found

Set **Settings > WorkIQ > Work IQ executable** to the full native executable
path. On Windows, use `workiq.exe`, not `workiq.ps1`.

### Sign-in or permission failure

Run this command in a terminal:

```powershell
workiq ask -q "Reply with exactly: Work IQ is ready."
```

Resolve the sign-in, tenant consent, billing-plan assignment, or Conditional
Access error in the official client first. The Obsidian plugin cannot bypass
Microsoft 365 permissions or tenant policies.

### Diagnostic log

The plugin writes ISO 8601 timestamped stages to:

```text
<vault>/.obsidian/plugins/workiq/workiq.log
```

A successful request records:

```text
request started
CLI returned <number> characters
answer saved to note
```

Failures record `request failed` followed by the CLI or plugin error. The log
does not contain access tokens, but it can contain error text returned by the
CLI. Review it before sharing.

The plugin limits a request to five minutes and 10 MB of CLI output. Exceeding
either limit opens an error modal and records the failure in the diagnostic log.

## Security

* The plugin invokes Microsoft's official Work IQ executable directly
* It does not use a command shell to interpolate questions
* It does not store Microsoft access or refresh tokens in the vault
* Work IQ can only access Microsoft 365 content available to the signed-in user
* Answers are written into the active vault and inherit that vault's storage,
  synchronization, and sharing controls

Review generated answers before sharing a note outside your organization.

## Development

Run the test suite and production build with:

```powershell
npm install
npm test
npm run build
```

The focused tests cover CLI JSON parsing, Markdown formatting, executable path
resolution, error handling, and the CLI process lifecycle.
