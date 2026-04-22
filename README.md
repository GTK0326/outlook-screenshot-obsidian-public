# Outlook Screenshot -> Obsidian

Outlook calendar export may be blocked by organization policy. This package reads a weekly Outlook screenshot, sends it to the OpenAI Responses API, and writes weekly and daily meeting notes into an Obsidian vault.

## Features

- Reads a screenshot from the clipboard by default.
- Supports Outlook weekly grid screenshots.
- Extracts meeting date, time, title, and color category with OpenAI vision.
- Snaps times to a configurable grid. Default is `30` minutes.
- Writes a weekly meeting note.
- Updates existing daily notes.
- Updates the daily template so future daily notes auto-fill meetings from the weekly note.
- Reads meeting color -> prefix rules from an Obsidian note on every run.

## Files

- [Import-OutlookMeetingsFromScreenshot.ps1](Import-OutlookMeetingsFromScreenshot.ps1)
- [Install-OutlookScreenshot.ps1](Install-OutlookScreenshot.ps1)
- [OutlookScreenshotLauncher.cs](OutlookScreenshotLauncher.cs)
- [outlook-screenshot.sample.json](outlook-screenshot.sample.json)

## Required setup

You must provide these values outside the codebase.

- `OPENAI_API_KEY`
- Obsidian vault root

You can set the vault root in either place:

- `outlook-screenshot.json`
- `OUTLOOK_SCREENSHOT_VAULT_ROOT`

Optional environment variables:

- `OUTLOOK_SCREENSHOT_CONFIG_PATH`
- `OUTLOOK_SCREENSHOT_SCRIPT_PATH`

## Install

1. Clone or download this folder.
2. Run:

```powershell
powershell.exe -ExecutionPolicy Bypass -File .\Install-OutlookScreenshot.ps1
```

This does three things:

- copies `outlook-screenshot.sample.json` to `outlook-screenshot.json` if needed
- builds `outlook-screenshot.exe`
- adds this folder to the user `PATH`

After that:

1. Edit `outlook-screenshot.json`
2. Set `OPENAI_API_KEY`
3. Restart Explorer or sign out/in once if `PATH` changed
4. Run `outlook-screenshot`

## Config

Start from [outlook-screenshot.sample.json](outlook-screenshot.sample.json).

Important fields:

- `vaultRoot`
- `weeklyNotePattern`
- `dailyNotePattern`
- `dailyTemplatePath`
- `meetingCategoryRulesPath`
- `openAiModel`
- `openAiImageDetail`
- `openAiApiKeyEnvVar`

Example:

```json
{
  "vaultRoot": "C:\\path\\to\\Obsidian\\Vault",
  "weeklyNotePattern": "91_MeetingSchedule/{weekYear}-W{week}.md",
  "dailyNotePattern": "01_Daily/{date}.md",
  "dailyTemplatePath": "90_Templates/Daily template.md",
  "meetingCategoryRulesPath": "91_MeetingSchedule/00_MeetingCategoryRules.md"
}
```

## Category rules

The script reads `meetingCategoryRulesPath` on every run and uses it to rewrite prefixes by meeting color.

Expected table:

```md
| colorKey | prefix | project | description |
| --- | --- | --- | --- |
| red | [Internal] | Internal cost center | Red meetings |
| lightblue | [Other] | Other / offhours | Light blue meetings |
| unknown |  | Uncategorized | Unknown color |
```

## Run

Fastest flow:

1. Copy the Outlook screenshot to the clipboard.
2. Press `Win + R`.
3. Run `outlook-screenshot`

Manual run:

```powershell
powershell.exe -ExecutionPolicy Bypass -File .\Import-OutlookMeetingsFromScreenshot.ps1
```

With debug output:

```powershell
powershell.exe -ExecutionPolicy Bypass -File .\Import-OutlookMeetingsFromScreenshot.ps1 `
  -DebugOutputDir .\debug
```

## Time handling

The default time grid is `30` minutes.

- The prompt tells the model to use only valid grid boundaries.
- The script also snaps start times down and end times up to the same grid.

With the default setting, `:15` and `:45` are not kept.

## Notes

- The default model is `gpt-5.4-mini`.
- If you want a cheaper fallback, switch the config to `gpt-4o-mini`.
- Generated debug and output files are intentionally ignored by git.
