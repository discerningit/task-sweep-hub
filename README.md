# TaskSweep Hub

A minimal, solo-developer-friendly PWA that sweeps tasks from many sources into one central list — with deduplication, beacon-based discovery, and optional Microsoft 365 connectivity.

**Built for:** IT consultants juggling irregular hours, a home build ("Cedar Ridge"), nonprofit leadership, and family tasks.

---

## Quick start (5 minutes)

### Prerequisites

- [Node.js](https://nodejs.org/) 20 or newer
- A terminal (Command Prompt, Terminal, or VS Code terminal)

### Run locally

```bash
cd task-sweep-hub
npm install
npm run dev
```

Open the URL shown (usually `http://localhost:5173`) in your browser.

### Install as PWA

In Chrome or Edge: click the install icon in the address bar, or use **Menu → Install TaskSweep Hub**. The app works offline for pasted tasks after the first load.

---

## How to use (no coding required)

### 1. Paste tasks (works immediately)

1. Copy text from email, Teams, Reminders, or any app.
2. Paste into the **Add tasks** box.
3. Click **Sweep pasted text**.
4. Tasks appear in **Your tasks**.

Example paste:

```
- Call Cedar Ridge contractor re: permit
- Board meeting prep (nonprofit) — due 3/15
- Review client ticket #4521 — urgent
```

### 2. Plant a beacon (discover connectivity)

1. Go to the **Beacon** tab.
2. Click **Copy beacon text**.
3. Create a task in Outlook, MS To Do, or any list using that title.
4. Later, copy/paste from that app into TaskSweep and click **Scan for beacon** — the app suggests how to connect.

Default marker: `[TaskSweep-Beacon]`

### 3. Connect Microsoft 365 (optional)

1. Go to **Settings**.
2. Register an app in [Azure Portal](https://portal.azure.com) → **App registrations** → **New registration**.
3. Add redirect URI: `http://localhost:5173` (and your production URL later).
4. Under **API permissions**, add delegated: `Tasks.Read`, `Mail.Read`, `Notes.Read.All`, `User.Read`.
5. Copy the **Application (client) ID** into Settings.
6. Click **Sign in to Microsoft 365**.
7. On the **Tasks** tab, click **Sweep all sources**.

**VDI / locked-down work PC?** Use **Paste** or **File upload** instead — no admin rights needed.

### 4. Export tasks

**Settings → Export tasks CSV** — import into Todoist, Apple Reminders, or Excel.

---

## Project structure

```
task-sweep-hub/
├── src/
│   ├── components/       # UI (task list, input, beacon, settings)
│   ├── connectors/       # Paste, file upload, M365 (add more here)
│   ├── db/               # IndexedDB — all data stays local
│   ├── hooks/            # React data hooks
│   ├── services/         # Extraction, dedup, AI, sync-back
│   └── types/            # Shared TypeScript types
├── public/               # Icons, PWA assets
├── vite.config.ts        # Build + PWA config
└── README.md             # You are here
```

### Adding a new connector

1. Create `src/services/connectors/yourSource.ts`
2. Implement the `Connector` interface (see `src/types/task.ts`)
3. Register it in `src/services/connectors/index.ts`

---

## What works in this MVP

| Feature | Status |
|---------|--------|
| Paste text sweep | ✅ Works offline |
| File upload (.txt, .csv, .eml) | ✅ |
| Local task extraction (no API) | ✅ |
| Deduplication (hash + similarity) | ✅ |
| Beacon copy / scan | ✅ |
| IndexedDB local storage | ✅ |
| M365 To Do + flagged Outlook | ✅ With Azure app + sign-in |
| AI providers (Copilot, Claude, Grok…) | 🔧 Stubbed — uses local extraction |
| Sync-back to source / MS To Do | 🔧 Stubbed — marks complete locally |
| Proton Mail, Jira, Siri | 📋 Planned — use paste/file for now |

---

## Recommended next prompts (paste into your AI assistant)

Copy one of these when you're ready to iterate:

**M365 sync-back**
> Wire TaskSweep Hub M365 sync-back: when I mark a task complete in the hub, PATCH the Graph API todoTask or clear the Outlook flag. Use `src/services/syncBack.ts` and `src/services/connectors/m365.ts`.

**Real AI extraction**
> Add Grok API extraction to TaskSweep Hub: in `src/services/aiOrchestrator.ts`, call the xAI API with `buildContextPrompt()` and parse JSON tasks. Use `VITE_GROK_API_KEY` from `.env`.

**Proton Mail connector**
> Add a Proton Mail connector to TaskSweep Hub using Proton's export format or Bridge IMAP — implement the Connector interface and register in `connectors/index.ts`.

**iOS Shortcuts**
> Document an iOS Shortcut that exports Apple Reminders to a text file and uploads it to TaskSweep Hub's file connector. Include the Shortcut steps for a non-coder.

**Todoist push**
> When primary task tool is Todoist, push new tasks from TaskSweep Hub via Todoist REST API. Add API token field in Settings.

**Improve dedup**
> Improve TaskSweep deduplication to catch tasks like "Call contractor" and "Call Cedar Ridge contractor" as duplicates using fuzzy matching.

**Deploy PWA**
> Help me deploy TaskSweep Hub to Cloudflare Pages or GitHub Pages with HTTPS so I can install the PWA on my phone.

---

## Privacy

- All tasks and settings are stored in **IndexedDB on your device**.
- No TaskSweep server exists — data leaves your browser only when **you** connect M365 or a future AI API.
- M365 tokens are stored in browser localStorage via MSAL.

---

## Scripts

| Command | Purpose |
|---------|---------|
| `npm run dev` | Start dev server |
| `npm run build` | Production build |
| `npm run preview` | Preview production build |

---

## License

MIT — use freely for personal and business task management.