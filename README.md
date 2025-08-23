# OfficeJS (General Office.js Add-in)

This repo hosts a general Office.js add-in. The first feature is **Outlook → Format Tables**:
- One-click: sets **1pt light-grey borders** and **0.1in padding** on all tables in the draft (Compose).

## Quick start (GitHub Pages hosting)
1. Push this folder as a GitHub repo named **officejs**.
2. Enable **Settings → Pages** (deploy from `main`, root).
3. Edit `manifest.xml` and replace **__GITHUB_USER__** with your GitHub username.
4. In Outlook web/PWA: **Get Add-ins → My add-ins → Custom add-ins → Upload my add-in** and choose `manifest.xml`.
5. Compose a message → click **Format Tables** on the ribbon.

## Structure
```
officejs/
├── manifest.xml
├── src/
│   ├── outlook/
│   │   └── formatTables.html
│   └── shared/
│       └── utils.js
├── assets/
│   └── outlook/
│       ├── icon-16.png
│       ├── icon-32.png
│       └── icon-80.png
└── README.md
```

## Notes
- Manifest uses absolute HTTPS URLs pointing at `https://<username>.github.io/officejs/...`.
- Requirement set: Mailbox 1.10, Permission: ReadWriteItem.
- Add more buttons by creating new function files and registering them in `manifest.xml`.
