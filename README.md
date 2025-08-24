# OfficeJS Add-in
---
## How to Use
1. Download `manifest.xml`
2. In Outlook web/PWA: **Add-ins → Get Add-ins → My add-ins → Custom add-ins → Upload**  → `manifest.xml`.
35. Compose a message: **Add-ins → OfficeJS Tools** on the ribbon.

## Structure
```
officejs/
├── manifest.xml
├── src/
│   ├── outlook/
│   │   └── blank.html        ← schema stub
│   │   └── taskpane.html     ← cache buster
│   │   └── taskpane.css      ← styles
│   │   └── features.js       ← all features
│   │   └── ui.js             ← UI
│   └── shared/
│       └── utils.js
├── assets/
│   └── outlook/
│       ├── icon-16.png
│       ├── icon-32.png
│       └── icon-80.png
└── README.md

```
---
## Notes
- Manifest uses absolute HTTPS URLs pointing at `https://<username>.github.io/officejs/...`.
- Add more buttons by creating new functions in taskpane.js.
