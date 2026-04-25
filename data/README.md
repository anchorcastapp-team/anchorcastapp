# Bible Data Files

## Required: kjv.json

Place the complete KJV Bible JSON file here as `kjv.json`.

### Format expected:
```json
[
  {"b": 1, "c": 1, "v": 1, "t": "In the beginning God created the heaven and the earth."},
  {"b": 1, "c": 1, "v": 2, "t": "And the earth was without form..."},
  ...
]
```

Where: b = book number (1-66), c = chapter, v = verse, t = text

### Download sources (free, public domain KJV):
1. https://github.com/scrollmapper/bible_databases  → json/t_kjv.json
2. https://github.com/aruljohn/Bible-kjv           → JSON per book
3. https://bolls.life/api/                          → REST API

### After downloading:
The app auto-detects kjv.json on startup and loads all 31,102 verses.
No app restart required after first load (cached in memory).

## Optional: nkjv.json, niv.json, esv.json, nlt.json, nasb.json
Same format. Place in this folder. App detects them automatically.
