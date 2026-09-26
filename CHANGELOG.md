# Changelog

## 2026-09-26
- Added: this changelog section, rendered from CHANGELOG.md.
- Added 💾 Projects: named saves in this browser, portable JSON export/import, and automatic session restore (your last session reloads on revisit). Saves capture host groups, existing inventory, port profile, and TOR + FC fabric configs — nothing uploaded.
- Fixed: Target hosts could show 0 while demand showed 21 hosts — an emptied input no longer sizes for 0 hosts; it falls back to the detected host count. “↺ Start over” now also clears the autosaved session so a reload can’t resurrect it.
