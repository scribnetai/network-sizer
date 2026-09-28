# Changelog

## 2026-09-28
- Added a floating Feedback button (bottom-right) that opens a dialog to send feedback via email — topic chips, optional name, and message, addressed to the site owner with the app name in the subject.

## 2026-09-28
- TLS certificate provisioned for the `network-sizer.scribnet.io` custom domain (GitHub's stuck DNS check was reset 2026-09-28); HTTPS is now enforced on the site. App-switcher menu links switched from legacy `scribnetai.github.io` URLs to direct `https://<app>.scribnet.io` URLs for all 10 apps (footer/launcher links updated likewise). This entry also covers the net-zero CNAME delete/re-add commits from the DNS-check reset, which carried no changelog entries. Touched: index.html, js/app-switcher.js.


## 2026-09-28
- Migrated legacy `scribnetai.github.io` links to `https://<app>.scribnet.io` for the HTTPS-enforced apps (se-command-center, server-sizer, network-sizer); links to the remaining apps left on the legacy URLs until their TLS certs are issued. Touched: index.html, js/app-switcher.js.

## 2026-09-26
- Restyled with a Physgun-inspired vibe: Outfit display font, glowing blue sliders with live value badges and min/mid/max scales, blue-glyph section tiles, gradient key numbers, animated bandwidth bars plus a stacked oversubscription meter on the Plan tab, and scroll-reveal on the landing. Segmented pill buttons now drive the dual-homing, N+1 spare, breakout, and FC-fabric toggles; Target hosts is now a slider (max auto-scales with detected hosts). All sizing math, projects, and saved data untouched.

## 2026-09-26
- Added: this changelog section, rendered from CHANGELOG.md.
- Added 💾 Projects: named saves in this browser, portable JSON export/import, and automatic session restore (your last session reloads on revisit). Saves capture host groups, existing inventory, port profile, and TOR + FC fabric configs — nothing uploaded.
- Fixed: Target hosts could show 0 while demand showed 21 hosts — an emptied input no longer sizes for 0 hosts; it falls back to the detected host count. “↺ Start over” now also clears the autosaved session so a reload can’t resurrect it.

## 2026-09-27
- Added top-left app-switcher dropdown on the brand mark: one-click jumps to every app in the suite (full index, this page marked).
- Fixed: shared-100G breakout switch sizing — sub-ports now share physical ports 4:1 in the demand math (previously undersized mixed-speed profiles). Corrected privacy copy: Projects/autosave use browser localStorage, machine-only, never uploaded.
