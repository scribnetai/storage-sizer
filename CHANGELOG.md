# Changelog

## 2026-09-28
- TLS certificate provisioned for the `storage-sizer.scribnet.io` custom domain (GitHub's stuck DNS check was reset 2026-09-28); HTTPS is now enforced on the site. App-switcher menu links switched from legacy `scribnetai.github.io` URLs to direct `https://<app>.scribnet.io` URLs for all 10 apps (footer/launcher links updated likewise). This entry also covers the net-zero CNAME delete/re-add commits from the DNS-check reset, which carried no changelog entries. Touched: index.html, js/app-switcher.js.


## 2026-09-26
- Added 💾 Projects: named saves in this browser, portable JSON export/import, and automatic session restore (your last session reloads on revisit). Saves capture pools, per-pool configs, and global growth settings — nothing uploaded.
- Added: this changelog section, rendered from CHANGELOG.md.
- 🎨 Physgun-style reskin: Outfit typeface, deeper navy palette, blue→cyan gradient accents, glyph tiles on section headings, scroll-reveal on the landing.
- ✨ All key tunables are now glowing sliders with live pill-badge values and min/mid/max scales (step 1 on integer sliders; reduction keeps its 0.1 step). Every output updates in real time.
- 📊 Growth plan tab: animated capacity-curve bars per year plus per-pool stacked meters at the horizon year (blue = covered by owned usable, red = shortfall); headline totals in big gradient numbers.
- ❓ FAQ items now carry emoji prefixes.

## 2026-09-27
- Added top-left app-switcher dropdown on the brand mark: one-click jumps to every app in the suite (full index, this page marked).
- Fixed: RAID overhead factors — RAID-5 4+1 now 1.25x (was 1.33), RAID-6 6+2 now 1.333x (was 1.5).
