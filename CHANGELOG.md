# Changelog

All notable changes to this project will be documented in this file.

## [1.0.1] - 2026-03-19

### Fixed
- Fixed critical JavaScript crash caused by missing HTML elements (`ans-toggle`, `ans-list`, `ans-arrow`) in `index.html`.
- Corrected script path in `index.html` from absolute (`/main.js`) to relative (`./main.js`) to ensure compatibility with GitHub Pages sub-path deployments.
- Added defensive checks in `main.js` for DOM element access.
- Restored missing answer summary functionality in the result view.

### Added
- Created `CHANGELOG.md` to track project updates.
