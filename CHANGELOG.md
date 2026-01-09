# Changelog

All notable changes to this project will be documented in this file.

## [1.2.0] - 2026-01-09
### Added
- Windows-friendly distribution with **scheduler_ui.exe** (GUI) and **scheduler.exe** (CLI).
- Sequential “screen section” handling based on **all unit columns after Comments** (standard first, then premium sections appended).

### Changed
- Sheet-name normalization to prevent duplicate tabs caused by truncation/trailing punctuation.
- Formatting polish: centered alignment, bold NEW titles, bold red FINAL box-score lines, and `m/d/yy` date display.

## [1.1.0] - 2026-01-02
### Added
- Initial release: reads a bookings export and updates a single master schedule workbook.
- Week banding, header styling, borders, wrapped text layout.
