# Changelog

All notable changes to ExcelWizard are documented in this file.

## 3.1.0 - 2026-08-25

### Added

- Added `WrapText` to `ExcelSheetColumnAttribute` for model-bound grid columns.
- Added `WrapText` to `ExcelTableColumnAttribute` for model-bound table columns.
- Added `CellStyle.WrapText` for manually built cells and retained `CellStyle.Wordwrap` as a compatible alias.
- Added automatic row-height expansion for wrapped content when no explicit row height is configured.
- Added a generated-workbook integration test project alongside the unit test suite.

### Changed

- Updated ClosedXML to 0.105.1 and BlazorDownloadFile to 2.4.0.2.
- Overrode BlazorDownloadFile's legacy JSON dependency chain with System.Text.Json 10.0.11 to remove a critical transitive security advisory.
- Updated the test stack while retaining FluentAssertions 7 under its Apache 2.0 license.
- Updated the API and Blazor sample applications, unit tests, and integration tests to .NET 10.
- Preserved the published library target at .NET Standard 2.1 for consumer compatibility.
- Updated cell value handling for the current ClosedXML typed-value API.

### Fixed

- Preserved explicit row heights when wrapped text is enabled.
- Corrected cell border style mapping to ClosedXML border values.
- Made cell-level styles consistently override row-level styles.
- Validated invalid number, currency, and date values with clear exceptions.

### Testing

- Added unit coverage for wrap defaults, the legacy alias, column attributes, and builder mapping.
- Added integration coverage for grid, table, and manual wrapping; automatic and explicit row heights; multiline and Persian/English text; cell content types; metadata; multiple sheets; merged cells; RTL layout; protection; fills; and borders.
