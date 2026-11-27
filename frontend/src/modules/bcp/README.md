# BCP Module

BCP is a browser-local module for Buku Catatan Pemasukan dan Pengeluaran. It is isolated under `frontend/src/modules/bcp` and does not use the legacy EXIM backend database.

## Current implementation

### M1 — Project Foundation
- Project creation, opening, recent projects, and project switching.
- `.bcp` ZIP package with JSON registries and original source files.
- Dedicated IndexedDB database: `reza-bcp-local`.
- Auto-Recovery and Save Reminder.
- SHA-256 checksums for package payload files.
- ZIP preflight checks for entry count, uncompressed size, encryption, ZIP64, duplicate internal names, local headers, central directory consistency, and unsafe paths.
- Archive loading requires every core package entry and a valid SHA-256 checksum for every non-manifest entry.

### M2 — Registry Foundation
- Common Document Registry and Item Registry models.
- Separate Raw Material and Raw Material Tariff registries.
- Original/current value separation.
- Correction history, business status, document relationships, search, and registry integrity checks.
- Document detail joins items, raw materials, and tariff children.
- Registry integrity checks enforce unique IDs and valid document → item → raw material → tariff parent-child references, period assignments, and source references.

### M3 — BC 2.7 Extraction
- Unified Upload for mixed single-document and multi-document Excel workbooks.
- BC 2.7 IN/OUT classification using configured company identity plus entity role codes 3 and 8.
- Header/row extraction based on actual worksheet cell addresses, not only worksheet dimension metadata.
- Duplicate column names are preserved as positional raw fields.
- Document rows group by Nomor Aju; item children link by Nomor Aju + Seri Barang; raw material tariffs link by Nomor Aju + Seri Barang + Seri Bahan Baku.
- Original Excel bytes are included in the project package after at least one document from that file is committed.
- Partial Import imports eligible unique documents while holding errors and duplicates.

## Current CEISA field mapping

| CEISA worksheet | Current normalized content | Preservation |
|---|---|---|
| HEADER | Nomor Aju, kode dokumen/kantor, Nomor Daftar, tanggal, valuta, NDPBM, CIF, harga penyerahan, bruto/netto/volume | Full positional raw fields |
| ENTITAS | Entity code, identity, name, address, NIB | Full raw fields and source row |
| BARANG | Seri Barang, HS, item code, description, specification, unit, quantity, package, weight, CIF/value | Full raw fields and per-field source locations |
| BAHANBAKU | Seri Bahan Baku, origin code, item code, description, quantity, unit, CIF/value | Separate Raw Material Registry |
| BAHANBAKUTARIF | Kode Pungutan, Kode Tarif, tariff/facility and payment values | Separate Raw Material Tariff Registry |
| BAHANBAKUDOKUMEN | Document references linked to item and raw material series | Raw material extension |
| DOKUMEN, PENGANGKUT, KEMASAN, PUNGUTAN, and other sheets | Document-level supporting details | Raw fields retained in document extension |
| BARANGTARIF, BARANGDOKUMEN, BARANGENTITAS, BARANGSPEKKHUSUS, BARANGVD | Item-level supporting details | Raw fields retained in item extension |

The current parser normalizes common fields and preserves all remaining worksheet fields in raw positional form. It does not yet claim that every CEISA field has a dedicated normalized column.

## Import behavior

1. Select multiple Excel files or drop them onto Unified Upload.
2. Identify each workbook and classify its document direction.
3. Split workbooks containing multiple Nomor Aju into independent staged documents.
4. Review errors, direction classification, and duplicate candidates.
5. Confirm Partial Import.
6. Save the project as a `.bcp` file.

Documents with an unresolved direction require a manual IN/OUT selection in Import Preview. Validation errors and duplicate candidates remain held; they are not silently discarded.

## Project package

The package contains `manifest.json`, document/item/raw-material registries, periods, corrections, validation, relationships, audit history, project settings, source registry, and original source files under `sources/ceisa`, `sources/historical`, or `sources/other`.

Current package limits:
- Maximum compressed project file: 250 MB.
- Maximum total uncompressed archive payload: 750 MB.
- Maximum archive entries: 10,000.
- ZIP64 and encrypted ZIP archives are not supported.
- Web Crypto is required for SHA-256 integrity checks.

Auto-Recovery is browser-local and is not a backup. The `.bcp` file remains the user's permanent project file. BCP does not upload project files to a server, encrypt projects, or create backups.

## Tests

Run from `frontend`:

```bash
npm run build
npm run test:bcp
```

The current test suite contains 36 passing tests in CI and covers:
- BC 2.7 IN single-document parsing.
- BC 2.7 OUT single-document parsing.
- Multi-document workbook splitting.
- Partial Import and original source inclusion.
- Document/item/raw-material/tariff relationships.
- Duplicate detection.
- Project archive round-trip and checksum validation.
- Unsafe ZIP path rejection.
- Missing required archive entry and registry orphan/reference rejection.

### M4 — Duplicate Review & Validation
- Duplicate Review Center with separate file-hash and document-identity candidates.
- Decisions: Use Existing, Use New, Keep Both, and Defer Review; each requires a reason and creates an audit event.
- Use Existing links a new source file to the existing document without creating a new transaction.
- Use New and Keep Both preserve both document records and create an explicit relationship between them.
- Validation Center supports controlled staging corrections for selected document and item fields.
- Original normalized values remain preserved; corrected values become Current values after import.
- Staging corrections are committed to correction history and Audit Trail only when the document is imported.
- Import decisions are stored in `data/import_decisions.json` inside the `.bcp` project.

### M5 — Monthly Period Foundation
- Create monthly periods with unique YYYY-MM keys.
- Lifecycle OPEN → REVIEW → CLOSED, with reopening from CLOSED to OPEN.
- Closing and reopening require a reason and write Audit Trail events.
- Documents can be assigned or moved between periods; moving requires a reason.
- CLOSED periods cannot receive new assignments or have documents moved out until reopened.
- Period summary reports document count, item count, business statuses, amendments, and document types.
- Reporting Date is not inferred automatically; per-document-type date mapping remains pending explicit CEISA mapping approval.

### M6 — Reports & Export Center Foundation
- Standard Excel presets: DOC_REGISTRY, ITEM_REGISTRY, MONTHLY_SUMMARY, BC_TYPE_REPORT, TRANSACTION_REPORT, SOURCE_REPORT, CORRECTION_REPORT, AUDIT_REPORT.
- Custom Export supports Document or Item Registry, user-selected columns, text filters, grouping, sorting, and column-label overrides.
- Custom column profiles (including filter/group/sort/label settings) are stored in project settings.
- Export Center shows a first-10-row content preview for registry exports before download; summary presets preview their first data rows.
- Data Version: ORIGINAL, CURRENT, and COMPARISON.
- Workbook includes README, VALIDATION_SUMMARY, and EXPORT_METADATA sheets.
- Export history and Audit Trail events are stored in the project; exported workbooks are not retained by BCP.
- Identifier values remain text; numeric fields remain numeric; formula-leading strings are escaped.
- Large registry sheets are split before Excel's worksheet row limit.

## Remaining work

- M6 completion: report reconciliation tests and validation of exports against larger real-world datasets.
- M4/M5 completion: real-file browser acceptance tests, broader validation rules, period assignment review, and duplicate edge-case coverage.
- M5: Monthly Period management and historical workbook migration.
- M6: Reports and Export Center.
- M7: Additional BC document types.
- M8: Performance, accessibility, regression, and user acceptance testing.

The BC 2.7 parser must still be validated against additional real-world variants before being treated as complete. The current implementation is a development baseline, not a production release.
