# BCP Module

BCP is a browser-local module for Buku Catatan Pemasukan dan Pengeluaran. It is isolated under `frontend/src/modules/bcp` and does not use the legacy EXIM backend database.

## Current implementation

### M1 — Project Foundation
- Project creation, opening, recent projects, and project switching.
- `.bcp` ZIP package with JSON registries and original source files.
- Dedicated IndexedDB database: `reza-bcp-local`.
- Auto-Recovery and Save Reminder.
- SHA-256 checksums for package payload files.
- ZIP preflight checks for entry count, uncompressed size, encryption, ZIP64, and unsafe paths.

### M2 — Registry Foundation
- Common Document Registry and Item Registry models.
- Separate Raw Material and Raw Material Tariff registries.
- Original/current value separation.
- Correction history, business status, document relationships, search, and registry integrity checks.
- Document detail joins items, raw materials, and tariff children.

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

The current test suite covers:
- BC 2.7 IN single-document parsing.
- BC 2.7 OUT single-document parsing.
- Multi-document workbook splitting.
- Partial Import and original source inclusion.
- Document/item/raw-material/tariff relationships.
- Duplicate detection.
- Project archive round-trip and checksum validation.
- Unsafe ZIP path rejection.

## Remaining work

- M4: Duplicate Review Center, staging corrections, full validation review, and import decision history.
- M5: Monthly Period management and historical workbook migration.
- M6: Reports and Export Center.
- M7: Additional BC document types.
- M8: Performance, accessibility, regression, and user acceptance testing.

The BC 2.7 parser must still be validated against additional real-world variants before being treated as complete. The current implementation is a development baseline, not a production release.
