import test from 'node:test';
import assert from 'node:assert/strict';
import { createEmptyProject } from '../../src/modules/bcp/projectSchema.js';
import { buildExportWorkbook, buildExportPreview, recordExportHistory, EXPORT_PRESETS, saveExportProfile } from '../../src/modules/bcp/exportEngine.js';
import * as XLSX from 'xlsx';

function fixture() {
  const project = createEmptyProject({ projectName: 'EXPORT TEST', companyName: 'PT Chinli' });
  project.data.documents.push({
    document_id: 'DOC-1', document_type: 'BC27_IN', aju_number: '00002700000000000001',
    registration_number: '000123', business_status: 'CONFIRMED', amendment_status: 'ORIGINAL',
    period_id: null, original_data: { nomor_aju: '00002700000000000001', nomor_daftar: '000123' },
    current_data: { nomor_aju: '00002700000000000001', nomor_daftar: '000123' }, dates: {}
  });
  project.data.items.push({
    item_id: 'ITEM-1', document_id: 'DOC-1', item_sequence: '1', item_code: '00123',
    description: '=HYPERLINK("x")', quantity: 10, unit: 'PRS',
    original_data: { item_code: '00123', description: 'Original' },
    current_data: { item_code: '00123', description: '=HYPERLINK("x")' }
  });
  return project;
}

test('standard export presets are available and create workbook sheets', () => {
  const project = fixture();
  for (const preset of EXPORT_PRESETS) {
    const workbook = buildExportWorkbook(project, { preset });
    assert.ok(workbook.SheetNames.length >= 3);
  }
});

test('export preserves identifier text and protects formula-like strings', () => {
  const workbook = buildExportWorkbook(fixture(), { preset: 'ITEM_REGISTRY' });
  const sheet = workbook.Sheets.ITEM_REGISTRY;
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true });
  assert.equal(rows[1][2], '00123');
  assert.ok(String(rows[1][3]).startsWith("'="));
});

test('comparison export includes original and current values', () => {
  const workbook = buildExportWorkbook(fixture(), { preset: 'ITEM_REGISTRY', dataVersion: 'COMPARISON' });
  const rows = XLSX.utils.sheet_to_json(workbook.Sheets.ITEM_REGISTRY, { header: 1 });
  assert.ok(rows[0].includes('Uraian Barang - Original'));
  assert.ok(rows[0].includes('Uraian Barang - Current'));
});

test('custom export requires selected columns and exports only selected fields', () => {
  const project = fixture();
  assert.throws(() => buildExportWorkbook(project, { preset: 'CUSTOM', custom: { entity_type: 'ITEM', columns: [] } }), /minimal satu kolom/);
  const workbook = buildExportWorkbook(project, { preset: 'CUSTOM', custom: { entity_type: 'ITEM', columns: ['item_code'] } });
  const rows = XLSX.utils.sheet_to_json(workbook.Sheets.CUSTOM_ITEMS, { header: 1 });
  assert.deepEqual(rows[0], ['Kode Barang']);
  assert.equal(rows[1][0], '00123');
});

test('export history is stored in project data and audit log', () => {
  const next = recordExportHistory(fixture(), { preset: 'DOC_REGISTRY', data_version: 'CURRENT', file_name: 'docs.xlsx', record_count: 1 });
  assert.equal(next.data.export_history.length, 1);
  assert.equal(next.data.audit_log.at(-1).event_type, 'EXPORT_CREATED');
});

test('custom export profiles are stored with validated field selections', () => {
  const saved = saveExportProfile(fixture(), { name: 'Item Basic', entity_type: 'ITEM', columns: ['item_code','quantity','not-a-field'] });
  assert.equal(saved.settings.export.profiles.length, 1);
  assert.deepEqual(saved.settings.export.profiles[0].columns, ['item_code','quantity']);
  const advanced = saveExportProfile(fixture(), { name: 'Advanced', entity_type: 'ITEM', columns: ['item_code','quantity'], filters: { item_code: '001' }, group_by: 'item_code', sort_by: 'quantity', sort_direction: 'DESC', label_overrides: { item_code: 'Internal Code' } });
  assert.equal(advanced.settings.export.profiles[0].filters.item_code, '001');
  assert.equal(advanced.settings.export.profiles[0].group_by, 'item_code');
  assert.throws(() => saveExportProfile(fixture(), { name: 'Empty', entity_type: 'ITEM', columns: [] }), /minimal satu kolom/);
});

test('formula guard prefixes values beginning with plus, minus, and at signs', () => {
  const project = fixture();
  project.data.items[0].description = '-CMD|x';
  project.data.items[0].current_data.description = '-CMD|x';
  const workbook = buildExportWorkbook(project, { preset: 'ITEM_REGISTRY' });
  const rows = XLSX.utils.sheet_to_json(workbook.Sheets.ITEM_REGISTRY, { header: 1 });
  assert.ok(String(rows[1][3]).startsWith("'-CMD"));
});

test('custom export applies filters, grouping, sorting, and label overrides', () => {
  const workbook = buildExportWorkbook(fixture(), {
    preset: 'CUSTOM',
    custom: {
      entity_type: 'ITEM',
      columns: ['item_code','description'],
      filters: { item_code: '001' },
      group_by: 'item_code',
      sort_by: 'item_code',
      sort_direction: 'DESC',
      label_overrides: { item_code: 'Internal Code' }
    }
  });
  const rows = XLSX.utils.sheet_to_json(workbook.Sheets.CUSTOM_ITEMS, { header: 1 });
  assert.deepEqual(rows[0], ['Group: Internal Code','Internal Code','Uraian Barang']);
  assert.equal(rows.length, 2);
  assert.equal(rows[1][1], '00123');
});

test('export preview matches the first rows of the generated workbook', () => {
  const project = fixture();
  const preview = buildExportPreview(project, { preset: 'ITEM_REGISTRY', dataVersion: 'CURRENT' }, 10);
  const workbook = buildExportWorkbook(project, { preset: 'ITEM_REGISTRY', dataVersion: 'CURRENT' });
  const rows = XLSX.utils.sheet_to_json(workbook.Sheets.ITEM_REGISTRY, { header: 1 });
  assert.deepEqual(preview.headers, rows[0]);
  assert.deepEqual(preview.rows, rows.slice(1, 11));
  assert.equal(preview.total_rows, 1);
});

test('custom export preview respects selected columns and text filter', () => {
  const preview = buildExportPreview(fixture(), {
    preset: 'CUSTOM',
    custom: { entity_type: 'ITEM', columns: ['item_code'], filters: { item_code: '001' } }
  });
  assert.deepEqual(preview.headers, ['Kode Barang']);
  assert.equal(preview.total_rows, 1);
  assert.deepEqual(preview.rows, [['00123']]);
});
