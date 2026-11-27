import test from 'node:test';
import assert from 'node:assert/strict';
import { createEmptyProject } from '../../src/modules/bcp/projectSchema.js';
import { buildExportWorkbook, recordExportHistory, EXPORT_PRESETS } from '../../src/modules/bcp/exportEngine.js';
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
    current_data: { item_code: '00123', description: 'Current' }
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
