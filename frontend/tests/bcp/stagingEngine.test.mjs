import test from 'node:test';
import assert from 'node:assert/strict';
import { applyStagingCorrection, getStagingValidationSummary } from '../../src/modules/bcp/stagingEngine.js';

function fixture() {
  return [{
    source: { source_id: 'SRC-1', original_file_name: 'ceisa.xlsx' },
    issues: [],
    documents: [{
      document: { document_id: 'DOC-1', aju_number: 'AJU-1', registration_number: '001', dates: { registration: '2026-11-01' }, current_data: { nomor_daftar: '001', tanggal_daftar: '2026-11-01' } },
      items: [{ item_id: 'ITEM-1', item_code: 'OLD', description: 'Old description', quantity: 5, current_data: { item_code: 'OLD', description: 'Old description', quantity: 5 } }],
      validation: { status: 'READY' }
    }]
  }];
}

test('staging correction changes normalized value but preserves original raw source data', () => {
  const input = fixture();
  const result = applyStagingCorrection(input, {
    document_id: 'DOC-1', entity_type: 'ITEM', entity_id: 'ITEM-1',
    field_path: 'current_data.description', new_value: 'Corrected description',
    reason: 'Koreksi uraian sesuai dokumen pendukung.'
  });
  assert.equal(result[0].documents[0].items[0].description, 'Corrected description');
  assert.equal(result[0].documents[0].items[0].current_data.description, 'Corrected description');
  assert.equal(result[0].documents[0].items[0].staging_original_data.description, 'Old description');
  assert.equal(result[0].documents[0].staging_corrections[0].old_value, 'Old description');
});

test('staging correction requires reason and rejects unsafe or unknown fields', () => {
  const input = fixture();
  assert.throws(() => applyStagingCorrection(input, {
    document_id: 'DOC-1', entity_type: 'DOCUMENT', field_path: 'current_data.nomor_daftar', new_value: '002', reason: ''
  }), /Alasan koreksi/);
  assert.throws(() => applyStagingCorrection(input, {
    document_id: 'DOC-1', entity_type: 'DOCUMENT', field_path: 'current_data.aju_number', new_value: 'AJU-2', reason: 'test'
  }), /tidak dapat dikoreksi/);
});

test('staging validation summary exposes issues and correction history', () => {
  const input = fixture();
  const changed = applyStagingCorrection(input, {
    document_id: 'DOC-1', entity_type: 'DOCUMENT', field_path: 'current_data.nomor_daftar',
    new_value: '002', reason: 'Nomor daftar diperbaiki.'
  });
  const summary = getStagingValidationSummary(changed);
  assert.equal(summary[0].documents[0].corrections.length, 1);
  assert.equal(summary[0].documents[0].status, 'READY');
});
