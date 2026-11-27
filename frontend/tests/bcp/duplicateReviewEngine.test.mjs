import test from 'node:test';
import assert from 'node:assert/strict';
import { createEmptyProject } from '../../src/modules/bcp/projectSchema.js';
import { buildDuplicateReview, recordDuplicateDecision, getDuplicateDecision, canProceedWithDuplicateDecision } from '../../src/modules/bcp/duplicateReviewEngine.js';

const existingDocument = {
  document_id: 'DOC-EXISTING',
  document_type: 'BC27_OUT',
  aju_number: '000027CHD85720261127001942',
  registration_number: '027413',
  current_data: {}
};

function staged({ sourceId = 'SRC-1', hash = 'hash-1', aju = existingDocument.aju_number, documentId = 'DOC-STAGED' } = {}) {
  return {
    source: { source_id: sourceId, sha256: hash, original_file_name: 'new.xlsx' },
    documents: [{
      document: {
        document_id: documentId,
        document_type: 'BC27_OUT',
        aju_number: aju,
        registration_number: '027413'
      }
    }]
  };
}

test('duplicate review detects exact document identity by Nomor Aju', () => {
  const project = createEmptyProject({ projectName: 'TEST' });
  project.data.documents.push(existingDocument);
  const rows = buildDuplicateReview(project, [staged()]);
  assert.equal(rows.length, 1);
  assert.equal(rows[0].candidate_type, 'DOCUMENT');
  assert.equal(rows[0].match_type, 'NOMOR_AJU');
  assert.equal(rows[0].existing_document_id, 'DOC-EXISTING');
});

test('duplicate review detects exact source file hash separately', () => {
  const project = createEmptyProject({ projectName: 'TEST' });
  project.sources.source_registry.push({ source_id: 'SRC-SAVED', sha256: 'hash-1', original_file_name: 'old.xlsx' });
  const rows = buildDuplicateReview(project, [staged()]);
  assert.ok(rows.some(row => row.candidate_type === 'FILE' && row.match_type === 'SHA256'));
});

test('duplicate review identifies duplicate Nomor Aju within a mixed batch', () => {
  const project = createEmptyProject({ projectName: 'TEST' });
  const rows = buildDuplicateReview(project, [
    staged({ sourceId: 'SRC-1', hash: 'hash-1', documentId: 'DOC-A' }),
    staged({ sourceId: 'SRC-2', hash: 'hash-2', documentId: 'DOC-B' })
  ]);
  assert.ok(rows.some(row => row.candidate_type === 'DOCUMENT' && row.match_type === 'NOMOR_AJU_IN_BATCH'));
});

test('duplicate decision requires reason and is persisted with audit event', () => {
  const project = createEmptyProject({ projectName: 'TEST' });
  const candidate = {
    candidate_id: 'DUP-DOCUMENT-DOC-X',
    candidate_type: 'DOCUMENT',
    staged_document_id: 'DOC-X',
    existing_document_id: 'DOC-EXISTING'
  };
  assert.throws(() => recordDuplicateDecision(project, candidate, 'KEEP_BOTH', ''), /Alasan keputusan wajib diisi/);
  const next = recordDuplicateDecision(project, candidate, 'KEEP_BOTH', 'Dua transaksi perlu dipertahankan.');
  assert.equal(getDuplicateDecision(next, candidate.candidate_id).decision, 'KEEP_BOTH');
  assert.equal(next.data.audit_log.at(-1).event_type, 'DUPLICATE_DECISION_RECORDED');
});

test('duplicate decision choices are explicitly constrained', () => {
  assert.equal(canProceedWithDuplicateDecision('USE_NEW'), true);
  assert.equal(canProceedWithDuplicateDecision('KEEP_BOTH'), true);
  assert.equal(canProceedWithDuplicateDecision('DEFER_REVIEW'), false);
});
