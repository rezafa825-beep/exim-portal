import test from 'node:test';
import assert from 'node:assert/strict';
import { createEmptyProject } from '../../src/modules/bcp/projectSchema.js';
import { createMonthlyPeriod, changeMonthlyPeriodStatus, assignDocumentToPeriod, getMonthlyPeriodSummary } from '../../src/modules/bcp/periodEngine.js';

function projectWithDocument() {
  const project = createEmptyProject({ projectName: 'PERIOD TEST' });
  project.data.documents.push({
    document_id: 'DOC-1', document_type: 'BC27_IN', aju_number: 'AJU-1',
    business_status: 'CONFIRMED', amendment_status: 'ORIGINAL', period_id: null, updated_at: new Date().toISOString()
  });
  project.data.items.push({ item_id: 'ITEM-1', document_id: 'DOC-1' });
  return project;
}

test('monthly period creation validates month and prevents duplicate period keys', () => {
  const project = projectWithDocument();
  const created = createMonthlyPeriod(project, { year: 2026, month: 11 });
  assert.equal(created.data.periods[0].period_key, '2026-11');
  assert.equal(created.data.periods[0].status, 'OPEN');
  assert.throws(() => createMonthlyPeriod(created, { year: 2026, month: 11 }), /sudah tersedia/);
  assert.throws(() => createMonthlyPeriod(project, { year: 2026, month: 13 }), /Bulan periode/);
});

test('period status transitions require reasons for close and reopen', () => {
  let project = createMonthlyPeriod(projectWithDocument(), { year: 2026, month: 11 });
  const id = project.data.periods[0].period_id;
  project = changeMonthlyPeriodStatus(project, id, 'REVIEW');
  assert.equal(project.data.periods[0].status, 'REVIEW');
  assert.throws(() => changeMonthlyPeriodStatus(project, id, 'CLOSED'), /Alasan wajib/);
  project = changeMonthlyPeriodStatus(project, id, 'CLOSED', 'Rekonsiliasi periode selesai.');
  assert.throws(() => assignDocumentToPeriod(project, 'DOC-1', id), /CLOSED/);
  assert.throws(() => changeMonthlyPeriodStatus(project, id, 'OPEN'), /Alasan wajib/);
  project = changeMonthlyPeriodStatus(project, id, 'OPEN', 'Perlu koreksi data.');
  assert.equal(project.data.periods[0].status, 'OPEN');
});

test('document period assignment and summary preserve document/item counts', () => {
  let project = createMonthlyPeriod(projectWithDocument(), { year: 2026, month: 11 });
  const periodId = project.data.periods[0].period_id;
  project = assignDocumentToPeriod(project, 'DOC-1', periodId);
  const summary = getMonthlyPeriodSummary(project, periodId);
  assert.equal(summary.document_count, 1);
  assert.equal(summary.item_count, 1);
  assert.equal(summary.documents_by_type.BC27_IN, 1);
  assert.throws(() => assignDocumentToPeriod(project, 'DOC-1', periodId), /sudah berada/);
});
