import { refreshProjectCounts } from './projectSchema.js';

function clone(value) { return structuredClone(value); }
function makeId() { return `PERIOD-${crypto.randomUUID()}`; }
function audit(project, eventType, description) {
  project.data.audit_log.push({
    event_id: `AUDIT-${crypto.randomUUID()}`,
    event_type: eventType,
    description,
    created_at: new Date().toISOString(),
    operator_id: project.settings?.operator?.operator_id || 'LOCAL-USER-001'
  });
}

export function createMonthlyPeriod(project, { year, month, label = '' }) {
  const y = Number(year), m = Number(month);
  if (!Number.isInteger(y) || y < 2000 || y > 2100) throw new Error('Tahun periode harus antara 2000 dan 2100.');
  if (!Number.isInteger(m) || m < 1 || m > 12) throw new Error('Bulan periode harus antara 1 dan 12.');
  const key = `${y}-${String(m).padStart(2, '0')}`;
  if (project.data.periods.some(row => row.period_key === key)) throw new Error(`Periode ${key} sudah tersedia.`);

  const next = clone(project);
  const now = new Date().toISOString();
  next.data.periods.push({
    period_id: makeId(),
    period_key: key,
    year: y,
    month: m,
    label: String(label || `Periode ${key}`).trim(),
    status: 'OPEN',
    document_ids: [],
    created_at: now,
    updated_at: now,
    review_started_at: null,
    closed_at: null,
    reopened_at: null
  });
  audit(next, 'MONTHLY_PERIOD_CREATED', `Periode ${key} dibuat.`);
  return refreshProjectCounts(next);
}

export function changeMonthlyPeriodStatus(project, periodId, nextStatus, reason = '') {
  const allowed = new Set(['OPEN', 'REVIEW', 'CLOSED']);
  if (!allowed.has(nextStatus)) throw new Error('Status periode tidak dikenal.');
  const next = clone(project);
  const period = next.data.periods.find(row => row.period_id === periodId);
  if (!period) throw new Error('Periode tidak ditemukan.');
  const previous = period.status;
  if (previous === nextStatus) throw new Error('Status periode sudah sesuai.');
  const transitions = {
    OPEN: ['REVIEW', 'CLOSED'],
    REVIEW: ['OPEN', 'CLOSED'],
    CLOSED: ['OPEN']
  };
  if (!transitions[previous]?.includes(nextStatus)) throw new Error(`Perubahan status ${previous} ke ${nextStatus} tidak diizinkan.`);
  if ((nextStatus === 'CLOSED' || previous === 'CLOSED') && !String(reason || '').trim()) {
    throw new Error('Alasan wajib diisi untuk menutup atau membuka kembali periode.');
  }

  const now = new Date().toISOString();
  period.status = nextStatus;
  period.updated_at = now;
  if (nextStatus === 'REVIEW') period.review_started_at = now;
  if (nextStatus === 'CLOSED') period.closed_at = now;
  if (previous === 'CLOSED' && nextStatus === 'OPEN') period.reopened_at = now;
  audit(next, 'MONTHLY_PERIOD_STATUS_CHANGED', `Periode ${period.period_key}: ${previous} → ${nextStatus}. ${String(reason || '').trim()}`);
  return refreshProjectCounts(next);
}

export function assignDocumentToPeriod(project, documentId, periodId, reason = '') {
  const next = clone(project);
  const document = next.data.documents.find(row => row.document_id === documentId);
  const period = next.data.periods.find(row => row.period_id === periodId);
  if (!document) throw new Error('Dokumen tidak ditemukan.');
  if (!period) throw new Error('Periode tidak ditemukan.');
  if (period.status === 'CLOSED') throw new Error('Periode CLOSED harus dibuka kembali sebelum menerima perubahan.');
  if (document.period_id === periodId) throw new Error('Dokumen sudah berada pada periode tersebut.');
  if (document.period_id && !String(reason || '').trim()) throw new Error('Alasan wajib diisi untuk memindahkan dokumen antarperiode.');

  const previousPeriod = next.data.periods.find(row => row.period_id === document.period_id);
  if (previousPeriod?.status === 'CLOSED') throw new Error('Periode asal CLOSED harus dibuka kembali sebelum dokumen dipindahkan.');
  if (previousPeriod) previousPeriod.document_ids = previousPeriod.document_ids.filter(id => id !== documentId);
  if (!period.document_ids.includes(documentId)) period.document_ids.push(documentId);
  document.period_id = periodId;
  document.updated_at = new Date().toISOString();
  period.updated_at = document.updated_at;
  audit(next, 'DOCUMENT_PERIOD_ASSIGNED', `Dokumen ${document.aju_number || document.document_id} dimasukkan ke periode ${period.period_key}. ${String(reason || '').trim()}`);
  return refreshProjectCounts(next);
}

export function getMonthlyPeriodSummary(project, periodId) {
  const period = project.data.periods.find(row => row.period_id === periodId);
  if (!period) return null;
  const documents = project.data.documents.filter(row => row.period_id === periodId);
  const documentIds = new Set(documents.map(row => row.document_id));
  const items = project.data.items.filter(row => documentIds.has(row.document_id));
  const byType = {};
  for (const document of documents) byType[document.document_type] = (byType[document.document_type] || 0) + 1;
  return {
    period: clone(period),
    document_count: documents.length,
    item_count: items.length,
    confirmed_count: documents.filter(row => row.business_status === 'CONFIRMED').length,
    void_count: documents.filter(row => row.business_status === 'VOID').length,
    cancelled_count: documents.filter(row => row.business_status === 'CANCELLED').length,
    amended_count: documents.filter(row => row.amendment_status === 'AMENDED').length,
    documents_by_type: byType
  };
}
