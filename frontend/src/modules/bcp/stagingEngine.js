const DOCUMENT_FIELDS = new Set([
  'current_data.nomor_daftar',
  'current_data.tanggal_daftar',
  'current_data.kode_kantor',
  'current_data.kode_kantor_tujuan',
  'current_data.tanggal_pernyataan'
]);
const ITEM_FIELDS = new Set([
  'current_data.item_code',
  'current_data.description',
  'current_data.specification',
  'current_data.quantity',
  'current_data.unit',
  'current_data.net_weight',
  'current_data.gross_weight'
]);

function clone(value) {
  return structuredClone(value);
}

function setPath(target, path, value) {
  const parts = path.split('.');
  const leaf = parts.pop();
  const parent = parts.reduce((node, part) => {
    if (!node || typeof node !== 'object' || !(part in node)) throw new Error('Field staging tidak ditemukan.');
    return node[part];
  }, target);
  if (!parent || typeof parent !== 'object' || !(leaf in parent)) throw new Error('Field staging tidak ditemukan.');
  parent[leaf] = value;
}

function syncCommonField(entry, entityType, entity, fieldPath, value) {
  if (entityType === 'DOCUMENT') {
    if (fieldPath === 'current_data.nomor_daftar') entry.document.registration_number = value;
    if (fieldPath === 'current_data.nomor_aju') entry.document.aju_number = value;
    if (fieldPath === 'current_data.tanggal_daftar') entry.document.dates.registration = value;
  } else {
    const key = fieldPath.split('.').at(-1);
    const topLevel = {
      item_code: 'item_code', description: 'description', specification: 'specification',
      quantity: 'quantity', unit: 'unit', net_weight: 'net_weight', gross_weight: 'gross_weight'
    }[key];
    if (topLevel) entity[topLevel] = value;
  }
}

/**
 * Corrects only normalized staging values. Original CEISA raw fields and source bytes remain unchanged.
 */
export function applyStagingCorrection(stagedFiles, {
  document_id, entity_type, entity_id, field_path, new_value, reason, operator_id = 'LOCAL-USER-001'
}) {
  if (!String(reason || '').trim()) throw new Error('Alasan koreksi staging wajib diisi.');
  if (!['DOCUMENT', 'ITEM'].includes(entity_type)) throw new Error('Jenis record staging tidak dikenal.');
  const allowed = entity_type === 'DOCUMENT' ? DOCUMENT_FIELDS : ITEM_FIELDS;
  if (!allowed.has(field_path)) throw new Error('Field ini tidak dapat dikoreksi pada tahap staging.');

  const next = clone(stagedFiles);
  let found = null;
  for (const file of next) {
    for (const entry of file.documents || []) {
      if (entry.document.document_id !== document_id) continue;
      const entity = entity_type === 'DOCUMENT'
        ? entry.document
        : entry.items.find(row => row.item_id === entity_id);
      if (!entity) continue;
      found = { entry, entity };
      break;
    }
    if (found) break;
  }
  if (!found) throw new Error('Dokumen atau item staging tidak ditemukan.');

  const oldValue = field_path.split('.').reduce((value, key) => value?.[key], found.entity);
  if (Object.is(oldValue, new_value)) throw new Error('Nilai baru sama dengan nilai sebelumnya.');
  if (['current_data.quantity', 'current_data.net_weight', 'current_data.gross_weight'].includes(field_path)) {
    if (new_value !== null && (!Number.isFinite(Number(new_value)) || Number(new_value) < 0)) {
      throw new Error('Nilai kuantitas atau berat harus berupa angka nol atau lebih.');
    }
    new_value = new_value === null || new_value === '' ? null : Number(new_value);
  } else if (new_value !== null) {
    new_value = String(new_value).trim();
    if (!new_value) throw new Error('Nilai teks tidak boleh kosong.');
  }

  setPath(found.entity, field_path, new_value);
  syncCommonField(found.entry, entity_type, found.entity, field_path, new_value);
  if (!Array.isArray(found.entry.staging_corrections)) found.entry.staging_corrections = [];
  found.entry.staging_corrections.push({
    correction_id: `STAGE-CORR-${crypto.randomUUID()}`,
    entity_type,
    entity_id: entity_type === 'DOCUMENT' ? document_id : entity_id,
    document_id,
    field_path,
    old_value: clone(oldValue),
    new_value: clone(new_value),
    reason: String(reason).trim(),
    operator_id,
    changed_at: new Date().toISOString()
  });
  return next;
}

export function getStagingValidationSummary(stagedFiles) {
  return (stagedFiles || []).map(file => ({
    source_id: file.source?.source_id || null,
    file_name: file.source?.original_file_name || file.file_name || 'File tidak dikenal',
    file_error: file.parse_error || null,
    documents: (file.documents || []).map(entry => ({
      document_id: entry.document.document_id,
      aju_number: entry.document.aju_number,
      document_type: entry.document.document_type,
      status: entry.validation?.status || 'UNKNOWN',
      issues: (file.issues || []).filter(issue =>
        !issue.document_number || issue.document_number === entry.document.aju_number
      ),
      corrections: entry.staging_corrections || []
    }))
  }));
}
