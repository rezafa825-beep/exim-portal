import { refreshProjectCounts } from './projectSchema.js';

const BUSINESS_STATUSES = new Set(['CONFIRMED', 'VOID', 'CANCELLED']);
const SENSITIVE_FIELDS = new Set([
  'current_data.document_number',
  'current_data.registration_number',
  'current_data.dates.aju',
  'current_data.dates.registration'
]);
const FORBIDDEN_PATH_PARTS = new Set(['__proto__', 'prototype', 'constructor']);

function clone(value) {
  return value === undefined ? undefined : structuredClone(value);
}

function makeId(prefix) {
  return `${prefix}-${crypto.randomUUID()}`;
}

function addAudit(project, eventType, description, operatorId) {
  project.data.audit_log.push({
    event_id: makeId('AUDIT'),
    event_type: eventType,
    description,
    created_at: new Date().toISOString(),
    operator_id: operatorId || project.settings?.operator?.operator_id || 'LOCAL-USER-001'
  });
}

function normalizeDocument(input, sourceId) {
  const now = new Date().toISOString();
  const current = clone(input.current_data || input.data || {});
  return {
    document_id: input.document_id || makeId('DOC'),
    document_type: input.document_type || 'UNKNOWN',
    aju_number: input.aju_number ?? current.aju_number ?? current.nomor_aju ?? null,
    document_number: input.document_number ?? current.document_number ?? null,
    registration_number: input.registration_number ?? current.registration_number ?? null,
    dates: clone(input.dates || current.dates || {}),
    parties: clone(input.parties || current.parties || {}),
    period_id: input.period_id || null,
    source_refs: [...new Set([...(input.source_refs || []), ...(sourceId ? [sourceId] : [])])],
    processing_status: 'COMPLETED',
    business_status: input.business_status || 'CONFIRMED',
    amendment_status: 'ORIGINAL',
    original_data: clone(input.staging_original_data || current),
    current_data: clone(current),
    extension: clone(input.extension || {}),
    created_at: now,
    updated_at: now
  };
}

function normalizeItem(input, documentId, sourceId) {
  const now = new Date().toISOString();
  const current = clone(input.current_data || input.data || {});
  return {
    item_id: input.item_id || makeId('ITEM'),
    document_id: documentId,
    item_sequence: input.item_sequence ?? input.item_sequence_number ?? null,
    item_code: input.item_code ?? current.item_code ?? null,
    description: input.description ?? current.description ?? null,
    specification: input.specification ?? current.specification ?? null,
    quantity: input.quantity ?? current.quantity ?? null,
    unit: input.unit ?? current.unit ?? null,
    net_weight: input.net_weight ?? current.net_weight ?? null,
    gross_weight: input.gross_weight ?? current.gross_weight ?? null,
    currency: input.currency ?? current.currency ?? null,
    item_value: input.item_value ?? current.item_value ?? null,
    source_refs: [...new Set([...(input.source_refs || []), ...(sourceId ? [sourceId] : [])])],
    source_locations: clone(input.source_locations || []),
    original_data: clone(input.staging_original_data || current),
    current_data: clone(current),
    extension: clone(input.extension || {}),
    created_at: now,
    updated_at: now
  };
}

export function appendImportedDocument(project, documentInput, itemInputs = [], sourceId = null, { allowDuplicateIdentity = false } = {}) {
  if (!documentInput || typeof documentInput !== 'object') throw new Error('Data dokumen tidak valid.');
  if (!Array.isArray(itemInputs)) throw new Error('Daftar barang harus berupa array.');

  const next = clone(project);
  const document = normalizeDocument(documentInput, sourceId);
  if (next.data.documents.some(row => row.document_id === document.document_id)) {
    throw new Error('Document ID sudah digunakan.');
  }
  if (!allowDuplicateIdentity && next.data.documents.some(row =>
    (document.aju_number && row.aju_number === document.aju_number) ||
    (row.document_type === document.document_type && document.registration_number && row.registration_number === document.registration_number)
  )) {
    throw new Error('Kandidat duplikasi nomor daftar ditemukan. Jalankan Duplicate Review sebelum commit.');
  }

  const items = itemInputs.map(item => normalizeItem(item, document.document_id, sourceId));
  const sequences = new Set();
  for (const item of items) {
    if (!item.item_id || next.data.items.some(row => row.item_id === item.item_id)) {
      throw new Error('Item ID tidak valid atau sudah digunakan.');
    }
    if (item.item_sequence !== null && item.item_sequence !== undefined) {
      const sequence = String(item.item_sequence);
      if (sequences.has(sequence)) throw new Error('Seri barang ganda ditemukan di dalam satu dokumen.');
      sequences.add(sequence);
    }
  }

  next.data.documents.push(document);
  next.data.items.push(...items);
  addAudit(next, 'DOCUMENT_IMPORTED', `Dokumen ${document.registration_number || document.document_id} dan ${items.length} detail barang diimpor.`);
  return refreshProjectCounts(next);
}

function getByPath(object, path) {
  return path.split('.').reduce((value, key) => value?.[key], object);
}

function setByPath(object, path, value) {
  const parts = path.split('.');
  if (parts.some(part => FORBIDDEN_PATH_PARTS.has(part))) throw new Error('Field path tidak diizinkan.');
  const last = parts.pop();
  const target = parts.reduce((value, key) => {
    if (!value || typeof value !== 'object' || !(key in value)) throw new Error('Field path tidak ditemukan.');
    return value[key];
  }, object);
  if (!target || typeof target !== 'object' || !(last in target)) throw new Error('Field path tidak ditemukan.');
  target[last] = clone(value);
}

export function applyCorrection(project, {
  entity_type, entity_id, field_path, new_value, reason, operator_id, confirmed_sensitive = false
}) {
  if (!['DOCUMENT', 'ITEM'].includes(entity_type)) throw new Error('Jenis record koreksi tidak dikenal.');
  if (!String(reason || '').trim()) throw new Error('Alasan koreksi wajib diisi.');
  if (!String(field_path || '').startsWith('current_data.')) throw new Error('Koreksi hanya boleh mengubah current_data.');
  if (SENSITIVE_FIELDS.has(field_path) && !confirmed_sensitive) {
    throw new Error('Field sensitif memerlukan konfirmasi tambahan sebelum dikoreksi.');
  }

  const next = clone(project);
  const collection = entity_type === 'DOCUMENT' ? next.data.documents : next.data.items;
  const entity = collection.find(row => (entity_type === 'DOCUMENT' ? row.document_id : row.item_id) === entity_id);
  if (!entity) throw new Error('Record yang akan dikoreksi tidak ditemukan.');

  const oldValue = getByPath(entity, field_path);
  if (JSON.stringify(oldValue) === JSON.stringify(new_value)) throw new Error('Nilai baru sama dengan nilai saat ini.');
  setByPath(entity, field_path, new_value);
  entity.updated_at = new Date().toISOString();

  const document = entity_type === 'DOCUMENT'
    ? entity
    : next.data.documents.find(row => row.document_id === entity.document_id);
  if (document) {
    document.amendment_status = 'AMENDED';
    document.updated_at = new Date().toISOString();
  }

  const correction = {
    correction_id: makeId('CORR'),
    entity_type,
    entity_id,
    document_id: entity_type === 'DOCUMENT' ? entity_id : entity.document_id,
    field_path,
    old_value: clone(oldValue),
    new_value: clone(new_value),
    reason: String(reason).trim(),
    operator_id: operator_id || next.settings?.operator?.operator_id || 'LOCAL-USER-001',
    changed_at: new Date().toISOString()
  };
  next.data.corrections.push(correction);
  addAudit(next, 'DATA_CORRECTED', `${entity_type} ${entity_id}: ${field_path} dikoreksi. Alasan: ${correction.reason}`, correction.operator_id);
  return refreshProjectCounts(next);
}

export function setDocumentBusinessStatus(project, documentId, status, reason, operatorId) {
  if (!BUSINESS_STATUSES.has(status)) throw new Error('Business status tidak dikenal.');
  if (status !== 'CONFIRMED' && !String(reason || '').trim()) {
    throw new Error('Alasan wajib diisi untuk status VOID atau CANCELLED.');
  }
  const next = clone(project);
  const document = next.data.documents.find(row => row.document_id === documentId);
  if (!document) throw new Error('Dokumen tidak ditemukan.');
  const previous = document.business_status;
  document.business_status = status;
  document.updated_at = new Date().toISOString();
  addAudit(next, 'BUSINESS_STATUS_CHANGED', `Status dokumen ${document.registration_number || documentId}: ${previous} → ${status}. ${reason || ''}`, operatorId);
  return refreshProjectCounts(next);
}

export function addDocumentRelationship(project, {
  source_document_id, target_document_id, relationship_type, match_status = 'RELATED', resolution = 'UNRESOLVED', reason = ''
}) {
  if (source_document_id === target_document_id) throw new Error('Dokumen tidak dapat dihubungkan dengan dirinya sendiri.');
  const next = clone(project);
  const ids = new Set(next.data.documents.map(row => row.document_id));
  if (!ids.has(source_document_id) || !ids.has(target_document_id)) throw new Error('Dokumen sumber atau tujuan tidak ditemukan.');
  const exists = next.data.relationships.some(row =>
    row.source_document_id === source_document_id &&
    row.target_document_id === target_document_id &&
    row.relationship_type === relationship_type
  );
  if (exists) throw new Error('Hubungan antar dokumen tersebut sudah tercatat.');
  next.data.relationships.push({
    relationship_id: makeId('REL'),
    source_document_id,
    target_document_id,
    relationship_type,
    match_status,
    resolution,
    reason,
    created_at: new Date().toISOString(),
    operator_id: next.settings?.operator?.operator_id || 'LOCAL-USER-001'
  });
  addAudit(next, 'DOCUMENT_RELATIONSHIP_ADDED', `Hubungan ${relationship_type} ditambahkan.`);
  return refreshProjectCounts(next);
}

export function getDocumentView(project, documentId) {
  const document = project.data.documents.find(row => row.document_id === documentId);
  if (!document) return null;
  const items = project.data.items.filter(row => row.document_id === documentId).map(item => ({
    ...clone(item),
    raw_materials: project.data.raw_materials.filter(row => row.item_id === item.item_id).map(material => ({
      ...clone(material),
      tariffs: project.data.raw_material_tariffs.filter(row => row.raw_material_id === material.raw_material_id).map(clone)
    }))
  }));
  return { ...clone(document), items };
}

export function searchDocuments(project, query = '') {
  const term = String(query).trim().toLowerCase();
  if (!term) return [...project.data.documents];
  return project.data.documents.filter(row => {
    const text = [
      row.document_type, row.aju_number, row.document_number, row.registration_number,
      row.current_data?.nomor_aju, row.current_data?.aju_number, row.current_data?.document_number, row.current_data?.registration_number,
      row.current_data?.parties?.sender?.name, row.current_data?.parties?.receiver?.name
    ].filter(Boolean).join(' ').toLowerCase();
    return text.includes(term);
  });
}

export function searchItems(project, query = '') {
  const term = String(query).trim().toLowerCase();
  if (!term) return [...project.data.items];
  return project.data.items.filter(row => {
    const text = [row.item_code, row.description, row.specification, row.current_data?.item_code, row.current_data?.description].filter(Boolean).join(' ').toLowerCase();
    return text.includes(term);
  });
}

export function getItemTransactionHistory(project, itemCode) {
  const code = String(itemCode || '').trim().toLowerCase();
  if (!code) return [];
  return project.data.items
    .filter(item => String(item.item_code || item.current_data?.item_code || '').trim().toLowerCase() === code)
    .map(item => ({
      ...clone(item),
      document: project.data.documents.find(doc => doc.document_id === item.document_id) || null
    }))
    .sort((a, b) => new Date(a.document?.dates?.registration || a.document?.current_data?.dates?.registration || 0) - new Date(b.document?.dates?.registration || b.document?.current_data?.dates?.registration || 0));
}


export function appendImportedDocumentBundle(project, {
  document, items = [], raw_materials = [], raw_material_tariffs = [], source_id = null, allowDuplicateIdentity = false
}) {
  let next = appendImportedDocument(project, document, items, source_id, { allowDuplicateIdentity });
  const documentId = next.data.documents.at(-1).document_id;
  const itemBySequence = new Map(
    next.data.items.filter(item => item.document_id === documentId)
      .map(item => [String(item.item_sequence), item])
  );
  const rawMaterialByKey = new Map();

  for (const input of raw_materials) {
    const item = itemBySequence.get(String(input.item_sequence));
    if (!item) throw new Error(`Bahan baku seri ${input.raw_material_sequence} tidak memiliki parent Seri Barang ${input.item_sequence}.`);
    const current = clone(input.current_data || input.data || {});
    const rawMaterial = {
      raw_material_id: input.raw_material_id || makeId('BB'),
      document_id: documentId,
      item_id: item.item_id,
      item_sequence: item.item_sequence,
      raw_material_sequence: input.raw_material_sequence ?? current.seri_bahan_baku ?? null,
      origin_code: input.origin_code ?? current.kode_asal_bahan_baku ?? null,
      item_code: input.item_code ?? current.item_code ?? null,
      description: input.description ?? current.description ?? null,
      quantity: input.quantity ?? current.quantity ?? null,
      unit: input.unit ?? current.unit ?? null,
      source_refs: [...new Set([...(input.source_refs || []), ...(source_id ? [source_id] : [])])],
      source_locations: clone(input.source_locations || []),
      original_data: clone(current),
      current_data: clone(current),
      extension: clone(input.extension || {}),
      created_at: new Date().toISOString(),
      updated_at: new Date().toISOString()
    };
    const key = `${String(item.item_sequence)}::${String(rawMaterial.raw_material_sequence)}`;
    if (rawMaterialByKey.has(key)) throw new Error('Seri bahan baku ganda ditemukan pada satu seri barang.');
    rawMaterialByKey.set(key, rawMaterial);
    next.data.raw_materials.push(rawMaterial);
  }

  for (const input of raw_material_tariffs) {
    const key = `${String(input.item_sequence)}::${String(input.raw_material_sequence)}`;
    const rawMaterial = rawMaterialByKey.get(key) || next.data.raw_materials.find(row =>
      row.document_id === documentId &&
      String(row.item_sequence) === String(input.item_sequence) &&
      String(row.raw_material_sequence) === String(input.raw_material_sequence)
    );
    if (!rawMaterial) throw new Error('Tarif bahan baku tidak memiliki parent bahan baku.');
    const current = clone(input.current_data || input.data || {});
    next.data.raw_material_tariffs.push({
      raw_material_tariff_id: input.raw_material_tariff_id || makeId('BBT'),
      document_id: documentId,
      item_id: rawMaterial.item_id,
      raw_material_id: rawMaterial.raw_material_id,
      item_sequence: input.item_sequence,
      raw_material_sequence: input.raw_material_sequence,
      tax_code: input.tax_code ?? current.kode_pungutan ?? null,
      tariff_code: input.tariff_code ?? current.kode_tarif ?? null,
      tariff: input.tariff ?? current.tarif ?? null,
      source_refs: [...new Set([...(input.source_refs || []), ...(source_id ? [source_id] : [])])],
      source_location: clone(input.source_location || null),
      original_data: clone(current),
      current_data: clone(current),
      extension: clone(input.extension || {}),
      created_at: new Date().toISOString()
    });
  }

  return refreshProjectCounts(next);
}

export function validateRegistryIntegrity(project) {
  const errors = [];
  const documentIds = new Set(project.data.documents.map(row => row.document_id));
  const itemIds = new Set();
  for (const item of project.data.items) {
    if (!documentIds.has(item.document_id)) errors.push(`Item ${item.item_id} tidak memiliki dokumen induk.`);
    if (itemIds.has(item.item_id)) errors.push(`Item ID duplikat: ${item.item_id}`);
    itemIds.add(item.item_id);
  }
  const seenDocuments = new Set();
  for (const document of project.data.documents) {
    if (seenDocuments.has(document.document_id)) errors.push(`Document ID duplikat: ${document.document_id}`);
    seenDocuments.add(document.document_id);
  }
  const rawMaterialIds = new Set();
  for (const material of project.data.raw_materials) {
    if (!itemIds.has(material.item_id)) errors.push(`Bahan baku ${material.raw_material_id} tidak memiliki item induk.`);
    if (rawMaterialIds.has(material.raw_material_id)) errors.push(`Raw material ID duplikat: ${material.raw_material_id}`);
    rawMaterialIds.add(material.raw_material_id);
  }
  for (const tariff of project.data.raw_material_tariffs) {
    if (!rawMaterialIds.has(tariff.raw_material_id)) errors.push(`Tarif ${tariff.raw_material_tariff_id} tidak memiliki bahan baku induk.`);
  }
  for (const relationship of project.data.relationships) {
    if (!documentIds.has(relationship.source_document_id) || !documentIds.has(relationship.target_document_id)) {
      errors.push(`Relationship ${relationship.relationship_id} mengarah ke dokumen yang tidak tersedia.`);
    }
  }
  return errors;
}
