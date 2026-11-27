const DECISIONS = new Set(['USE_EXISTING', 'USE_NEW', 'KEEP_BOTH', 'DEFER_REVIEW']);

function text(value) {
  return String(value ?? '').trim();
}

function makeCandidateId(type, key) {
  return `DUP-${type}-${key}`;
}

function findDocumentMatch(project, stagedDocument) {
  const aju = text(stagedDocument.aju_number);
  const registration = text(stagedDocument.registration_number);
  const type = text(stagedDocument.document_type);
  if (aju) {
    const match = project.data.documents.find(row => text(row.aju_number || row.current_data?.nomor_aju || row.current_data?.aju_number) === aju);
    if (match) return { match, match_type: 'NOMOR_AJU', confidence: 'EXACT' };
  }
  if (registration && type) {
    const match = project.data.documents.find(row =>
      text(row.document_type) === type &&
      text(row.registration_number || row.current_data?.registration_number) === registration
    );
    if (match) return { match, match_type: 'NOMOR_DAFTAR_AND_TYPE', confidence: 'EXACT' };
  }
  return null;
}

/**
 * Produces review candidates without changing the project or staged records.
 * File hashes and document identities are reviewed separately.
 */
export function buildDuplicateReview(project, stagedFiles) {
  const candidates = [];
  const seenHashes = new Map();
  const seenAju = new Map();

  for (const staged of stagedFiles || []) {
    if (staged.parse_error || !staged.source) continue;
    const source = staged.source;
    const savedSource = project.sources.source_registry.find(row => row.sha256 && row.sha256 === source.sha256);
    const earlierFile = seenHashes.get(source.sha256);
    if (savedSource || earlierFile) {
      candidates.push({
        candidate_id: makeCandidateId('FILE', source.sha256),
        candidate_type: 'FILE',
        source_id: source.source_id,
        file_name: source.original_file_name,
        existing_source_id: savedSource?.source_id || earlierFile?.source_id || null,
        existing_file_name: savedSource?.original_file_name || earlierFile?.original_file_name || null,
        match_type: 'SHA256',
        confidence: 'EXACT',
        reason: 'File memiliki SHA-256 yang sama dengan sumber yang sudah terdaftar atau file lain dalam batch.'
      });
    } else {
      seenHashes.set(source.sha256, source);
    }

    for (const entry of staged.documents || []) {
      const document = entry.document || {};
      const aju = text(document.aju_number);
      const existing = findDocumentMatch(project, document);
      const batchMatch = aju ? seenAju.get(aju) : null;
      if (aju && !batchMatch) seenAju.set(aju, { document, source });
      if (!existing && !batchMatch) continue;

      const match = existing || {
        match: batchMatch.document,
        match_type: 'NOMOR_AJU_IN_BATCH',
        confidence: 'EXACT'
      };
      const matchedDocument = match.match;
      const key = entry.document.document_id;
      candidates.push({
        candidate_id: makeCandidateId('DOCUMENT', key),
        candidate_type: 'DOCUMENT',
        staged_document_id: key,
        staged_source_id: source.source_id,
        file_name: source.original_file_name,
        aju_number: document.aju_number || null,
        registration_number: document.registration_number || null,
        document_type: document.document_type || null,
        existing_document_id: matchedDocument.document_id || null,
        existing_aju_number: matchedDocument.aju_number || null,
        existing_registration_number: matchedDocument.registration_number || null,
        match_type: match.match_type,
        confidence: match.confidence,
        reason: match.match_type === 'NOMOR_AJU'
          ? 'Nomor Aju sudah ditemukan pada proyek.'
          : match.match_type === 'NOMOR_DAFTAR_AND_TYPE'
            ? 'Nomor Daftar dan jenis BC sudah ditemukan pada proyek.'
            : 'Nomor Aju yang sama ditemukan pada file lain dalam batch.'
      });
    }
  }

  return candidates.map(candidate => ({
    ...candidate,
    decision: project.data.import_decisions?.find(row => row.candidate_id === candidate.candidate_id) || null
  }));
}

export function recordDuplicateDecision(project, candidate, decision, reason) {
  if (!candidate?.candidate_id) throw new Error('Kandidat duplikasi tidak valid.');
  if (!DECISIONS.has(decision)) throw new Error('Keputusan duplikasi tidak dikenal.');
  if (!text(reason)) throw new Error('Alasan keputusan wajib diisi.');

  const next = structuredClone(project);
  if (!Array.isArray(next.data.import_decisions)) next.data.import_decisions = [];
  const now = new Date().toISOString();
  const record = {
    decision_id: `DEC-${crypto.randomUUID()}`,
    candidate_id: candidate.candidate_id,
    candidate_type: candidate.candidate_type,
    staged_document_id: candidate.staged_document_id || null,
    staged_source_id: candidate.staged_source_id || candidate.source_id || null,
    existing_document_id: candidate.existing_document_id || null,
    existing_source_id: candidate.existing_source_id || null,
    decision,
    reason: text(reason),
    decided_at: now,
    operator_id: next.settings?.operator?.operator_id || 'LOCAL-USER-001'
  };
  next.data.import_decisions.push(record);
  next.data.audit_log.push({
    event_id: `AUDIT-${crypto.randomUUID()}`,
    event_type: 'DUPLICATE_DECISION_RECORDED',
    description: `Keputusan duplikasi ${candidate.candidate_id}: ${decision}. Alasan: ${record.reason}`,
    created_at: now,
    operator_id: record.operator_id
  });
  next.manifest.updated_at = now;
  return next;
}

export function getDuplicateDecision(project, candidateId) {
  return [...(project.data.import_decisions || [])].reverse().find(row => row.candidate_id === candidateId) || null;
}

export function canProceedWithDuplicateDecision(decision) {
  return ['USE_NEW', 'KEEP_BOTH', 'USE_EXISTING'].includes(decision);
}
