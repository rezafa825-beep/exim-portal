import { appendImportedDocumentBundle } from './registryEngine.js';
import { refreshProjectCounts } from './projectSchema.js';

function findDuplicateDocument(project, document) {
  return project.data.documents.find(existing =>
    (document.aju_number && (existing.aju_number === document.aju_number ||
      existing.current_data?.nomor_aju === document.aju_number ||
      existing.current_data?.aju_number === document.aju_number)) ||
    (existing.document_type === document.document_type &&
      document.registration_number &&
      (existing.registration_number === document.registration_number ||
       existing.current_data?.registration_number === document.registration_number))
  ) || null;
}

function latestDecision(project, documentId) {
  return [...(project.data.import_decisions || [])].reverse()
    .find(row => row.candidate_id === `DUP-DOCUMENT-${documentId}`) || null;
}

export function evaluateStagedImport(project, stagedFiles) {
  const seenFileHashes = new Set();
  const seenAju = new Set();
  const rows = [];

  for (const staged of stagedFiles) {
    if (staged.parse_error) {
      rows.push({ file_name: staged.file_name, status: 'FILE_ERROR', reason: staged.parse_error, documents: [] });
      continue;
    }
    const file = staged;
    const source = file.source;
    const repeatedInBatch = seenFileHashes.has(source.sha256);
    seenFileHashes.add(source.sha256);
    const existingSource = project.sources.source_registry.find(row => row.sha256 === source.sha256);
    const duplicateFile = repeatedInBatch || !!existingSource;
    const documentRows = [];

    for (const entry of file.documents) {
      const document = entry.document;
      const aju = String(document.aju_number || '').trim();
      const duplicateExisting = findDuplicateDocument(project, document);
      const duplicateInBatch = aju && seenAju.has(aju);
      const duplicateMatch = duplicateExisting || (duplicateInBatch ? { document_id: seenAju.get(aju) } : null);
      const decision = latestDecision(project, document.document_id);
      if (aju) seenAju.add(aju);

      let status = 'READY';
      let reason = 'Dokumen unik dan siap diimpor.';
      const hasError = entry.validation?.status === 'ERROR' ||
        file.issues.some(issue => issue.severity === 'ERROR' && issue.document_number === aju);
      if (hasError) {
        status = 'ERROR';
        reason = 'Dokumen memiliki validation error dan belum dapat diimpor.';
      } else if (duplicateFile) {
        status = 'DUPLICATE_FILE';
        reason = 'File dengan SHA-256 yang sama sudah terdaftar atau muncul lebih awal dalam batch.';
      } else if (duplicateMatch) {
        if (decision?.decision === 'USE_EXISTING') {
          status = 'USE_EXISTING';
          reason = 'Keputusan review: gunakan dokumen yang sudah ada.';
        } else if (['USE_NEW', 'KEEP_BOTH'].includes(decision?.decision)) {
          status = 'READY_WITH_DUPLICATE_DECISION';
          reason = `Keputusan review: ${decision.decision}.`;
        } else {
          status = decision?.decision === 'DEFER_REVIEW' ? 'POSSIBLE_DUPLICATE' : 'POSSIBLE_DUPLICATE';
          reason = duplicateExisting ? 'Nomor Aju atau Nomor Daftar sudah ditemukan di proyek.' : 'Nomor Aju yang sama ditemukan pada file lain dalam batch.';
        }
      } else if (document.document_type === 'BC27_UNKNOWN') {
        status = 'NEEDS_REVIEW';
        reason = document.extension?.classification?.reason || 'Arah transaksi perlu dikonfirmasi.';
      }
      documentRows.push({
        document_id: document.document_id,
        aju_number: document.aju_number,
        registration_number: document.registration_number,
        document_type: document.document_type,
        item_count: entry.items.length,
        existing_document_id: duplicateExisting?.document_id || (duplicateInBatch ? seenAju.get(aju) : null),
        decision: decision?.decision || null,
        status,
        reason
      });
    }

    const allDuplicate = documentRows.length > 0 && documentRows.every(row =>
      ['POSSIBLE_DUPLICATE', 'DUPLICATE_FILE', 'USE_EXISTING'].includes(row.status)
    );
    rows.push({
      source_id: source.source_id,
      file_name: source.original_file_name,
      file_format: source.file_format,
      sha256: source.sha256,
      existing_source: !!existingSource,
      status: duplicateFile ? 'DUPLICATE_FILE' : allDuplicate ? 'ALL_DOCUMENTS_DUPLICATE' : 'STAGED',
      documents: documentRows,
      issues: file.issues || [],
      summary: file.summary
    });
  }
  return rows;
}

export function commitPartialImport(project, stagedFiles, selectedDocumentIds, directionOverrides = {}) {
  const selected = new Set(selectedDocumentIds || []);
  const evaluation = evaluateStagedImport(project, stagedFiles);
  let next = structuredClone(project);
  const importedIds = [];
  const held = [];
  const importedBySource = new Map();

  for (const staged of stagedFiles) {
    if (staged.parse_error) {
      held.push({ file_name: staged.file_name, status: 'FILE_ERROR', reason: staged.parse_error });
      continue;
    }
    const evaluatedFile = evaluation.find(row => row.source_id === staged.source.source_id);
    if (!evaluatedFile || evaluatedFile.status === 'DUPLICATE_FILE') {
      held.push({ file_name: staged.source.original_file_name, status: 'DUPLICATE_FILE', reason: 'File duplikat dalam batch.' });
      continue;
    }

    for (const entry of staged.documents) {
      const check = evaluatedFile.documents.find(row => row.document_id === entry.document.document_id);
      if (!selected.has(entry.document.document_id)) {
        if (check?.status !== 'READY') held.push({ document_id: entry.document.document_id, aju_number: entry.document.aju_number, status: check?.status || 'NOT_SELECTED', reason: check?.reason || 'Tidak dipilih untuk impor.' });
        continue;
      }
      if (check?.status === 'USE_EXISTING') {
        const existingDocument = next.data.documents.find(row => row.document_id === check.existing_document_id);
        if (!existingDocument) {
          held.push({ document_id: entry.document.document_id, aju_number: entry.document.aju_number, status: 'ERROR', reason: 'Dokumen existing yang dipilih tidak ditemukan lagi.' });
          continue;
        }
        if (!importedBySource.has(staged.source.sha256)) importedBySource.set(staged.source.sha256, []);
        importedBySource.get(staged.source.sha256).push(existingDocument.document_id);
        const sourceId = staged.source.source_id;
        existingDocument.source_refs = [...new Set([...(existingDocument.source_refs || []), sourceId])];
        held.push({ document_id: entry.document.document_id, existing_document_id: existingDocument.document_id, aju_number: entry.document.aju_number, status: 'USE_EXISTING', reason: 'Sumber baru ditautkan ke dokumen existing; tidak dibuat transaksi baru.' });
        continue;
      }
      if (!check || !['READY', 'NEEDS_REVIEW', 'READY_WITH_DUPLICATE_DECISION'].includes(check.status)) {
        held.push({ document_id: entry.document.document_id, aju_number: entry.document.aju_number, status: check?.status || 'ERROR', reason: check?.reason || 'Dokumen tidak lolos pemeriksaan.' });
        continue;
      }

      const override = directionOverrides[entry.document.document_id];
      const direction = override || entry.document.extension?.classification?.direction;
      if (!['IN', 'OUT'].includes(direction)) {
        held.push({ document_id: entry.document.document_id, aju_number: entry.document.aju_number, status: 'NEEDS_REVIEW', reason: 'Pilih arah BC 2.7 IN atau OUT sebelum impor.' });
        continue;
      }

      const sourceExisting = next.sources.source_registry.find(row => row.sha256 === staged.source.sha256);
      const sourceId = sourceExisting?.source_id || staged.source.source_id;
      const document = structuredClone(entry.document);
      document.document_type = `BC27_${direction}`;
      document.extension = {
        ...document.extension,
        classification: {
          ...document.extension?.classification,
          direction,
          status: override ? 'MANUAL_OVERRIDE' : 'CLASSIFIED',
          reason: override ? 'Arah transaksi dikonfirmasi pengguna pada Import Preview.' : document.extension?.classification?.reason
        }
      };

      try {
        next = appendImportedDocumentBundle(next, {
          document,
          items: entry.items,
          raw_materials: entry.raw_materials,
          raw_material_tariffs: entry.raw_material_tariffs,
          source_id: sourceId,
          allowDuplicateIdentity: check.status === 'READY_WITH_DUPLICATE_DECISION'
        });
        if (check.status === 'READY_WITH_DUPLICATE_DECISION' && check.existing_document_id) {
          next.data.relationships.push({
            relationship_id: `REL-${crypto.randomUUID()}`,
            source_document_id: check.existing_document_id,
            target_document_id: document.document_id,
            relationship_type: check.decision === 'USE_NEW' ? 'SUPERSEDED_BY' : 'RELATED_TRANSACTION',
            match_status: 'EXACT_DUPLICATE', resolution: check.decision,
            reason: latestDecision(project, document.document_id)?.reason || 'Keputusan Duplicate Review.',
            created_at: new Date().toISOString(), operator_id: next.settings?.operator?.operator_id || 'LOCAL-USER-001'
          });
        }
        importedIds.push(document.document_id);
        if (!importedBySource.has(staged.source.sha256)) importedBySource.set(staged.source.sha256, []);
        importedBySource.get(staged.source.sha256).push(document.document_id);
      } catch (error) {
        held.push({ document_id: document.document_id, aju_number: document.aju_number, status: 'ERROR', reason: error.message });
      }
    }
  }

  for (const staged of stagedFiles) {
    if (staged.parse_error) continue;
    const importedFromFile = importedBySource.get(staged.source.sha256) || [];
    if (!importedFromFile.length) continue;
    const existing = next.sources.source_registry.find(row => row.sha256 === staged.source.sha256);
    const allDocumentIds = [...new Set([...(existing?.document_ids || []), ...importedFromFile])];
    const sourceRecord = {
      ...(existing || staged.source),
      document_ids: allDocumentIds,
      import_status: allDocumentIds.length >= staged.source.document_count ? 'IMPORTED' : 'PARTIAL'
    };
    if (existing) {
      next.sources.source_registry = next.sources.source_registry.map(row => row.source_id === existing.source_id ? sourceRecord : row);
    } else {
      next.sources.source_registry.push(sourceRecord);
      next.sources.files[sourceRecord.internal_path] = staged.source_bytes;
    }
  }

  return {
    project: refreshProjectCounts(next),
    imported_document_ids: importedIds,
    held,
    evaluation,
    summary: {
      imported_documents: importedIds.length,
      held_documents: held.length,
      total_staged_documents: stagedFiles.reduce((sum, file) => sum + (file.documents?.length || 0), 0)
    }
  };
}
