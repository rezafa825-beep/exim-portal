import { appendImportedDocumentBundle } from './registryEngine.js';
import { refreshProjectCounts } from './projectSchema.js';

function isDuplicateDocument(project, document) {
  return project.data.documents.some(existing =>
    (document.aju_number && existing.aju_number === document.aju_number) ||
    (existing.document_type === document.document_type &&
      document.registration_number &&
      existing.registration_number === document.registration_number)
  );
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
    const documentRows = [];

    for (const entry of file.documents) {
      const document = entry.document;
      const aju = String(document.aju_number || '').trim();
      const duplicateExisting = isDuplicateDocument(project, document);
      const duplicateInBatch = aju && seenAju.has(aju);
      if (aju) seenAju.add(aju);

      let status = 'READY';
      let reason = 'Dokumen unik dan siap diimpor.';
      const hasError = entry.validation?.status === 'ERROR' ||
        file.issues.some(issue => issue.severity === 'ERROR' && issue.document_number === aju);
      if (repeatedInBatch) {
        status = 'DUPLICATE_FILE';
        reason = 'File dengan SHA-256 yang sama sudah ada lebih awal dalam batch ini.';
      } else if (duplicateExisting || duplicateInBatch) {
        status = 'POSSIBLE_DUPLICATE';
        reason = duplicateExisting
          ? 'Nomor Aju atau Nomor Daftar sudah ditemukan di proyek.'
          : 'Nomor Aju yang sama ditemukan pada file lain dalam batch ini.';
      } else if (hasError) {
        status = 'ERROR';
        reason = 'Dokumen memiliki validation error dan belum dapat diimpor.';
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
        status,
        reason
      });
    }

    const allDuplicate = documentRows.length > 0 && documentRows.every(row =>
      ['POSSIBLE_DUPLICATE', 'DUPLICATE_FILE'].includes(row.status)
    );
    rows.push({
      source_id: source.source_id,
      file_name: source.original_file_name,
      file_format: source.file_format,
      sha256: source.sha256,
      existing_source: !!existingSource,
      status: repeatedInBatch ? 'DUPLICATE_FILE' : allDuplicate ? 'ALL_DOCUMENTS_DUPLICATE' : 'STAGED',
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
      if (!check || !['READY', 'NEEDS_REVIEW'].includes(check.status)) {
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
          source_id: sourceId
        });
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
