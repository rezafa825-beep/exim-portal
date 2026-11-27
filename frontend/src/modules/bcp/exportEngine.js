import * as XLSX from 'xlsx';
import { refreshProjectCounts } from './projectSchema.js';

export const EXPORT_PRESETS = [
  'DOC_REGISTRY','ITEM_REGISTRY','MONTHLY_SUMMARY','BC_TYPE_REPORT',
  'TRANSACTION_REPORT','SOURCE_REPORT','CORRECTION_REPORT','AUDIT_REPORT'
];

const DOC_FIELDS = [
  { key: 'document_type', label: 'Jenis BC', type: 'text' },
  { key: 'nomor_aju', label: 'Nomor Aju', type: 'text' },
  { key: 'nomor_daftar', label: 'Nomor Daftar', type: 'text' },
  { key: 'business_status', label: 'Status Dokumen', type: 'text' },
  { key: 'amendment_status', label: 'Status Amendment', type: 'text' },
  { key: 'period_id', label: 'Period ID', type: 'text' },
  { key: 'tanggal_daftar', label: 'Tanggal Daftar', type: 'date' },
  { key: 'tanggal_pernyataan', label: 'Tanggal Pernyataan', type: 'date' }
];
const ITEM_FIELDS = [
  { key: 'document_id', label: 'Document ID', type: 'text' },
  { key: 'item_sequence', label: 'Seri Barang', type: 'text' },
  { key: 'item_code', label: 'Kode Barang', type: 'text' },
  { key: 'description', label: 'Uraian Barang', type: 'text' },
  { key: 'specification', label: 'Spesifikasi', type: 'text' },
  { key: 'quantity', label: 'Jumlah', type: 'number' },
  { key: 'unit', label: 'Satuan', type: 'text' },
  { key: 'net_weight', label: 'Netto', type: 'number' },
  { key: 'gross_weight', label: 'Bruto', type: 'number' },
  { key: 'item_value', label: 'Nilai Barang', type: 'number' }
];

function getPath(object, path) {
  return path.split('.').reduce((value, key) => value?.[key], object);
}
function safeCell(value, type = 'text') {
  if (value === null || value === undefined || value === '') return '';
  if (type === 'number') {
    const number = Number(value);
    return Number.isFinite(number) ? number : '';
  }
  if (type === 'date') {
    const date = new Date(value);
    return Number.isNaN(date.getTime()) ? String(value) : date;
  }
  const string = String(value);
  return /^[\s]*[=+\-@]/.test(string) ? `'${string}` : string;
}
function rowsFromFields(records, fields, dataVersion = 'CURRENT', custom = {}) {
  const label = field => custom.label_overrides?.[field.key] || field.label;
  const header = fields.map(label);
  const rows = records.map(record => fields.map(field => {
    const source = dataVersion === 'ORIGINAL' ? record.original_data : record.current_data;
    const value = field.key.includes('.') ? getPath(source || {}, field.key) : (source?.[field.key] ?? record[field.key]);
    return safeCell(value, field.type);
  }));
  return [header, ...rows];
}
function addSplitSheets(workbook, baseName, matrix) {
  const headers = matrix[0] || [];
  const rows = matrix.slice(1);
  const maxRows = 1048575;
  if (!rows.length) {
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(matrix), baseName.slice(0,31));
    return;
  }
  for (let start = 0, part = 1; start < rows.length; start += maxRows, part++) {
    const name = part === 1 ? baseName : `${baseName}_${part}`;
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([headers, ...rows.slice(start, start + maxRows)]), name.slice(0,31));
  }
}
function currentDocumentRows(project) {
  return project.data.documents.filter(row => row.business_status === 'CONFIRMED');
}
function buildMonthlySummary(project) {
  const headers = ['Periode','Status Periode','Jenis BC','Jumlah Dokumen','Jumlah Item','Amended','VOID','CANCELLED'];
  const rows = [];
  for (const period of project.data.periods) {
    const docs = currentDocumentRows(project).filter(doc => doc.period_id === period.period_id);
    const types = [...new Set(docs.map(doc => doc.document_type))];
    for (const type of types.length ? types : ['—']) {
      const subset = docs.filter(doc => doc.document_type === type);
      const ids = new Set(subset.map(doc => doc.document_id));
      rows.push([
        period.period_key, period.status, type, subset.length,
        project.data.items.filter(item => ids.has(item.document_id)).length,
        subset.filter(doc => doc.amendment_status === 'AMENDED').length,
        project.data.documents.filter(doc => doc.period_id === period.period_id && doc.business_status === 'VOID' && doc.document_type === type).length,
        project.data.documents.filter(doc => doc.period_id === period.period_id && doc.business_status === 'CANCELLED' && doc.document_type === type).length
      ]);
    }
  }
  return [headers, ...rows];
}
function comparisonRows(records, fields, custom = {}) {
  const label = field => custom.label_overrides?.[field.key] || field.label;
  const headers = ['Record ID', ...fields.flatMap(field => [`${label(field)} - Original`, `${label(field)} - Current`])];
  return [headers, ...records.map(record => [
    record.document_id || record.item_id,
    ...fields.flatMap(field => {
      const original = field.key.includes('.') ? getPath(record.original_data || {}, field.key) : (record.original_data?.[field.key] ?? record[field.key]);
      const current = field.key.includes('.') ? getPath(record.current_data || {}, field.key) : (record.current_data?.[field.key] ?? record[field.key]);
      return [safeCell(original, field.type), safeCell(current, field.type)];
    })
  ])];
}

export function buildExportWorkbook(project, {
  preset = 'DOC_REGISTRY',
  dataVersion = 'CURRENT',
  custom = null,
  filters = {}
} = {}) {
  if (!EXPORT_PRESETS.includes(preset) && preset !== 'CUSTOM') throw new Error('Preset export tidak dikenal.');
  if (!['ORIGINAL','CURRENT','COMPARISON'].includes(dataVersion)) throw new Error('Data Version tidak dikenal.');
  const workbook = XLSX.utils.book_new();
  const docs = project.data.documents.filter(row => filters.business_status ? row.business_status === filters.business_status : (filters.include_inactive || row.business_status === 'CONFIRMED'));
  const items = project.data.items.filter(row => docs.some(doc => doc.document_id === row.document_id));
  const docFields = DOC_FIELDS;
  const itemFields = ITEM_FIELDS;
  const addData = (name, records, fields, entityType, config = {}) => {
    const normalizedFields = fields.map(field => ({ ...field, label: config.label_overrides?.[field.key] || field.label }));
    let selectedRecords = [...records];
    for (const [key, value] of Object.entries(config.filters || {})) {
      if (value === '' || value === null || value === undefined) continue;
      const needle = String(value).toLowerCase();
      selectedRecords = selectedRecords.filter(record => {
        const current = record.current_data || {};
        const actual = getPath(current, key) ?? getPath(record, key);
        return String(actual ?? '').toLowerCase().includes(needle);
      });
    }
    if (config.sort_by || config.group_by) {
      const key = config.sort_by || config.group_by;
      const direction = config.sort_by && config.sort_direction === 'DESC' ? -1 : 1;
      selectedRecords.sort((a,b) => {
        const av = getPath(a.current_data || {}, key) ?? getPath(a, key) ?? '';
        const bv = getPath(b.current_data || {}, key) ?? getPath(b, key) ?? '';
        const an = Number(av), bn = Number(bv);
        const compare = av !== '' && bv !== '' && Number.isFinite(an) && Number.isFinite(bn)
          ? an - bn : String(av).localeCompare(String(bv), 'id', { numeric: true, sensitivity: 'base' });
        return compare * direction;
      });
    }
    const matrix = dataVersion === 'COMPARISON'
      ? comparisonRows(selectedRecords, normalizedFields, config)
      : rowsFromFields(selectedRecords, normalizedFields, dataVersion, config);
    if (config.group_by) {
      const groupField = normalizedFields.find(field => field.key === config.group_by);
      if (groupField) {
        matrix[0].unshift('Group: ' + (config.label_overrides?.[groupField.key] || groupField.label));
        for (let index = 1; index < matrix.length; index++) {
          const record = selectedRecords[index - 1];
          matrix[index].unshift(String(getPath(record.current_data || {}, config.group_by) ?? getPath(record, config.group_by) ?? '—'));
        }
      }
    }
    addSplitSheets(workbook, name, matrix);
  };

  if (preset === 'DOC_REGISTRY') addData('DOCUMENT_REGISTRY', docs, docFields, 'DOCUMENT');
  if (preset === 'ITEM_REGISTRY') addData('ITEM_REGISTRY', items, itemFields, 'ITEM');
  if (preset === 'MONTHLY_SUMMARY') addSplitSheets(workbook, 'MONTHLY_SUMMARY', buildMonthlySummary(project));
  if (preset === 'BC_TYPE_REPORT') {
    const rows = [['Jenis BC','Jumlah Dokumen','Jumlah Barang']];
    for (const type of [...new Set(docs.map(row => row.document_type))].sort()) {
      const subset = docs.filter(row => row.document_type === type);
      const ids = new Set(subset.map(row => row.document_id));
      rows.push([type, subset.length, items.filter(item => ids.has(item.document_id)).length]);
    }
    addSplitSheets(workbook, 'BC_TYPE_REPORT', rows);
  }
  if (preset === 'TRANSACTION_REPORT') {
    const rows = [['Jenis BC','Nomor Aju','Nomor Daftar','Seri Barang','Kode Barang','Uraian','Jumlah','Satuan','Status']];
    for (const item of items) {
      const doc = docs.find(row => row.document_id === item.document_id);
      if (doc) rows.push([doc.document_type, safeCell(doc.aju_number), safeCell(doc.registration_number), safeCell(item.item_sequence), safeCell(item.item_code), safeCell(item.description), safeCell(item.quantity,'number'), safeCell(item.unit), doc.business_status]);
    }
    addSplitSheets(workbook, 'TRANSACTION_REPORT', rows);
  }
  if (preset === 'SOURCE_REPORT') {
    const rows = [['Source ID','Nama File','Tipe Sumber','SHA-256','Ukuran Byte','Status Impor','Jumlah Dokumen']];
    for (const source of project.sources.source_registry) rows.push([source.source_id, safeCell(source.original_file_name), source.source_type, source.sha256, source.file_size || 0, source.import_status, (source.document_ids || []).length]);
    addSplitSheets(workbook, 'SOURCE_REPORT', rows);
  }
  if (preset === 'CORRECTION_REPORT') {
    const rows = [['Correction ID','Jenis Record','Record ID','Document ID','Field','Nilai Lama','Nilai Baru','Alasan','Waktu','Operator']];
    for (const row of project.data.corrections) rows.push([row.correction_id,row.entity_type,row.entity_id,row.document_id,row.field_path,safeCell(row.old_value),safeCell(row.new_value),safeCell(row.reason),row.changed_at,row.operator_id]);
    addSplitSheets(workbook, 'CORRECTION_REPORT', rows);
  }
  if (preset === 'AUDIT_REPORT') {
    const rows = [['Event ID','Jenis Aktivitas','Deskripsi','Waktu','Operator']];
    for (const row of project.data.audit_log) rows.push([row.event_id,row.event_type,safeCell(row.description),row.created_at,row.operator_id]);
    addSplitSheets(workbook, 'AUDIT_REPORT', rows);
  }
  if (preset === 'CUSTOM') {
    const entityType = custom?.entity_type === 'ITEM' ? 'ITEM' : 'DOCUMENT';
    const available = entityType === 'ITEM' ? itemFields : docFields;
    const fields = available.filter(field => (custom?.columns || []).includes(field.key));
    if (!fields.length) throw new Error('Pilih minimal satu kolom untuk Custom Export.');
    const records = entityType === 'ITEM' ? items : docs;
    addData(entityType === 'ITEM' ? 'CUSTOM_ITEMS' : 'CUSTOM_DOCUMENTS', records, fields, entityType, custom || {});
  }

  const readme = [
    ['BCP EXPORT'],
    ['Project', project.manifest.project_name],
    ['Company', project.manifest.company_name || ''],
    ['Preset', preset],
    ['Data Version', dataVersion],
    ['Exported At', new Date().toISOString()],
    ['Project ID', project.manifest.project_id],
    ['Filter Business Status', filters.business_status || (filters.include_inactive ? 'ALL' : 'CONFIRMED')],
    ['Record counts', `Documents: ${docs.length}; Items: ${items.length}`]
  ];
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(readme), 'README');
  const validation = [['Kategori','Jumlah'],['Documents',project.data.documents.length],['Items',project.data.items.length],['Sources',project.sources.source_registry.length],['Corrections',project.data.corrections.length],['Audit Events',project.data.audit_log.length]];
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(validation), 'VALIDATION_SUMMARY');
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([['Project ID','Preset','Data Version','Exported At'],[project.manifest.project_id,preset,dataVersion,new Date().toISOString()]]), 'EXPORT_METADATA');
  return workbook;
}

export function recordExportHistory(project, metadata) {
  const next = structuredClone(project);
  if (!Array.isArray(next.data.export_history)) next.data.export_history = [];
  const row = {
    export_id: `EXPORT-${crypto.randomUUID()}`,
    preset: metadata.preset,
    data_version: metadata.data_version,
    file_name: metadata.file_name,
    record_count: metadata.record_count || 0,
    exported_at: new Date().toISOString(),
    operator_id: next.settings?.operator?.operator_id || 'LOCAL-USER-001'
  };
  next.data.export_history.push(row);
  next.data.audit_log.push({
    event_id: `AUDIT-${crypto.randomUUID()}`,
    event_type: 'EXPORT_CREATED',
    description: `Export ${row.preset} dibuat: ${row.file_name}.`,
    created_at: row.exported_at,
    operator_id: row.operator_id
  });
  return refreshProjectCounts(next);
}

export function downloadExportWorkbook(workbook, fileName) {
  const bytes = XLSX.write(workbook, { bookType: 'xlsx', type: 'array', cellDates: true });
  const blob = new Blob([bytes], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement('a');
  anchor.href = url;
  anchor.download = fileName;
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
  window.setTimeout(() => URL.revokeObjectURL(url), 1500);
  return bytes.byteLength;
}

export function getExportFields(entityType = 'DOCUMENT') {
  return (entityType === 'ITEM' ? ITEM_FIELDS : DOC_FIELDS).map(field => ({ ...field }));
}

export function saveExportProfile(project, { name, entity_type = 'DOCUMENT', columns = [], filters = {}, group_by = '', sort_by = '', sort_direction = 'ASC', label_overrides = {} }) {
  const label = String(name || '').trim();
  if (!label) throw new Error('Nama profil export wajib diisi.');
  const fields = getExportFields(entity_type);
  const selected = [...new Set(columns)].filter(key => fields.some(field => field.key === key));
  if (!selected.length) throw new Error('Pilih minimal satu kolom profil export.');
  const next = structuredClone(project);
  if (!Array.isArray(next.settings.export.profiles)) next.settings.export.profiles = [];
  const profileId = `EXPORT-PROFILE-${crypto.randomUUID()}`;
  next.settings.export.profiles.push({
    profile_id: profileId,
    name: label,
    entity_type,
    columns: selected,
    filters: structuredClone(filters),
    group_by: selected.includes(group_by) ? group_by : '',
    sort_by: selected.includes(sort_by) ? sort_by : '',
    sort_direction: sort_direction === 'DESC' ? 'DESC' : 'ASC',
    label_overrides: Object.fromEntries(Object.entries(label_overrides).filter(([key]) => selected.includes(key))),
    created_at: new Date().toISOString(),
    updated_at: new Date().toISOString()
  });
  next.data.audit_log.push({
    event_id: `AUDIT-${crypto.randomUUID()}`,
    event_type: 'EXPORT_PROFILE_SAVED',
    description: `Profil Custom Export "${label}" disimpan.`,
    created_at: new Date().toISOString(),
    operator_id: next.settings?.operator?.operator_id || 'LOCAL-USER-001'
  });
  return next;
}
