export const BCP_FORMAT = 'BCP_PROJECT';
export const BCP_FORMAT_VERSION = '1.0.0';
export const BCP_SCHEMA_VERSION = '1.0.0';
export const BCP_APP_VERSION = '0.1.0';

export function createEmptyProject({ projectName, companyName = '' }) {
  const now = new Date().toISOString();
  const projectId = crypto.randomUUID();
  return {
    manifest: {
      format: BCP_FORMAT,
      format_version: BCP_FORMAT_VERSION,
      schema_version: BCP_SCHEMA_VERSION,
      app_version: BCP_APP_VERSION,
      project_id: projectId,
      project_name: projectName.trim(),
      company_name: companyName.trim(),
      created_at: now,
      updated_at: now,
      last_saved_at: null,
      project_status: 'OPEN',
      counts: { documents: 0, items: 0, source_files: 0 }
    },
    data: {
      documents: [],
      items: [],
      periods: [],
      corrections: [],
      validation: [],
      audit_log: []
    },
    sources: { source_registry: [] },
    settings: {
      operator: { operator_id: 'LOCAL-USER-001', display_name: 'Default Operator' },
      project: { reminder_minutes: 5, reminder_change_count: 10 },
      export: { default_data_version: 'CURRENT' }
    }
  };
}

export function validateProject(project) {
  const errors = [];
  if (!project || typeof project !== 'object') return ['Struktur proyek tidak valid.'];
  const m = project.manifest;
  if (!m || m.format !== BCP_FORMAT) errors.push('Format proyek bukan BCP_PROJECT.');
  if (!m?.project_id || !m?.project_name) errors.push('Identitas proyek tidak lengkap.');
  if (!project.data || !Array.isArray(project.data.documents) || !Array.isArray(project.data.items)) {
    errors.push('Data documents/items tidak valid.');
  }
  for (const key of ['periods', 'corrections', 'validation', 'audit_log']) {
    if (!Array.isArray(project.data?.[key])) errors.push(`Data ${key} harus berupa array.`);
  }
  if (!Array.isArray(project.sources?.source_registry)) errors.push('Source registry tidak valid.');
  if (!project.settings || typeof project.settings !== 'object') errors.push('Pengaturan proyek tidak valid.');
  return errors;
}

export function refreshProjectCounts(project) {
  const next = structuredClone(project);
  next.manifest.counts = {
    documents: next.data.documents.length,
    items: next.data.items.length,
    source_files: next.sources.source_registry.length
  };
  next.manifest.updated_at = new Date().toISOString();
  return next;
}
