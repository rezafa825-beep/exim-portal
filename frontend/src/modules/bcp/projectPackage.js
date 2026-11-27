import { strToU8, strFromU8, zipSync, unzipSync } from 'fflate';
import { BCP_FORMAT, BCP_FORMAT_VERSION, validateProject, refreshProjectCounts } from './projectSchema';

const MAX_PROJECT_FILE_BYTES = 250 * 1024 * 1024;
const MAX_UNCOMPRESSED_BYTES = 750 * 1024 * 1024;
const MAX_ARCHIVE_ENTRIES = 10000;

function jsonBytes(value) {
  return strToU8(JSON.stringify(value, null, 2));
}

async function sha256Hex(bytes) {
  if (!globalThis.crypto?.subtle) {
    throw new Error('Web Crypto API tidak tersedia. Gunakan browser modern melalui HTTPS atau localhost.');
  }
  const digest = await globalThis.crypto.subtle.digest('SHA-256', bytes);
  return Array.from(new Uint8Array(digest), value => value.toString(16).padStart(2, '0')).join('');
}

function safeInternalPath(path) {
  return typeof path === 'string' &&
    path.startsWith('sources/') &&
    !path.startsWith('/') &&
    !path.includes('\\') &&
    !path.split('/').some(part => part === '..' || part === '.');
}

export async function buildProjectArchive(project) {
  const errors = validateProject(project);
  if (errors.length) throw new Error(errors.join(' '));

  const normalized = refreshProjectCounts(project);
  const manifest = {
    ...normalized.manifest,
    format: BCP_FORMAT,
    format_version: BCP_FORMAT_VERSION,
    last_saved_at: new Date().toISOString()
  };

  const entries = {
    'data/documents.json': jsonBytes(normalized.data.documents),
    'data/items.json': jsonBytes(normalized.data.items),
    'data/raw_materials.json': jsonBytes(normalized.data.raw_materials),
    'data/raw_material_tariffs.json': jsonBytes(normalized.data.raw_material_tariffs),
    'data/periods.json': jsonBytes(normalized.data.periods),
    'data/corrections.json': jsonBytes(normalized.data.corrections),
    'data/validation.json': jsonBytes(normalized.data.validation),
    'data/relationships.json': jsonBytes(normalized.data.relationships),
    'data/audit_log.json': jsonBytes(normalized.data.audit_log),
    'sources/source_registry.json': jsonBytes(normalized.sources.source_registry),
    'settings/project_settings.json': jsonBytes(normalized.settings)
  };

  for (const [path, bytes] of Object.entries(normalized.sources.files || {})) {
    if (!safeInternalPath(path)) throw new Error('Lokasi file sumber di dalam proyek tidak valid.');
    entries[path] = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
  }

  const checksums = {};
  for (const [path, bytes] of Object.entries(entries)) {
    checksums[path] = await sha256Hex(bytes);
  }
  manifest.checksums = checksums;
  entries['manifest.json'] = jsonBytes(manifest);

  return { bytes: zipSync(entries, { level: 6 }), manifest };
}

function readJson(entries, path, fallback) {
  const bytes = entries[path];
  if (!bytes) return fallback;
  try {
    return JSON.parse(strFromU8(bytes));
  } catch {
    throw new Error(`File internal ${path} bukan JSON yang valid.`);
  }
}

function assertSafeArchive(entries) {
  const names = Object.keys(entries);
  if (names.length > MAX_ARCHIVE_ENTRIES) throw new Error('Proyek berisi terlalu banyak file internal.');
  let total = 0;
  for (const name of names) {
    if (name.startsWith('/') || name.includes('\\') || name.split('/').some(part => part === '..' || part === '.')) {
      throw new Error('Arsip berisi jalur file yang tidak aman.');
    }
    total += entries[name]?.byteLength || 0;
    if (total > MAX_UNCOMPRESSED_BYTES) throw new Error('Ukuran hasil ekstraksi proyek melewati batas keamanan 750 MB.');
  }
}

export async function readProjectArchive(file) {
  if (!file || !/\.bcp$/i.test(file.name)) throw new Error('Pilih file proyek dengan ekstensi .bcp.');
  if (file.size > MAX_PROJECT_FILE_BYTES) throw new Error('Ukuran file proyek melebihi batas 250 MB.');

  const input = new Uint8Array(await file.arrayBuffer());
  let entries;
  try { entries = unzipSync(input); }
  catch { throw new Error('File .bcp rusak atau bukan arsip proyek yang valid.'); }

  assertSafeArchive(entries);
  const manifest = readJson(entries, 'manifest.json', null);
  if (!manifest || manifest.format !== BCP_FORMAT) throw new Error('Manifest tidak ditemukan atau format proyek tidak dikenali.');
  if (manifest.format_version !== BCP_FORMAT_VERSION) {
    throw new Error(`Versi format proyek ${manifest.format_version || 'tidak diketahui'} belum didukung oleh aplikasi ini.`);
  }
  if (!manifest.checksums || typeof manifest.checksums !== 'object') throw new Error('Manifest tidak memiliki checksum integritas.');

  for (const [path, expected] of Object.entries(manifest.checksums)) {
    if (!entries[path]) throw new Error('File internal hilang: ' + path);
    const actual = await sha256Hex(entries[path]);
    if (actual !== expected) throw new Error('Pemeriksaan integritas gagal pada ' + path + '. File proyek mungkin berubah atau rusak.');
  }

  const sourceRegistry = readJson(entries, 'sources/source_registry.json', []);
  const sourceFiles = {};
  for (const source of sourceRegistry) {
    if (!source.internal_path || !safeInternalPath(source.internal_path)) {
      throw new Error('Lokasi file sumber tidak valid di Source Registry.');
    }
    if (!entries[source.internal_path]) {
      throw new Error('File sumber asli tidak ditemukan di dalam proyek: ' + source.original_file_name);
    }
    if (!manifest.checksums[source.internal_path]) {
      throw new Error('Checksum file sumber tidak ditemukan: ' + source.original_file_name);
    }
    sourceFiles[source.internal_path] = entries[source.internal_path];
  }

  const project = {
    manifest,
    data: {
      documents: readJson(entries, 'data/documents.json', []),
      items: readJson(entries, 'data/items.json', []),
      raw_materials: readJson(entries, 'data/raw_materials.json', []),
      raw_material_tariffs: readJson(entries, 'data/raw_material_tariffs.json', []),
      periods: readJson(entries, 'data/periods.json', []),
      corrections: readJson(entries, 'data/corrections.json', []),
      validation: readJson(entries, 'data/validation.json', []),
      relationships: readJson(entries, 'data/relationships.json', []),
      audit_log: readJson(entries, 'data/audit_log.json', [])
    },
    sources: { source_registry: sourceRegistry, files: sourceFiles },
    settings: readJson(entries, 'settings/project_settings.json', {})
  };

  const errors = validateProject(project);
  if (errors.length) throw new Error(errors.join(' '));
  return project;
}

export async function downloadProjectArchive(project) {
  const { bytes, manifest } = await buildProjectArchive(project);
  const blob = new Blob([bytes], { type: 'application/vnd.reza.bcp+zip' });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement('a');
  const safeName = (manifest.project_name || 'BCP_PROJECT').replace(/[^a-z0-9_-]+/gi, '_');
  anchor.href = url;
  anchor.download = `${safeName}.bcp`;
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
  window.setTimeout(() => URL.revokeObjectURL(url), 1500);
  return manifest;
}
