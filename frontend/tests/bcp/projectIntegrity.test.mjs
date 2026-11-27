import test from 'node:test';
import assert from 'node:assert/strict';
import { File } from 'node:buffer';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
import { createEmptyProject } from '../../src/modules/bcp/projectSchema.js';
import { buildProjectArchive, readProjectArchive } from '../../src/modules/bcp/projectPackage.js';
import { validateRegistryIntegrity } from '../../src/modules/bcp/registryEngine.js';

test('registry integrity rejects orphaned items, tariffs, and period references', () => {
  const project = createEmptyProject({ projectName: 'Integrity' });
  project.data.documents.push({ document_id: 'DOC-1', period_id: 'PERIOD-MISSING', source_refs: [] });
  project.data.items.push({ item_id: 'ITEM-1', document_id: 'DOC-MISSING', source_refs: [] });
  project.data.raw_materials.push({ raw_material_id: 'BB-1', document_id: 'DOC-1', item_id: 'ITEM-MISSING', source_refs: [] });
  project.data.raw_material_tariffs.push({ raw_material_tariff_id: 'BBT-1', document_id: 'DOC-1', item_id: 'ITEM-MISSING', raw_material_id: 'BB-MISSING' });
  const errors = validateRegistryIntegrity(project);
  assert.ok(errors.some(message => message.includes('tidak memiliki dokumen induk')));
  assert.ok(errors.some(message => message.includes('periode yang tidak tersedia')));
  assert.ok(errors.some(message => message.includes('tidak memiliki item induk')));
  assert.ok(errors.some(message => message.includes('tidak memiliki bahan baku induk')));
});

test('project archive refuses an orphaned registry record before saving', async () => {
  const project = createEmptyProject({ projectName: 'Invalid Project' });
  project.data.items.push({ item_id: 'ITEM-ORPHAN', document_id: 'DOC-MISSING' });
  await assert.rejects(() => buildProjectArchive(project), /tidak memiliki dokumen induk/);
});

test('project archive rejects missing required data entries even if manifest checksums are edited', async () => {
  const project = createEmptyProject({ projectName: 'Missing Entry' });
  const archive = await buildProjectArchive(project);
  const entries = unzipSync(archive.bytes);
  delete entries['data/items.json'];
  const manifest = JSON.parse(strFromU8(entries['manifest.json']));
  delete manifest.checksums['data/items.json'];
  entries['manifest.json'] = strToU8(JSON.stringify(manifest));
  const tampered = zipSync(entries);
  await assert.rejects(
    () => readProjectArchive(new File([tampered], 'missing-entry.bcp')),
    /File internal wajib tidak ditemukan: data\/items\.json/
  );
});
