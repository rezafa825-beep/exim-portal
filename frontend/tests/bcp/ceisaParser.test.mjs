import test from 'node:test';
import assert from 'node:assert/strict';
import { File } from 'node:buffer';
import * as XLSX from 'xlsx';
import { parseBC27File } from '../../src/modules/bcp/ceisaParser.js';
import { createEmptyProject } from '../../src/modules/bcp/projectSchema.js';
import { commitPartialImport, evaluateStagedImport } from '../../src/modules/bcp/importEngine.js';
import { getDocumentView, validateRegistryIntegrity } from '../../src/modules/bcp/registryEngine.js';
import { buildProjectArchive, readProjectArchive } from '../../src/modules/bcp/projectPackage.js';

const COMPANY = 'PT Chinli Plastic Materials Indonesia';

function makeWorkbook({ documentCount = 1, direction = 'IN' } = {}) {
  const workbook = XLSX.utils.book_new();
  const ajuNumbers = Array.from({ length: documentCount }, (_, i) =>
    `000027CHD8572026112700${String(1942 + i).padStart(4, '0')}`
  );

  const headers = [
    ['HEADER', ['NOMOR AJU','KODE DOKUMEN','KODE KANTOR','KODE KANTOR TUJUAN','NOMOR DAFTAR','TANGGAL DAFTAR','TANGGAL PERNYATAAN','KODE VALUTA','NDPBM','CIF','HARGA_PENYERAHAN','BRUTO','NETTO','VOLUME']],
    ['ENTITAS', ['NOMOR AJU','SERI','KODE ENTITAS','NAMA ENTITAS','NOMOR IDENTITAS']],
    ['BARANG', ['NOMOR AJU','SERI BARANG','HS','KODE BARANG','URAIAN','SPESIFIKASI LAIN','KODE SATUAN','JUMLAH SATUAN','KODE KEMASAN','JUMLAH KEMASAN','NETTO','BRUTO','CIF','HARGA PENYERAHAN','DUPLIKAT','DUPLIKAT']],
    ['BAHANBAKU', ['NOMOR AJU','SERI BARANG','SERI BAHAN BAKU','KODE ASAL BAHAN BAKU','HS','KODE BARANG','URAIAN','KODE SATUAN','JUMLAH SATUAN','CIF','CIF RUPIAH','NDPBM','HARGA PENYERAHAN','HARGA PEROLEHAN']],
    ['BAHANBAKUTARIF', ['NOMOR AJU','SERI BARANG','SERI BAHAN BAKU','KODE ASAL BAHAN BAKU','KODE PUNGUTAN','KODE TARIF','TARIF','KODE FASILITAS','TARIF FASILITAS','NILAI BAYAR','NILAI FASILITAS']],
    ['BAHANBAKUDOKUMEN', ['NOMOR AJU','SERI BARANG','SERI BAHAN BAKU','KODE DOKUMEN','NOMOR DOKUMEN','TANGGAL DOKUMEN']],
    ['KEMASAN', ['NOMOR AJU','SERI','KODE KEMASAN','JUMLAH KEMASAN','MEREK']],
    ['DOKUMEN', ['NOMOR AJU','SERI','KODE DOKUMEN','NOMOR DOKUMEN','TANGGAL DOKUMEN']],
    ['PENGANGKUT', ['NOMOR AJU','SERI','NAMA PENGANGKUT']],
    ['PUNGUTAN', ['NOMOR AJU','KODE FASILITAS TARIF','KODE JENIS PUNGUTAN','NILAI PUNGUTAN']]
  ];

  const rows = Object.fromEntries(headers.map(([name, columns]) => [name, [columns]]));
  for (let i = 0; i < documentCount; i++) {
    const aju = ajuNumbers[i];
    const reg = String(27413 + i).padStart(6, '0');
    rows.HEADER.push([aju,'27','050700','150300',reg,'2026-11-27','2026-11-27','USD',18263,912.24,16660239.12,117.93,114.03,0]);
    const companyCode = direction === 'IN' ? '8' : '3';
    const partnerCode = direction === 'IN' ? '3' : '8';
    const companyName = 'CHINLI PLASTIC MATERIALS INDONESIA';
    const partnerName = direction === 'IN' ? 'SHOETOWN LIGUNG INDONESIA' : 'CHINLI PLASTIC TECHNOLOGY INDONESIA';
    rows.ENTITAS.push([aju,'1',companyCode,companyName,'0929190064455000000000']);
    rows.ENTITAS.push([aju,'2',partnerCode,partnerName,'0929190064455000000001']);
    rows.ENTITAS.push([aju,'3',direction === 'IN' ? '7' : '7',partnerName,'0929190064455000000001']);
    rows.BARANG.push([aju,1,'64069049','PSO20148','INSOLE TEST ITEM','SPEC TEST','NPR',100,'CT',10,12.5,13.5,100,1800000,'DUP-A','DUP-B']);
    for (let m = 1; m <= 2; m++) {
      rows.BAHANBAKU.push([aju,1,m,'0',m === 1 ? '39211399' : '48169040',`RM-${m}`,`RAW MATERIAL ${m}`,'ST',m * 10,100 + m,1826300 + m,18263,1800000 + m,0]);
      rows.BAHANBAKUDOKUMEN.push([aju,1,m,'380',`SUP-${m}`,'2026-11-20']);
      for (const code of ['BM','PPN','PPH']) {
        rows.BAHANBAKUTARIF.push([aju,1,m,'0',code,'1',code === 'BM' ? 5 : code === 'PPN' ? 11 : 2.5,'3',100,0,200]);
      }
    }
    rows.KEMASAN.push([aju,1,'BG',10,'']);
    rows.DOKUMEN.push([aju,1,'380',`DOC-${i+1}`,'2026-11-27']);
    rows.PENGANGKUT.push([aju,1,'MOBIL']);
    rows.PUNGUTAN.push([aju,'3','PPN',1000]);
  }

  for (const [name, matrix] of Object.entries(rows)) {
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(matrix), name);
  }
  for (const [name, columns] of [['KONTAINER',['NOMOR AJU','SERI','NOMOR KONTAINER']],['BARANGTARIF',['NOMOR AJU','SERI BARANG','KODE PUNGUTAN']]]) {
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([columns]), name);
  }
  return workbook;
}

function asFile(workbook, name = 'ceisa-test.xlsx') {
  const bytes = XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer' });
  return new File([bytes], name, { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
}

test('BC 2.7 IN single-document parses document, item, raw materials and tariffs', async () => {
  const parsed = await parseBC27File(asFile(makeWorkbook({ documentCount: 1, direction: 'IN' })), COMPANY);
  assert.equal(parsed.file_format, 'SINGLE_DOCUMENT');
  assert.equal(parsed.documents.length, 1);
  assert.equal(parsed.documents[0].document.document_type, 'BC27_IN');
  assert.equal(parsed.documents[0].document.aju_number, '000027CHD85720261127001942');
  assert.equal(parsed.documents[0].items.length, 1);
  assert.equal(parsed.documents[0].raw_materials.length, 2);
  assert.equal(parsed.documents[0].raw_material_tariffs.length, 6);
  assert.equal(parsed.documents[0].raw_materials[0].extension.related_documents.length, 1);
  assert.equal(parsed.documents[0].items[0].extension.raw_fields.filter(field => field.name === 'DUPLIKAT').length, 2);
  assert.equal(parsed.documents[0].items[0].source_locations[0].sheet, 'BARANG');
  assert.equal(parsed.summary.error_count, 0);
});

test('BC 2.7 OUT single-document is classified by configured company entity role', async () => {
  const parsed = await parseBC27File(asFile(makeWorkbook({ documentCount: 1, direction: 'OUT' })), COMPANY);
  assert.equal(parsed.documents[0].document.document_type, 'BC27_OUT');
  assert.equal(parsed.documents[0].document.extension.classification.status, 'CLASSIFIED');
});

test('one workbook containing multiple Nomor Aju is split into separate documents', async () => {
  const parsed = await parseBC27File(asFile(makeWorkbook({ documentCount: 2, direction: 'IN' }), 'weekly.xlsx'), COMPANY);
  assert.equal(parsed.file_format, 'MULTI_DOCUMENT');
  assert.equal(parsed.documents.length, 2);
  assert.equal(parsed.summary.item_count, 2);
  assert.notEqual(parsed.documents[0].document.aju_number, parsed.documents[1].document.aju_number);
});

test('Partial Import commits valid unique documents and includes original source file', async () => {
  const parsed = await parseBC27File(asFile(makeWorkbook({ documentCount: 1, direction: 'IN' })), COMPANY);
  const project = createEmptyProject({ projectName: 'TEST BCP', companyName: COMPANY });
  const evaluation = evaluateStagedImport(project, [parsed]);
  const ready = evaluation[0].documents.filter(row => row.status === 'READY').map(row => row.document_id);
  const result = commitPartialImport(project, [parsed], ready);
  assert.equal(result.summary.imported_documents, 1);
  assert.equal(result.project.data.documents.length, 1);
  assert.equal(result.project.data.items.length, 1);
  assert.equal(result.project.data.raw_materials.length, 2);
  assert.equal(result.project.data.raw_material_tariffs.length, 6);
  assert.equal(result.project.sources.source_registry.length, 1);
  assert.ok(result.project.sources.files[parsed.source.internal_path] instanceof Uint8Array);
  assert.deepEqual(validateRegistryIntegrity(result.project), []);
  const detail = getDocumentView(result.project, result.project.data.documents[0].document_id);
  assert.equal(detail.items[0].raw_materials.length, 2);
  assert.equal(detail.items[0].raw_materials[0].tariffs.length, 3);
});

test('BCP project package round-trips original source bytes with checksum validation', async () => {
  const project = createEmptyProject({ projectName: 'PACKAGE TEST', companyName: COMPANY });
  const sourcePath = 'sources/ceisa/SRC-test_test.xlsx';
  project.sources.source_registry.push({
    source_id: 'SRC-test', source_type: 'CEISA', original_file_name: 'test.xlsx',
    internal_path: sourcePath, sha256: 'source-hash', file_size: 3,
    document_ids: [], import_status: 'IMPORTED'
  });
  project.sources.files[sourcePath] = new Uint8Array([1, 2, 3]);
  const archive = await buildProjectArchive(project);
  const restored = await readProjectArchive(new File([archive.bytes], 'package-test.bcp'));
  assert.equal(restored.manifest.project_id, project.manifest.project_id);
  assert.deepEqual([...restored.sources.files[sourcePath]], [1, 2, 3]);
});
