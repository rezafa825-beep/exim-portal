/* EXIM WhatsApp Report Generator - multi Excel + pengecualian nomor daftar + label pihak manual */
const $ = (id) => document.getElementById(id);

const DOC_LABELS = {
  "27": "BC 2.7", "40": "BC 4.0", "41": "BC 4.1", "261": "BC 2.6.1",
  "262": "BC 2.6.2", "25": "BC 2.5", "23": "BC 2.3"
};
const PURPOSE_MASTER = {
  "27": {"1":"PENYERAHAN BKP","2":"PENYERAHAN JKP","3":"RETUR","4":"NON PENYERAHAN","5":"LAINNYA"},
  "40": {"1":"PENYERAHAN BKP","2":"PENYERAHAN JKP","3":"RETUR","4":"NON PENYERAHAN","5":"LAINNYA","6":"PENYERAHAN BKP DENGAN FP101"},
  "41": {"1":"PENYERAHAN BKP","2":"PENYERAHAN JKP","3":"RETUR","4":"NON PENYERAHAN","5":"LAINNYA"},
  "261":{"1":"DIPERBAIKI","2":"DISUBKONTRAKKAN","3":"DIPINJAMKAN","4":"BARANG RETURN","5":"LAINNYA"},
  "262":{"1":"EKS-DIPERBAIKI","2":"EKS-DISUBKONTRAKKAN","3":"EKS-DIPINJAMKAN","4":"RETURN BAHAN BAKU","5":"PEMASUKAN MELEWATI JANGKA WAKTU","6":"LAINNYA"},
  "25": {"1":"PENYERAHAN BKP","2":"PENYERAHAN JKP","3":"RETUR","4":"NON PENYERAHAN","5":"LAINNYA"},
  "23": {"1":"PEMBELIAN BKP","2":"PENERIMA JASA BKP"}
};
const DOC_TYPES = [
  {value:"27-out", label:"BC 2.7 Keluar", code:"27", direction:"out"},
  {value:"27-in", label:"BC 2.7 Masuk", code:"27", direction:"in"},
  {value:"40-in", label:"BC 4.0 Masuk", code:"40", direction:"in"},
  {value:"41-out", label:"BC 4.1 Keluar", code:"41", direction:"out"},
  {value:"261-out", label:"BC 2.6.1", code:"261", direction:"out"},
  {value:"262-in", label:"BC 2.6.2", code:"262", direction:"in"},
  {value:"25-out", label:"BC 2.5", code:"25", direction:"out"},
  {value:"23-in", label:"BC 2.3", code:"23", direction:"in"}
];
let records = [];
let uploadedFiles = new Map();
let fileSequence = 0;
let redNumbers = new Set();
let excludedNumbers = new Set();
let selectedExportIds = new Set();
function norm(v) { return String(v ?? "").trim(); }
function key(v) { return norm(v).replace(/\.0$/, ""); }
function excelDate(v) {
  if (v instanceof Date && !isNaN(v)) return v;
  if (typeof v === "number" && v > 20000 && v < 80000) {
    const p = XLSX.SSF.parse_date_code(v);
    if (p) return new Date(p.y, p.m - 1, p.d);
  }
  const s = norm(v);
  if (!s) return null;
  const d = new Date(s);
  return isNaN(d) ? null : d;
}
function dateText(v) {
  const d = v instanceof Date ? v : excelDate(v);
  if (!d) return "";
  return `${String(d.getDate()).padStart(2,"0")}/${String(d.getMonth()+1).padStart(2,"0")}/${d.getFullYear()}`;
}
function dateInputToday() {
  const d = new Date();
  return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}-${String(d.getDate()).padStart(2,"0")}`;
}
function codeLabel(docCode, purposeCode) {
  return PURPOSE_MASTER[docCode]?.[key(purposeCode)] || (key(purposeCode) ? `KODE ${key(purposeCode)} (PERLU DIPERIKSA)` : "TUJUAN TIDAK TERISI");
}
function partyDefault(direction) { return direction === "out" ? "customer" : "supplier"; }
function docTypeFor(header, entitas) {
  const code = key(header[1]);
  if (code === "27") {
    const code3 = entitas.find(e => key(e[2]) === "3");
    const name = norm(code3?.[5]).toUpperCase();
    return name === "CHINLI PLASTIC MATERIALS INDONESIA" ? "27-out" : "27-in";
  }
  const mapping = {"40":"40-in","41":"41-out","261":"261-out","262":"262-in","25":"25-out","23":"23-in"};
  return mapping[code] || `unknown-${code}`;
}
function headerDateLabel(typeValue) {
  const t = DOC_TYPES.find(x => x.value === typeValue);
  return t?.direction === "in" ? "Tanggal Masuk" : "Tanggal Keluar";
}
function expectedPartyMode(typeValue) {
  const t = DOC_TYPES.find(x => x.value === typeValue);
  return partyDefault(t?.direction || "out");
}
function uniqueSortedDates(items) {
  const map = new Map();
  items.forEach(d => { if (d) map.set(dateText(d), d); });
  return [...map.keys()].sort((a,b) => {
    const [da,ma,ya]=a.split("/").map(Number), [db,mb,yb]=b.split("/").map(Number);
    return new Date(ya,ma-1,da)-new Date(yb,mb-1,db);
  });
}
// Urutkan nomor daftar dan nomor segel secara natural: angka 2 sebelum 10,
// huruf A-Z tanpa membedakan kapitalisasi.
function naturalSort(a, b) {
  return String(a).localeCompare(String(b), "id", {
    numeric: true,
    sensitivity: "base"
  });
}
function fmtQty(n) { return new Intl.NumberFormat("id-ID", {maximumFractionDigits: 3}).format(n); }
function readSheet(wb, name) {
  const ws = wb.Sheets[name];
  if (!ws) throw new Error(`Sheet ${name} tidak ditemukan.`);
  return XLSX.utils.sheet_to_json(ws, {header:1, defval:null, raw:true});
}
function parseWorkbook(wb) {
  const headerRows = readSheet(wb,"HEADER");
  const entityRows = readSheet(wb,"ENTITAS");
  const packageRows = readSheet(wb,"KEMASAN");
  const goodsRows = readSheet(wb,"BARANG");
  const entByAju = new Map(), packByAju = new Map(), goodsByAju = new Map();
  for (const r of entityRows.slice(1)) {
    const aju = key(r[0]); if (!aju) continue;
    if (!entByAju.has(aju)) entByAju.set(aju,[]);
    entByAju.get(aju).push(r);
  }
  for (const r of packageRows.slice(1)) {
    const aju = key(r[0]); if (!aju) continue;
    if (!packByAju.has(aju)) packByAju.set(aju,[]);
    packByAju.get(aju).push(r);
  }
  for (const r of goodsRows.slice(1)) {
    const aju = key(r[0]); if (!aju) continue;
    if (!goodsByAju.has(aju)) goodsByAju.set(aju,[]);
    goodsByAju.get(aju).push(r);
  }
  const result = [];
  for (const h of headerRows.slice(1)) {
    const aju = key(h[0]); if (!aju) continue;
    const entities = entByAju.get(aju) || [];
    const typeValue = docTypeFor(h, entities);
    const type = DOC_TYPES.find(x => x.value === typeValue);
    const docCode = key(h[1]);
    // BC 2.7 memakai posisi entitas yang berbeda menurut arah dokumen:
    // Keluar: Chinli = kode 3, customer = kode 8.
    // Masuk: customer/supplier eksternal = kode 3, Chinli = kode 8.
    // Untuk laporan, nama pihak eksternal BC 2.7 Masuk diambil dari kode 3.
    const isBc27 = docCode === "27";
    const customerEntity = entities.find(e => key(e[2]) === "8");
    // Pengecualian BC 2.3: supplier menggunakan Kode Entitas 5 (kolom C),
    // dan nama diambil dari kolom F pada baris ENTITAS yang sama.
    const supplierEntity = docCode === "23"
      ? entities.find(e => key(e[2]) === "5")
      : isBc27 && type?.direction === "in"
        ? entities.find(e => key(e[2]) === "3")
        : entities.find(e => key(e[2]) === "9");
    const defaultPartyCode = docCode === "23"
      ? "5"
      : isBc27
        ? (type?.direction === "out" ? "8" : "3")
        : (type ? (type.direction === "out" ? "8" : "9") : "");
    const partyEntity = entities.find(e => key(e[2]) === defaultPartyCode);
    const regNo = key(h[93]); // HEADER kolom CP = NOMOR DAFTAR (0-based index 93)
    const regDate = h[94];    // HEADER kolom CQ = TANGGAL DAFTAR (0-based index 94)
    // Sumber Jenis Tujuan Transaksi mengikuti field pada HEADER:
    // BC 2.6.2: kolom M (KODE TUJUAN PEMASUKAN).
    // Dokumen lainnya: kolom N (KODE TUJUAN PENGIRIMAN).
    // Pengecualian BC 2.3: BARANG kolom AZ (KODE PERHITUNGAN),
    // ambil kode pertama yang terisi pada barang dengan nomor aju yang sama.
    const goodsForAju = goodsByAju.get(aju) || [];
    const purposeCode = docCode === "23"
      ? (goodsForAju.find(g => key(g[51]) !== "")?.[51] ?? "")
      : docCode === "262"
        ? h[12]
        : h[13];
    const packages = new Map(), seals = [];
    for (const p of (packByAju.get(aju) || [])) {
      const pkgCode = norm(p[2]), qty = Number(p[3] || 0), seal = norm(p[5]);
      if (pkgCode) packages.set(pkgCode,(packages.get(pkgCode)||0)+qty);
      if (seal && !seals.includes(seal)) seals.push(seal);
    }
    const quantities = new Map(), descriptions = [];
    for (const g of goodsForAju) {
      const unit = norm(g[9]), qty = Number(g[10] || 0), desc = norm(g[4]);
      if (unit && Number.isFinite(qty) && qty !== 0) quantities.set(unit,(quantities.get(unit)||0)+qty);
      if (desc && !descriptions.includes(desc)) descriptions.push(desc);
    }
    result.push({
      aju, typeValue, typeLabel: type?.label || `Kode dokumen ${docCode}`,
      direction: type?.direction || "out", docCode, regNo: regNo || "(Nomor daftar kosong)",
      regDate, purposeCode: key(purposeCode), purpose: codeLabel(docCode,purposeCode),
      partyCode: defaultPartyCode,
      partyName: norm(partyEntity?.[5]) || "(Nama pihak tidak ditemukan)",
      customerName: norm(customerEntity?.[5]) || "",
      supplierName: norm(supplierEntity?.[5]) || "",
      packages, seals, quantities, descriptions
    });
  }
  return result;
}
function selectedTypeRecords() {
  const type = $("docType").value;
  return records.filter(r => r.typeValue === type);
}
// Mode sumber data tetap mengikuti pemetaan otomatis sistem.
// Dropdown hanya mengubah label yang ditampilkan pada laporan, bukan sumber nama pihak.
function currentPartyMode() {
  return expectedPartyMode($("docType").value);
}
function currentPartyLabelMode() {
  const selected = $("partyMode").value;
  return selected === "auto" ? currentPartyMode() : selected;
}
// Pisahkan input banyak nomor (baris baru, koma, titik koma), buang spasi tepi
// dan nomor kosong, lalu hilangkan duplikat.
function parseNumberList(text) {
  return new Set(norm(text).split(/[\r\n,;]+/).map(s => key(s)).filter(Boolean));
}
function manualExcludedNumbers() {
  return parseNumberList($("excludeManual").value);
}
function manualRedNumbers() {
  return parseNumberList($("redManual").value);
}
// Nomor daftar kosong "(Nomor daftar kosong)" atau "0" tidak pernah dicocokkan.
function isMatchableRegNo(regNo) {
  const k = key(regNo);
  return !!k && !k.startsWith("(") && k !== "0";
}
// Jalur merah dicocokkan lewat NOMOR DAFTAR (bukan nomor aju), sehingga semua
// record dengan nomor daftar yang sama mendapat jalur yang sama.
function isManualRedRecord(r) {
  return isMatchableRegNo(r.regNo) && manualRedNumbers().has(key(r.regNo));
}
function isRedRecord(r) {
  return redNumbers.has(r.recordId) || isManualRedRecord(r);
}
function listPreview(items, max = 20) {
  const shown = items.slice(0, max).join(", ");
  return items.length > max ? `${shown}, dan ${items.length - max} lainnya` : shown;
}
function isExcluded(r) {
  return excludedNumbers.has(key(r.regNo)) || manualExcludedNumbers().has(key(r.regNo));
}
function updateExcludeStatus() {
  const manual = manualExcludedNumbers();
  const all = new Set([...excludedNumbers, ...manual]);
  const count = records.filter(r => all.has(key(r.regNo))).length;
  let text = `${all.size} nomor daftar unik dikecualikan; ${count} dokumen cocok akan dikeluarkan dari laporan.`;
  if (records.length) {
    const known = new Set(records.map(r => key(r.regNo)));
    const missing = [...manual].filter(n => !known.has(n));
    if (missing.length) text += ` Tidak ditemukan di file Excel (${missing.length}): ${listPreview(missing)}.`;
  }
  $("excludeStatus").textContent = text;
}
function updateRedStatus() {
  const manual = manualRedNumbers();
  const el = $("redStatus");
  if (!manual.size) { el.textContent = "Belum ada nomor daftar jalur merah manual."; return; }
  if (!records.length) { el.textContent = `${manual.size} nomor daftar unik dicatat; menunggu file Excel untuk dicocokkan.`; return; }
  const countByReg = new Map();
  records.forEach(r => { if (isMatchableRegNo(r.regNo)) { const k = key(r.regNo); countByReg.set(k, (countByReg.get(k) || 0) + 1); } });
  const found = [...manual].filter(n => countByReg.has(n));
  const missing = [...manual].filter(n => !countByReg.has(n));
  const docCount = found.reduce((sum, n) => sum + countByReg.get(n), 0);
  let text = `${manual.size} nomor daftar unik; ${found.length} ditemukan (${docCount} dokumen ditandai MERAH).`;
  if (missing.length) text += ` Tidak ditemukan di file Excel (${missing.length}): ${listPreview(missing)}.`;
  el.textContent = text;
}
function refreshPartyOptions() {
  const mode = currentPartyMode();
  const previous = $("partyName").value;
  const names = [...new Set(selectedTypeRecords()
    .map(r => mode === "customer" ? r.customerName : r.supplierName)
    .filter(n => n && !n.startsWith("(")))].sort((a,b)=>a.localeCompare(b));
  $("partyName").innerHTML = "";
  const all = document.createElement("option"); all.value="__ALL__"; all.textContent=`Semua ${mode === "customer" ? "Customer" : "Supplier"}`; $("partyName").append(all);
  names.forEach(name => { const o=document.createElement("option"); o.value=name; o.textContent=name; $("partyName").append(o); });
  if ([...$("partyName").options].some(o => o.value === previous)) $("partyName").value = previous;
}
function renderRedList() {
  const q = norm($("redSearch").value).toLowerCase();
  const list = $("redList"); list.innerHTML = "";
  const items = selectedTypeRecords().filter(r => {
    const currentName = currentPartyMode() === "customer" ? r.customerName : r.supplierName;
    const partyMatch = $("partyName").value === "__ALL__" || currentName === $("partyName").value;
    return partyMatch && `${r.regNo} ${currentName} ${r.aju} ${r.sourceName}`.toLowerCase().includes(q);
  });
  if (!items.length) { list.className="red-list empty"; list.textContent="Tidak ada dokumen sesuai pilihan."; return; }
  list.className="red-list";
  items.forEach(r => {
    const label=document.createElement("label"); label.className="red-row";
    const cb=document.createElement("input"); cb.type="checkbox"; cb.checked=isRedRecord(r);
    if (isManualRedRecord(r)) { cb.disabled=true; cb.title="MERAH dari input nomor daftar manual"; }
    cb.addEventListener("change",()=>{ if(cb.checked) redNumbers.add(r.recordId); else redNumbers.delete(r.recordId); renderRedList(); updateOutput(); });
    const meta=document.createElement("span"); meta.className="meta";
    const title=document.createElement("strong"); title.textContent=r.regNo;
    const currentName = currentPartyMode() === "customer" ? r.customerName : r.supplierName;
    const small=document.createElement("small"); small.textContent=`${currentName || "(Nama pihak tidak ditemukan)"} · ${dateText(r.regDate)||"Tanggal daftar kosong"} · ${r.sourceName}`;
    meta.append(title,small); label.append(cb,meta);
    if (isRedRecord(r)) { const badge=document.createElement("span"); badge.className="red-flag"; badge.textContent="MERAH"; label.append(badge); }
    list.append(label);
  });
}
function renderExcludeList() {
  const q = norm($("excludeSearch").value).toLowerCase();
  const list = $("excludeList"); list.innerHTML = "";
  const items = selectedTypeRecords().filter(r => {
    const name = currentPartyMode() === "customer" ? r.customerName : r.supplierName;
    return `${r.regNo} ${name} ${r.aju} ${r.sourceName}`.toLowerCase().includes(q);
  });
  if (!items.length) { list.className="exclude-list empty"; list.textContent="Tidak ada nomor daftar sesuai pencarian."; updateExcludeStatus(); return; }
  list.className="exclude-list";
  const seen = new Set();
  items.forEach(r => {
    const regNo = key(r.regNo);
    if (seen.has(regNo)) return;
    seen.add(regNo);
    const label=document.createElement("label"); label.className="exclude-row";
    const cb=document.createElement("input"); cb.type="checkbox"; cb.checked=excludedNumbers.has(regNo);
    cb.addEventListener("change",()=>{
      if(cb.checked) excludedNumbers.add(regNo); else excludedNumbers.delete(regNo);
      renderExcludeList(); renderExportList(); updateOutput();
    });
    const meta=document.createElement("span"); meta.className="meta";
    const title=document.createElement("strong"); title.textContent=regNo;
    const currentName = currentPartyMode() === "customer" ? r.customerName : r.supplierName;
    const small=document.createElement("small"); small.textContent=`${currentName || "(Nama pihak tidak ditemukan)"} · ${dateText(r.regDate)||"Tanggal daftar kosong"}`;
    meta.append(title,small); label.append(cb,meta); list.append(label);
  });
  updateExcludeStatus();
}

function exportEligibleRecords() {
  const mode = currentPartyMode();
  const partyFilter = $("partyName").value;
  return selectedTypeRecords().filter(r => {
    const party = mode === "customer" ? r.customerName : r.supplierName;
    return party && (partyFilter === "__ALL__" || party === partyFilter) && !isExcluded(r);
  });
}
function updateExportSelectionStatus() {
  const eligible = exportEligibleRecords();
  const selected = eligible.filter(r => selectedExportIds.has(r.recordId));
  $("exportSelectionStatus").textContent = `${selected.length} dari ${eligible.length} dokumen terpilih untuk diekspor.`;
  $("exportExcelBtn").disabled = selected.length === 0;
  $("selectAllExportBtn").disabled = eligible.length === 0;
  $("clearExportSelectionBtn").disabled = selected.length === 0;
}
function renderExportList() {
  const q = norm($("exportSearch").value).toLowerCase();
  const list = $("exportList");
  list.innerHTML = "";
  const eligibleRecords = exportEligibleRecords();
  const eligibleIds = new Set(eligibleRecords.map(r => r.recordId));
  selectedExportIds = new Set([...selectedExportIds].filter(id => eligibleIds.has(id)));
  const items = eligibleRecords.filter(r => {
    const party = currentPartyMode() === "customer" ? r.customerName : r.supplierName;
    return `${r.regNo} ${party} ${r.aju} ${r.sourceName}`.toLowerCase().includes(q);
  });
  if (!items.length) {
    list.className = "export-list empty";
    list.textContent = "Tidak ada dokumen sesuai filter. Periksa jenis dokumen, pihak, atau pengecualian nomor daftar.";
    updateExportSelectionStatus();
    return;
  }
  list.className = "export-list";
  items.forEach(r => {
    const party = currentPartyMode() === "customer" ? r.customerName : r.supplierName;
    const label = document.createElement("label");
    label.className = "export-row";
    const cb = document.createElement("input");
    cb.type = "checkbox";
    cb.checked = selectedExportIds.has(r.recordId);
    cb.addEventListener("change", () => {
      if (cb.checked) selectedExportIds.add(r.recordId);
      else selectedExportIds.delete(r.recordId);
      updateExportSelectionStatus();
    });
    const meta = document.createElement("span");
    meta.className = "meta";
    const title = document.createElement("strong");
    title.textContent = r.regNo;
    const small = document.createElement("small");
    small.textContent = `${r.aju} · ${party} · ${dateText(r.regDate) || "Tanggal daftar kosong"} · ${r.sourceName}`;
    meta.append(title, small);
    label.append(cb, meta);
    list.append(label);
  });
  updateExportSelectionStatus();
}
function getAjuDocumentDate(aju) {
  const match = key(aju).match(/(20\d{6})/);
  if (!match) return null;
  const value = match[1];
  const year = Number(value.slice(0, 4));
  const month = Number(value.slice(4, 6));
  const day = Number(value.slice(6, 8));
  const date = new Date(year, month - 1, day);
  return date.getFullYear() === year && date.getMonth() === month - 1 && date.getDate() === day ? date : null;
}
function exportSelectedExcel() {
  const type = DOC_TYPES.find(x => x.value === $("docType").value);
  if (!type) return;
  const chosen = exportEligibleRecords().filter(r => selectedExportIds.has(r.recordId));
  if (!chosen.length) {
    $("exportSelectionStatus").textContent = "Pilih minimal satu dokumen yang akan diekspor.";
    return;
  }
  const isInbound = type.direction === "in";
  const movementHeader = isInbound ? "TANGGAL MASUK" : "TANGGAL KELUAR";
  const headers = ["JENIS DOKUMEN", "NO PENGAJUAN", "NO PENDAFTARAN", "TANGGAL DOKUMEN", movementHeader, "KAWASAN BERIKAT"];
  const rows = chosen.map(r => {
    const docDate = getAjuDocumentDate(r.aju);
    const movementDate = excelDate(r.regDate);
    const party = currentPartyMode() === "customer" ? r.customerName : r.supplierName;
    const dateValue = date => {
      if (!date) return "";
      return isInbound
        ? `${date.getFullYear()}-${String(date.getMonth()+1).padStart(2,"0")}-${String(date.getDate()).padStart(2,"0")}`
        : date;
    };
    return [Number(r.docCode) || r.docCode, String(r.aju || ""), String(r.regNo || ""), dateValue(docDate), dateValue(movementDate), party || ""];
  });
  const wb = XLSX.utils.book_new();
  const mainWs = XLSX.utils.aoa_to_sheet([headers, ...rows]);
  mainWs["!cols"] = [{wch:18},{wch:36},{wch:22},{wch:20},{wch:20},{wch:42}];
  mainWs["!autofilter"] = {ref:`A1:F${chosen.length + 1}`};
  mainWs["!freeze"] = {xSplit:0,ySplit:1,topLeftCell:"A2",activePane:"bottomLeft",state:"frozen"};
  XLSX.utils.book_append_sheet(wb, mainWs, "Sheet1");
  if (isInbound) {
    const ajuWs = XLSX.utils.aoa_to_sheet([["DATA"], ...chosen.map(r => [String(r.aju || "")])]);
    ajuWs["!cols"] = [{wch:36}];
    XLSX.utils.book_append_sheet(wb, ajuWs, "Sheet2");
    const detailWs = XLSX.utils.aoa_to_sheet(chosen.map(r => {
      const docDate = getAjuDocumentDate(r.aju);
      const movementDate = excelDate(r.regDate);
      const party = currentPartyMode() === "customer" ? r.customerName : r.supplierName;
      return [String(r.aju || ""), String(r.regNo || ""), docDate || "", movementDate || "", party || ""];
    }));
    detailWs["!cols"] = [{wch:36},{wch:20},{wch:18},{wch:18},{wch:42}];
    XLSX.utils.book_append_sheet(wb, detailWs, "Sheet3");
    wb.Workbook = {Sheets:[{Hidden:0},{Hidden:1},{Hidden:1}]};
  } else {
    const helperWs = XLSX.utils.aoa_to_sheet(chosen.map(r => [String(r.aju || ""), String(r.regNo || "")]));
    helperWs["!cols"] = [{wch:36},{wch:20}];
    XLSX.utils.book_append_sheet(wb, helperWs, "Sheet2");
    wb.Workbook = {Sheets:[{Hidden:0},{Hidden:0}]};
  }
  const today = dateInputToday().replaceAll("-", "");
  const fileName = `BC ${isInbound ? "IN" : "OUT"} - ${today}.xlsx`;
  XLSX.writeFile(wb, fileName, {compression:true});
  $("exportSelectionStatus").textContent = `Berhasil menyiapkan ${chosen.length} dokumen. File: ${fileName}`;
}

function renderUploadedFiles() {
  const box = $("uploadedFiles"); box.innerHTML = "";
  if (!uploadedFiles.size) {
    box.className = "uploaded-files empty"; box.textContent = "Belum ada file Excel yang ditambahkan.";
    $("clearFilesBtn").disabled = true; return;
  }
  box.className = "uploaded-files";
  for (const [fileId, item] of uploadedFiles) {
    const row = document.createElement("div"); row.className = "uploaded-file";
    const info = document.createElement("span"); const skippedDocs = item.allRecords.length - item.records.length;
    info.textContent = `${item.name} — ${item.records.length} dokumen` + (skippedDocs ? ` (${skippedDocs} dokumen dilewati: nomor aju sudah ada dari file lain)` : "");
    const remove = document.createElement("button"); remove.type = "button"; remove.textContent = "Hapus";
    remove.addEventListener("click", () => removeUploadedFile(fileId));
    row.append(info, remove); box.append(row);
  }
  $("clearFilesBtn").disabled = false;
}
function removeUploadedFile(fileId) {
  const item = uploadedFiles.get(fileId);
  if (!item) return;
  item.allRecords.forEach(r => redNumbers.delete(r.recordId));
  uploadedFiles.delete(fileId);
  selectedExportIds.clear();
  rebuildRecords();
  if (!records.length) {
    $("docType").disabled = true; $("partyMode").disabled = true; $("partyName").disabled = true;
    $("generateBtn").disabled = true; $("copyBtn").disabled = true;
    $("docType").innerHTML = '<option value="">Unggah Excel terlebih dahulu</option>';
    $("partyName").innerHTML = '<option value="__ALL__">Semua pihak</option>';
    $("redList").className = "red-list empty"; $("redList").textContent = "Unggah Excel untuk menampilkan daftar dokumen.";
    $("excludeList").className = "exclude-list empty"; $("excludeList").textContent = "Unggah Excel untuk menampilkan nomor daftar.";
    $("exportList").className = "export-list empty"; $("exportList").textContent = "Unggah Excel untuk menampilkan dokumen.";
    $("exportSelectionStatus").textContent = "Belum ada dokumen dipilih."; $("exportExcelBtn").disabled = true; $("selectAllExportBtn").disabled = true; $("clearExportSelectionBtn").disabled = true;
    $("output").value = ""; $("fileStatus").textContent = "Semua file telah dihapus.";
  } else {
    setupDocTypes(); refreshAll();
    $("fileStatus").textContent = `${uploadedFiles.size} file tersisa, ${records.length} dokumen dibaca.`;
  }
  renderUploadedFiles(); updateExcludeStatus(); updateRedStatus();
}
// Satu nomor aju hanya dipakai satu kali (file/record pertama menang), supaya
// dokumen yang sama dari file berbeda tidak terhitung dua kali. Dihitung ulang
// dari data asli setiap file, sehingga menghapus satu file tidak menghilangkan
// dokumen yang sebelumnya dilewati.
function rebuildRecords() {
  const seen = new Set();
  for (const item of uploadedFiles.values()) {
    item.records = item.allRecords.filter(r => {
      if (seen.has(r.aju)) return false;
      seen.add(r.aju);
      return true;
    });
  }
  records = [...uploadedFiles.values()].flatMap(x => x.records);
}
// Sidik jari isi file, untuk mengenali file yang sama walau namanya berbeda.
async function fileHash(buffer) {
  if (globalThis.crypto?.subtle) {
    try {
      const digest = await globalThis.crypto.subtle.digest("SHA-256", buffer);
      return [...new Uint8Array(digest)].map(b => b.toString(16).padStart(2, "0")).join("");
    } catch (e) { /* lanjut ke cadangan */ }
  }
  const bytes = new Uint8Array(buffer);
  let h1 = 0x811c9dc5, h2 = 5381;
  for (let i = 0; i < bytes.length; i++) {
    h1 = Math.imul(h1 ^ bytes[i], 16777619) >>> 0;
    h2 = (Math.imul(h2, 33) + bytes[i]) >>> 0;
  }
  return `fb-${bytes.length}-${h1}-${h2}`;
}
function sumMap(map) {
  return [...map.entries()].filter(([,v])=>v!==0).map(([k,v])=>`${fmtQty(v)} ${k}`).join(" + ") || "-";
}
function packageText(map) {
  return [...map.entries()].filter(([,v])=>v!==0).map(([k,v])=>`${fmtQty(v)} ${k}`).join(" + ") || "-";
}
function buildReport() {
  const type = DOC_TYPES.find(x=>x.value===$("docType").value);
  if (!type) return "";
  const mode = currentPartyMode();
  const labelMode = currentPartyLabelMode();
  const partyFilter = $("partyName").value;
  const docs = selectedTypeRecords().map(r => ({...r, reportPartyName: mode === "customer" ? r.customerName : r.supplierName}))
    .filter(r => r.reportPartyName && (partyFilter === "__ALL__" || r.reportPartyName === partyFilter) && !isExcluded(r));
  if (!docs.length) return "Tidak ada dokumen yang sesuai dengan filter atau semua nomor daftar telah dikecualikan.";
  const movementDateRaw = $("movementDate").value;
  const movementDate = movementDateRaw ? dateText(new Date(`${movementDateRaw}T12:00:00`)) : dateText(new Date());
  const goodsType = norm($("goodsType").value) || "PERLU DIISI";
  const output = [];
  const partyGroups = new Map();
  docs.forEach(r => {
    const party = r.reportPartyName || "(Nama pihak tidak ditemukan)";
    if (!partyGroups.has(party)) partyGroups.set(party, []);
    partyGroups.get(party).push(r);
  });
  for (const [partyName, partyDocs] of partyGroups) {
    // Satu blok per jenis dokumen + Customer/Supplier. Jalur HIJAU/MERAH
    // menjadi rincian di dalam blok yang sama, bukan blok laporan terpisah.
    output.push(`${type.label.toUpperCase()}`);
    output.push(`${labelMode === "customer" ? "Customer" : "Supplier"} : ${partyName}`);

    const routePurposeGroups = new Map();
    partyDocs.forEach(r => {
      const route = isRedRecord(r) ? "MERAH" : "HIJAU";
      const groupKey = `${route}\u0000${r.purpose}`;
      if (!routePurposeGroups.has(groupKey)) routePurposeGroups.set(groupKey, {route, purpose:r.purpose, docs:[]});
      routePurposeGroups.get(groupKey).docs.push(r);
    });
    for (const {route, purpose, docs:purposeDocs} of routePurposeGroups.values()) {
      const numbers = purposeDocs.map(r => r.regNo).filter(n => n && !n.startsWith("(") && n !== "0").sort(naturalSort);
      output.push(`${type.label.toUpperCase()} (${route} | ${purpose}) : ${numbers.join(", ") || "-"}`);
    }

    const allSeals = [...new Set(partyDocs.flatMap(r => r.seals))].sort(naturalSort);
    if (["27-out","27-in","23-in"].includes(type.value)) output.push(`No. Segel : ${allSeals.join(", ") || "-"}`);
    output.push(`Jumlah Dokumen : ${partyDocs.length} Dokumen`);
    output.push(`Jenis Barang : ${goodsType}`);
    const quantities = new Map(), packages = new Map();
    partyDocs.forEach(r => {
      r.quantities.forEach((v,k) => quantities.set(k,(quantities.get(k)||0)+v));
      r.packages.forEach((v,k) => packages.set(k,(packages.get(k)||0)+v));
    });
    output.push(`Jumlah Barang : ${sumMap(quantities)}`);
    output.push(`Jumlah Kemasan : ${packageText(packages)}`);
    const dates = uniqueSortedDates(partyDocs.map(r => excelDate(r.regDate)));
    output.push(`Tanggal Dokumen : ${dates.join(" & ") || "-"}`);
    output.push(`${headerDateLabel(type.value)} : ${movementDate}`);
    output.push("");
  }
  return output.join("\n").trim();
}
function updateOutput() {
  if (records.length && !$ ("generateBtn").disabled) $("output").value = buildReport();
}
function refreshAll() {
  refreshPartyOptions();
  renderRedList();
  renderExcludeList();
  renderExportList();
  updateRedStatus();
  const hasRecords = records.length > 0;
  $("generateBtn").disabled = !hasRecords;
  $("copyBtn").disabled = !hasRecords;
  $("output").value = hasRecords ? buildReport() : "";
}
function setupDocTypes() {
  const previous = $("docType").value;
  const typesPresent = new Set(records.map(r => r.typeValue));
  $("docType").innerHTML = "";
  DOC_TYPES.filter(t => typesPresent.has(t.value)).forEach(t=>{const o=document.createElement("option");o.value=t.value;o.textContent=t.label;$("docType").append(o);});
  if (![...$("docType").options].some(o => o.value === previous)) $("docType").selectedIndex = 0;
}
$("excelFile").addEventListener("change", async (event) => {
  const input = event.target;
  const files = [...(input.files || [])];
  if (!files.length) return;
  input.disabled = true; // cegah pemilihan ganda selagi file diproses
  const added = [], skipped = [], errors = [];
  try {
    for (const file of files) {
      $("fileStatus").textContent = `Membaca ${file.name}...`;
      try {
        if (!/\.(xlsx|xls|xlsm)$/i.test(file.name)) throw new Error("format file tidak didukung (gunakan .xlsx, .xls, atau .xlsm)");
        const data = await file.arrayBuffer();
        const hash = await fileHash(data);
        const same = [...uploadedFiles.values()].find(x => x.hash === hash);
        if (same) { skipped.push(`${file.name} (isi sama dengan ${same.name}, tidak diproses ulang)`); continue; }
        const wb = XLSX.read(data,{type:"array",cellDates:true});
        const parsed = parseWorkbook(wb);
        const fileId = `file-${++fileSequence}`;
        const tagged = parsed.map((r,index)=>({...r, recordId:`${fileId}-${index}-${r.aju}`, sourceId:fileId, sourceName:file.name}));
        uploadedFiles.set(fileId, {name:file.name, hash, allRecords:tagged, records:[]});
        rebuildRecords();
        const item = uploadedFiles.get(fileId);
        const dupCount = tagged.length - item.records.length;
        added.push(`${file.name} (${item.records.length} dokumen${dupCount ? `; ${dupCount} dilewati karena nomor aju sudah ada` : ""})`);
      } catch (err) {
        errors.push(`${file.name}: ${err?.message || String(err)}`);
        console.error(err);
      }
    }
  } finally {
    input.disabled = false;
  }
  if (records.length) {
    setupDocTypes();
    $("docType").disabled = false; $("partyMode").disabled = false; $("partyName").disabled = false;
    if (!$("movementDate").value) $("movementDate").value = dateInputToday();
    $("dateLabel").firstChild.textContent = `${headerDateLabel($("docType").value)} `;
    refreshAll();
  }
  renderUploadedFiles();
  const summary = [];
  if (added.length) summary.push(`Berhasil ditambahkan (${added.length} file): ${added.join("; ")}.`);
  if (skipped.length) summary.push(`Dilewati (${skipped.length} file): ${skipped.join("; ")}.`);
  if (errors.length) summary.push(`Gagal dibaca (${errors.length} file):`, ...errors.map(e => `\u2022 ${e}`));
  summary.push(`Total saat ini: ${uploadedFiles.size} file, ${records.length} dokumen.`);
  $("fileStatus").textContent = summary.join("\n");
  input.value = ""; // izinkan memilih ulang nama file yang sama bila diperlukan
});
$("clearFilesBtn").addEventListener("click",()=>{
  uploadedFiles.clear(); records=[]; redNumbers.clear(); excludedNumbers.clear(); selectedExportIds.clear();
  $("excludeManual").value=""; $("redManual").value=""; $("docType").disabled=true; $("partyMode").disabled=true; $("partyName").disabled=true;
  $("generateBtn").disabled=true; $("copyBtn").disabled=true;
  $("docType").innerHTML='<option value="">Unggah Excel terlebih dahulu</option>';
  $("partyName").innerHTML='<option value="__ALL__">Semua pihak</option>';
  $("redList").className="red-list empty"; $("redList").textContent="Unggah Excel untuk menampilkan daftar dokumen.";
  $("excludeList").className="exclude-list empty"; $("excludeList").textContent="Unggah Excel untuk menampilkan nomor daftar.";
  $("exportList").className="export-list empty"; $("exportList").textContent="Unggah Excel untuk menampilkan dokumen.";
  $("exportSelectionStatus").textContent="Belum ada dokumen dipilih."; $("exportExcelBtn").disabled=true; $("selectAllExportBtn").disabled=true; $("clearExportSelectionBtn").disabled=true;
  $("output").value=""; $("fileStatus").textContent="Semua file telah dihapus.";
  renderUploadedFiles(); updateExcludeStatus(); updateRedStatus();
});
$("docType").addEventListener("change",()=>{
  selectedExportIds.clear();
  $("dateLabel").firstChild.textContent = `${headerDateLabel($("docType").value)} `;
  refreshAll();
});
$("partyMode").addEventListener("change",()=>{selectedExportIds.clear(); refreshAll();});
$("partyName").addEventListener("change",()=>{selectedExportIds.clear(); refreshAll();});
$("goodsType").addEventListener("input",()=>{if(records.length)$("output").value=buildReport();});
$("movementDate").addEventListener("change",()=>{if(records.length)$("output").value=buildReport();});
$("redSearch").addEventListener("input",renderRedList);
$("excludeSearch").addEventListener("input",renderExcludeList);
$("redManual").addEventListener("input",()=>{updateRedStatus(); if(records.length){renderRedList(); $("output").value=buildReport();}});
$("excludeManual").addEventListener("input",()=>{updateExcludeStatus(); if(records.length){renderExportList(); $("output").value=buildReport();}});
$("exportSearch").addEventListener("input",renderExportList);
$("selectAllExportBtn").addEventListener("click",()=>{
  exportEligibleRecords().forEach(r=>selectedExportIds.add(r.recordId));
  renderExportList();
});
$("clearExportSelectionBtn").addEventListener("click",()=>{
  selectedExportIds.clear();
  renderExportList();
});
$("exportExcelBtn").addEventListener("click",exportSelectedExcel);
$("generateBtn").addEventListener("click",()=>{$("output").value=buildReport();$("outputStatus").textContent="Laporan diperbarui. Periksa isi sebelum dikirim.";});
$("copyBtn").addEventListener("click",async()=>{
  try { await navigator.clipboard.writeText($("output").value); $("outputStatus").textContent="Laporan berhasil disalin."; }
  catch { $("output").focus(); $("output").select(); $("outputStatus").textContent="Browser tidak mengizinkan salin otomatis. Tekan Ctrl+C setelah teks dipilih."; }
});
$("movementDate").value = dateInputToday();
updateExcludeStatus();
updateRedStatus();
