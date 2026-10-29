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
    // Pengecualian BC 2.3: Jenis Tujuan Transaksi diambil dari BARANG kolom AZ
    // (KODE PERHITUNGAN), bukan HEADER kolom N. Ambil kode pertama yang terisi
    // pada barang dengan nomor aju yang sama; master BC 2.3 memetakan kode 1 dan 2.
    const goodsForAju = goodsByAju.get(aju) || [];
    const purposeCode = docCode === "23"
      ? (goodsForAju.find(g => key(g[51]) !== "")?.[51] ?? "")
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
function manualExcludedNumbers() {
  return new Set(norm($("excludeManual").value).split(/[\n,;]+/).map(s => key(s)).filter(Boolean));
}
function isExcluded(r) {
  return excludedNumbers.has(key(r.regNo)) || manualExcludedNumbers().has(key(r.regNo));
}
function updateExcludeStatus() {
  const manual = manualExcludedNumbers();
  const all = new Set([...excludedNumbers, ...manual]);
  const count = records.filter(r => all.has(key(r.regNo))).length;
  $("excludeStatus").textContent = `${all.size} nomor daftar unik dikecualikan; ${count} dokumen cocok akan dikeluarkan dari laporan.`;
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
    const cb=document.createElement("input"); cb.type="checkbox"; cb.checked=redNumbers.has(r.recordId);
    cb.addEventListener("change",()=>{ if(cb.checked) redNumbers.add(r.recordId); else redNumbers.delete(r.recordId); renderRedList(); updateOutput(); });
    const meta=document.createElement("span"); meta.className="meta";
    const title=document.createElement("strong"); title.textContent=r.regNo;
    const currentName = currentPartyMode() === "customer" ? r.customerName : r.supplierName;
    const small=document.createElement("small"); small.textContent=`${currentName || "(Nama pihak tidak ditemukan)"} · ${dateText(r.regDate)||"Tanggal daftar kosong"} · ${r.sourceName}`;
    meta.append(title,small); label.append(cb,meta);
    if (redNumbers.has(r.recordId)) { const badge=document.createElement("span"); badge.className="red-flag"; badge.textContent="MERAH"; label.append(badge); }
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
      renderExcludeList(); updateOutput();
    });
    const meta=document.createElement("span"); meta.className="meta";
    const title=document.createElement("strong"); title.textContent=regNo;
    const currentName = currentPartyMode() === "customer" ? r.customerName : r.supplierName;
    const small=document.createElement("small"); small.textContent=`${currentName || "(Nama pihak tidak ditemukan)"} · ${dateText(r.regDate)||"Tanggal daftar kosong"}`;
    meta.append(title,small); label.append(cb,meta); list.append(label);
  });
  updateExcludeStatus();
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
    const info = document.createElement("span"); info.textContent = `${item.name} — ${item.records.length} dokumen`;
    const remove = document.createElement("button"); remove.type = "button"; remove.textContent = "Hapus";
    remove.addEventListener("click", () => removeUploadedFile(fileId));
    row.append(info, remove); box.append(row);
  }
  $("clearFilesBtn").disabled = false;
}
function removeUploadedFile(fileId) {
  const item = uploadedFiles.get(fileId);
  if (!item) return;
  item.records.forEach(r => redNumbers.delete(r.recordId));
  uploadedFiles.delete(fileId);
  records = [...uploadedFiles.values()].flatMap(x => x.records);
  if (!records.length) {
    $("docType").disabled = true; $("partyMode").disabled = true; $("partyName").disabled = true;
    $("generateBtn").disabled = true; $("copyBtn").disabled = true;
    $("docType").innerHTML = '<option value="">Unggah Excel terlebih dahulu</option>';
    $("partyName").innerHTML = '<option value="__ALL__">Semua pihak</option>';
    $("redList").className = "red-list empty"; $("redList").textContent = "Unggah Excel untuk menampilkan daftar dokumen.";
    $("excludeList").className = "exclude-list empty"; $("excludeList").textContent = "Unggah Excel untuk menampilkan nomor daftar.";
    $("output").value = ""; $("fileStatus").textContent = "Semua file telah dihapus.";
  } else {
    setupDocTypes(); refreshAll();
    $("fileStatus").textContent = `${uploadedFiles.size} file tersisa, ${records.length} dokumen dibaca.`;
  }
  renderUploadedFiles(); updateExcludeStatus();
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
      const route = redNumbers.has(r.recordId) ? "MERAH" : "HIJAU";
      const groupKey = `${route}\u0000${r.purpose}`;
      if (!routePurposeGroups.has(groupKey)) routePurposeGroups.set(groupKey, {route, purpose:r.purpose, docs:[]});
      routePurposeGroups.get(groupKey).docs.push(r);
    });
    for (const {route, purpose, docs:purposeDocs} of routePurposeGroups.values()) {
      const numbers = purposeDocs.map(r => r.regNo).filter(n => n && !n.startsWith("(") && n !== "0");
      output.push(`${type.label.toUpperCase()} (${route} | ${purpose}) : ${numbers.join(", ") || "-"}`);
    }

    const allSeals = [...new Set(partyDocs.flatMap(r => r.seals))];
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
  const files = [...(event.target.files || [])];
  if (!files.length) return;
  const added = [], errors = [];
  for (const file of files) {
    const fileId = `file-${++fileSequence}`;
    $("fileStatus").textContent = `Membaca ${file.name}...`;
    try {
      const data = await file.arrayBuffer();
      const wb = XLSX.read(data,{type:"array",cellDates:true});
      const parsed = parseWorkbook(wb);
      const tagged = parsed.map((r,index)=>({...r, recordId:`${fileId}-${index}-${r.aju}`, sourceId:fileId, sourceName:file.name}));
      uploadedFiles.set(fileId, {name:file.name, records:tagged});
      added.push(`${file.name} (${tagged.length} dokumen)`);
    } catch (err) {
      errors.push(`${file.name}: ${err.message}`);
      console.error(err);
    }
  }
  records = [...uploadedFiles.values()].flatMap(x => x.records);
  if (records.length) {
    setupDocTypes();
    $("docType").disabled = false; $("partyMode").disabled = false; $("partyName").disabled = false;
    if (!$("movementDate").value) $("movementDate").value = dateInputToday();
    $("dateLabel").firstChild.textContent = `${headerDateLabel($("docType").value)} `;
    refreshAll(); renderUploadedFiles();
  }
  const summary = [];
  if (added.length) summary.push(`Berhasil menambahkan: ${added.join("; ")}.`);
  if (errors.length) summary.push(`Gagal membaca: ${errors.join("; ")}.`);
  summary.push(`Total saat ini: ${uploadedFiles.size} file, ${records.length} dokumen.`);
  $("fileStatus").textContent = summary.join(" ");
  event.target.value = ""; // izinkan memilih ulang nama file yang sama bila diperlukan
});
$("clearFilesBtn").addEventListener("click",()=>{
  uploadedFiles.clear(); records=[]; redNumbers.clear(); excludedNumbers.clear();
  $("excludeManual").value=""; $("docType").disabled=true; $("partyMode").disabled=true; $("partyName").disabled=true;
  $("generateBtn").disabled=true; $("copyBtn").disabled=true;
  $("docType").innerHTML='<option value="">Unggah Excel terlebih dahulu</option>';
  $("partyName").innerHTML='<option value="__ALL__">Semua pihak</option>';
  $("redList").className="red-list empty"; $("redList").textContent="Unggah Excel untuk menampilkan daftar dokumen.";
  $("excludeList").className="exclude-list empty"; $("excludeList").textContent="Unggah Excel untuk menampilkan nomor daftar.";
  $("output").value=""; $("fileStatus").textContent="Semua file telah dihapus.";
  renderUploadedFiles(); updateExcludeStatus();
});
$("docType").addEventListener("change",()=>{
  $("dateLabel").firstChild.textContent = `${headerDateLabel($("docType").value)} `;
  refreshAll();
});
$("partyMode").addEventListener("change",refreshAll);
$("partyName").addEventListener("change",refreshAll);
$("goodsType").addEventListener("input",()=>{if(records.length)$("output").value=buildReport();});
$("movementDate").addEventListener("change",()=>{if(records.length)$("output").value=buildReport();});
$("redSearch").addEventListener("input",renderRedList);
$("excludeSearch").addEventListener("input",renderExcludeList);
$("excludeManual").addEventListener("input",()=>{updateExcludeStatus(); if(records.length)$("output").value=buildReport();});
$("generateBtn").addEventListener("click",()=>{$("output").value=buildReport();$("outputStatus").textContent="Laporan diperbarui. Periksa isi sebelum dikirim.";});
$("copyBtn").addEventListener("click",async()=>{
  try { await navigator.clipboard.writeText($("output").value); $("outputStatus").textContent="Laporan berhasil disalin."; }
  catch { $("output").focus(); $("output").select(); $("outputStatus").textContent="Browser tidak mengizinkan salin otomatis. Tekan Ctrl+C setelah teks dipilih."; }
});
$("movementDate").value = dateInputToday();
updateExcludeStatus();
