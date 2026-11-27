const DB_NAME = 'reza-bcp-local';
const DB_VERSION = 1;
const RECOVERY_STORE = 'recovery';
const RECENT_STORE = 'recent';

function openDatabase() {
  return new Promise((resolve, reject) => {
    if (!('indexedDB' in window)) {
      reject(new Error('Browser ini tidak mendukung IndexedDB. BCP membutuhkan penyimpanan lokal browser.'));
      return;
    }
    const request = indexedDB.open(DB_NAME, DB_VERSION);
    request.onupgradeneeded = () => {
      const db = request.result;
      if (!db.objectStoreNames.contains(RECOVERY_STORE)) db.createObjectStore(RECOVERY_STORE, { keyPath: 'project_id' });
      if (!db.objectStoreNames.contains(RECENT_STORE)) db.createObjectStore(RECENT_STORE, { keyPath: 'project_id' });
    };
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error || new Error('Gagal membuka penyimpanan lokal BCP.'));
  });
}

function runTransaction(storeName, mode, operation) {
  return openDatabase().then(db => new Promise((resolve, reject) => {
    const tx = db.transaction(storeName, mode);
    const store = tx.objectStore(storeName);
    let result;
    try { result = operation(store); } catch (error) { reject(error); return; }
    tx.oncomplete = () => { db.close(); resolve(result); };
    tx.onerror = () => { db.close(); reject(tx.error || new Error('Transaksi penyimpanan BCP gagal.')); };
    tx.onabort = () => { db.close(); reject(tx.error || new Error('Transaksi penyimpanan BCP dibatalkan.')); };
  }));
}

export function saveRecovery(project) {
  const snapshot = {
    project_id: project.manifest.project_id,
    project_name: project.manifest.project_name,
    updated_at: new Date().toISOString(),
    project
  };
  return runTransaction(RECOVERY_STORE, 'readwrite', store => store.put(snapshot));
}

export function getRecovery(projectId) {
  return openDatabase().then(db => new Promise((resolve, reject) => {
    const request = db.transaction(RECOVERY_STORE, 'readonly').objectStore(RECOVERY_STORE).get(projectId);
    request.onsuccess = () => { db.close(); resolve(request.result || null); };
    request.onerror = () => { db.close(); reject(request.error); };
  }));
}

export function deleteRecovery(projectId) {
  return runTransaction(RECOVERY_STORE, 'readwrite', store => store.delete(projectId));
}

export function saveRecentProject(project) {
  const row = {
    project_id: project.manifest.project_id,
    project_name: project.manifest.project_name,
    company_name: project.manifest.company_name || '',
    updated_at: project.manifest.updated_at,
    last_saved_at: project.manifest.last_saved_at || null
  };
  return runTransaction(RECENT_STORE, 'readwrite', store => store.put(row));
}

export function listRecentProjects() {
  return openDatabase().then(db => new Promise((resolve, reject) => {
    const request = db.transaction(RECENT_STORE, 'readonly').objectStore(RECENT_STORE).getAll();
    request.onsuccess = () => {
      db.close();
      resolve((request.result || []).sort((a, b) => new Date(b.updated_at) - new Date(a.updated_at)));
    };
    request.onerror = () => { db.close(); reject(request.error); };
  }));
}
