// Armazenamento local (IndexedDB): vistorias (JSON) e fotos (blobs)
const NAME = 'qmc-vistoria';
let dbp = null;

function open() {
  if (!dbp) {
    dbp = new Promise((resolve, reject) => {
      const r = indexedDB.open(NAME, 1);
      r.onupgradeneeded = () => {
        const d = r.result;
        d.createObjectStore('vistorias', { keyPath: 'id' });
        const p = d.createObjectStore('photos', { keyPath: 'id' });
        p.createIndex('vid', 'vid');
      };
      r.onsuccess = () => resolve(r.result);
      r.onerror = () => reject(r.error);
    });
  }
  return dbp;
}

async function run(store, mode, fn) {
  const db = await open();
  return new Promise((resolve, reject) => {
    const t = db.transaction(store, mode);
    const req = fn(t.objectStore(store), t);
    t.oncomplete = () => resolve(req && 'result' in req ? req.result : undefined);
    t.onerror = () => reject(t.error);
    t.onabort = () => reject(t.error);
  });
}

export const listVistorias = () => run('vistorias', 'readonly', (s) => s.getAll());
export const getVistoria = (id) => run('vistorias', 'readonly', (s) => s.get(id));
export const putVistoria = (v) => run('vistorias', 'readwrite', (s) => s.put(v));
export const putPhoto = (rec) => run('photos', 'readwrite', (s) => s.put(rec));
export const getPhoto = (id) => run('photos', 'readonly', (s) => s.get(id));
export const deletePhoto = (id) => run('photos', 'readwrite', (s) => s.delete(id));
export const photosOf = (vid) => run('photos', 'readonly', (s) => s.index('vid').getAll(vid));

export async function deleteVistoria(id) {
  const ph = await photosOf(id);
  const db = await open();
  await new Promise((resolve, reject) => {
    const t = db.transaction(['vistorias', 'photos'], 'readwrite');
    t.objectStore('vistorias').delete(id);
    for (const p of ph) t.objectStore('photos').delete(p.id);
    t.oncomplete = resolve;
    t.onerror = () => reject(t.error);
  });
}
