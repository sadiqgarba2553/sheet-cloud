import { BookRecord } from './model';

function database(): Promise<IDBDatabase> { return new Promise((resolve, reject) => { const request = indexedDB.open('sheetcloud', 1); request.onupgradeneeded = () => { request.result.createObjectStore('workbooks', { keyPath: 'id' }); request.result.createObjectStore('originals'); }; request.onsuccess = () => resolve(request.result); request.onerror = () => reject(request.error); }); }
async function run<T>(name: string, mode: IDBTransactionMode, action: (store: IDBObjectStore) => IDBRequest<T>): Promise<T> { const db = await database(); return new Promise((resolve, reject) => { const tx = db.transaction(name, mode); const request = action(tx.objectStore(name)); tx.oncomplete = () => { db.close(); resolve(request.result); }; tx.onabort = tx.onerror = () => { db.close(); reject(tx.error || new Error('Local storage operation failed.')); }; }); }
export const store = {
  list: (): Promise<BookRecord[]> => run('workbooks', 'readonly', s => s.getAll()),
  get: (id: string): Promise<BookRecord | undefined> => run('workbooks', 'readonly', s => s.get(id)),
  save: (book: BookRecord) => run('workbooks', 'readwrite', s => s.put(book)),
  remove: async (id: string) => {
    const db = await database();
    return new Promise<void>((resolve, reject) => {
      const tx = db.transaction(['workbooks', 'originals'], 'readwrite');
      tx.objectStore('workbooks').delete(id); tx.objectStore('originals').delete(id);
      tx.oncomplete = () => { db.close(); resolve(); };
      tx.onabort = tx.onerror = () => { db.close(); reject(tx.error); };
    });
  },
  original: (id: string, name: string, data: Uint8Array) => run('originals', 'readwrite', s => s.put({ name, data }, id)),
  getOriginal: (id: string): Promise<{ name: string; data: Uint8Array } | undefined> => run('originals', 'readonly', s => s.get(id)),
};
