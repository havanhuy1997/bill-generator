import type { Batch } from "./types";

function openDatabase(): Promise<IDBDatabase> {
  return new Promise((resolve, reject) => {
    const request = indexedDB.open("chi-invoice-studio", 1);
    request.onupgradeneeded = () =>
      request.result.createObjectStore("batches", { keyPath: "id" });
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
    request.onblocked = () =>
      reject(new Error("Bộ nhớ trình duyệt đang bị chặn."));
  });
}

async function transaction<T>(
  mode: IDBTransactionMode,
  action: (store: IDBObjectStore) => IDBRequest<T>,
): Promise<T> {
  const db = await openDatabase();
  return new Promise((resolve, reject) => {
    const tx = db.transaction("batches", mode);
    const request = action(tx.objectStore("batches"));
    tx.oncomplete = () => {
      db.close();
      resolve(request.result);
    };
    tx.onerror = () => {
      db.close();
      reject(tx.error);
    };
    tx.onabort = () => {
      db.close();
      reject(tx.error);
    };
  });
}

export const listBatches = async () =>
  ((await transaction("readonly", (store) => store.getAll())) as Batch[]).sort(
    (a, b) => b.createdAt.localeCompare(a.createdAt),
  );
export const saveBatch = (batch: Batch) =>
  transaction("readwrite", (store) => store.put(batch));
export const deleteBatch = (id: string) =>
  transaction("readwrite", (store) => store.delete(id));

export function download(blob: Blob, name: string): void {
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = name;
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
  setTimeout(() => URL.revokeObjectURL(url), 30000);
}
