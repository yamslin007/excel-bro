export type ThemePreset =
  | "default"
  | "warm-orange"
  | "calm-blue"
  | "vivid-green";

export interface ThemeSettings {
  preset: ThemePreset;
  opacity: number;
  autoMask: boolean;
  hasCustomBackground: boolean;
}

const DB_NAME = "excel_bro_theme";
const DB_VERSION = 1;
const STORE_BACKGROUNDS = "backgrounds";
const STORE_SETTINGS = "settings";

function openDB(): Promise<IDBDatabase> {
  return new Promise((resolve, reject) => {
    const request = indexedDB.open(DB_NAME, DB_VERSION);

    request.onerror = () => reject(request.error ?? new Error("无法打开主题数据库"));
    request.onsuccess = () => resolve(request.result);
    request.onupgradeneeded = () => {
      const db = request.result;
      if (!db.objectStoreNames.contains(STORE_BACKGROUNDS)) {
        db.createObjectStore(STORE_BACKGROUNDS);
      }
      if (!db.objectStoreNames.contains(STORE_SETTINGS)) {
        db.createObjectStore(STORE_SETTINGS);
      }
    };
  });
}

async function withStore<T>(
  storeName: string,
  mode: IDBTransactionMode,
  action: (store: IDBObjectStore) => IDBRequest<T>
): Promise<T> {
  const db = await openDB();
  return new Promise((resolve, reject) => {
    const transaction = db.transaction(storeName, mode);
    const request = action(transaction.objectStore(storeName));
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error ?? new Error("主题数据读写失败"));
  });
}

export function saveBackgroundImage(blob: Blob): Promise<IDBValidKey> {
  return withStore(STORE_BACKGROUNDS, "readwrite", (store) =>
    store.put(blob, "current")
  );
}

export function loadBackgroundImage(): Promise<Blob | null> {
  return withStore(STORE_BACKGROUNDS, "readonly", (store) =>
    store.get("current")
  ).then((value) => (value instanceof Blob ? value : null));
}

export function clearBackgroundImage(): Promise<undefined> {
  return withStore(STORE_BACKGROUNDS, "readwrite", (store) =>
    store.delete("current")
  );
}

export function saveThemeSettings(settings: ThemeSettings): Promise<IDBValidKey> {
  return withStore(STORE_SETTINGS, "readwrite", (store) =>
    store.put(settings, "theme_config")
  );
}

export function loadThemeSettings(): Promise<ThemeSettings | null> {
  return withStore(STORE_SETTINGS, "readonly", (store) =>
    store.get("theme_config")
  ).then((value) => (value && typeof value === "object" ? (value as ThemeSettings) : null));
}
