// app/db/db.js
const DB_NAME = "odl_academic_ops";
const DB_VERSION = 1;

const STORES = {
  students: "students",
  courses: "courses",
  results: "results",
};

function openDB() {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(DB_NAME, DB_VERSION);

    req.onupgradeneeded = (e) => {
      const db = req.result;

      // Students
      if (!db.objectStoreNames.contains(STORES.students)) {
        const s = db.createObjectStore(STORES.students, { keyPath: "studentId" });
        s.createIndex("name", "name", { unique: false });
      }

      // Courses
      if (!db.objectStoreNames.contains(STORES.courses)) {
        db.createObjectStore(STORES.courses, { keyPath: "courseCode" });
      }

      // Results (one row per student-course-session)
      if (!db.objectStoreNames.contains(STORES.results)) {
        const r = db.createObjectStore(STORES.results, { keyPath: "resultId" });
        r.createIndex("studentId", "studentId", { unique: false });
        r.createIndex("courseCode", "courseCode", { unique: false });
        r.createIndex("semester", "semester", { unique: false });
      }
    };

    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}

function tx(db, storeName, mode = "readonly") {
  return db.transaction(storeName, mode).objectStore(storeName);
}

export async function upsertMany(storeName, records) {
  const db = await openDB();
  return new Promise((resolve, reject) => {
    const transaction = db.transaction(storeName, "readwrite");
    const store = transaction.objectStore(storeName);

    for (const rec of records) store.put(rec);

    transaction.oncomplete = () => resolve(true);
    transaction.onerror = () => reject(transaction.error);
  });
}

export async function countStore(storeName) {
  const db = await openDB();
  return new Promise((resolve, reject) => {
    const store = tx(db, storeName, "readonly");
    const req = store.count();
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}

export async function getAllRecords(storeName) {
  const db = await openDB();
  return new Promise((resolve, reject) => {
    const store = tx(db, storeName, "readonly");
    const req = store.getAll();
    req.onsuccess = () => resolve(req.result ?? []);
    req.onerror = () => reject(req.error);
  });
}

export async function deleteRecord(storeName, key) {
  const db = await openDB();
  return new Promise((resolve, reject) => {
    const transaction = db.transaction(storeName, "readwrite");
    const store = transaction.objectStore(storeName);
    store.delete(key);
    transaction.oncomplete = () => resolve(true);
    transaction.onerror = () => reject(transaction.error);
  });
}

export async function replaceMany(storeName, deletes = [], records = []) {
  const db = await openDB();
  return new Promise((resolve, reject) => {
    const transaction = db.transaction(storeName, "readwrite");
    const store = transaction.objectStore(storeName);

    for (const key of deletes) {
      if (key !== null && key !== undefined) store.delete(key);
    }
    for (const rec of records) store.put(rec);

    transaction.oncomplete = () => resolve(true);
    transaction.onerror = () => reject(transaction.error);
  });
}

export async function getCounts() {
  const [students, courses, results] = await Promise.all([
    countStore(STORES.students),
    countStore(STORES.courses),
    countStore(STORES.results),
  ]);
  return { students, courses, results };
}

export { STORES };
