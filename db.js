const DB_NAME = 'ICFES_Extractor_DB';
const DB_VERSION = 1;

let dbInstance = null;

export function initDB() {
  return new Promise((resolve, reject) => {
    if (dbInstance) {
      resolve(dbInstance);
      return;
    }

    const request = indexedDB.open(DB_NAME, DB_VERSION);

    request.onerror = (event) => {
      console.error('Database error:', event.target.error);
      reject(event.target.error);
    };

    request.onsuccess = (event) => {
      dbInstance = event.target.result;
      resolve(dbInstance);
    };

    request.onupgradeneeded = (event) => {
      const db = event.target.result;

      // Store for files uploaded
      if (!db.objectStoreNames.contains('files')) {
        const fileStore = db.createObjectStore('files', { keyPath: 'id', autoIncrement: true });
        fileStore.createIndex('area', 'area', { unique: false });
        fileStore.createIndex('name', 'name', { unique: false });
      }

      // Store for extracted questions
      if (!db.objectStoreNames.contains('questions')) {
        const questionStore = db.createObjectStore('questions', { keyPath: 'id', autoIncrement: true });
        questionStore.createIndex('area', 'area', { unique: false });
        questionStore.createIndex('fileId', 'fileId', { unique: false });
      }
    };
  });
}

// Files operations
export async function addFile(fileRecord) {
  const db = await initDB();
  return new Promise((resolve, reject) => {
    const transaction = db.transaction(['files'], 'readwrite');
    const store = transaction.objectStore('files');
    const request = store.add({
      name: fileRecord.name,
      size: fileRecord.size,
      area: fileRecord.area,
      uploadedAt: new Date().toISOString(),
      questionCount: fileRecord.questionCount || 0
    });

    request.onsuccess = (event) => resolve(event.target.result);
    request.onerror = (event) => reject(event.target.error);
  });
}

export async function getFiles() {
  const db = await initDB();
  return new Promise((resolve, reject) => {
    const transaction = db.transaction(['files'], 'readonly');
    const store = transaction.objectStore('files');
    const request = store.getAll();

    request.onsuccess = (event) => resolve(event.target.result);
    request.onerror = (event) => reject(event.target.error);
  });
}

export async function deleteFile(fileId) {
  const db = await initDB();
  
  // First delete all questions associated with this file
  const questions = await getQuestionsByFile(fileId);
  const deletePromises = questions.map(q => deleteQuestion(q.id));
  await Promise.all(deletePromises);

  return new Promise((resolve, reject) => {
    const transaction = db.transaction(['files'], 'readwrite');
    const store = transaction.objectStore('files');
    const request = store.delete(Number(fileId));

    request.onsuccess = () => resolve();
    request.onerror = (event) => reject(event.target.error);
  });
}

// Questions operations
export async function addQuestion(questionRecord) {
  const db = await initDB();
  return new Promise((resolve, reject) => {
    const transaction = db.transaction(['questions'], 'readwrite');
    const store = transaction.objectStore('questions');
    const request = store.add({
      fileId: questionRecord.fileId ? Number(questionRecord.fileId) : null,
      area: questionRecord.area, // 'matematicas', 'lectura_critica', 'sociales_ciudadanas', 'ciencias_naturales', 'ingles'
      headerText: questionRecord.headerText || '',
      bodyText: questionRecord.bodyText || '',
      options: {
        A: questionRecord.options.A || '',
        B: questionRecord.options.B || '',
        C: questionRecord.options.C || '',
        D: questionRecord.options.D || ''
      },
      correctOption: questionRecord.correctOption || '', // 'A', 'B', 'C', 'D' or empty
      solutionExplanation: questionRecord.solutionExplanation || '',
      images: questionRecord.images || [], // Array of base64 strings or Object URLs
      createdAt: new Date().toISOString()
    });

    request.onsuccess = (event) => resolve(event.target.result);
    request.onerror = (event) => reject(event.target.error);
  });
}

export async function getQuestionsByArea(area) {
  const db = await initDB();
  return new Promise((resolve, reject) => {
    const transaction = db.transaction(['questions'], 'readonly');
    const store = transaction.objectStore('questions');
    const index = store.index('area');
    const request = index.getAll(area);

    request.onsuccess = (event) => resolve(event.target.result);
    request.onerror = (event) => reject(event.target.error);
  });
}

export async function getQuestionsByFile(fileId) {
  const db = await initDB();
  return new Promise((resolve, reject) => {
    const transaction = db.transaction(['questions'], 'readonly');
    const store = transaction.objectStore('questions');
    const index = store.index('fileId');
    const request = index.getAll(Number(fileId));

    request.onsuccess = (event) => resolve(event.target.result);
    request.onerror = (event) => reject(event.target.error);
  });
}

export async function deleteQuestion(id) {
  const db = await initDB();
  return new Promise((resolve, reject) => {
    const transaction = db.transaction(['questions'], 'readwrite');
    const store = transaction.objectStore('questions');
    const request = store.delete(Number(id));

    request.onsuccess = () => resolve();
    request.onerror = (event) => reject(event.target.error);
  });
}

export async function updateQuestionCount(fileId, count) {
  const db = await initDB();
  return new Promise((resolve, reject) => {
    const transaction = db.transaction(['files'], 'readwrite');
    const store = transaction.objectStore('files');
    const getReq = store.get(Number(fileId));

    getReq.onsuccess = () => {
      const record = getReq.result;
      if (record) {
        record.questionCount = count;
        const updateReq = store.put(record);
        updateReq.onsuccess = () => resolve();
        updateReq.onerror = (e) => reject(e.target.error);
      } else {
        resolve();
      }
    };
    getReq.onerror = (e) => reject(e.target.error);
  });
}
