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

export async function deleteAllQuestionsByArea(area) {
  const questions = await getQuestionsByArea(area);

  // Collect unique fileIds associated with these questions
  const fileIds = [...new Set(questions.map(q => q.fileId).filter(id => id != null))];

  // Delete the questions
  const deletePromises = questions.map(q => deleteQuestion(q.id));
  await Promise.all(deletePromises);

  // Also delete the associated file records from the 'files' store
  if (fileIds.length > 0) {
    const db = await initDB();
    const fileDeletePromises = fileIds.map(fileId =>
      new Promise((resolve, reject) => {
        const transaction = db.transaction(['files'], 'readwrite');
        const store = transaction.objectStore('files');
        const request = store.delete(Number(fileId));
        request.onsuccess = () => resolve();
        request.onerror = (event) => reject(event.target.error);
      })
    );
    await Promise.all(fileDeletePromises);
  }

  return questions.length;
}

/**
 * Deletes ALL records from both object stores (files and questions).
 * This completely empties the database.
 */
export async function borrarTodoElBanco() {
  const db = await initDB();
  const tx = db.transaction(['files', 'questions'], 'readwrite');
  tx.objectStore('files').clear();
  tx.objectStore('questions').clear();
  return new Promise((resolve, reject) => {
    tx.oncomplete = () => resolve();
    tx.onerror = (event) => reject(event.target.error);
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

// ── Backup / Restore ──────────────────────────────────────────────────────────

/**
 * Exports the entire database (files + questions) as a JSON backup file.
 * The file is downloaded automatically in the browser.
 */
export async function exportarBanco() {
  const db = await initDB();

  const files = await new Promise((resolve, reject) => {
    const tx = db.transaction(['files'], 'readonly');
    const req = tx.objectStore('files').getAll();
    req.onsuccess = e => resolve(e.target.result);
    req.onerror = e => reject(e.target.error);
  });

  const questions = await new Promise((resolve, reject) => {
    const tx = db.transaction(['questions'], 'readonly');
    const req = tx.objectStore('questions').getAll();
    req.onsuccess = e => resolve(e.target.result);
    req.onerror = e => reject(e.target.error);
  });

  const backup = {
    version: 1,
    exportedAt: new Date().toISOString(),
    files,
    questions
  };

  const blob = new Blob([JSON.stringify(backup, null, 2)], { type: 'application/json' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `banco_icfes_${new Date().toISOString().slice(0, 10)}.json`;
  a.click();
  URL.revokeObjectURL(url);

  return { files: files.length, questions: questions.length };
}

/**
 * Restores the database from a JSON backup file.
 * Merges with existing data — does not delete what's already there.
 * Returns counts of records imported.
 */
export async function restaurarBanco(jsonFile) {
  const text = await jsonFile.text();
  let backup;
  try {
    backup = JSON.parse(text);
  } catch {
    throw new Error('El archivo no es un respaldo válido.');
  }

  if (!backup.files || !backup.questions) {
    throw new Error('El archivo no tiene el formato esperado.');
  }

  const db = await initDB();

  // Build a map of old fileId → new fileId to relink questions
  const fileIdMap = {};

  for (const file of backup.files) {
    const oldId = file.id;
    const { id: _id, ...fileData } = file; // strip old id, let autoIncrement assign new one
    const newId = await new Promise((resolve, reject) => {
      const tx = db.transaction(['files'], 'readwrite');
      const req = tx.objectStore('files').add(fileData);
      req.onsuccess = e => resolve(e.target.result);
      req.onerror = e => reject(e.target.error);
    });
    fileIdMap[oldId] = newId;
  }

  let questionsImported = 0;
  for (const question of backup.questions) {
    const { id: _id, fileId, ...qData } = question;
    const newFileId = fileIdMap[fileId] ?? null;
    await new Promise((resolve, reject) => {
      const tx = db.transaction(['questions'], 'readwrite');
      const req = tx.objectStore('questions').add({ ...qData, fileId: newFileId });
      req.onsuccess = () => resolve();
      req.onerror = e => reject(e.target.error);
    });
    questionsImported++;
  }

  return { files: backup.files.length, questions: questionsImported };
}
