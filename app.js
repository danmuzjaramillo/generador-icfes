import { initDB, addQuestion, getQuestionsByArea, deleteQuestion, addFile, getFiles, deleteFile } from './db.js';
import { parsePDF, parseDocx } from './parser.js';
import { seleccionarPreguntasAleatorias, exportarWord, exportarExcel, exportarPDF } from './evaluacion.js';

// DOM Elements
const views = {
  dashboard: document.getElementById('view-dashboard'),
  upload: document.getElementById('view-upload'),
  browser: document.getElementById('view-area-browser')
};

const navItems = {
  dashboard: document.getElementById('nav-dashboard'),
  upload: document.getElementById('nav-upload')
};

const pageTitle = document.getElementById('page-title-text');
const dropzone = document.getElementById('upload-dropzone');
const fileInput = document.getElementById('file-input');
const selectedFileInfo = document.getElementById('selected-file-info');
const fileNameSpan = document.getElementById('file-info-name');
const btnClearFile = document.getElementById('btn-clear-file');
const selectUploadArea = document.getElementById('select-upload-area');
const btnProcessFile = document.getElementById('btn-process-file');

const progressContainer = document.getElementById('upload-progress-container');
const progressStatusText = document.getElementById('progress-status-text');
const progressBarFill = document.getElementById('progress-bar-fill');
const progressPercentage = document.getElementById('progress-percentage');

const extractionResultsWrapper = document.getElementById('extraction-results-wrapper');
const extractedCountSummary = document.getElementById('extracted-count-summary');
const extractedQuestionsList = document.getElementById('extracted-questions-list');
const btnSaveAllExtracted = document.getElementById('btn-save-all-extracted');

const browserAreaTitle = document.getElementById('browser-area-title');
const browserAreaCountSummary = document.getElementById('browser-area-count-summary');
const savedQuestionsList = document.getElementById('saved-questions-list');
const btnBackToDashboard = document.getElementById('btn-back-to-dashboard');
const btnGenerateTextArea = document.getElementById('btn-generate-text-area');
const btnGenerateTextAll = document.getElementById('btn-generate-text-all');

// Modal Elements
const textModal = document.getElementById('text-modal');
const modalTitle = document.getElementById('modal-title');
const generatedTextBox = document.getElementById('generated-text-box');
const btnCloseModal = document.getElementById('btn-close-modal');
const btnCloseModalOk = document.getElementById('btn-close-modal-ok');
const btnCopyText = document.getElementById('btn-copy-text');

// Evaluación Modal Elements
const evaluacionModal = document.getElementById('evaluacion-modal');
const btnGenerarEvaluacion = document.getElementById('btn-generar-evaluacion');
const btnCloseEvalModal = document.getElementById('btn-close-eval-modal');
const evalSelectArea = document.getElementById('eval-select-area');
const evalNumPreguntas = document.getElementById('eval-num-preguntas');
const evalIncluirClaves = document.getElementById('eval-incluir-claves');
const evalWarningText = document.getElementById('eval-warning-text');
const btnEvalWord = document.getElementById('btn-eval-word');
const btnEvalExcel = document.getElementById('btn-eval-excel');
const btnEvalPdf = document.getElementById('btn-eval-pdf');

// State variables
let activeView = 'dashboard';
let currentBrowsingArea = '';
let currentFile = null;
let extractedQuestions = []; // Temporary questions parsed but not yet saved

// Area Labels Mapping
const areaLabels = {
  matematicas: 'Matemáticas',
  lectura_critica: 'Lectura Crítica',
  sociales_ciudadanas: 'Sociales y Ciudadanas',
  ciencias_naturales: 'Ciencias Naturales',
  ingles: 'Inglés'
};

// Initialize Application
async function init() {
  await initDB();
  setupEventListeners();
  await updateCounts();
}

// Setup Event Listeners
function setupEventListeners() {
  // Navigation
  Object.keys(navItems).forEach(key => {
    navItems[key].addEventListener('click', () => switchView(key));
  });

  // Category Cards Click
  document.querySelectorAll('.area-card').forEach(card => {
    card.addEventListener('click', () => {
      const area = card.getAttribute('data-area');
      browseArea(area);
    });
  });

  // Dropzone Events
  dropzone.addEventListener('click', () => fileInput.click());
  fileInput.addEventListener('change', handleFileSelect);

  dropzone.addEventListener('dragover', (e) => {
    e.preventDefault();
    dropzone.classList.add('dragover');
  });

  dropzone.addEventListener('dragleave', () => {
    dropzone.classList.remove('dragover');
  });

  dropzone.addEventListener('drop', (e) => {
    e.preventDefault();
    dropzone.classList.remove('dragover');
    if (e.dataTransfer.files.length > 0) {
      setFile(e.dataTransfer.files[0]);
    }
  });

  btnClearFile.addEventListener('click', clearFile);

  // File Processing
  btnProcessFile.addEventListener('click', processUploadedFile);

  // Save Extracted Questions
  btnSaveAllExtracted.addEventListener('click', saveExtractedQuestions);

  // Back to Dashboard from Browser
  btnBackToDashboard.addEventListener('click', () => switchView('dashboard'));

  // Modal actions
  btnCloseModal.addEventListener('click', hideModal);
  btnCloseModalOk.addEventListener('click', hideModal);
  btnCopyText.addEventListener('click', copyGeneratedText);

  // Generate text buttons
  btnGenerateTextArea.addEventListener('click', () => generateTextForArea(currentBrowsingArea));
  btnGenerateTextAll.addEventListener('click', () => generateTextForAllAreas());

  // ── Evaluación Modal ──────────────────────────────────────────
  btnGenerarEvaluacion.addEventListener('click', () => {
    evaluacionModal.classList.add('active');
    evalWarningText.style.display = 'none';
  });

  btnCloseEvalModal.addEventListener('click', () => evaluacionModal.classList.remove('active'));

  evaluacionModal.addEventListener('click', (e) => {
    if (e.target === evaluacionModal) evaluacionModal.classList.remove('active');
  });

  // Pre-fill area selector with currently browsed area if in browser view
  btnGenerarEvaluacion.addEventListener('click', () => {
    if (activeView === 'browser' && currentBrowsingArea) {
      evalSelectArea.value = currentBrowsingArea;
    }
  });

  async function ejecutarExportacion(formato) {
    const area = evalSelectArea.value;
    const count = parseInt(evalNumPreguntas.value, 10) || 30;
    const incluirClaves = evalIncluirClaves.checked;

    // Disable buttons during generation
    [btnEvalWord, btnEvalExcel, btnEvalPdf].forEach(b => b.setAttribute('disabled', 'true'));
    evalWarningText.style.display = 'none';

    try {
      const { questions, warning } = await seleccionarPreguntasAleatorias(area, count);

      if (warning) {
        evalWarningText.textContent = warning;
        evalWarningText.style.display = 'block';
      }

      if (formato === 'word') {
        if (!window.docx) throw new Error('La librería docx no está cargada. Verifica tu conexión a internet.');
        await exportarWord(questions, area, incluirClaves);
      } else if (formato === 'excel') {
        if (!window.XLSX) throw new Error('La librería XLSX no está cargada. Verifica tu conexión a internet.');
        exportarExcel(questions, area, incluirClaves);
      } else if (formato === 'pdf') {
        exportarPDF(questions, area, incluirClaves);
      }

    } catch (err) {
      alert(`Error al generar la evaluación: ${err.message}`);
    } finally {
      [btnEvalWord, btnEvalExcel, btnEvalPdf].forEach(b => b.removeAttribute('disabled'));
    }
  }

  btnEvalWord.addEventListener('click', () => ejecutarExportacion('word'));
  btnEvalExcel.addEventListener('click', () => ejecutarExportacion('excel'));
  btnEvalPdf.addEventListener('click', () => ejecutarExportacion('pdf'));
}

// Switch between SPA views
function switchView(viewName) {
  activeView = viewName;
  
  // Update sidebar highlight
  Object.keys(navItems).forEach(key => {
    if (key === viewName) {
      navItems[key].classList.add('active');
    } else {
      navItems[key].classList.remove('active');
    }
  });

  // Update DOM panels visibility
  Object.keys(views).forEach(key => {
    if (key === viewName) {
      views[key].classList.remove('hidden');
    } else {
      views[key].classList.add('hidden');
    }
  });

  // Clean elements if entering/leaving views
  if (viewName === 'dashboard') {
    pageTitle.textContent = 'Dashboard de Áreas';
    updateCounts();
  } else if (viewName === 'upload') {
    pageTitle.textContent = 'Carga y Extracción';
  }
}

// Update DB question counts per area
async function updateCounts() {
  for (const area of Object.keys(areaLabels)) {
    const list = await getQuestionsByArea(area);
    const countEl = document.getElementById(`count-${area}`);
    if (countEl) {
      countEl.textContent = list.length;
    }
  }
}

// Browse questions in a specific area
async function browseArea(area) {
  currentBrowsingArea = area;
  pageTitle.textContent = `Banco de Preguntas > ${areaLabels[area]}`;
  
  // Render questions
  await renderSavedQuestions();

  switchView('browser');
}

// Set uploaded file
function setFile(file) {
  const extension = file.name.split('.').pop().toLowerCase();
  if (extension !== 'pdf' && extension !== 'docx') {
    alert('Por favor, selecciona un archivo válido (.pdf o .docx)');
    return;
  }
  currentFile = file;
  fileNameSpan.textContent = `${file.name} (${(file.size / 1024 / 1024).toFixed(2)} MB)`;
  selectedFileInfo.classList.add('visible');
  btnProcessFile.removeAttribute('disabled');

  // Reset progress and results UI
  progressContainer.classList.remove('visible');
  progressBarFill.style.width = '0%';
  progressPercentage.textContent = '0%';
  extractionResultsWrapper.classList.add('hidden');
  extractedQuestions = [];
}

// Clear selected file
function clearFile() {
  currentFile = null;
  fileInput.value = '';
  selectedFileInfo.classList.remove('visible');
  btnProcessFile.setAttribute('disabled', 'true');
  progressContainer.classList.remove('visible');
  extractionResultsWrapper.classList.add('hidden');
  extractedQuestions = [];
}

// Handle traditional file input select
function handleFileSelect(e) {
  if (e.target.files.length > 0) {
    setFile(e.target.files[0]);
  }
}

// Process the file (parse and extract questions)
async function processUploadedFile() {
  if (!currentFile) return;

  const area = selectUploadArea.value;
  progressContainer.classList.add('visible');
  progressBarFill.style.width = '10%';
  progressPercentage.textContent = '10%';
  progressStatusText.textContent = 'Leyendo archivo...';
  btnProcessFile.setAttribute('disabled', 'true');

  try {
    const arrayBuffer = await currentFile.arrayBuffer();
    const extension = currentFile.name.split('.').pop().toLowerCase();
    
    let questionsList = [];

    const onProgress = (percentage) => {
      progressBarFill.style.width = `${percentage}%`;
      progressPercentage.textContent = `${percentage}%`;
      if (percentage < 50) {
        progressStatusText.textContent = 'Extrayendo texto del documento...';
      } else if (percentage < 80) {
        progressStatusText.textContent = 'Filtrando y procesando preguntas...';
      } else if (percentage < 100) {
        progressStatusText.textContent = 'Estructurando preguntas e imágenes...';
      } else {
        progressStatusText.textContent = '¡Extracción completada!';
      }
    };

    if (extension === 'pdf') {
      questionsList = await parsePDF(arrayBuffer, onProgress);
    } else {
      questionsList = await parseDocx(arrayBuffer, onProgress);
    }

    extractedQuestions = questionsList.map((q, idx) => ({
      ...q,
      area: area,
      tempId: idx
    }));

    renderExtractedQuestions();

  } catch (error) {
    console.error('Error al procesar el archivo:', error);
    alert(`Error de procesamiento: ${error.message}`);
    progressContainer.classList.remove('visible');
  } finally {
    btnProcessFile.removeAttribute('disabled');
  }
}

// Render extracted questions in draft review list
function renderExtractedQuestions() {
  extractedQuestionsList.innerHTML = '';
  
  if (extractedQuestions.length === 0) {
    extractedCountSummary.textContent = 'No se detectaron preguntas individuales válidas (o se descartaron por ser de múltiples preguntas).';
    extractedQuestionsList.innerHTML = '<div class="empty-state">No se extrajeron preguntas. Intenta con otro archivo.</div>';
    extractionResultsWrapper.classList.remove('hidden');
    return;
  }

  extractedCountSummary.textContent = `Se detectaron ${extractedQuestions.length} preguntas individuales listas para guardar.`;

  extractedQuestions.forEach(q => {
    const card = document.createElement('div');
    card.className = 'question-item';
    card.id = `extracted-card-${q.tempId}`;

    // Options mapping
    const optLetters = ['A', 'B', 'C', 'D'];
    let optionsHtml = '';
    optLetters.forEach(letter => {
      optionsHtml += `
        <div class="option-input-group">
          <span class="option-letter">${letter}</span>
          <input type="text" class="option-text-field" data-q-temp-id="${q.tempId}" data-opt="${letter}" value="${q.options[letter]}">
        </div>
      `;
    });

    // Images HTML
    let imagesHtml = '';
    if (q.images && q.images.length > 0) {
      imagesHtml = '<div class="question-images">';
      q.images.forEach((imgSrc, imgIdx) => {
        imagesHtml += `
          <div class="extracted-img-container" id="extracted-img-${q.tempId}-${imgIdx}">
            <img src="${imgSrc}" alt="Pregunta ${q.number}">
            <button class="remove-img-btn" onclick="window.removeExtractedImage(${q.tempId}, ${imgIdx})">×</button>
          </div>
        `;
      });
      imagesHtml += '</div>';
    }

    card.innerHTML = `
      <div class="question-meta">
        <span>Pregunta original: #${q.number}</span>
        <span>Área: ${areaLabels[q.area]}</span>
      </div>
      
      <div style="margin-bottom: 8px;">
        ${imagesHtml}
      </div>

      <div style="margin-bottom: 12px;">
        <label class="form-label">Enunciado de la Pregunta</label>
        <textarea class="question-text-edit" style="height: 100px;" data-q-temp-id="${q.tempId}" data-field="bodyText" required>${q.bodyText}</textarea>
      </div>

      <label class="form-label">Opciones de Respuesta</label>
      <div class="question-options-grid">
        ${optionsHtml}
      </div>

      <div class="question-action-bar">
        <div class="correct-select">
          <span class="form-label" style="margin-bottom: 0;">Clave de Respuesta:</span>
          <select class="select-control" style="width: 80px; padding: 4px 8px;" data-q-temp-id="${q.tempId}" data-field="correctOption">
            <option value="">Ninguna</option>
            <option value="A" ${q.correctOption === 'A' ? 'selected' : ''}>A</option>
            <option value="B" ${q.correctOption === 'B' ? 'selected' : ''}>B</option>
            <option value="C" ${q.correctOption === 'C' ? 'selected' : ''}>C</option>
            <option value="D" ${q.correctOption === 'D' ? 'selected' : ''}>D</option>
          </select>
        </div>

        <button class="btn btn-danger" style="padding: 6px 12px; font-size: 12px;" onclick="window.removeExtractedQuestion(${q.tempId})">
          Descartar
        </button>
      </div>
    `;

    extractedQuestionsList.appendChild(card);
  });

  // Attach event listeners for real-time state synchronization
  document.querySelectorAll('[data-q-temp-id]').forEach(el => {
    el.addEventListener('input', (e) => {
      const tempId = parseInt(e.target.getAttribute('data-q-temp-id'), 10);
      const field = e.target.getAttribute('data-field');
      const optLetter = e.target.getAttribute('data-opt');
      
      const q = extractedQuestions.find(item => item.tempId === tempId);
      if (!q) return;

      if (optLetter) {
        q.options[optLetter] = e.target.value;
      } else if (field) {
        q[field] = e.target.value;
      }
    });

    el.addEventListener('change', (e) => {
      const tempId = parseInt(e.target.getAttribute('data-q-temp-id'), 10);
      const field = e.target.getAttribute('data-field');
      if (field === 'correctOption') {
        const q = extractedQuestions.find(item => item.tempId === tempId);
        if (q) q.correctOption = e.target.value;
      }
    });
  });

  extractionResultsWrapper.classList.remove('hidden');
}

// Global window functions for inline onclick handlers
window.removeExtractedImage = function(tempId, imgIdx) {
  const q = extractedQuestions.find(item => item.tempId === tempId);
  if (q && q.images) {
    q.images.splice(imgIdx, 1);
    const imgEl = document.getElementById(`extracted-img-${tempId}-${imgIdx}`);
    if (imgEl) imgEl.remove();
  }
};

window.removeExtractedQuestion = function(tempId) {
  extractedQuestions = extractedQuestions.filter(item => item.tempId !== tempId);
  const cardEl = document.getElementById(`extracted-card-${tempId}`);
  if (cardEl) cardEl.remove();
  
  extractedCountSummary.textContent = `Se detectaron ${extractedQuestions.length} preguntas individuales listas para guardar.`;
  if (extractedQuestions.length === 0) {
    renderExtractedQuestions();
  }
};

// Save temporary extracted questions to database
async function saveExtractedQuestions() {
  if (extractedQuestions.length === 0) return;

  btnSaveAllExtracted.setAttribute('disabled', 'true');
  btnSaveAllExtracted.textContent = 'Guardando...';

  try {
    // Write new file entry to file storage
    const fileId = await addFile({
      name: currentFile.name,
      size: currentFile.size,
      area: selectUploadArea.value,
      questionCount: extractedQuestions.length
    });

    for (const q of extractedQuestions) {
      await addQuestion({
        fileId: fileId,
        area: q.area,
        bodyText: q.bodyText,
        options: q.options,
        correctOption: q.correctOption,
        images: q.images
      });
    }

    alert('¡Preguntas guardadas con éxito en la base de datos!');
    clearFile();
    switchView('dashboard');
  } catch (err) {
    console.error('Error al guardar preguntas:', err);
    alert('Hubo un error al guardar en la base de datos local.');
  } finally {
    btnSaveAllExtracted.removeAttribute('disabled');
    btnSaveAllExtracted.textContent = 'Guardar Seleccionadas a la BBDD';
  }
}

// Render saved questions in browser view
async function renderSavedQuestions() {
  savedQuestionsList.innerHTML = '';
  const questions = await getQuestionsByArea(currentBrowsingArea);
  
  browserAreaCountSummary.textContent = `${questions.length} preguntas guardadas en esta categoría.`;

  if (questions.length === 0) {
    savedQuestionsList.innerHTML = `
      <div class="empty-state">
        <svg fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 6.253v13m0-13C10.832 5.477 9.246 5 7.5 5S4.168 5.477 3 6.253v13C4.168 18.477 5.754 18 7.5 18s3.332.477 4.5 1.253m0-13C13.168 5.477 14.754 5 16.5 5c1.747 0 3.332.477 4.5 1.253v13C19.832 18.477 18.247 18 16.5 18c-1.746 0-3.332.477-4.5 1.253"/></svg>
        <p>Aún no hay preguntas guardadas en esta categoría.</p>
        <button class="btn btn-primary" style="margin-top: 16px;" onclick="window.goToUploadView()">Subir archivo</button>
      </div>
    `;
    btnGenerateTextArea.setAttribute('disabled', 'true');
    return;
  }

  btnGenerateTextArea.removeAttribute('disabled');

  questions.forEach(q => {
    const card = document.createElement('div');
    card.className = 'saved-q-card';
    card.id = `saved-card-${q.id}`;

    // Images
    let imagesHtml = '';
    if (q.images && q.images.length > 0) {
      imagesHtml = '<div class="question-images">';
      q.images.forEach(imgSrc => {
        imagesHtml += `
          <div class="extracted-img-container">
            <img src="${imgSrc}" alt="Imagen de pregunta">
          </div>
        `;
      });
      imagesHtml += '</div>';
    }

    card.innerHTML = `
      <div class="question-meta">
        <span>ID #${q.id}</span>
        <span>Fecha de creación: ${new Date(q.createdAt).toLocaleDateString()}</span>
      </div>

      
      <div class="saved-q-body">${q.bodyText}</div>

      ${imagesHtml}

      <div class="saved-q-options">
        <div class="saved-q-option ${q.correctOption === 'A' ? 'correct' : ''}"><strong>A.</strong> ${q.options.A}</div>
        <div class="saved-q-option ${q.correctOption === 'B' ? 'correct' : ''}"><strong>B.</strong> ${q.options.B}</div>
        <div class="saved-q-option ${q.correctOption === 'C' ? 'correct' : ''}"><strong>C.</strong> ${q.options.C}</div>
        <div class="saved-q-option ${q.correctOption === 'D' ? 'correct' : ''}"><strong>D.</strong> ${q.options.D}</div>
      </div>

      <div class="question-action-bar" style="margin-top: 12px; border-top: 1px solid rgba(255,255,255,0.05); padding-top: 8px;">
        <span style="font-size: 13px; color: var(--accent); font-weight: 600;">
          ${q.correctOption ? `Respuesta Correcta: ${q.correctOption}` : 'Sin clave especificada'}
        </span>
        <button class="btn btn-danger" style="padding: 4px 8px; font-size: 11px;" onclick="window.deleteSavedQuestion(${q.id})">
          Eliminar
        </button>
      </div>
    `;

    savedQuestionsList.appendChild(card);
  });
}

// Redirect helpers
window.goToUploadView = function() {
  switchView('upload');
};

window.deleteSavedQuestion = async function(id) {
  if (confirm('¿Estás seguro de que deseas eliminar esta pregunta del banco?')) {
    await deleteQuestion(id);
    const card = document.getElementById(`saved-card-${id}`);
    if (card) card.remove();
    await updateCounts();
    
    // Refresh count label
    const questions = await getQuestionsByArea(currentBrowsingArea);
    browserAreaCountSummary.textContent = `${questions.length} preguntas guardadas en esta categoría.`;
    if (questions.length === 0) {
      renderSavedQuestions();
    }
  }
};

// Generate formatted plain-text output highlighting the correct options in bold
function formatQuestionsText(questions, title) {
  let output = `==================================================\n`;
  output += `PREGUNTAS EXTRAÍDAS: ${title.toUpperCase()}\n`;
  output += `==================================================\n\n`;

  if (questions.length === 0) {
    output += `No hay preguntas registradas.\n`;
    return output;
  }

  questions.forEach((q, idx) => {
    output += `PREGUNTA ${idx + 1}.\n`;
    output += `${q.bodyText}\n\n`;
    
    // Formatting options with bolding for correct answers
    const formatOpt = (letter) => {
      const text = q.options[letter] || '';
      return q.correctOption === letter ? `**${letter}. ${text}** (Correcta)` : `${letter}. ${text}`;
    };

    output += `${formatOpt('A')}\n`;
    output += `${formatOpt('B')}\n`;
    output += `${formatOpt('C')}\n`;
    output += `${formatOpt('D')}\n\n`;
    
    if (q.correctOption) {
      output += `Solución correcta: Opción ${q.correctOption}\n`;
    }
    
    if (q.images && q.images.length > 0) {
      output += `[Imágenes asociadas: ${q.images.length}]\n`;
    }
    
    output += `--------------------------------------------------\n\n`;
  });

  return output;
}

// Show generated questions modal
function showModal(title, text) {
  modalTitle.textContent = title;
  generatedTextBox.value = text;
  textModal.classList.add('active');
}

// Hide generated questions modal
function hideModal() {
  textModal.classList.remove('active');
}

// Copy generated content to clipboard
function copyGeneratedText() {
  generatedTextBox.select();
  generatedTextBox.setSelectionRange(0, 99999); // For mobile devices
  navigator.clipboard.writeText(generatedTextBox.value);
  
  const originalText = btnCopyText.textContent;
  btnCopyText.textContent = '¡Copiado!';
  setTimeout(() => {
    btnCopyText.textContent = originalText;
  }, 2000);
}

// Generate formatted questions for single Area
async function generateTextForArea(area) {
  const list = await getQuestionsByArea(area);
  const formatted = formatQuestionsText(list, areaLabels[area]);
  showModal(`Preguntas de ${areaLabels[area]}`, formatted);
}

// Generate formatted questions for All Areas combined
async function generateTextForAllAreas() {
  let combinedText = '';
  for (const area of Object.keys(areaLabels)) {
    const list = await getQuestionsByArea(area);
    if (list.length > 0) {
      combinedText += formatQuestionsText(list, areaLabels[area]) + '\n\n';
    }
  }

  if (!combinedText) {
    combinedText = 'La base de datos está vacía. Sube y extrae preguntas primero.';
  }

  showModal('Banco Completo de Preguntas ICFES', combinedText);
}

// Start application
window.addEventListener('DOMContentLoaded', init);