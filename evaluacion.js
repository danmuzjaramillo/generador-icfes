/**
 * evaluacion.js
 * Genera evaluaciones de 30 preguntas aleatorias y las exporta a Word, Excel o PDF.
 * 
 * Dependencias (CDN en index.html):
 *   - docx UMD:  https://unpkg.com/docx@8.5.0/build/index.umd.js
 *   - SheetJS:   https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js
 */

import { getQuestionsByArea } from './db.js';

export const areaLabels = {
  matematicas: 'Matemáticas',
  lectura_critica: 'Lectura Crítica',
  sociales_ciudadanas: 'Sociales y Ciudadanas',
  ciencias_naturales: 'Ciencias Naturales',
  ingles: 'Inglés'
};

/**
 * Fisher-Yates shuffle — returns a NEW array, does not mutate input.
 */
function shuffle(arr) {
  const a = [...arr];
  for (let i = a.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [a[i], a[j]] = [a[j], a[i]];
  }
  return a;
}

/**
 * Selects up to `count` random questions from the given area.
 * Returns the selected questions and a warning if fewer than requested.
 */
export async function seleccionarPreguntasAleatorias(area, count = 30) {
  const all = await getQuestionsByArea(area);
  if (all.length === 0) {
    throw new Error(`No hay preguntas guardadas en el área de ${areaLabels[area]}.`);
  }
  const shuffled = shuffle(all);
  const selected = shuffled.slice(0, Math.min(count, shuffled.length));
  const warning = selected.length < count
    ? `Nota: Solo hay ${selected.length} preguntas disponibles (se solicitaron ${count}).`
    : null;
  return { questions: selected, warning };
}

// ─────────────────────────────────────────────
// WORD EXPORT  (uses docx UMD loaded via CDN)
// ─────────────────────────────────────────────
export async function exportarWord(questions, area, incluirClaves) {
  const {
    Document, Packer, Paragraph, TextRun, ImageRun, AlignmentType,
    HeadingLevel, BorderStyle, LevelFormat, PageNumber,
    Header, Footer, PageBreak, WidthType
  } = window.docx;

  // Converts a base64 data URL to Uint8Array for ImageRun
  function base64ToUint8Array(dataUrl) {
    const base64 = dataUrl.includes(',') ? dataUrl.split(',')[1] : dataUrl;
    const binary = atob(base64);
    const bytes = new Uint8Array(binary.length);
    for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
    return bytes;
  }

  function imgType(dataUrl) {
    if (dataUrl.startsWith('data:image/png')) return 'png';
    if (dataUrl.startsWith('data:image/gif')) return 'gif';
    if (dataUrl.startsWith('data:image/bmp')) return 'bmp';
    return 'jpg';
  }

  const areaLabel = areaLabels[area] || area;
  const fecha = new Date().toLocaleDateString('es-CO', {
    year: 'numeric', month: 'long', day: 'numeric'
  });

  // ── Numbering config ──────────────────────────────────────────
  const numberingConfig = {
    config: [
      {
        reference: 'opciones',
        levels: [
          {
            level: 0,
            format: LevelFormat.UPPER_LETTER,
            text: '%1.',
            alignment: AlignmentType.LEFT,
            style: {
              paragraph: { indent: { left: 720, hanging: 360 } }
            }
          }
        ]
      }
    ]
  };

  // ── Helpers ───────────────────────────────────────────────────
  const spacer = (pts = 120) =>
    new Paragraph({ children: [], spacing: { before: pts, after: 0 } });

  const ruleLine = () =>
    new Paragraph({
      children: [],
      border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: 'CCCCCC' } },
      spacing: { before: 80, after: 80 }
    });

  // ── Header ────────────────────────────────────────────────────
  const docHeader = new Header({
    children: [
      new Paragraph({
        children: [
          new TextRun({
            text: `EVALUACIÓN ICFES — ${areaLabel.toUpperCase()}`,
            font: 'Arial', size: 18, bold: true, color: '3B3B3B'
          }),
          new TextRun({
            text: `\t${fecha}`,
            font: 'Arial', size: 18, color: '888888'
          })
        ],
        tabStops: [{ type: 'right', position: 9360 }],
        border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: '6366F1' } }
      })
    ]
  });

  // ── Footer ────────────────────────────────────────────────────
  const docFooter = new Footer({
    children: [
      new Paragraph({
        children: [
          new TextRun({ text: 'ICFES Extractor  •  Página ', font: 'Arial', size: 16, color: '888888' }),
          new TextRun({ children: [PageNumber.CURRENT], font: 'Arial', size: 16, color: '888888' }),
          new TextRun({ text: ' de ', font: 'Arial', size: 16, color: '888888' }),
          new TextRun({ children: [PageNumber.TOTAL_PAGES], font: 'Arial', size: 16, color: '888888' })
        ],
        alignment: AlignmentType.CENTER,
        border: { top: { style: BorderStyle.SINGLE, size: 4, color: 'DDDDDD' } }
      })
    ]
  });

  // ── Build question blocks ─────────────────────────────────────
  const children = [
    // Title block
    new Paragraph({
      children: [new TextRun({ text: `EVALUACIÓN — ${areaLabel.toUpperCase()}`, font: 'Arial', size: 36, bold: true, color: '1E1E2E' })],
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 200 }
    }),
    new Paragraph({
      children: [new TextRun({ text: `Fecha: ${fecha}    |    Número de preguntas: ${questions.length}`, font: 'Arial', size: 20, color: '555555' })],
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 60 }
    }),
    new Paragraph({
      children: [new TextRun({ text: 'Nombre: ________________________________________    Grado: __________', font: 'Arial', size: 22, color: '333333' })],
      spacing: { before: 100, after: 200 }
    }),
    ruleLine()
  ];

  questions.forEach((q, idx) => {
    // Question number + body
    const bodyLines = (q.bodyText || '').split('\n').filter(l => l.trim());

    children.push(spacer(280));
    children.push(
      new Paragraph({
        children: [
          new TextRun({ text: `${idx + 1}. `, font: 'Arial', size: 24, bold: true, color: '6366F1' }),
          new TextRun({ text: bodyLines[0] || '', font: 'Arial', size: 22 })
        ],
        spacing: { before: 0, after: 80 }
      })
    );

    // Extra body lines
    for (let li = 1; li < bodyLines.length; li++) {
      children.push(
        new Paragraph({
          children: [new TextRun({ text: bodyLines[li], font: 'Arial', size: 22 })],
          indent: { left: 360 },
          spacing: { before: 0, after: 60 }
        })
      );
    }

    // Header/context block
    if (q.headerText && q.headerText.trim()) {
      children.push(
        new Paragraph({
          children: [new TextRun({ text: q.headerText.trim(), font: 'Arial', size: 20, italics: true, color: '555555' })],
          indent: { left: 360 },
          spacing: { before: 80, after: 80 },
          border: { left: { style: BorderStyle.SINGLE, size: 8, color: '6366F1' } }
        })
      );
    }

    // Images (shown before options, as per document format)
    if (q.images && q.images.length > 0) {
      for (const imgSrc of q.images) {
        try {
          children.push(
            new Paragraph({
              children: [
                new ImageRun({
                  data: base64ToUint8Array(imgSrc),
                  type: imgType(imgSrc),
                  transformation: { width: 400, height: 280 }
                })
              ],
              alignment: AlignmentType.CENTER,
              spacing: { before: 100, after: 100 }
            })
          );
        } catch (e) {
          console.warn('No se pudo insertar imagen en Word:', e);
        }
      }
    }

    // Options
    const optLetters = ['A', 'B', 'C', 'D'];
    optLetters.forEach(letter => {
      const optText = (q.options && q.options[letter]) ? q.options[letter] : '';
      if (!optText) return;
      const isCorrect = incluirClaves && q.correctOption === letter;
      children.push(
        new Paragraph({
          children: [
            new TextRun({
              text: `${letter}.  ${optText}`,
              font: 'Arial',
              size: 21,
              bold: isCorrect,
              color: isCorrect ? '10B981' : '222222'
            })
          ],
          indent: { left: 520, hanging: 300 },
          spacing: { before: 60, after: 0 }
        })
      );
    });

    // Correct answer badge (if incluirClaves)
    if (incluirClaves && q.correctOption) {
      children.push(
        new Paragraph({
          children: [new TextRun({ text: `✔ Respuesta: ${q.correctOption}`, font: 'Arial', size: 19, bold: true, color: '10B981' })],
          indent: { left: 520 },
          spacing: { before: 80, after: 0 }
        })
      );
    }
  });

  // ── Answer key table (at end) ─────────────────────────────────
  if (incluirClaves) {
    children.push(spacer(400));
    children.push(new Paragraph({ children: [new PageBreak()] }));
    children.push(
      new Paragraph({
        children: [new TextRun({ text: 'TABLA DE RESPUESTAS', font: 'Arial', size: 28, bold: true, color: '1E1E2E' })],
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 300 }
      })
    );

    // Build rows of 10 per row
    const { Table, TableRow, TableCell, ShadingType } = window.docx;
    const border = { style: BorderStyle.SINGLE, size: 4, color: 'DDDDDD' };
    const borders = { top: border, bottom: border, left: border, right: border };
    const cellWidth = { size: 936, type: WidthType.DXA }; // 9360 / 10 cols

    // Header row
    const headerCells = ['N°', ...Array.from({ length: 9 }, (_, i) => String(i + 1))].map(txt =>
      new TableCell({
        borders,
        width: cellWidth,
        shading: { fill: '6366F1', type: ShadingType.CLEAR },
        margins: { top: 80, bottom: 80, left: 80, right: 80 },
        children: [new Paragraph({
          children: [new TextRun({ text: txt, font: 'Arial', size: 18, bold: true, color: 'FFFFFF' })],
          alignment: AlignmentType.CENTER
        })]
      })
    );

    const tableRows = [];
    for (let row = 0; row < Math.ceil(questions.length / 10); row++) {
      const startQ = row * 10;
      const rowNums = Array.from({ length: 10 }, (_, i) => startQ + i);

      // Number row
      const numCells = [
        new TableCell({
          borders, width: cellWidth,
          shading: { fill: 'F3F4F6', type: ShadingType.CLEAR },
          margins: { top: 80, bottom: 80, left: 80, right: 80 },
          children: [new Paragraph({
            children: [new TextRun({ text: 'N°', font: 'Arial', size: 17, bold: true })],
            alignment: AlignmentType.CENTER
          })]
        }),
        ...rowNums.map(qi => new TableCell({
          borders, width: cellWidth,
          shading: { fill: 'F9FAFB', type: ShadingType.CLEAR },
          margins: { top: 80, bottom: 80, left: 80, right: 80 },
          children: [new Paragraph({
            children: [new TextRun({ text: qi < questions.length ? String(qi + 1) : '', font: 'Arial', size: 17 })],
            alignment: AlignmentType.CENTER
          })]
        }))
      ];

      // Answer row
      const ansCells = [
        new TableCell({
          borders, width: cellWidth,
          shading: { fill: 'F3F4F6', type: ShadingType.CLEAR },
          margins: { top: 80, bottom: 80, left: 80, right: 80 },
          children: [new Paragraph({
            children: [new TextRun({ text: 'Rta.', font: 'Arial', size: 17, bold: true })],
            alignment: AlignmentType.CENTER
          })]
        }),
        ...rowNums.map(qi => new TableCell({
          borders, width: cellWidth,
          shading: { fill: 'ECFDF5', type: ShadingType.CLEAR },
          margins: { top: 80, bottom: 80, left: 80, right: 80 },
          children: [new Paragraph({
            children: [new TextRun({
              text: qi < questions.length ? (questions[qi].correctOption || '—') : '',
              font: 'Arial', size: 18, bold: true, color: '10B981'
            })],
            alignment: AlignmentType.CENTER
          })]
        }))
      ];

      tableRows.push(new TableRow({ children: numCells }));
      tableRows.push(new TableRow({ children: ansCells }));
    }

    children.push(new Table({
      width: { size: 9360, type: WidthType.DXA },
      columnWidths: Array(10).fill(936),
      rows: tableRows
    }));
  }

  // ── Assemble document ─────────────────────────────────────────
  const doc = new Document({
    numbering: numberingConfig,
    styles: {
      default: { document: { run: { font: 'Arial', size: 22 } } }
    },
    sections: [{
      properties: {
        page: {
          size: { width: 12240, height: 15840 },
          margin: { top: 1440, right: 1260, bottom: 1440, left: 1260 }
        }
      },
      headers: { default: docHeader },
      footers: { default: docFooter },
      children
    }]
  });

  const blob = await Packer.toBlob(doc);
  descargarBlob(blob, `Evaluacion_${areaLabel}_${getFechaArchivo()}.docx`);
}

// ─────────────────────────────────────────────
// EXCEL EXPORT  (uses SheetJS loaded via CDN)
// ─────────────────────────────────────────────
export function exportarExcel(questions, area, incluirClaves) {
  const XLSX = window.XLSX;
  const areaLabel = areaLabels[area] || area;

  const wsData = [
    ['N°', 'Enunciado', 'Opción A', 'Opción B', 'Opción C', 'Opción D', 'Imagen', ...(incluirClaves ? ['Respuesta Correcta'] : [])]
  ];

  questions.forEach((q, idx) => {
    const row = [
      idx + 1,
      (q.bodyText || '').replace(/\n/g, ' ').trim(),
      q.options?.A || '',
      q.options?.B || '',
      q.options?.C || '',
      q.options?.D || '',
      (q.images && q.images.length > 0) ? `Sí (${q.images.length})` : '',
      ...(incluirClaves ? [q.correctOption || ''] : [])
    ];
    wsData.push(row);
  });

  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet(wsData);

  // Column widths
  ws['!cols'] = [
    { wch: 5 },  // N°
    { wch: 60 }, // Enunciado
    { wch: 30 }, // A
    { wch: 30 }, // B
    { wch: 30 }, // C
    { wch: 30 }, // D
    { wch: 10 }, // Imagen
    ...(incluirClaves ? [{ wch: 18 }] : [])
  ];

  XLSX.utils.book_append_sheet(wb, ws, areaLabel.substring(0, 31));
  XLSX.writeFile(wb, `Evaluacion_${areaLabel}_${getFechaArchivo()}.xlsx`);
}

// ─────────────────────────────────────────────
// PDF EXPORT  (styled print dialog)
// ─────────────────────────────────────────────
export function exportarPDF(questions, area, incluirClaves) {
  const areaLabel = areaLabels[area] || area;
  const fecha = new Date().toLocaleDateString('es-CO', {
    year: 'numeric', month: 'long', day: 'numeric'
  });

  let html = `
    <!DOCTYPE html><html lang="es">
    <head>
      <meta charset="UTF-8">
      <title>Evaluación ${areaLabel}</title>
      <style>
        * { box-sizing: border-box; margin: 0; padding: 0; }
        body {
          font-family: Arial, sans-serif;
          font-size: 11pt;
          color: #1a1a1a;
          background: #fff;
          padding: 0;
        }
        .page-header {
          display: flex;
          justify-content: space-between;
          align-items: flex-start;
          border-bottom: 3px solid #6366f1;
          padding-bottom: 12px;
          margin-bottom: 16px;
        }
        .page-header h1 { font-size: 16pt; color: #1e1e2e; }
        .page-header span { font-size: 10pt; color: #888; }
        .meta-row {
          display: flex;
          gap: 40px;
          margin-bottom: 10px;
          font-size: 10pt;
          color: #555;
        }
        .name-line {
          border: none;
          border-bottom: 1px solid #aaa;
          padding-bottom: 2px;
          display: inline-block;
          width: 260px;
        }
        .divider { border: none; border-top: 1px solid #ddd; margin: 10px 0; }
        .question-block { margin-bottom: 18px; page-break-inside: avoid; }
        .q-num { font-weight: bold; color: #6366f1; }
        .q-body { margin: 4px 0 6px 0; line-height: 1.5; }
        .q-header {
          font-style: italic;
          color: #555;
          border-left: 3px solid #6366f1;
          padding-left: 8px;
          margin: 4px 0 8px 0;
          font-size: 10pt;
        }
        .options { margin-left: 16px; }
        .option {
          padding: 2px 0;
          line-height: 1.5;
        }
        .option.correct { font-weight: bold; color: #10b981; }
        .correct-badge {
          font-size: 9pt;
          color: #10b981;
          font-weight: bold;
          margin-left: 16px;
          margin-top: 2px;
        }
        .answer-table-section { page-break-before: always; margin-top: 20px; }
        .answer-table-section h2 { font-size: 14pt; text-align: center; margin-bottom: 16px; color: #1e1e2e; }
        table.claves {
          border-collapse: collapse;
          width: 100%;
          font-size: 10pt;
        }
        table.claves th {
          background: #6366f1;
          color: white;
          padding: 6px;
          text-align: center;
        }
        table.claves td {
          border: 1px solid #ddd;
          padding: 6px;
          text-align: center;
        }
        table.claves tr:nth-child(even) td { background: #f9f9f9; }
        .q-image {
          text-align: center;
          margin: 6px 0 10px 0;
        }
        .q-image img {
          max-width: 100%;
          max-height: 220px;
          object-fit: contain;
        }
        .ans-cell { color: #10b981; font-weight: bold; }
        .footer { position: fixed; bottom: 12px; width: 100%; text-align: center; font-size: 9pt; color: #aaa; }
        @media print {
          @page { size: Letter; margin: 2cm 2cm 2.5cm 2cm; }
          .no-print { display: none; }
        }
      </style>
    </head>
    <body>
      <div class="page-header">
        <div>
          <h1>EVALUACIÓN — ${areaLabel.toUpperCase()}</h1>
          <div class="meta-row" style="margin-top:8px;">
            <span>Fecha: ${fecha}</span>
            <span>Preguntas: ${questions.length}</span>
          </div>
          <div class="meta-row">
            <span>Nombre: <span class="name-line"></span></span>
            <span>Grado: __________</span>
          </div>
        </div>
      </div>
      <hr class="divider">
  `;

  questions.forEach((q, idx) => {
    html += `<div class="question-block">`;
    html += `<p><span class="q-num">${idx + 1}.</span> <span class="q-body">${escapeHtml(q.bodyText || '')}</span></p>`;
    if (q.headerText && q.headerText.trim()) {
      html += `<div class="q-header">${escapeHtml(q.headerText)}</div>`;
    }
    if (q.images && q.images.length > 0) {
      q.images.forEach(src => {
        html += `<div class="q-image"><img src="${src}" alt="imagen pregunta ${idx + 1}"></div>`;
      });
    }
    html += `<div class="options">`;
    ['A', 'B', 'C', 'D'].forEach(letter => {
      const txt = q.options?.[letter] || '';
      if (!txt) return;
      const isCorrect = incluirClaves && q.correctOption === letter;
      html += `<div class="option ${isCorrect ? 'correct' : ''}">${letter}.&nbsp;${escapeHtml(txt)}</div>`;
    });
    html += `</div>`;
    if (incluirClaves && q.correctOption) {
      html += `<div class="correct-badge">✔ Respuesta: ${q.correctOption}</div>`;
    }
    html += `</div>`;
  });

  // Answer key table
  if (incluirClaves) {
    html += `<div class="answer-table-section"><h2>TABLA DE RESPUESTAS</h2>`;
    html += `<table class="claves"><thead><tr><th>N°</th><th>Respuesta</th><th>N°</th><th>Respuesta</th><th>N°</th><th>Respuesta</th></tr></thead><tbody>`;
    const rows = Math.ceil(questions.length / 3);
    for (let r = 0; r < rows; r++) {
      html += `<tr>`;
      for (let col = 0; col < 3; col++) {
        const qi = r + col * rows;
        if (qi < questions.length) {
          html += `<td>${qi + 1}</td><td class="ans-cell">${questions[qi].correctOption || '—'}</td>`;
        } else {
          html += `<td></td><td></td>`;
        }
      }
      html += `</tr>`;
    }
    html += `</tbody></table></div>`;
  }

  html += `<div class="footer">ICFES Extractor — ${areaLabel} — ${fecha}</div>`;
  html += `</body></html>`;

  // Open print window
  const win = window.open('', '_blank', 'width=900,height=750');
  win.document.write(html);
  win.document.close();
  win.onload = () => win.print();
}

// ─────────────────────────────────────────────
// Helpers
// ─────────────────────────────────────────────
function descargarBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 5000);
}

function getFechaArchivo() {
  const d = new Date();
  return `${d.getFullYear()}${String(d.getMonth() + 1).padStart(2, '0')}${String(d.getDate()).padStart(2, '0')}`;
}

function escapeHtml(str) {
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/\n/g, '<br>');
}