/**
 * Parser for PDF and Word documents to extract ICFES questions.
 */

// We assume PDF.js (pdfjsLib) and Mammoth are loaded via CDN in index.html
// global pdfjsLib, mammoth

/**
 * Normalizes text to make extraction easier.
 */
function normalizeText(text) {
  return text
    .replace(/\r\n/g, '\n')
    .replace(/\r/g, '\n')
    .replace(/\u00a0/g, ' ') // Replace non-breaking spaces
    .trim();
}

/**
 * Scans the full text to find all question numbers that belong to multi-question reference texts.
 * (e.g. "Responda las preguntas 6 a 9 de acuerdo con...")
 */
function getExcludedQuestionNumbers(text) {
  const excluded = new Set();
  // Match patterns like "preguntas 6,7,8 y 9" or "preguntas del 6 al 9" or "preguntas 6 y 7"
  // Note: Only matches plural "preguntas", not singular "pregunta"
  const regex = /\bpreguntas\s+([0-9\s,ay\-a|de|al|del|la|las]+)/gi;
  let match;
  while ((match = regex.exec(text)) !== null) {
    const sequenceStr = match[1];
    const numbers = sequenceStr.match(/\d+/g);
    if (numbers && numbers.length >= 2) {
      // Verify context to avoid matches with informational texts (e.g. "este módulo contiene preguntas de la 1 a la 20")
      const matchIndex = match.index;
      const startContext = Math.max(0, matchIndex - 80);
      const contextText = text.substring(startContext, matchIndex + match[0].length).toLowerCase();
      
      const isInstruction = /responda|conteste|lea|con base|de acuerdo|siguiente|grafic|tabl|texto|imagen|figura/i.test(contextText);
      const isInformational = /contiene|consta|presenta|evalúa|prueba de/i.test(contextText) && !/responda|conteste/i.test(contextText);
      
      if (isInstruction && !isInformational) {
        const parsedNums = numbers.map(Number);
        const min = Math.min(...parsedNums);
        const max = Math.max(...parsedNums);
        // Safety limit: ICFES reading passages apply to 2-6 questions at most.
        // If range is larger (e.g., 20 questions), it is likely a module description, not a shared reading text.
        if (max - min < 8) {
          for (let n = min; n <= max; n++) {
            excluded.add(n);
          }
        }

      }
    }
  }
  return excluded;
}


/**
 * Parses raw text content and splits it into structured ICFES questions.
 */
function parseQuestionsFromText(text, imagesByPage = {}, fileType = 'pdf') {
  const normalized = normalizeText(text);
  console.log("=== TEXTO EXTRAÍDO DEL DOCUMENTO (PRIMEROS 1000 CARACTERES) ===");
  console.log(normalized.substring(0, 1000));
  console.log("===============================================================");
  
  // Find all question numbers that belong to multi-question contexts globally
  const excludedSet = getExcludedQuestionNumbers(normalized);
  console.log('Números de preguntas excluidos globalmente:', Array.from(excludedSet));

  // Allow whitespace after newline or start of string.
  // We match two formats:
  // 1. "Pregunta X" or "N° X" with optional punctuation at the end.
  // 2. Just a number "X" requiring punctuation (to avoid random number matching).
  const questionSplitRegex = /(?:^|\n)\s*(?:(?:[Pp]regunta\s+|[Nn]°\s*)([1-9][0-9]?|1[0-9]{2})\b\s*[\.\-\)\:]*|([1-9][0-9]?|1[0-9]{2})\s*[\.\-\)\:]+)(?!\d)\s*/g;
  
  let match;
  const matches = [];
  while ((match = questionSplitRegex.exec(normalized)) !== null) {
    matches.push({
      index: match.index,
      number: match[1] || match[2],
      fullMatch: match[0]
    });
  }

  console.log(`Coincidencias encontradas con regex principal: ${matches.length}`);

  if (matches.length === 0) {
    // Fallback: search simple numbers at line starts
    const simpleSplitRegex = /(?:^|\n)\s*([1-9][0-9]?)\s*[\.\-\)\:]+(?!\d)\s*/g;
    while ((match = simpleSplitRegex.exec(normalized)) !== null) {
      matches.push({
        index: match.index,
        number: match[1],
        fullMatch: match[0]
      });
    }
    console.log(`Coincidencias encontradas con regex fallback: ${matches.length}`);
  }

  const parsedQuestions = [];
  let pendingHeader = '';

  for (let i = 0; i < matches.length; i++) {
    const startIdx = matches[i].index;
    const endIdx = (i + 1 < matches.length) ? matches[i + 1].index : normalized.length;
    const blockText = normalized.substring(startIdx, endIdx).trim();

    // The question number
    const qNumber = parseInt(matches[i].number, 10);

    // Skip if it is excluded because it references multiple questions
    if (excludedSet.has(qNumber)) {
      console.log(`Skipped question ${qNumber} because it belongs to a multi-question block.`);
      continue;
    }

    let headerText = pendingHeader;
    pendingHeader = ''; // Reset

    // Let's look if there was some context text before the first question
    if (i === 0 && startIdx > 0) {
      headerText = normalized.substring(0, startIdx).trim();
    }

    // Try to extract options A, B, C, D
    const optionARegex = /(?:^|\s|\n)[Aa][\.\-\)]\s+([\s\S]*?)(?=(?:[Bb][\.\-\)]\s+)|$)/;
    const optionBRegex = /(?:^|\s|\n)[Bb][\.\-\)]\s+([\s\S]*?)(?=(?:[Cc][\.\-\)]\s+)|$)/;
    const optionCRegex = /(?:^|\s|\n)[Cc][\.\-\)]\s+([\s\S]*?)(?=(?:[Dd][\.\-\)]\s+)|$)/;
    const optionDRegex = /(?:^|\s|\n)[Dd][\.\-\)]\s+([\s\S]*?)(?=(?:[A-Ea-e\d][\.\-\)]\s+)|$)/;

    let optionsStartIdx = blockText.search(/(?:^|\s|\n)[Aa][\.\-\)]\s+/);
    let questionBody = blockText;
    let optA = '', optB = '', optC = '', optD = '';

    if (optionsStartIdx !== -1) {
      questionBody = blockText.substring(0, optionsStartIdx).trim();
      const optionsPart = blockText.substring(optionsStartIdx);
      
      const matchA = optionsPart.match(optionARegex);
      const matchB = optionsPart.match(optionBRegex);
      const matchC = optionsPart.match(optionCRegex);
      const matchD = optionsPart.match(optionDRegex);

      if (matchA) optA = matchA[1].trim();
      if (matchB) optB = matchB[1].trim();
      if (matchC) optC = matchC[1].trim();
      if (matchD) optD = matchD[1].trim();
    }

    // Check if Option D contains a header/context block for the NEXT question
    const nextHeaderPattern = /(?:^|\n)(Responda la pregunta \d+|De acuerdo con|Con base en|Lea el siguiente|Texto:?|INFORMACIÓN)\b[\s\S]*/i;
    const headerMatch = optD.match(nextHeaderPattern);
    if (headerMatch) {
      const headerIdx = optD.indexOf(headerMatch[0]);
      pendingHeader = optD.substring(headerIdx).trim();
      optD = optD.substring(0, headerIdx).trim();
    }

    // Clean the questionBody from the number prefix
    let cleanBody = questionBody.replace(/^(?:[Pp]regunta\s+|[Nn]°\s*|)\d+\s*[\.\-\)\:]*\s*/, '').trim();

    // FALLBACK FOR BUBBLE LAYOUT IN ICFES:
    // If options are empty but we have multiple paragraphs in the body, the last 4 paragraphs are the options.
    if (!optA && !optB && !optC && !optD) {
      const paragraphs = cleanBody.split('\n').map(l => l.trim()).filter(l => l.length > 0);
      if (paragraphs.length >= 5) {
        optD = paragraphs.pop();
        optC = paragraphs.pop();
        optB = paragraphs.pop();
        optA = paragraphs.pop();
        cleanBody = paragraphs.join('\n');
      }
    }

    // Double check if body or header references multiple questions
    const combinedText = (headerText + ' ' + cleanBody).toLowerCase();
    const multiPreguntasRegex = /\bpreguntas\s+\d+/i;
    if (multiPreguntasRegex.test(combinedText)) {
      console.log(`Skipped question ${qNumber} during detailed checks.`);
      continue;
    }

    // Extract "Respuesta Correcta: X" from the full block (may fall after optD or outside its match)
    const respCorrectaRegex = /\bRespuesta\s+Correcta\s*:\s*([A-Da-d])\b/i;
    const respMatch = blockText.match(respCorrectaRegex);

    // Also clean it from optD in case it landed there
    if (respMatch) optD = optD.replace(respCorrectaRegex, '').trim();

    // Try to auto-detect correct answer (if marked with * or bold)
    let correctOption = respMatch ? respMatch[1].toUpperCase() : '';
    if (!correctOption) {
      if (optA.startsWith('*') || optA.endsWith('*')) { correctOption = 'A'; optA = optA.replace(/^\*|\*$/g, '').trim(); }
      else if (optB.startsWith('*') || optB.endsWith('*')) { correctOption = 'B'; optB = optB.replace(/^\*|\*$/g, '').trim(); }
      else if (optC.startsWith('*') || optC.endsWith('*')) { correctOption = 'C'; optC = optC.replace(/^\*|\*$/g, '').trim(); }
      else if (optD.startsWith('*') || optD.endsWith('*')) { correctOption = 'D'; optD = optD.replace(/^\*|\*$/g, '').trim(); }
    }

    parsedQuestions.push({
      number: qNumber,
      headerText: headerText,
      bodyText: cleanBody,
      options: {
        A: optA,
        B: optB,
        C: optC,
        D: optD
      },
      correctOption: correctOption,
      images: []
    });
  }


  // Parse trailing answer keys if available
  const answerKeyRegex = /(?:[Cc]lave de [Rr]espuestas|[Tt]abla de [Rr]espuestas|[Rr]espuestas):?\s*([\s\S]*)$/i;
  const keyMatch = normalized.match(answerKeyRegex);
  if (keyMatch) {
    const keySection = keyMatch[1];
    const pairsRegex = /(\d+)\s*[\.\-\:\s]*\s*([A-Da-d])\b/g;
    let pair;
    const answerKeys = {};
    while ((pair = pairsRegex.exec(keySection)) !== null) {
      answerKeys[parseInt(pair[1], 10)] = pair[2].toUpperCase();
    }
    
    parsedQuestions.forEach(q => {
      if (!q.correctOption && answerKeys[q.number]) {
        q.correctOption = answerKeys[q.number];
      }
    });
  }

  return parsedQuestions;
}

/**
 * Extracts questions from a PDF file.
 */
export async function parsePDF(arrayBuffer, progressCallback) {
  if (typeof pdfjsLib === 'undefined') {
    throw new Error('PDF.js library is not loaded. Please check CDN.');
  }

  pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';

  const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
  const numPages = pdf.numPages;
  let fullText = '';
  const pageTexts = [];
  const imagesByPage = {};

  for (let i = 1; i <= numPages; i++) {
    if (progressCallback) {
      progressCallback(Math.round((i / numPages) * 50));
    }
    
    const page = await pdf.getPage(i);
    const textContent = await page.getTextContent();
    
    // Reconstruct lines based on Y coordinate shifts to preserve layouts and split questions correctly
    let lastY;
    let pageText = '';
    for (const item of textContent.items) {
      if (lastY !== undefined && Math.abs(item.transform[5] - lastY) > 4) {
        pageText += '\n';
      }
      pageText += item.str + ' ';
      lastY = item.transform[5];
    }

    pageTexts.push(pageText);
    fullText += '\n' + pageText;

    // Extract images from page
    try {
      const pageImages = await extractImagesFromPDFPage(page);
      if (pageImages && pageImages.length > 0) {
        imagesByPage[i] = pageImages;
      }
    } catch (err) {
      console.warn(`Failed to extract images from page ${i}:`, err);
    }
  }

  if (progressCallback) progressCallback(75);
  const questions = parseQuestionsFromText(fullText, imagesByPage, 'pdf');



  // Match questions to page images
  // Heuristic: If page text contains the question text, associate page's images with it.
  questions.forEach(q => {
    // Find which page this question is on
    for (let pNum = 1; pNum <= numPages; pNum++) {
      const pText = pageTexts[pNum - 1] || '';
      // If the body text is found on this page
      if (pText.includes(q.bodyText.substring(0, Math.min(30, q.bodyText.length)))) {
        if (imagesByPage[pNum]) {
          q.images = q.images.concat(imagesByPage[pNum]);
        }
      }
    }
  });

  if (progressCallback) progressCallback(100);
  return questions;
}

/**
 * Extracts image resources from a PDF.js page object.
 */
async function extractImagesFromPDFPage(page) {
  const operatorList = await page.getOperatorList();
  const images = [];
  
  // We iterate through operations to find images
  for (let i = 0; i < operatorList.fnArray.length; i++) {
    const fn = operatorList.fnArray[i];
    // PaintImageXObject or PaintInlineImageXObject
    if (fn === pdfjsLib.OPS.paintImageXObject || fn === pdfjsLib.OPS.paintInlineImageXObject) {
      const imgKey = operatorList.argsArray[i][0];
      try {
        const imgObj = await new Promise((resolve, reject) => {
          page.objs.get(imgKey, (obj) => {
            if (obj) resolve(obj);
            else reject(new Error('Image object not found'));
          });
        });

        if (imgObj && imgObj.data) {
          // Convert Uint8ClampedArray/Uint8Array to canvas and Base64
          const width = imgObj.width;
          const height = imgObj.height;
          const canvas = document.createElement('canvas');
          canvas.width = width;
          canvas.height = height;
          const ctx = canvas.getContext('2d');
          
          const imgData = ctx.createImageData(width, height);
          
          // PDF.js image data format can vary. Usually it is RGB or RGBA.
          if (imgObj.data.length === width * height * 4) {
            imgData.data.set(imgObj.data);
          } else if (imgObj.data.length === width * height * 3) {
            // RGB format, need to add Alpha
            let srcIdx = 0;
            let destIdx = 0;
            for (let j = 0; j < width * height; j++) {
              imgData.data[destIdx] = imgObj.data[srcIdx];     // R
              imgData.data[destIdx + 1] = imgObj.data[srcIdx + 1]; // G
              imgData.data[destIdx + 2] = imgObj.data[srcIdx + 2]; // B
              imgData.data[destIdx + 3] = 255;                   // A
              srcIdx += 3;
              destIdx += 4;
            }
          } else {
            continue; // Unsupported image format for direct canvas rendering
          }
          
          ctx.putImageData(imgData, 0, 0);
          images.push(canvas.toDataURL('image/png'));
        }
      } catch (err) {
        console.warn('Error reading PDF image object:', err);
      }
    }
  }

  return images;
}

/**
 * Extracts questions from a DOCX file using Mammoth.js.
 */
export async function parseDocx(arrayBuffer, progressCallback) {
  if (typeof mammoth === 'undefined') {
    throw new Error('Mammoth.js library is not loaded. Please check CDN.');
  }

  if (progressCallback) progressCallback(30);

  // Convert docx to HTML to capture layout and images
  const result = await mammoth.convertToHtml({ arrayBuffer: arrayBuffer });
  const html = result.value; // The generated HTML
  
  if (progressCallback) progressCallback(60);

  // Create a temporary DOM parser to inspect the structure
  const domParser = new DOMParser();
  const doc = domParser.parseFromString(html, 'text/html');

  // Build a positional map: for each <img>, find its question number.
  // Rule: images always appear above their question in the document.
  // The number may be in the same <p> as the image, or in the immediately
  // preceding sibling (Mammoth puts "14." in its own <p> then the <img> in the next).
  const imagesByQuestionNumber = {};

  const qNumInText = (text) => {
    // Only match if the number appears at the VERY START of the trimmed text.
    // Requires either a space/end after the separator (prevents matching "12."
    // at the end of a sentence like "...aterriza en el extremo 12.").
    const m = text.trimStart().match(/^(?:[Pp]regunta\s+|[Nn]°\s*)?([1-9][0-9]?)\s*[\.\-\)\:](\s|$)/);
    return m ? parseInt(m[1], 10) : null;
  };

  doc.body.querySelectorAll('img[src^="data:image"]').forEach(imgEl => {
    const src = imgEl.getAttribute('src');

    // Find closest <p> ancestor
    let para = imgEl.parentElement;
    while (para && para.nodeName !== 'P' && para !== doc.body) {
      para = para.parentElement;
    }

    let assignedQ = null;

    if (para && para.nodeName === 'P') {
      // 1. Check the paragraph's own text (covers "4. [img]" in same <p>)
      assignedQ = qNumInText(para.textContent);

      // 2. Walk backwards through siblings (covers "<p>14.</p><p>[img]</p>"
      //    and cases where image is embedded mid-question body text)
      if (assignedQ === null) {
        let prevSib = para.previousElementSibling;
        while (prevSib) {
          assignedQ = qNumInText(prevSib.textContent);
          if (assignedQ !== null) break;
          prevSib = prevSib.previousElementSibling;
        }
      }

      // 3. If still not found, walk forward through siblings (image before question number)
      if (assignedQ === null) {
        let sibling = para.nextElementSibling;
        while (sibling) {
          assignedQ = qNumInText(sibling.textContent);
          if (assignedQ !== null) break;
          sibling = sibling.nextElementSibling;
        }
      }
    }

    if (assignedQ !== null) {
      if (!imagesByQuestionNumber[assignedQ]) imagesByQuestionNumber[assignedQ] = [];
      imagesByQuestionNumber[assignedQ].push(src);
    }
  });

  // Extract text and parse
  const textResult = await mammoth.extractRawText({ arrayBuffer: arrayBuffer });
  const text = textResult.value;

  if (progressCallback) progressCallback(80);
  const questions = parseQuestionsFromText(text, {}, 'docx');

  // Assign images to questions using the positional map built above
  questions.forEach(q => {
    if (imagesByQuestionNumber[q.number] && imagesByQuestionNumber[q.number].length > 0) {
      q.images = q.images.concat(imagesByQuestionNumber[q.number]);
    }
  });

  if (progressCallback) progressCallback(100);
  return questions;
}
