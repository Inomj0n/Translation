const fs = require('fs');
const os = require('os');
const path = require('path');
const { execFile } = require('child_process');
const { promisify } = require('util');
const express = require('express');
const multer = require('multer');
const pdfParse = require('pdf-parse');
const JSZip = require('jszip');
const { PDFDocument, StandardFonts } = require('pdf-lib');
const { transliterate } = require('./translit');

const execFileAsync = promisify(execFile);

const app = express();
const PORT = 3000;

const uploadsDir = path.join(__dirname, 'uploads');
const outputDir = path.join(__dirname, 'output');
const cleanupIntervalMs = 30 * 60 * 1000;
const maxFileAgeMs = 24 * 60 * 60 * 1000;
const conversionLinks = new Map();
const sofficeCandidates = [
  'soffice',
  'C:\\Program Files\\LibreOffice\\program\\soffice.exe',
  'C:\\Program Files (x86)\\LibreOffice\\program\\soffice.exe'
];

function normalizeFilename(name) {
  try {
    return Buffer.from(name, 'latin1').toString('utf8');
  } catch (error) {
    return name;
  }
}

if (!fs.existsSync(uploadsDir)) fs.mkdirSync(uploadsDir);
if (!fs.existsSync(outputDir)) fs.mkdirSync(outputDir);

const storage = multer.diskStorage({
  destination: uploadsDir,
  filename: (req, file, cb) => {
    const originalName = normalizeFilename(file.originalname);
    const uniqueName = `${Date.now()}-${originalName}`;
    cb(null, uniqueName);
  }
});

const upload = multer({
  storage,
  fileFilter: (req, file, cb) => {
    const originalName = normalizeFilename(file.originalname);
    const ext = path.extname(originalName).toLowerCase();
    const allowed = ['.pdf', '.docx', '.txt', '.ppt', '.pptx'];

    if (!allowed.includes(ext)) {
      return cb(new Error('Разрешены только PDF, DOCX, TXT, PPT или PPTX файлы.'));
    }

    cb(null, true);
  }
});

app.use((req, res, next) => {
  res.setHeader('Cache-Control', 'no-store');
  next();
});
app.use(express.static(path.join(__dirname, 'public')));

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function logError(error, context = {}) {
  const time = new Date().toISOString();
  const contextText = Object.keys(context).length ? ` | context=${JSON.stringify(context)}` : '';
  const message = `[${time}] ${error.stack || error.message || error}${contextText}${os.EOL}`;
  fs.appendFileSync(errorLogPath, message, 'utf8');
}

function cleanupDirectory(dirPath, maxAgeMsValue) {
  const now = Date.now();
  const entries = fs.readdirSync(dirPath, { withFileTypes: true });

  for (const entry of entries) {
    if (!entry.isFile()) continue;

    const filePath = path.join(dirPath, entry.name);
    const stat = fs.statSync(filePath);
    const isOld = now - stat.mtimeMs > maxAgeMsValue;

    if (isOld) {
      fs.unlinkSync(filePath);
    }
  }
}

function startCleanupJob() {
  const runCleanup = () => {
    try {
      cleanupDirectory(uploadsDir, maxFileAgeMs);
      cleanupDirectory(outputDir, maxFileAgeMs);
    } catch (error) {
      logError(error, { scope: 'cleanup-job' });
    }
  };

  runCleanup();
  setInterval(runCleanup, cleanupIntervalMs);
}

function computeMinimumWaitMs(fileSizeBytes) {
  const minBaseMs = 2000 + Math.floor(Math.random() * 1000); // 2-3 seconds base
  const sizeMb = fileSizeBytes / (1024 * 1024);
  const extraBySizeMs = Math.min(15000, Math.floor(sizeMb * 800)); // larger files wait longer
  return minBaseMs + extraBySizeMs;
}

function decodeXmlEntities(value) {
  return value
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, '&');
}

function encodeXmlEntities(value) {
  return value
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

function transliterateXmlContent(xmlText, tagRegex) {
  return xmlText.replace(tagRegex, (match, openTag, value, closeTag) => {
    const decoded = decodeXmlEntities(value);
    const transliterated = transliterate(decoded);
    const encoded = encodeXmlEntities(transliterated);
    return `${openTag}${encoded}${closeTag}`;
  });
}

async function processDocx(filePath) {
  const zip = await JSZip.loadAsync(fs.readFileSync(filePath));

  const files = Object.keys(zip.files).filter(
    (name) => name.startsWith('word/') && name.endsWith('.xml')
  );

  for (const fileName of files) {
    const xml = await zip.file(fileName).async('string');
    const updated = transliterateXmlContent(xml, /(<w:t[^>]*>)([\s\S]*?)(<\/w:t>)/g);
    zip.file(fileName, updated);
  }

  return zip.generateAsync({ type: 'nodebuffer' });
}

async function processPptx(filePath) {
  const zip = await JSZip.loadAsync(fs.readFileSync(filePath));

  const files = Object.keys(zip.files).filter(
    (name) => name.startsWith('ppt/') && name.endsWith('.xml')
  );

  for (const fileName of files) {
    const xml = await zip.file(fileName).async('string');
    const updated = transliterateXmlContent(xml, /(<a:t[^>]*>)([\s\S]*?)(<\/a:t>)/g);
    zip.file(fileName, updated);
  }

  return zip.generateAsync({ type: 'nodebuffer' });
}

function splitTextToLines(text, maxChars) {
  const lines = [];
  const paragraphs = text.split(/\r?\n/);

  for (const paragraph of paragraphs) {
    if (!paragraph.trim()) {
      lines.push('');
      continue;
    }

    let current = '';
    for (const word of paragraph.split(/\s+/)) {
      const candidate = current ? `${current} ${word}` : word;
      if (candidate.length <= maxChars) {
        current = candidate;
      } else {
        if (current) lines.push(current);
        current = word;
      }
    }

    if (current) lines.push(current);
  }

  return lines;
}

async function processPdf(filePath) {
  const dataBuffer = fs.readFileSync(filePath);
  const parsed = await pdfParse(dataBuffer);
  const transliteratedText = transliterate(parsed.text || '');

  const pdfDoc = await PDFDocument.create();
  const font = await pdfDoc.embedFont(StandardFonts.Helvetica);

  const fontSize = 12;
  const pageWidth = 595;
  const pageHeight = 842;
  const margin = 50;
  const lineHeight = 16;
  const maxChars = 90;
  const lines = splitTextToLines(transliteratedText, maxChars);

  let page = pdfDoc.addPage([pageWidth, pageHeight]);
  let y = pageHeight - margin;

  for (const line of lines) {
    if (y < margin) {
      page = pdfDoc.addPage([pageWidth, pageHeight]);
      y = pageHeight - margin;
    }

    page.drawText(line, {
      x: margin,
      y,
      size: fontSize,
      font
    });

    y -= lineHeight;
  }

  return Buffer.from(await pdfDoc.save());
}

async function runLibreOfficeConvert(inputPath, targetExt, outDir) {
  const args = ['--headless', '--convert-to', targetExt, '--outdir', outDir, inputPath];
  let lastError;

  for (const candidate of sofficeCandidates) {
    try {
      await execFileAsync(candidate, args);
      return;
    } catch (error) {
      lastError = error;
    }
  }

  throw lastError || new Error('LibreOffice command not found.');
}

async function processPpt(filePath) {
  const tempDir = fs.mkdtempSync(path.join(os.tmpdir(), 'translit-'));

  try {
    const baseName = path.parse(filePath).name;

    await runLibreOfficeConvert(filePath, 'pptx', tempDir);
    const convertedPptxPath = path.join(tempDir, `${baseName}.pptx`);

    if (!fs.existsSync(convertedPptxPath)) {
      throw new Error('Не удалось конвертировать PPT в PPTX.');
    }

    const processedPptxBuffer = await processPptx(convertedPptxPath);
    const processedPptxPath = path.join(tempDir, `${baseName}-latin.pptx`);
    fs.writeFileSync(processedPptxPath, processedPptxBuffer);

    await runLibreOfficeConvert(processedPptxPath, 'ppt', tempDir);
    const finalPptPath = path.join(tempDir, `${baseName}-latin.ppt`);

    if (!fs.existsSync(finalPptPath)) {
      throw new Error('Не удалось собрать итоговый PPT.');
    }

    return fs.readFileSync(finalPptPath);
  } catch (error) {
    if (error.code === 'ENOENT') {
      throw new Error('Для обработки .ppt установите LibreOffice (команда soffice должна быть доступна в PATH).');
    }

    throw error;
  }
}

async function processFile(filePath, ext) {
  if (ext === '.txt') {
    const source = fs.readFileSync(filePath, 'utf8');
    return Buffer.from(transliterate(source), 'utf8');
  }

  if (ext === '.docx') {
    return processDocx(filePath);
  }

  if (ext === '.pptx') {
    return processPptx(filePath);
  }

  if (ext === '.pdf') {
    return processPdf(filePath);
  }

  if (ext === '.ppt') {
    return processPpt(filePath);
  }

  throw new Error('Неподдерживаемый формат файла.');
}

app.post('/convert', upload.single('file'), async (req, res) => {
  const requestStartedAt = Date.now();

  try {
    if (!req.file) {
      return res.status(400).json({ error: 'Файл не загружен.' });
    }

    const uploadedPath = req.file.path;
    const originalName = normalizeFilename(req.file.originalname);
    const originalExt = path.extname(originalName).toLowerCase();
    const originalBase = path.parse(originalName).name;

    const outputBuffer = await processFile(uploadedPath, originalExt);
    const internalOutputName = `${Date.now()}-${originalBase}-latin${originalExt}`;
    const outputPath = path.join(outputDir, internalOutputName);

    fs.writeFileSync(outputPath, outputBuffer);
    conversionLinks.set(internalOutputName, {
      uploadedPath,
      downloadName: originalName
    });

    const minWaitMs = computeMinimumWaitMs(req.file.size || 0);
    const elapsedMs = Date.now() - requestStartedAt;
    const remainingMs = Math.max(0, minWaitMs - elapsedMs);

    if (remainingMs > 0) {
      await sleep(remainingMs);
    }

    return res.json({
      message: 'Файл успешно обработан.',
      outputName: originalName,
      downloadUrl: `/download/${path.basename(outputPath)}`
    });
  } catch (error) {
    logError(error, {
      scope: 'convert',
      originalName: req.file ? normalizeFilename(req.file.originalname) : null
    });
    return res.status(500).json({ error: error.message || 'Ошибка обработки файла.' });
  }
});

app.get('/download/:filename', (req, res) => {
  const internalName = req.params.filename;
  const filePath = path.join(outputDir, internalName);

  if (!fs.existsSync(filePath)) {
    return res.status(404).send('Файл не найден.');
  }

  const link = conversionLinks.get(internalName);
  const downloadName = link && link.downloadName ? link.downloadName : internalName;

  return res.download(filePath, downloadName, (downloadError) => {
    conversionLinks.delete(internalName);

    try {
      if (fs.existsSync(filePath)) {
        fs.unlinkSync(filePath);
      }
    } catch (error) {
      logError(error, { scope: 'download-output-cleanup', filename: internalName });
    }

    try {
      if (link && link.uploadedPath && fs.existsSync(link.uploadedPath)) {
        fs.unlinkSync(link.uploadedPath);
      }
    } catch (error) {
      logError(error, { scope: 'download-upload-cleanup', filename: internalName });
    }

    if (downloadError) {
      logError(downloadError, { scope: 'download-send', filename: internalName });
    }
  });
});

app.use((err, req, res, next) => {
  if (err instanceof multer.MulterError || err.message) {
    logError(err, { scope: 'multer-or-request' });
    return res.status(400).json({ error: err.message });
  }

  return next(err);
});

app.use((err, req, res, next) => {
  logError(err, { scope: 'unhandled-express' });
  return res.status(500).json({ error: 'Внутренняя ошибка сервера.' });
});

startCleanupJob();
app.listen(PORT, () => {
  console.log(`Server started: http://localhost:${PORT}`);
});
