const cors = require('cors');
const crypto = require('node:crypto');
const express = require('express');
const ffmpeg = require('fluent-ffmpeg');
const ffmpegInstaller = require('@ffmpeg-installer/ffmpeg');
const fs = require('node:fs');
const fsp = require('node:fs/promises');
const mammoth = require('mammoth');
const multer = require('multer');
const os = require('node:os');
const path = require('node:path');
const PDFDocument = require('pdfkit');
const pdfParseModule = require('pdf-parse');
const sharp = require('sharp');
const XlsxPopulate = require('xlsx-populate');
const zlib = require('node:zlib');
const { Document, Packer, Paragraph, TextRun } = require('docx');
const WordExtractor = require('word-extractor');

const PDFParse = pdfParseModule.PDFParse;
const wordExtractor = new WordExtractor();

ffmpeg.setFfmpegPath(ffmpegInstaller.path);

const PORT = Number(process.env.CONVERTER_PORT || 8787);
const ROOT = path.resolve(__dirname, '..');
const TMP = path.join(os.tmpdir(), 'simple-file-converter');
const upload = multer({
  dest: TMP,
  limits: { fileSize: 350 * 1024 * 1024 }
});

const app = express();
app.use(cors());
app.use(express.json());

const FORMATS = {
  png: { ext: 'png', mime: 'image/png', family: 'image' },
  jpg: { ext: 'jpg', mime: 'image/jpeg', family: 'image' },
  jpeg: { ext: 'jpg', mime: 'image/jpeg', family: 'image' },
  webp: { ext: 'webp', mime: 'image/webp', family: 'image' },
  avif: { ext: 'avif', mime: 'image/avif', family: 'image' },
  heif: { ext: 'heif', mime: 'image/heif', family: 'image' },
  tiff: { ext: 'tiff', mime: 'image/tiff', family: 'image' },
  gif: { ext: 'gif', mime: 'image/gif', family: 'image' },
  mp3: { ext: 'mp3', mime: 'audio/mpeg', family: 'audio' },
  wav: { ext: 'wav', mime: 'audio/wav', family: 'audio' },
  ogg: { ext: 'ogg', mime: 'audio/ogg', family: 'audio' },
  flac: { ext: 'flac', mime: 'audio/flac', family: 'audio' },
  opus: { ext: 'opus', mime: 'audio/opus', family: 'audio' },
  aac: { ext: 'aac', mime: 'audio/aac', family: 'audio' },
  m4a: { ext: 'm4a', mime: 'audio/mp4', family: 'audio' },
  webm_audio: { ext: 'webm', mime: 'audio/webm', family: 'audio' },
  mp4: { ext: 'mp4', mime: 'video/mp4', family: 'video' },
  webm: { ext: 'webm', mime: 'video/webm', family: 'video' },
  mov: { ext: 'mov', mime: 'video/quicktime', family: 'video' },
  avi: { ext: 'avi', mime: 'video/x-msvideo', family: 'video' },
  mkv: { ext: 'mkv', mime: 'video/x-matroska', family: 'video' },
  gif_video: { ext: 'gif', mime: 'image/gif', family: 'video' },
  txt: { ext: 'txt', mime: 'text/plain', family: 'document' },
  md: { ext: 'md', mime: 'text/markdown', family: 'document' },
  rtf: { ext: 'rtf', mime: 'application/rtf', family: 'document' },
  html: { ext: 'html', mime: 'text/html', family: 'document' },
  pdf: { ext: 'pdf', mime: 'application/pdf', family: 'document' },
  doc: { ext: 'doc', mime: 'application/msword', family: 'document' },
  docx: {
    ext: 'docx',
    mime: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    family: 'document'
  },
  json: { ext: 'json', mime: 'application/json', family: 'data' },
  csv: { ext: 'csv', mime: 'text/csv', family: 'data' },
  tsv: { ext: 'tsv', mime: 'text/tab-separated-values', family: 'data' },
  xlsx: {
    ext: 'xlsx',
    mime: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    family: 'data'
  },
  xml: { ext: 'xml', mime: 'application/xml', family: 'data' },
  yaml: { ext: 'yaml', mime: 'application/yaml', family: 'data' },
  base64: { ext: 'txt', mime: 'text/plain', family: 'data' },
  zip: { ext: 'zip', mime: 'application/zip', family: 'archive' },
  tar: { ext: 'tar', mime: 'application/x-tar', family: 'archive' },
  gzip: { ext: 'gz', mime: 'application/gzip', family: 'archive' },
  tgz: { ext: 'tgz', mime: 'application/gzip', family: 'archive' }
};

app.get('/api/health', (req, res) => {
  res.json({ ok: true, ffmpeg: Boolean(ffmpegInstaller.path), formats: Object.keys(FORMATS) });
});

app.post('/api/convert', upload.single('file'), async (req, res) => {
  const file = req.file;
  const requestedFormat = normalizeFormat(req.body.format);
  const target = FORMATS[requestedFormat];
  const quality = clamp(Number(req.body.quality || 92), 1, 100);
  const maxWidth = clamp(Number(req.body.maxWidth || 1600), 64, 4096);

  if (!file || !target) {
    res.status(400).json({ error: 'Choose a file and a supported output format.' });
    return;
  }

  const jobId = crypto.randomUUID();
  const jobDir = path.join(TMP, jobId);
  const sourcePath = path.join(jobDir, safeName(file.originalname || `source.${extensionFromMime(file.mimetype)}`));
  const outputName = normalizeFileName(req.body.fileName || 'converted-file', target.ext);
  const outputPath = path.join(jobDir, outputName);

  try {
    await fsp.mkdir(jobDir, { recursive: true });
    await fsp.rename(file.path, sourcePath);

    const source = {
      path: sourcePath,
      name: file.originalname || path.basename(sourcePath),
      mime: file.mimetype || '',
      ext: path.extname(file.originalname || '').replace('.', '').toLowerCase()
    };

    await convertToTarget(source, outputPath, requestedFormat, target, { quality, maxWidth, jobDir });

    res.setHeader('Content-Type', target.mime);
    res.setHeader('Content-Disposition', `attachment; filename="${encodeURIComponent(outputName)}"`);
    res.setHeader('X-Output-Name', encodeURIComponent(outputName));
    res.setHeader('X-Output-Mime', target.mime);

    const stream = fs.createReadStream(outputPath);
    stream.pipe(res);
    res.on('finish', () => cleanup(jobDir));
    res.on('close', () => cleanup(jobDir));
  } catch (error) {
    await cleanup(jobDir);
    if (file?.path) await cleanup(file.path);
    res.status(422).json({ error: readableError(error) });
  }
});

app.use(express.static(path.join(ROOT, 'dist')));
app.use((req, res) => {
  const indexPath = path.join(ROOT, 'dist', 'index.html');
  if (fs.existsSync(indexPath)) {
    res.sendFile(indexPath);
  } else {
    res.status(404).send('Build the frontend with npm run build, or run npm run dev for Vite.');
  }
});

app.listen(PORT, async () => {
  await fsp.mkdir(TMP, { recursive: true });
  console.log(`Converter backend listening on http://127.0.0.1:${PORT}`);
});

async function convertToTarget(source, outputPath, format, target, options) {
  if (target.family === 'image') return convertToImage(source, outputPath, format, options);
  if (target.family === 'audio') return convertToAudio(source, outputPath, target.ext);
  if (target.family === 'video') return convertToVideo(source, outputPath, format, target.ext, options);
  if (target.family === 'document') return convertToDocument(source, outputPath, format);
  if (target.family === 'data') return convertToData(source, outputPath, format);
  if (target.family === 'archive') return convertToArchive(source, outputPath, format);
}

async function convertToImage(source, outputPath, format, options) {
  if (isImage(source)) {
    await sharp(source.path, { animated: format === 'gif' })
      .resize({ width: options.maxWidth, withoutEnlargement: true })
      .toFormat(imageOutputFormat(format), imageOutputOptions(format, options.quality))
      .toFile(outputPath);
    return;
  }

  if (isVideo(source)) {
    await screenshotVideo(source.path, outputPath, options.jobDir);
    if (format !== 'png') {
      const converted = path.join(options.jobDir, `image.${FORMATS[format].ext}`);
      await sharp(outputPath).toFormat(imageOutputFormat(format), imageOutputOptions(format, options.quality)).toFile(converted);
      await fsp.rename(converted, outputPath);
    }
    return;
  }

  const text = await extractText(source);
  const svg = textCardSvg(source.name, text || `${source.name}\n${source.mime || source.ext || 'file'}`);
  await sharp(Buffer.from(svg)).toFormat(imageOutputFormat(format), imageOutputOptions(format, options.quality)).toFile(outputPath);
}

function imageOutputFormat(format) {
  if (format === 'jpg' || format === 'jpeg') return 'jpeg';
  return format;
}

function imageOutputOptions(format, quality) {
  if (format === 'heif') {
    return { quality, compression: 'av1' };
  }
  return { quality };
}

async function convertToAudio(source, outputPath, ext) {
  if (!isAudio(source) && !isVideo(source)) {
    throw new Error('Audio output needs audio or video input. Text-to-speech is not bundled.');
  }
  await runFfmpeg(source.path, outputPath, (command) => {
    command.noVideo();
    if (ext === 'mp3') command.audioCodec('libmp3lame');
    if (ext === 'wav') command.audioCodec('pcm_s16le');
    if (ext === 'ogg') command.audioCodec('libvorbis');
    if (ext === 'flac') command.audioCodec('flac');
    if (ext === 'opus') command.audioCodec('libopus');
    if (ext === 'm4a') command.audioCodec('aac');
  });
}

async function convertToVideo(source, outputPath, format, ext, options) {
  if (isVideo(source)) {
    await runFfmpeg(source.path, outputPath, (command) => {
      if (format === 'gif_video') command.outputOptions(['-vf', 'fps=12,scale=960:-1:flags=lanczos']);
    });
    return;
  }

  const poster = path.join(options.jobDir, 'poster.png');
  if (isImage(source)) {
    await sharp(source.path).resize({ width: 1280, height: 720, fit: 'contain', background: '#faf8f2' }).png().toFile(poster);
  } else {
    const text = await extractText(source);
    await sharp(Buffer.from(textCardSvg(source.name, text || source.name))).png().toFile(poster);
  }

  if (isAudio(source)) {
    await audioWithPosterToVideo(poster, source.path, outputPath, ext);
    return;
  }

  await runFfmpegWithInputs([poster], outputPath, (command) => {
    command.inputOptions(['-loop', '1']);
    command.duration(5);
    command.outputOptions(['-pix_fmt', 'yuv420p', '-r', '30']);
    if (format === 'gif_video') command.outputOptions(['-vf', 'fps=12,scale=960:-1:flags=lanczos']);
  });
}

async function convertToDocument(source, outputPath, format) {
  const text = await extractText(source);
  if (format === 'txt') return fsp.writeFile(outputPath, text || `${source.name}\nNo extractable text was found.`, 'utf8');
  if (format === 'md') return fsp.writeFile(outputPath, `# ${source.name}\n\n${text || 'No extractable text was found.'}\n`, 'utf8');
  if (format === 'rtf') return fsp.writeFile(outputPath, textToRtf(source.name, text), 'utf8');
  if (format === 'html') {
    return fsp.writeFile(
      outputPath,
      `<!doctype html><html><head><meta charset="utf-8"><title>${escapeHtml(source.name)}</title><style>body{font:16px/1.55 system-ui;margin:40px;max-width:920px}pre{white-space:pre-wrap}</style></head><body><h1>${escapeHtml(source.name)}</h1><pre>${escapeHtml(text)}</pre></body></html>`,
      'utf8'
    );
  }
  if (format === 'pdf') return writePdf(outputPath, source.name, text);
  if (format === 'doc') return fsp.writeFile(outputPath, wordCompatibleDoc(source.name, text), 'utf8');
  if (format === 'docx') return writeDocx(outputPath, source.name, text);
}

async function convertToData(source, outputPath, format) {
  const rows = await extractRows(source);
  if (format === 'json') return fsp.writeFile(outputPath, JSON.stringify(rows, null, 2), 'utf8');
  if (format === 'csv') return fsp.writeFile(outputPath, rowsToCsv(rows), 'utf8');
  if (format === 'tsv') return fsp.writeFile(outputPath, rowsToDelimited(rows, '\t'), 'utf8');
  if (format === 'xlsx') {
    const workbook = await XlsxPopulate.fromBlankAsync();
    const sheet = workbook.sheet(0).name('Converted');
    const keys = keysForRows(rows);
    sheet.cell(1, 1).value(keys);
    rows.forEach((row, rowIndex) => {
      sheet.cell(rowIndex + 2, 1).value(keys.map((key) => row[key] ?? ''));
    });
    keys.forEach((key, index) => sheet.column(index + 1).width(Math.max(12, Math.min(32, key.length + 4))));
    await workbook.toFileAsync(outputPath);
    return;
  }
  if (format === 'xml') return fsp.writeFile(outputPath, rowsToXml(rows), 'utf8');
  if (format === 'yaml') return fsp.writeFile(outputPath, rowsToYaml(rows), 'utf8');
  if (format === 'base64') {
    const data = await fsp.readFile(source.path);
    return fsp.writeFile(outputPath, data.toString('base64'), 'utf8');
  }
}

async function convertToArchive(source, outputPath, format) {
  if (format === 'gzip') {
    const data = await fsp.readFile(source.path);
    await fsp.writeFile(outputPath, await gzipBuffer(data));
    return;
  }

  const archiver = (await import('archiver')).default;
  await new Promise((resolve, reject) => {
    const output = fs.createWriteStream(outputPath);
    const archive =
      format === 'tar' || format === 'tgz'
        ? archiver('tar', format === 'tgz' ? { gzip: true, gzipOptions: { level: 9 } } : {})
        : archiver('zip', { zlib: { level: 9 } });
    output.on('close', resolve);
    archive.on('error', reject);
    archive.pipe(output);
    archive.file(source.path, { name: source.name });
    archive.finalize();
  });
}

async function extractText(source) {
  if (source.ext === 'doc') {
    const data = await fsp.readFile(source.path);
    if (looksLikeHtml(data)) {
      return htmlToPlainText(data.toString('utf8'));
    }

    try {
      const doc = await wordExtractor.extract(data);
      return doc.getBody();
    } catch (error) {
      if (looksLikeText(data)) return data.toString('utf8');
      throw error;
    }
  }

  if (source.ext === 'docx') {
    const result = await mammoth.extractRawText({ path: source.path });
    return result.value;
  }

  if (source.ext === 'pdf') {
    const data = await fsp.readFile(source.path);
    if (typeof pdfParseModule === 'function') {
      const result = await pdfParseModule(data);
      return result.text || '';
    }
    const parser = new PDFParse({ data });
    try {
      const result = await parser.getText();
      return result.text || '';
    } finally {
      await parser.destroy();
    }
  }

  if (['xlsx', 'xls'].includes(source.ext)) {
    const workbook = await XlsxPopulate.fromFileAsync(source.path);
    return workbook.sheets()
      .map((sheet) => {
        const values = sheet.usedRange()?.value() || [];
        const lines = values.map((row) => row.join(','));
        return `# ${sheet.name}\n${lines.join('\n')}`;
      })
      .join('\n\n');
  }

  if (isText(source) || ['json', 'csv', 'xml', 'html', 'md', 'svg'].includes(source.ext)) {
    return fsp.readFile(source.path, 'utf8');
  }

  if (isImage(source)) {
    const meta = await sharp(source.path).metadata();
    return `${source.name}\n${meta.width || '?'} x ${meta.height || '?'} ${meta.format || 'image'}`;
  }

  return `${source.name}\n${source.mime || source.ext || 'file'}`;
}

async function extractRows(source) {
  if (['xlsx', 'xls'].includes(source.ext)) {
    const workbook = await XlsxPopulate.fromFileAsync(source.path);
    return worksheetToRows(workbook.sheet(0));
  }
  if (source.ext === 'json') {
    const parsed = JSON.parse(await fsp.readFile(source.path, 'utf8'));
    if (Array.isArray(parsed)) return parsed.map(normalizeRow);
    if (parsed && typeof parsed === 'object') return [normalizeRow(parsed)];
  }
  if (source.ext === 'csv') {
    return csvToRows(await fsp.readFile(source.path, 'utf8'));
  }
  const text = await extractText(source);
  return text.split(/\r?\n/).filter(Boolean).map((line, index) => ({ line: index + 1, value: line }));
}

function runFfmpeg(input, output, configure = () => {}) {
  return runFfmpegWithInputs([input], output, configure);
}

function runFfmpegWithInputs(inputs, output, configure = () => {}) {
  return new Promise((resolve, reject) => {
    const command = ffmpeg();
    for (const input of inputs) command.input(input);
    configure(command);
    command.on('end', resolve);
    command.on('error', reject);
    command.save(output);
  });
}

function audioWithPosterToVideo(poster, audio, output, ext) {
  return new Promise((resolve, reject) => {
    const command = ffmpeg()
      .input(poster)
      .inputOptions(['-loop', '1'])
      .input(audio)
      .outputOptions(['-shortest', '-pix_fmt', 'yuv420p', '-r', '30']);

    if (ext === 'mp4') command.videoCodec('libx264').audioCodec('aac');
    if (ext === 'webm') command.videoCodec('libvpx-vp9').audioCodec('libopus');

    command.on('end', resolve);
    command.on('error', reject);
    command.save(output);
  });
}

function screenshotVideo(input, output, folder) {
  return new Promise((resolve, reject) => {
    const filename = `frame-${crypto.randomUUID()}.png`;
    const shotPath = path.join(folder, filename);
    ffmpeg(input)
      .on('end', async () => {
        try {
          await fsp.rename(shotPath, output);
          resolve();
        } catch (error) {
          reject(error);
        }
      })
      .on('error', reject)
      .screenshots({ count: 1, filename, folder });
  });
}

function writePdf(outputPath, title, text) {
  return new Promise((resolve, reject) => {
    const doc = new PDFDocument({ margin: 54 });
    const stream = fs.createWriteStream(outputPath);
    stream.on('finish', resolve);
    stream.on('error', reject);
    doc.pipe(stream);
    doc.fontSize(22).text(title, { underline: true });
    doc.moveDown();
    doc.fontSize(11).text(text || 'No extractable text was found.', { lineGap: 4 });
    doc.end();
  });
}

async function writeDocx(outputPath, title, text) {
  const doc = new Document({
    sections: [
      {
        children: [
          new Paragraph({ children: [new TextRun({ text: title, bold: true, size: 32 })] }),
          ...String(text || 'No extractable text was found.')
            .split(/\r?\n/)
            .map((line) => new Paragraph({ children: [new TextRun(line)] }))
        ]
      }
    ]
  });
  await fsp.writeFile(outputPath, await Packer.toBuffer(doc));
}

function wordCompatibleDoc(title, text) {
  const body = text || 'No extractable text was found.';
  return `<!doctype html>
<html xmlns:o="urn:schemas-microsoft-com:office:office"
  xmlns:w="urn:schemas-microsoft-com:office:word"
  xmlns="http://www.w3.org/TR/REC-html40">
<head>
  <meta charset="utf-8">
  <meta name="ProgId" content="Word.Document">
  <title>${escapeHtml(title)}</title>
  <style>
    body { font-family: Arial, sans-serif; font-size: 12pt; line-height: 1.45; }
    h1 { font-size: 18pt; margin-bottom: 18pt; }
    pre { white-space: pre-wrap; font-family: Arial, sans-serif; }
  </style>
</head>
<body>
  <h1>${escapeHtml(title)}</h1>
  <pre>${escapeHtml(body)}</pre>
</body>
</html>`;
}

function looksLikeHtml(data) {
  const sample = data.subarray(0, 4096).toString('utf8').trimStart().toLowerCase();
  return sample.startsWith('<!doctype') || sample.startsWith('<html') || sample.includes('progid" content="word.document"');
}

function looksLikeText(data) {
  const sample = data.subarray(0, Math.min(data.length, 4096));
  if (sample.includes(0)) return false;
  let printable = 0;
  for (const byte of sample) {
    if (byte === 9 || byte === 10 || byte === 13 || (byte >= 32 && byte <= 126) || byte >= 128) printable += 1;
  }
  return printable / Math.max(sample.length, 1) > 0.85;
}

function htmlToPlainText(html) {
  return decodeHtmlEntities(
    String(html)
      .replace(/<script[\s\S]*?<\/script>/gi, '')
      .replace(/<style[\s\S]*?<\/style>/gi, '')
      .replace(/<br\s*\/?>/gi, '\n')
      .replace(/<\/(p|div|h[1-6]|li|tr|pre)>/gi, '\n')
      .replace(/<[^>]+>/g, ' ')
      .replace(/[ \t]+/g, ' ')
      .replace(/\n\s+/g, '\n')
      .trim()
  );
}

function decodeHtmlEntities(value) {
  return String(value)
    .replace(/&nbsp;/g, ' ')
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&#039;/g, "'");
}

function textCardSvg(title, text) {
  const lines = String(text).replace(/\s+/g, ' ').slice(0, 500).match(/.{1,58}(\s|$)/g) || ['Converted file'];
  return `<svg xmlns="http://www.w3.org/2000/svg" width="1280" height="720" viewBox="0 0 1280 720">
    <rect width="1280" height="720" fill="#faf8f2"/>
    <rect x="44" y="44" width="1192" height="632" rx="12" fill="#fffdf8" stroke="#111" stroke-width="6"/>
    <text x="94" y="132" font-family="Arial, sans-serif" font-size="56" font-weight="900" fill="#111">${escapeHtml(title)}</text>
    ${lines.slice(0, 10).map((line, index) => `<text x="94" y="${220 + index * 42}" font-family="Arial, sans-serif" font-size="28" fill="#333">${escapeHtml(line.trim())}</text>`).join('')}
  </svg>`;
}

function rowsToCsv(rows) {
  const keys = keysForRows(rows);
  return [keys, ...rows.map((row) => keys.map((key) => row[key] ?? ''))].map((row) => row.map(csvEscape).join(',')).join('\n');
}

function rowsToDelimited(rows, delimiter) {
  const keys = keysForRows(rows);
  return [keys, ...rows.map((row) => keys.map((key) => row[key] ?? ''))]
    .map((row) => row.map((value) => String(value ?? '').replace(/\r?\n/g, ' ')).join(delimiter))
    .join('\n');
}

function keysForRows(rows) {
  return Array.from(
    rows.reduce((set, row) => {
      Object.keys(row).forEach((key) => set.add(key));
      return set;
    }, new Set())
  );
}

function worksheetToRows(sheet) {
  if (!sheet) return [];
  const values = sheet.usedRange()?.value() || [];
  const headers = (values.shift() || ['value']).map((value, index) => String(value || `column_${index + 1}`));
  return values.map((row) => Object.fromEntries(headers.map((header, index) => [header, row[index] ?? ''])));
}

function csvToRows(text) {
  const parsed = text.split(/\r?\n/).filter(Boolean).map(parseCsvLine);
  const headers = parsed.shift() || ['value'];
  return parsed.map((row) => Object.fromEntries(headers.map((header, index) => [header || `column_${index + 1}`, row[index] ?? ''])));
}

function parseCsvLine(line) {
  const values = [];
  let current = '';
  let quoted = false;
  for (let index = 0; index < line.length; index += 1) {
    const char = line[index];
    if (char === '"' && line[index + 1] === '"') {
      current += '"';
      index += 1;
    } else if (char === '"') {
      quoted = !quoted;
    } else if (char === ',' && !quoted) {
      values.push(current);
      current = '';
    } else {
      current += char;
    }
  }
  values.push(current);
  return values;
}

function rowsToXml(rows) {
  return `<?xml version="1.0" encoding="UTF-8"?>\n<rows>\n${rows
    .map((row) => `  <row>\n${Object.entries(row).map(([key, value]) => `    <${safeXmlName(key)}>${escapeHtml(value)}</${safeXmlName(key)}>`).join('\n')}\n  </row>`)
    .join('\n')}\n</rows>\n`;
}

function rowsToYaml(rows) {
  return rows
    .map((row) => {
      const body = Object.entries(row)
        .map(([key, value]) => `  ${safeYamlKey(key)}: ${yamlValue(value)}`)
        .join('\n');
      return `- ${body.trimStart()}`;
    })
    .join('\n');
}

function textToRtf(title, text) {
  const escapeRtf = (value) =>
    String(value || '')
      .replace(/\\/g, '\\\\')
      .replace(/{/g, '\\{')
      .replace(/}/g, '\\}')
      .replace(/\r?\n/g, '\\par\n');
  return `{\\rtf1\\ansi\\deff0{\\fonttbl{\\f0 Arial;}}\\fs28\\b ${escapeRtf(title)}\\b0\\par\\fs22 ${escapeRtf(
    text || 'No extractable text was found.'
  )}}`;
}

function gzipBuffer(data) {
  return new Promise((resolve, reject) => {
    zlib.gzip(data, { level: 9 }, (error, buffer) => {
      if (error) reject(error);
      else resolve(buffer);
    });
  });
}

function normalizeRow(value) {
  return value && typeof value === 'object' ? value : { value };
}

function normalizeFormat(format = '') {
  return String(format).replace(/-/g, '_').toLowerCase();
}

function normalizeFileName(value, extension) {
  const clean = safeName(String(value || '').trim() || `converted-file.${extension}`);
  return clean.toLowerCase().endsWith(`.${extension}`) ? clean : `${clean.replace(/\.[^.]+$/, '')}.${extension}`;
}

function safeName(value) {
  return String(value).replace(/[\\/:*?"<>|]+/g, '-').replace(/\s+/g, ' ').trim().slice(0, 180) || 'file';
}

function extensionFromMime(mime = '') {
  return mime.includes('/') ? mime.split('/').pop().replace('jpeg', 'jpg') : 'bin';
}

function readableError(error) {
  const message = error?.message || String(error);
  if (message.includes('Invalid data found')) return 'That file could not be decoded for the selected output.';
  if (message.includes('Input buffer contains unsupported image format')) return 'That file is not an image the converter can decode.';
  return message.split('\n')[0];
}

function isImage(source) {
  return source.mime.startsWith('image/') || ['png', 'jpg', 'jpeg', 'webp', 'avif', 'tiff', 'gif', 'bmp', 'svg'].includes(source.ext);
}

function isAudio(source) {
  return source.mime.startsWith('audio/') || ['mp3', 'wav', 'ogg', 'flac', 'aac', 'm4a', 'webm'].includes(source.ext);
}

function isVideo(source) {
  return source.mime.startsWith('video/') || ['mp4', 'webm', 'mov', 'avi', 'mkv', 'm4v'].includes(source.ext);
}

function isText(source) {
  return source.mime.startsWith('text/') || ['txt', 'md', 'log', 'js', 'ts', 'css'].includes(source.ext);
}

function clamp(value, min, max) {
  if (Number.isNaN(value)) return min;
  return Math.max(min, Math.min(max, value));
}

function csvEscape(value) {
  const text = String(value ?? '');
  return /[",\n]/.test(text) ? `"${text.replace(/"/g, '""')}"` : text;
}

function safeYamlKey(value) {
  const key = String(value || 'field').replace(/[^a-zA-Z0-9_-]/g, '_');
  return /^[a-zA-Z_]/.test(key) ? key : `field_${key}`;
}

function yamlValue(value) {
  if (value == null) return 'null';
  const text = String(value);
  if (/^[a-zA-Z0-9_.-]+$/.test(text)) return text;
  return JSON.stringify(text);
}

function escapeHtml(value) {
  return String(value ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

function safeXmlName(value) {
  const name = String(value || 'field').replace(/[^a-zA-Z0-9_-]/g, '_');
  return /^[a-zA-Z_]/.test(name) ? name : `field_${name}`;
}

async function cleanup(target) {
  try {
    await fsp.rm(target, { recursive: true, force: true });
  } catch {
    // Best-effort temp cleanup.
  }
}
