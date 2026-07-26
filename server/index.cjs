const cors = require('cors');
const crypto = require('node:crypto');
const cheerio = require('cheerio');
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
const sharp = require('sharp');
const XlsxPopulate = require('xlsx-populate');
const zlib = require('node:zlib');
const {
  Document,
  ExternalHyperlink,
  HeadingLevel,
  ImageRun,
  Packer,
  Paragraph,
  Table,
  TableCell,
  TableRow,
  TextRun,
  WidthType
} = require('docx');
const WordExtractor = require('word-extractor');

const wordExtractor = new WordExtractor();

ffmpeg.setFfmpegPath(ffmpegInstaller.path);

const PORT = Number(process.env.CONVERTER_PORT || 8787);
const ROOT = path.resolve(__dirname, '..');
const TMP = path.join(os.tmpdir(), 'simple-file-converter');
fs.mkdirSync(TMP, { recursive: true });
const upload = multer({
  dest: TMP,
  limits: { fileSize: (process.env.VERCEL ? 4 : 350) * 1024 * 1024 }
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

if (require.main === module) {
  app.listen(PORT, () => {
    console.log(`Converter backend listening on http://127.0.0.1:${PORT}`);
  });
}

module.exports = app;

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
  if (source.ext === FORMATS[format].ext) {
    return fsp.copyFile(source.path, outputPath);
  }

  const structured = await extractStructuredDocument(source);
  const text = structured.text || `${source.name}\nNo extractable text was found.`;
  if (format === 'txt') return fsp.writeFile(outputPath, text, 'utf8');
  if (format === 'md') return fsp.writeFile(outputPath, structuredHtmlToMarkdown(source.name, structured.html), 'utf8');
  if (format === 'rtf') return fsp.writeFile(outputPath, textToRtf(source.name, text), 'utf8');
  if (format === 'html') {
    return fsp.writeFile(outputPath, standaloneHtmlDocument(source.name, structured.html), 'utf8');
  }
  if (format === 'pdf') return writeStructuredPdf(outputPath, source.name, structured.html);
  if (format === 'doc') return fsp.writeFile(outputPath, wordCompatibleDoc(source.name, structured.html), 'utf8');
  if (format === 'docx') return writeStructuredDocx(outputPath, source.name, structured.html);
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

  const { TarArchive, ZipArchive } = await import('archiver');
  await new Promise((resolve, reject) => {
    const output = fs.createWriteStream(outputPath);
    const archive =
      format === 'tar' || format === 'tgz'
        ? new TarArchive(format === 'tgz' ? { gzip: true, gzipOptions: { level: 9 } } : {})
        : new ZipArchive({ zlib: { level: 9 } });
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
    require('@napi-rs/canvas');
    const pdfParseModule = require('pdf-parse');
    const data = await fsp.readFile(source.path);
    if (typeof pdfParseModule === 'function') {
      const result = await pdfParseModule(data);
      return result.text || '';
    }
    const PDFParse = pdfParseModule.PDFParse;
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

async function extractStructuredDocument(source) {
  if (source.ext === 'docx') {
    const result = await mammoth.convertToHtml(
      { path: source.path },
      {
        convertImage: mammoth.images.imgElement(async (image) => ({
          src: `data:${image.contentType};base64,${await image.read('base64')}`
        }))
      }
    );
    const html = sanitizeStructuredHtml(result.value);
    return { html, text: structuredHtmlToText(html) };
  }

  if (source.ext === 'html' || source.ext === 'htm') {
    const html = sanitizeStructuredHtml(await fsp.readFile(source.path, 'utf8'));
    return { html, text: structuredHtmlToText(html) };
  }

  if (source.ext === 'doc') {
    const data = await fsp.readFile(source.path);
    if (looksLikeHtml(data)) {
      const html = sanitizeStructuredHtml(data.toString('utf8'));
      return { html, text: structuredHtmlToText(html) };
    }
  }

  if (['xlsx', 'xls', 'csv', 'tsv', 'json'].includes(source.ext)) {
    const rows = await extractRows(source);
    const html = rowsToStructuredHtml(rows);
    return { html, text: structuredHtmlToText(html) };
  }

  if (isImage(source)) {
    const data = await fsp.readFile(source.path);
    const normalized = await normalizeEmbeddedImage(data);
    const html = `<figure><img src="data:image/png;base64,${normalized.buffer.toString('base64')}" alt="${escapeHtml(
      source.name
    )}"><figcaption>${escapeHtml(source.name)}</figcaption></figure>`;
    return { html, text: source.name };
  }

  const text = await extractText(source);
  return { html: textToStructuredHtml(text), text };
}

function sanitizeStructuredHtml(html) {
  const $ = cheerio.load(String(html || ''));
  $('script,style,iframe,object,embed,form,meta,link').remove();
  $('*').each((_, element) => {
    for (const attribute of Object.keys(element.attribs || {})) {
      if (attribute.toLowerCase().startsWith('on') || attribute === 'style') {
        $(element).removeAttr(attribute);
      }
    }
  });
  $('a').each((_, element) => {
    const href = String($(element).attr('href') || '');
    if (!/^(https?:|mailto:|#)/i.test(href)) $(element).removeAttr('href');
  });
  $('img').each((_, element) => {
    const src = String($(element).attr('src') || '');
    if (!/^data:image\/[a-z0-9.+-]+;base64,/i.test(src)) $(element).remove();
  });
  return $('body').html() || '';
}

function structuredHtmlToText(html) {
  const $ = cheerio.load(`<body>${html || ''}</body>`);
  $('table').each((_, table) => {
    const rows = $(table)
      .find('tr')
      .toArray()
      .map((row) =>
        $(row)
          .children('th,td')
          .toArray()
          .map((cell) => $(cell).text().replace(/\s+/g, ' ').trim())
          .join('\t')
      )
      .join('\n');
    $(table).replaceWith(`<pre>${escapeHtml(rows)}</pre>`);
  });
  $('br').replaceWith('\n');
  $('li').each((_, element) => $(element).prepend('- ').append('\n'));
  $('p,h1,h2,h3,h4,h5,h6,pre,figure,blockquote').each((_, element) => $(element).append('\n'));
  return $('body')
    .text()
    .replace(/\u00a0/g, ' ')
    .replace(/[ \t]+\n/g, '\n')
    .replace(/\n{3,}/g, '\n\n')
    .trim();
}

function textToStructuredHtml(text) {
  return String(text || 'No extractable text was found.')
    .split(/\r?\n\r?\n/)
    .map((block) => `<p>${escapeHtml(block).replace(/\r?\n/g, '<br>')}</p>`)
    .join('');
}

function rowsToStructuredHtml(rows) {
  const keys = keysForRows(rows);
  if (!keys.length) return '<p>No table data was found.</p>';
  const header = keys.map((key) => `<th scope="col">${escapeHtml(key)}</th>`).join('');
  const body = rows
    .map((row) => `<tr>${keys.map((key) => `<td>${escapeHtml(row[key] ?? '')}</td>`).join('')}</tr>`)
    .join('');
  return `<table><thead><tr>${header}</tr></thead><tbody>${body}</tbody></table>`;
}

function standaloneHtmlDocument(title, body) {
  return `<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width,initial-scale=1">
  <title>${escapeHtml(title)}</title>
  <style>
    body{font:16px/1.55 Arial,sans-serif;color:#171717;margin:40px auto;max-width:920px;padding:0 24px}
    h1,h2,h3,h4,h5,h6{line-height:1.2} table{border-collapse:collapse;width:100%;margin:20px 0}
    th,td{border:1px solid #999;padding:8px 10px;text-align:left;vertical-align:top} th{background:#eee}
    img{display:block;max-width:100%;height:auto;margin:18px auto} figcaption{text-align:center;color:#555}
    a{color:#075dcc} pre{white-space:pre-wrap} blockquote{border-left:4px solid #bbb;margin-left:0;padding-left:16px}
  </style>
</head>
<body>
  <h1>${escapeHtml(title)}</h1>
  ${body || '<p>No extractable content was found.</p>'}
</body>
</html>`;
}

function structuredHtmlToMarkdown(title, html) {
  const $ = cheerio.load(`<body>${html || ''}</body>`);
  const blocks = [`# ${title}`];
  for (const element of $('body').children().toArray()) {
    const tag = String(element.tagName || '').toLowerCase();
    if (/^h[1-6]$/.test(tag)) {
      blocks.push(`${'#'.repeat(Number(tag.slice(1)) + 1)} ${markdownInline($, element)}`);
      continue;
    }
    if (tag === 'table') {
      const rows = $(element)
        .find('tr')
        .toArray()
        .map((row) =>
          $(row)
            .children('th,td')
            .toArray()
            .map((cell) => $(cell).text().replace(/\|/g, '\\|').replace(/\s+/g, ' ').trim())
        );
      if (rows.length) {
        const width = Math.max(...rows.map((row) => row.length));
        const normalized = rows.map((row) => [...row, ...Array(Math.max(0, width - row.length)).fill('')]);
        blocks.push(
          [
            `| ${normalized[0].join(' | ')} |`,
            `| ${normalized[0].map(() => '---').join(' | ')} |`,
            ...normalized.slice(1).map((row) => `| ${row.join(' | ')} |`)
          ].join('\n')
        );
      }
      continue;
    }
    if (tag === 'ul' || tag === 'ol') {
      blocks.push(
        $(element)
          .children('li')
          .map((index, item) => `${tag === 'ol' ? `${index + 1}.` : '-'} ${markdownInline($, item)}`)
          .get()
          .join('\n')
      );
      continue;
    }
    if (tag === 'figure' || tag === 'img') {
      const image = tag === 'img' ? $(element) : $(element).find('img').first();
      if (image.attr('src')) blocks.push(`![${image.attr('alt') || title}](${image.attr('src')})`);
      const caption = $(element).find('figcaption').text().trim();
      if (caption) blocks.push(`_${caption}_`);
      continue;
    }
    const content = markdownInline($, element);
    if (content) blocks.push(tag === 'blockquote' ? `> ${content}` : content);
  }
  return `${blocks.filter(Boolean).join('\n\n')}\n`;
}

function markdownInline($, element, trim = true) {
  const value = $(element)
    .contents()
    .map((_, node) => {
      if (node.type === 'text') return node.data;
      const tag = String(node.tagName || '').toLowerCase();
      if (tag === 'br') return '  \n';
      const content = markdownInline($, node, false);
      const trailingSpace = content.match(/\s+$/)?.[0] || '';
      if (tag === 'strong' || tag === 'b') return `**${content.trim()}**${trailingSpace}`;
      if (tag === 'em' || tag === 'i') return `_${content.trim()}_${trailingSpace}`;
      if (tag === 'a' && $(node).attr('href')) return `[${content}](${$(node).attr('href')})`;
      return content;
    })
    .get()
    .join('')
    .replace(/[ \t]+/g, ' ');
  return trim ? value.trim() : value;
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
  if (source.ext === 'tsv') {
    const parsed = (await fsp.readFile(source.path, 'utf8'))
      .split(/\r?\n/)
      .filter(Boolean)
      .map((line) => line.split('\t'));
    const headers = parsed.shift() || ['value'];
    return parsed.map((row) =>
      Object.fromEntries(headers.map((header, index) => [header || `column_${index + 1}`, row[index] ?? '']))
    );
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

function writeStructuredPdf(outputPath, title, html) {
  return new Promise(async (resolve, reject) => {
    const doc = new PDFDocument({ margin: 54 });
    const stream = fs.createWriteStream(outputPath);
    stream.on('finish', resolve);
    stream.on('error', reject);
    doc.pipe(stream);
    try {
      doc.font('Helvetica-Bold').fontSize(22).text(title);
      doc.moveDown(0.8);
      const $ = cheerio.load(`<body>${html || '<p>No extractable content was found.</p>'}</body>`);
      for (const element of $('body').children().toArray()) {
        await appendPdfElement(doc, $, element);
      }
      doc.end();
    } catch (error) {
      doc.end();
      reject(error);
    }
  });
}

async function appendPdfElement(doc, $, element) {
  const tag = String(element.tagName || '').toLowerCase();
  if (/^h[1-6]$/.test(tag)) {
    ensurePdfSpace(doc, 44);
    const level = Number(tag.slice(1));
    doc.font('Helvetica-Bold').fontSize(Math.max(12, 22 - level * 2)).text($(element).text().trim(), { lineGap: 3 });
    doc.moveDown(0.35);
    return;
  }
  if (tag === 'table') {
    drawPdfTable(doc, $, element);
    doc.moveDown(0.6);
    return;
  }
  if (tag === 'img' || tag === 'figure' || $(element).find('img').length) {
    const image = tag === 'img' ? $(element) : $(element).find('img').first();
    const parsed = parseDataImage(image.attr('src'));
    if (parsed) {
      const normalized = await normalizeEmbeddedImage(parsed.buffer);
      const width = Math.min(normalized.width, doc.page.width - doc.page.margins.left - doc.page.margins.right);
      const height = width * (normalized.height / normalized.width);
      ensurePdfSpace(doc, Math.min(height, 360) + 20);
      doc.image(normalized.buffer, { fit: [width, 340], align: 'center' });
      const caption = $(element).find('figcaption').text().trim() || image.attr('alt');
      if (caption) doc.font('Helvetica-Oblique').fontSize(9).fillColor('#555555').text(caption, { align: 'center' }).fillColor('#000000');
      doc.moveDown(0.6);
    }
    return;
  }
  if (tag === 'ul' || tag === 'ol') {
    $(element)
      .children('li')
      .each((index, item) => {
        ensurePdfSpace(doc, 24);
        const marker = tag === 'ol' ? `${index + 1}.` : '-';
        doc.font('Helvetica').fontSize(11).text(`${marker} ${$(item).text().replace(/\s+/g, ' ').trim()}`, {
          indent: 12,
          lineGap: 3
        });
      });
    doc.moveDown(0.35);
    return;
  }
  const text = $(element).text().replace(/\s+/g, ' ').trim();
  if (text) {
    ensurePdfSpace(doc, 28);
    doc.font(tag === 'blockquote' ? 'Helvetica-Oblique' : 'Helvetica').fontSize(11).text(text, {
      indent: tag === 'blockquote' ? 14 : 0,
      lineGap: 4
    });
    doc.moveDown(0.45);
  }
}

function drawPdfTable(doc, $, table) {
  const rows = $(table)
    .find('tr')
    .toArray()
    .map((row) =>
      $(row)
        .children('th,td')
        .toArray()
        .map((cell) => $(cell).text().replace(/\s+/g, ' ').trim())
    );
  const columnCount = Math.max(1, ...rows.map((row) => row.length));
  const pageWidth = doc.page.width - doc.page.margins.left - doc.page.margins.right;
  const columnWidth = pageWidth / columnCount;
  rows.forEach((row, rowIndex) => {
    doc.font(rowIndex === 0 ? 'Helvetica-Bold' : 'Helvetica').fontSize(9);
    const rowHeight = Math.max(
      24,
      ...row.map((value) => doc.heightOfString(value || ' ', { width: columnWidth - 10, lineGap: 2 }) + 10)
    );
    ensurePdfSpace(doc, rowHeight);
    const y = doc.y;
    row.forEach((value, columnIndex) => {
      const x = doc.page.margins.left + columnIndex * columnWidth;
      if (rowIndex === 0) doc.rect(x, y, columnWidth, rowHeight).fillAndStroke('#eeeeee', '#777777');
      else doc.rect(x, y, columnWidth, rowHeight).stroke('#999999');
      doc.fillColor('#111111').text(value || '', x + 5, y + 5, {
        width: columnWidth - 10,
        height: rowHeight - 10,
        lineGap: 2
      });
    });
    doc.y = y + rowHeight;
  });
}

function ensurePdfSpace(doc, height) {
  const bottom = doc.page.height - doc.page.margins.bottom;
  if (doc.y + height > bottom) doc.addPage();
}

async function writeStructuredDocx(outputPath, title, html) {
  const children = [new Paragraph({ text: title, heading: HeadingLevel.TITLE })];
  const $ = cheerio.load(`<body>${html || '<p>No extractable content was found.</p>'}</body>`);
  for (const element of $('body').children().toArray()) {
    const converted = await htmlElementToDocx($, element);
    if (Array.isArray(converted)) children.push(...converted);
    else if (converted) children.push(converted);
  }
  const doc = new Document({
    sections: [
      {
        children
      }
    ]
  });
  await fsp.writeFile(outputPath, await Packer.toBuffer(doc));
}

async function htmlElementToDocx($, element) {
  const tag = String(element.tagName || '').toLowerCase();
  if (tag === 'table') {
    const rows = [];
    for (const row of $(element).find('tr').toArray()) {
      const cells = [];
      for (const cell of $(row).children('th,td').toArray()) {
        const sourceParagraphs = $(cell).children('p').toArray();
        const paragraphs = [];
        for (const sourceParagraph of sourceParagraphs.length ? sourceParagraphs : [cell]) {
          const runs = await inlineDocxRuns($, sourceParagraph, { bold: cell.tagName === 'th' });
          paragraphs.push(new Paragraph({ children: runs.length ? runs : [new TextRun('')] }));
        }
        cells.push(new TableCell({ children: paragraphs }));
      }
      if (cells.length) rows.push(new TableRow({ children: cells }));
    }
    return rows.length ? new Table({ rows, width: { size: 100, type: WidthType.PERCENTAGE } }) : null;
  }
  if (tag === 'ul' || tag === 'ol') {
    return $(element)
      .children('li')
      .toArray()
      .map(
        (item, index) =>
          new Paragraph({
            bullet: tag === 'ul' ? { level: 0 } : undefined,
            children: [new TextRun(tag === 'ol' ? `${index + 1}. ${$(item).text().trim()}` : $(item).text().trim())]
          })
      );
  }

  const children = await inlineDocxRuns($, element);
  const heading = /^h[1-6]$/.test(tag) ? HeadingLevel[`HEADING_${tag.slice(1)}`] : undefined;
  return new Paragraph({
    heading,
    children: children.length ? children : [new TextRun($(element).text().trim())]
  });
}

async function inlineDocxRuns($, element, styles = {}) {
  const runs = [];
  for (const node of $(element).contents().toArray()) {
    if (node.type === 'text') {
      if (node.data) runs.push(new TextRun({ text: node.data, ...styles }));
      continue;
    }
    const tag = String(node.tagName || '').toLowerCase();
    if (tag === 'br') {
      runs.push(new TextRun({ break: 1 }));
      continue;
    }
    if (tag === 'img') {
      const parsed = parseDataImage($(node).attr('src'));
      if (parsed) {
        const normalized = await normalizeEmbeddedImage(parsed.buffer);
        const width = Math.min(normalized.width, 560);
        runs.push(
          new ImageRun({
            data: normalized.buffer,
            transformation: { width, height: Math.max(1, Math.round(width * (normalized.height / normalized.width))) },
            type: 'png'
          })
        );
      }
      continue;
    }
    const nextStyles = {
      ...styles,
      bold: styles.bold || tag === 'strong' || tag === 'b',
      italics: styles.italics || tag === 'em' || tag === 'i',
      underline: styles.underline || tag === 'u' ? {} : undefined
    };
    const nested = await inlineDocxRuns($, node, nextStyles);
    if (tag === 'a' && $(node).attr('href') && nested.length) {
      runs.push(new ExternalHyperlink({ link: $(node).attr('href'), children: nested }));
    } else {
      runs.push(...nested);
    }
  }
  return runs;
}

function wordCompatibleDoc(title, body) {
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
    table { border-collapse: collapse; width: 100%; margin: 12pt 0; }
    th, td { border: 1px solid #777; padding: 6pt; vertical-align: top; }
    th { background: #eee; font-weight: bold; }
    img { max-width: 100%; height: auto; }
  </style>
</head>
<body>
  <h1>${escapeHtml(title)}</h1>
  ${body || '<p>No extractable content was found.</p>'}
</body>
</html>`;
}

function parseDataImage(value) {
  const match = String(value || '').match(/^data:(image\/[a-z0-9.+-]+);base64,([a-z0-9+/=\s]+)$/i);
  if (!match) return null;
  return { mime: match[1], buffer: Buffer.from(match[2].replace(/\s+/g, ''), 'base64') };
}

async function normalizeEmbeddedImage(buffer) {
  const image = sharp(buffer, { pages: 1 });
  const metadata = await image.metadata();
  const width = Math.max(1, metadata.width || 640);
  const height = Math.max(1, metadata.height || 480);
  return { buffer: await image.png().toBuffer(), width, height };
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
