import React, { useEffect, useMemo, useRef, useState } from 'react';
import { createRoot } from 'react-dom/client';
import JSZip from 'jszip';
import mimeDb from 'mime-db';
import {
  Archive,
  ArrowRight,
  CheckCircle2,
  ChevronDown,
  Database,
  Download,
  File,
  FileText,
  FolderOpen,
  Image as ImageIcon,
  Loader2,
  LockKeyhole,
  Moon,
  Music,
  RefreshCw,
  RotateCcw,
  Search,
  Settings2,
  ShieldCheck,
  Sparkles,
  Sun,
  Upload,
  Video,
  X
} from 'lucide-react';
import './styles.css';

const OUTPUT_TYPES = [
  {
    id: 'image',
    label: 'Image',
    icon: ImageIcon,
    formats: [
      { value: 'png', label: 'PNG (.png)', mime: 'image/png' },
      { value: 'jpg', label: 'JPEG (.jpg)', mime: 'image/jpeg' },
      { value: 'webp', label: 'WebP (.webp)', mime: 'image/webp' },
      { value: 'avif', label: 'AVIF (.avif)', mime: 'image/avif' },
      { value: 'heif', label: 'HEIF (.heif)', mime: 'image/heif' },
      { value: 'tiff', label: 'TIFF (.tiff)', mime: 'image/tiff' },
      { value: 'gif', label: 'GIF image (.gif)', mime: 'image/gif' }
    ]
  },
  {
    id: 'audio',
    label: 'Audio',
    icon: Music,
    formats: [
      { value: 'mp3', label: 'MP3 (.mp3)', mime: 'audio/mpeg' },
      { value: 'wav', label: 'WAV (.wav)', mime: 'audio/wav' },
      { value: 'ogg', label: 'OGG (.ogg)', mime: 'audio/ogg' },
      { value: 'flac', label: 'FLAC (.flac)', mime: 'audio/flac' },
      { value: 'opus', label: 'Opus (.opus)', mime: 'audio/opus' },
      { value: 'aac', label: 'AAC (.aac)', mime: 'audio/aac' },
      { value: 'm4a', label: 'M4A (.m4a)', mime: 'audio/mp4' },
      { value: 'webm-audio', label: 'WebM audio (.webm)', mime: 'audio/webm' }
    ]
  },
  {
    id: 'video',
    label: 'Video',
    icon: Video,
    formats: [
      { value: 'mp4', label: 'MP4 (.mp4)', mime: 'video/mp4' },
      { value: 'webm-video', label: 'WebM video (.webm)', mime: 'video/webm' },
      { value: 'mov', label: 'MOV (.mov)', mime: 'video/quicktime' },
      { value: 'avi', label: 'AVI (.avi)', mime: 'video/x-msvideo' },
      { value: 'mkv', label: 'MKV (.mkv)', mime: 'video/x-matroska' },
      { value: 'gif-video', label: 'Animated GIF (.gif)', mime: 'image/gif' }
    ]
  },
  {
    id: 'document',
    label: 'Document',
    icon: FileText,
    formats: [
      { value: 'txt', label: 'Plain text (.txt)', mime: 'text/plain' },
      { value: 'md', label: 'Markdown (.md)', mime: 'text/markdown' },
      { value: 'rtf', label: 'Rich Text (.rtf)', mime: 'application/rtf' },
      { value: 'html', label: 'HTML document (.html)', mime: 'text/html' },
      { value: 'pdf', label: 'PDF (.pdf)', mime: 'application/pdf' },
      { value: 'doc', label: 'Word DOC (.doc)', mime: 'application/msword' },
      {
        value: 'docx',
        label: 'Word DOCX (.docx)',
        mime: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
      }
    ]
  },
  {
    id: 'data',
    label: 'Data',
    icon: Database,
    formats: [
      { value: 'json', label: 'JSON (.json)', mime: 'application/json' },
      { value: 'csv', label: 'CSV table (.csv)', mime: 'text/csv' },
      { value: 'tsv', label: 'TSV table (.tsv)', mime: 'text/tab-separated-values' },
      {
        value: 'xlsx',
        label: 'Excel XLSX (.xlsx)',
        mime: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
      },
      { value: 'xml', label: 'XML (.xml)', mime: 'application/xml' },
      { value: 'yaml', label: 'YAML (.yaml)', mime: 'application/yaml' },
      { value: 'base64', label: 'Base64 text (.txt)', mime: 'text/plain' }
    ]
  },
  {
    id: 'archive',
    label: 'Archive',
    icon: Archive,
    formats: [
      { value: 'zip', label: 'ZIP archive (.zip)', mime: 'application/zip' },
      { value: 'tar', label: 'TAR archive (.tar)', mime: 'application/x-tar' },
      { value: 'gzip', label: 'GZIP (.gz)', mime: 'application/gzip' },
      { value: 'tgz', label: 'TAR.GZ (.tgz)', mime: 'application/gzip' }
    ]
  }
];

const PRESETS = [
  { id: 'auto', label: 'Auto', description: 'Recommended' },
  { id: 'small', label: 'Small', description: 'Lower size' },
  { id: 'sharp', label: 'Sharp', description: 'Higher quality' }
];

const FORMAT_ALIASES = {
  image: ['image', 'picture', 'photo', 'thumbnail', 'screenshot'],
  audio: ['audio', 'sound', 'music', 'song', 'voice'],
  video: ['video', 'movie', 'clip', 'reel', 'animation'],
  document: ['document', 'doc', 'word', 'text', 'paper', 'readable'],
  data: ['data', 'spreadsheet', 'excel', 'sheet', 'table'],
  archive: ['archive', 'compressed', 'folder', 'package'],
  jpg: ['jpeg', 'photo'],
  webp: ['web image', 'small image'],
  heif: ['heic', 'iphone image', 'apple image'],
  gif: ['animated image'],
  mp3: ['music', 'audio'],
  wav: ['lossless audio'],
  flac: ['lossless music'],
  opus: ['voice', 'web audio'],
  mp4: ['video', 'phone video', 'standard video'],
  'webm-video': ['web video'],
  'gif-video': ['animated gif', 'gif animation'],
  txt: ['plain text', 'notes'],
  md: ['markdown', 'readme'],
  rtf: ['rich text'],
  pdf: ['paper', 'print'],
  doc: ['word', 'legacy word', 'microsoft word', 'msword'],
  docx: ['word', 'doc', 'microsoft word'],
  csv: ['spreadsheet', 'table'],
  tsv: ['tab separated', 'spreadsheet', 'table'],
  xlsx: ['excel', 'sheet', 'spreadsheet'],
  yaml: ['yml', 'config'],
  zip: ['compressed', 'package'],
  tar: ['package'],
  gzip: ['gz', 'compressed'],
  tgz: ['tar gz', 'compressed']
};

const FORMAT_INDEX = OUTPUT_TYPES.flatMap((type) =>
  type.formats.map((format) => ({
    categoryId: type.id,
    categoryLabel: type.label,
    formatValue: format.value,
    formatLabel: format.label,
    mime: format.mime,
    search: normalizeSearch(
      [
        type.id,
        type.label,
        format.value,
        format.label,
        format.mime,
        extensionForFormat(format.value),
        ...(FORMAT_ALIASES[type.id] || []),
        ...(FORMAT_ALIASES[format.value] || [])
      ].join(' ')
    )
  }))
);

const POPULAR_FORMAT_VALUES = ['mp4', 'mp3', 'pdf', 'doc', 'docx', 'xlsx', 'webp', 'zip'];

const KNOWN_FILE_TYPES = Object.entries(mimeDb)
  .flatMap(([mime, meta]) =>
    (meta.extensions || []).map((extension) => {
      const suggested = suggestOutputForMime(mime, extension);
      return {
        extension,
        mime,
        family: familyLabelForMime(mime),
        suggested,
        search: normalizeSearch(`${extension} ${mime} ${familyLabelForMime(mime)} ${suggested.formatValue}`)
      };
    })
  )
  .sort((a, b) => a.extension.localeCompare(b.extension));

const DEFAULT_OPTIONS = {
  category: 'image',
  format: 'png',
  name: 'converted-file.png',
  preset: 'auto',
  quality: 0.92,
  maxWidth: 1600,
  includeMetadata: true
};

const FORMAT_DESCRIPTIONS = {
  png: 'Lossless image', jpg: 'Universal photo', webp: 'Small web image', avif: 'Modern compressed image',
  heif: 'High efficiency image', tiff: 'Print-quality image', gif: 'Animated or still image', mp3: 'Universal audio',
  wav: 'Uncompressed audio', ogg: 'Open audio', flac: 'Lossless audio', opus: 'Efficient voice and audio',
  aac: 'Compressed audio', m4a: 'Apple-compatible audio', 'webm-audio': 'Web audio', mp4: 'Universal video',
  'webm-video': 'Web video', mov: 'QuickTime video', avi: 'Legacy video', mkv: 'Matroska video',
  'gif-video': 'Looping animation', txt: 'Plain text', md: 'Markdown document', rtf: 'Formatted text',
  html: 'Web document', pdf: 'Portable document', doc: 'Legacy Word document', docx: 'Microsoft Word document',
  json: 'Structured data', csv: 'Comma-separated table', tsv: 'Tab-separated table', xlsx: 'Excel spreadsheet',
  xml: 'Structured markup', yaml: 'Readable structured data', base64: 'Encoded file contents', zip: 'Universal archive',
  tar: 'Uncompressed archive', gzip: 'Single-file compression', tgz: 'Compressed TAR archive'
};

function initialTheme() {
  const saved = localStorage.getItem('file-converter-theme');
  if (saved === 'light' || saved === 'dark') return saved;
  return window.matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light';
}

function App() {
  const [sourceFile, setSourceFile] = useState(null);
  const [sourceMeta, setSourceMeta] = useState(null);
  const [options, setOptions] = useState(DEFAULT_OPTIONS);
  const [output, setOutput] = useState(null);
  const [status, setStatus] = useState('Add a file to begin.');
  const [isConverting, setIsConverting] = useState(false);
  const [advancedOpen, setAdvancedOpen] = useState(false);
  const [isDragging, setIsDragging] = useState(false);
  const [formatSearch, setFormatSearch] = useState('');
  const [formatCategory, setFormatCategory] = useState('all');
  const [theme, setTheme] = useState(initialTheme);
  const inputRef = useRef(null);

  const formatMatches = useMemo(() => {
    const query = outputSearchQuery(formatSearch);
    if (!query) {
      if (formatCategory !== 'all') return FORMAT_INDEX.filter((entry) => entry.categoryId === formatCategory);
      return POPULAR_FORMAT_VALUES.map((value) => FORMAT_INDEX.find((entry) => entry.formatValue === value)).filter(Boolean);
    }

    const terms = query.split(' ').filter(Boolean);
    return FORMAT_INDEX.filter(
      (entry) =>
        (formatCategory === 'all' || entry.categoryId === formatCategory) &&
        terms.every((term) => entry.search.includes(term))
    ).slice(0, 16);
  }, [formatCategory, formatSearch]);

  const knownTypeMatches = useMemo(() => {
    const query = normalizeSearch(formatSearch);
    if (query.length < 2) {
      return [];
    }
    const terms = query.split(' ').filter(Boolean);
    const outputExtensions = new Set(FORMAT_INDEX.map((entry) => extensionForFormat(entry.formatValue)));
    return KNOWN_FILE_TYPES.filter(
      (entry) => terms.every((term) => entry.search.includes(term)) && !outputExtensions.has(entry.extension)
    ).slice(0, 8);
  }, [formatSearch]);

  useEffect(() => {
    document.documentElement.dataset.theme = theme;
    localStorage.setItem('file-converter-theme', theme);
  }, [theme]);

  useEffect(() => {
    if (!sourceFile) {
      return;
    }

    let cancelled = false;
    readFileMeta(sourceFile).then((meta) => {
      if (cancelled) {
        return;
      }

      setSourceMeta(meta);
      const smartCategory = guessCategory(sourceFile);
      const smartFormat = defaultFormatFor(smartCategory, sourceFile);
      setFormatCategory(smartCategory);
      setOptions((current) => ({
        ...current,
        category: smartCategory,
        format: smartFormat,
        name: buildOutputName(sourceFile.name, smartFormat)
      }));
      setStatus('Ready');
      setOutput(null);
    });

    return () => {
      cancelled = true;
    };
  }, [sourceFile]);

  useEffect(() => {
    return () => {
      if (output?.url) {
        URL.revokeObjectURL(output.url);
      }
    };
  }, [output]);

  const selectOutput = (categoryId, formatValue, clearSearch = false) => {
    setOptions((current) => ({
      ...current,
      category: categoryId,
      format: formatValue,
      name: sourceFile
        ? buildOutputName(sourceFile.name, formatValue)
        : replaceExtension(current.name || DEFAULT_OPTIONS.name, extensionForFormat(formatValue))
    }));
    setOutput(null);
    if (clearSearch) {
      setFormatSearch('');
    }
  };

  const reset = () => {
    if (output?.url) {
      URL.revokeObjectURL(output.url);
    }
    setSourceFile(null);
    setSourceMeta(null);
    setOptions(DEFAULT_OPTIONS);
    setOutput(null);
    setStatus('Add a file to begin.');
    setIsConverting(false);
    setAdvancedOpen(false);
    setIsDragging(false);
    setFormatSearch('');
    setFormatCategory('all');
    if (inputRef.current) {
      inputRef.current.value = '';
    }
  };

  const handleFiles = (files) => {
    const nextFile = files?.[0];
    if (!nextFile) {
      return;
    }
    setSourceFile(nextFile);
    setStatus('Reading file...');
  };

  const convert = async () => {
    if (!sourceFile || isConverting) {
      return;
    }

    setIsConverting(true);
    setStatus('Converting...');

    try {
      const result = await convertFile(sourceFile, options, sourceMeta);
      if (output?.url) {
        URL.revokeObjectURL(output.url);
      }
      setOutput(result);
      setStatus('Converted');
    } catch (error) {
      setStatus(error.message || 'Conversion failed');
      setOutput(null);
    } finally {
      setIsConverting(false);
    }
  };

  const compatible = isFormatCompatible(sourceFile, options.category);
  const canConvert = Boolean(sourceFile) && !isConverting && compatible;
  const conversionNote = describeConversion(sourceFile, options.category);

  return (
    <main className="app-shell">
      <header className="topbar">
        <div className="brand-lockup">
          <span className="brand-mark"><RefreshCw size={25} /></span>
          <div>
            <h1>File Converter</h1>
            <p>Any file. The right format. Done.</p>
          </div>
        </div>
        <div className="topbar-actions">
          <span className="privacy-note"><LockKeyhole size={16} /> Your file stays on this device</span>
          <button className="icon-button" type="button" onClick={() => setTheme((value) => (value === 'dark' ? 'light' : 'dark'))} aria-label={`Switch to ${theme === 'dark' ? 'light' : 'dark'} mode`} title={`Switch to ${theme === 'dark' ? 'light' : 'dark'} mode`}>
            {theme === 'dark' ? <Sun size={19} /> : <Moon size={19} />}
          </button>
          <button className="reset-button" type="button" onClick={reset}><RotateCcw size={17} /> Reset</button>
        </div>
      </header>

      <section className="converter-grid" aria-label="File converter workspace">
        <div className="source-panel">
          <div className="section-heading"><span>1</span><PanelTitle title="Add a file" /></div>
          <div
            className={`drop-zone ${isDragging ? 'is-dragging' : ''}`}
            role="button"
            tabIndex="0"
            onClick={() => inputRef.current?.click()}
            onKeyDown={(event) => {
              if (event.key === 'Enter' || event.key === ' ') inputRef.current?.click();
            }}
            onDragOver={(event) => {
              event.preventDefault();
              setIsDragging(true);
            }}
            onDragLeave={() => setIsDragging(false)}
            onDrop={(event) => {
              event.preventDefault();
              setIsDragging(false);
              handleFiles(event.dataTransfer.files);
            }}
          >
            <input
              ref={inputRef}
              type="file"
              onChange={(event) => handleFiles(event.target.files)}
              aria-label="Choose file"
            />
            <Upload size={40} />
            <span>Drop any file here.</span>
            <button className="blue-button" type="button" onClick={(event) => { event.stopPropagation(); inputRef.current?.click(); }}>
              <FolderOpen size={18} />
              Choose file
            </button>
          </div>

          <SourceCard file={sourceFile} meta={sourceMeta} status={status} />

          <div className="section-heading output-heading"><span>2</span><PanelTitle title="Choose output" /></div>

          <label className="field-label" htmlFor="format-search">
            Search formats
          </label>
          <div className="format-search">
            <Search size={18} aria-hidden="true" />
            <input
              id="format-search"
              value={formatSearch}
              onChange={(event) => setFormatSearch(event.target.value)}
              onKeyDown={(event) => {
                if (event.key === 'Enter' && formatMatches[0]) {
                  event.preventDefault();
                  selectOutput(formatMatches[0].categoryId, formatMatches[0].formatValue, true);
                }
              }}
              placeholder="Search PDF, Word, MP4, or 'PDF to Word'"
              autoComplete="off"
            />
            {formatSearch && (
              <button className="clear-search" type="button" onClick={() => setFormatSearch('')} aria-label="Clear search">
                <X size={16} />
              </button>
            )}
          </div>
          <div className="category-tabs" role="tablist" aria-label="Format categories">
            <button type="button" role="tab" aria-selected={formatCategory === 'all'} className={formatCategory === 'all' ? 'is-selected' : ''} onClick={() => setFormatCategory('all')}>All</button>
            {OUTPUT_TYPES.map((item) => (
              <button key={item.id} type="button" role="tab" aria-selected={formatCategory === item.id} className={formatCategory === item.id ? 'is-selected' : ''} onClick={() => setFormatCategory(item.id)}>{item.label}</button>
            ))}
          </div>
          <div className="format-results" aria-label={formatSearch ? 'Matching formats' : formatCategory === 'all' ? 'Popular formats' : `${formatCategory} formats`}>
            {formatMatches.length > 0 ? (
              formatMatches.map((match) => {
                const selected = options.category === match.categoryId && options.format === match.formatValue;
                const targetType = OUTPUT_TYPES.find((item) => item.id === match.categoryId);
                const Icon = targetType?.icon || File;
                const available = isFormatCompatible(sourceFile, match.categoryId);
                return (
                  <button
                    className={`format-result ${selected ? 'is-selected' : ''}`}
                    type="button"
                    key={`${match.categoryId}-${match.formatValue}`}
                    onClick={() => selectOutput(match.categoryId, match.formatValue, true)}
                    aria-pressed={selected}
                  >
                    <span className="format-icon"><Icon size={18} /></span>
                    <span className="format-copy"><strong>{shortFormatLabel(match.formatLabel)}</strong><small>{FORMAT_DESCRIPTIONS[match.formatValue] || match.categoryLabel}</small></span>
                    <span className="format-extension">.{extensionForFormat(match.formatValue)}</span>
                    {!available ? <span className="format-warning">Media input needed</span> : null}
                  </button>
                );
              })
            ) : (
              <span className="no-format-results">No output format matches that search.</span>
            )}
          </div>
          {knownTypeMatches.length > 0 && (
            <div className="known-type-results" aria-label="Known input file types">
              {knownTypeMatches.map((match) => (
                <button
                  className="known-type-result"
                  type="button"
                  key={`${match.mime}-${match.extension}`}
                  onClick={() => selectOutput(match.suggested.categoryId, match.suggested.formatValue, true)}
                  title={`${match.mime} → ${match.suggested.label}`}
                >
                  <span>.{match.extension}</span>
                  <small>{match.family}</small>
                </button>
              ))}
            </div>
          )}

          <div className="option-grid">
            <label className="field-group" htmlFor="name">
              <span>Output filename</span>
              <input id="name" className="text-input" value={options.name} onChange={(event) => setOptions((current) => ({ ...current, name: event.target.value }))} placeholder="converted-file.png" />
            </label>
            <div className={`conversion-note ${compatible ? '' : 'is-error'}`} aria-live="polite">
              {compatible ? <CheckCircle2 size={18} /> : <X size={18} />}
              <span><strong>{conversionNote.title}</strong><small>{conversionNote.detail}</small></span>
            </div>
          </div>

          {options.category === 'image' ? <><label className="field-label">Image quality</label>
          <div className="preset-row" role="radiogroup" aria-label="Image quality">
            {PRESETS.map((preset) => (
              <button
                key={preset.id}
                className={`preset-button ${options.preset === preset.id ? 'is-selected' : ''}`}
                type="button"
                role="radio"
                aria-checked={options.preset === preset.id}
                onClick={() =>
                  setOptions((current) => ({
                    ...current,
                    preset: preset.id,
                    quality: preset.id === 'small' ? 0.72 : preset.id === 'sharp' ? 0.98 : 0.92,
                    maxWidth: preset.id === 'small' ? 900 : preset.id === 'sharp' ? 2400 : 1600
                  }))
                }
              >
                <span className="radio-dot" />
                <span>
                  {preset.label}
                  <small>{preset.description}</small>
                </span>
              </button>
            ))}
          </div></> : null}

          {options.category === 'image' ? <><button className="advanced-toggle" type="button" onClick={() => setAdvancedOpen((value) => !value)} aria-expanded={advancedOpen}>
            <span>
              <Settings2 size={18} />
              Advanced options
            </span>
            <ChevronDown className={advancedOpen ? 'is-open' : ''} size={18} />
          </button>

          {advancedOpen && (
            <div className="advanced-panel">
              <label>
                Quality
                <input
                  type="range"
                  min="0.5"
                  max="1"
                  step="0.01"
                  value={options.quality}
                  onChange={(event) =>
                    setOptions((current) => ({ ...current, quality: Number(event.target.value) }))
                  }
                />
                <span>{Math.round(options.quality * 100)}%</span>
              </label>
              <label>
                Max image width
                <input
                  type="number"
                  min="64"
                  max="4096"
                  value={options.maxWidth}
                  onChange={(event) =>
                    setOptions((current) => ({ ...current, maxWidth: Number(event.target.value) || 1600 }))
                  }
                />
              </label>
            </div>
          )}</> : null}

          <button className="convert-button" type="button" onClick={convert} disabled={!canConvert}>
            {isConverting ? <Loader2 className="spin" size={20} /> : <RefreshCw size={20} />}
            {isConverting ? 'Converting...' : 'Convert file'}
          </button>
        </div>

        <div className="preview-panel">
          <div className="preview-header">
            <div className="section-heading result-heading"><span>3</span><PanelTitle title="Result" /></div>
            <a
              className={`download-button ${output ? '' : 'is-disabled'}`}
              href={output?.url || '#'}
              download={output?.fileName}
              aria-disabled={!output}
              onClick={(event) => {
                if (!output) {
                  event.preventDefault();
                }
              }}
            >
              <Download size={18} />
              Download
            </a>
          </div>
          <Preview output={output} file={sourceFile} status={status} />
        </div>
      </section>

      <footer className="utility-footer"><ShieldCheck size={18} /> Private, local conversion with bundled media and document tools.</footer>
    </main>
  );
}

function PanelTitle({ title }) {
  return <h2 className="panel-title">{title}</h2>;
}

function SourceCard({ file, meta, status }) {
  if (!file) {
    return (
      <div className="source-card is-empty">
        <File size={30} />
        <div>
          <strong>Empty</strong>
          <span>Add one file to start converting.</span>
        </div>
      </div>
    );
  }

  const isError = !['Ready', 'Converted', 'Converting...', 'Reading file...'].includes(status);
  const statusLabel = status === 'Ready' ? 'Ready' : status;

  return (
    <div className="source-card">
      <FileBadge name={file.name} />
      <div className="source-info">
        <strong>{file.name}</strong>
        <span>{[formatBytes(file.size), meta?.dimensions, file.type || 'unknown type'].filter(Boolean).join('  •  ')}</span>
      </div>
      <span className={`ready-pill ${isError ? 'is-error' : ''}`} title={statusLabel}>
        <CheckCircle2 size={18} />
        <span>{statusLabel}</span>
      </span>
    </div>
  );
}

function FileBadge({ name }) {
  const extension = getExtension(name).slice(0, 4).toUpperCase() || 'FILE';
  return <span className="file-badge">{extension}</span>;
}

function Preview({ output, file, status }) {
  if (!output) {
    return (
      <div className="empty-preview">
        <FileText size={82} strokeWidth={1.4} />
        <h2>No output yet</h2>
        <p>Your converted file will appear here.</p>
        <span>{file ? status : 'Choose options and click “Convert file” to get started.'}</span>
      </div>
    );
  }

  if (output.previewType === 'image') {
    return (
      <div className="result-preview">
        <img src={output.url} alt="Converted file preview" />
        <ResultDetails output={output} />
      </div>
    );
  }

  if (output.previewType === 'video') {
    return (
      <div className="result-preview">
        <video src={output.url} controls />
        <ResultDetails output={output} />
      </div>
    );
  }

  if (output.previewType === 'audio') {
    return (
      <div className="result-preview is-compact">
        <Music size={74} />
        <audio src={output.url} controls />
        <ResultDetails output={output} />
      </div>
    );
  }

  return (
    <div className="result-preview is-compact">
      <Sparkles size={74} />
      <pre>{output.preview}</pre>
      <ResultDetails output={output} />
    </div>
  );
}

function ResultDetails({ output }) {
  return (
    <div className="result-details">
      <strong>{output.fileName}</strong>
      <span>
        {formatBytes(output.blob.size)} · {output.mime}
      </span>
      <p>{output.message}</p>
    </div>
  );
}

async function convertFile(file, options, meta) {
  const category = OUTPUT_TYPES.find((item) => item.id === options.category);
  const format = category?.formats.find((item) => item.value === options.format);
  const requestedName = normalizeFileName(options.name, extensionForFormat(options.format));

  try {
    return await convertWithBackend(file, requestedName, format, options);
  } catch (error) {
    const canUseBrowserFallback =
      (options.category === 'image' && file.type.startsWith('image/') && ['png', 'jpg', 'jpeg', 'webp', 'avif'].includes(options.format)) ||
      (options.category === 'document' && ['txt', 'html'].includes(options.format)) ||
      (options.category === 'data' && ['json', 'csv', 'base64'].includes(options.format)) ||
      (options.category === 'archive' && options.format === 'zip');

    if (!canUseBrowserFallback) {
      throw error;
    }

    if (options.category === 'image') {
      return convertImage(file, requestedName, format, options);
    }
    if (options.category === 'document' && options.format === 'txt') {
      return textOutput(file, requestedName, meta);
    }
    if (options.category === 'document' && options.format === 'html') {
      return htmlOutput(file, requestedName, meta);
    }
    if (options.category === 'data' && options.format === 'json') {
      return jsonPackage(file, requestedName, meta, options);
    }
    if (options.category === 'data' && options.format === 'csv') {
      return csvOutput(file, requestedName, meta);
    }
    if (options.category === 'data' && options.format === 'base64') {
      return base64Output(file, requestedName);
    }
    return zipOutput(file, requestedName, options);
  }
}

async function convertWithBackend(file, fileName, format, options) {
  const form = new FormData();
  form.append('file', file);
  form.append('format', options.format);
  form.append('fileName', fileName);
  form.append('quality', String(Math.round(options.quality * 100)));
  form.append('maxWidth', String(options.maxWidth));

  const response = await fetch(`${apiBase()}/api/convert`, {
    method: 'POST',
    body: form
  });

  if (!response.ok) {
    let message = 'Conversion failed.';
    try {
      const payload = await response.json();
      message = payload.error || message;
    } catch {
      message = await response.text();
    }
    throw new Error(message);
  }

  const blob = await response.blob();
  const outputName = decodeURIComponent(response.headers.get('X-Output-Name') || encodeURIComponent(fileName));
  const mime = response.headers.get('X-Output-Mime') || format?.mime || blob.type || 'application/octet-stream';
  const previewType = previewTypeForMime(mime, outputName);
  const preview = previewType === 'text' ? await blob.text() : '';

  return makeOutput({
    blob,
    fileName: outputName,
    mime,
    previewType,
    preview: preview.slice(0, 1600),
    message: 'Converted with the local engine.'
  });
}

function apiBase() {
  if (window.location.hostname === '127.0.0.1' && window.location.port !== '8787') {
    return 'http://127.0.0.1:8787';
  }
  return '';
}

function previewTypeForMime(mime, fileName) {
  if (mime.startsWith('image/')) return 'image';
  if (mime.startsWith('video/')) return 'video';
  if (mime.startsWith('audio/')) return 'audio';
  if (mime.startsWith('text/') || /\.(json|csv|xml|html|txt)$/i.test(fileName)) return 'text';
  return 'binary';
}

async function convertImage(file, fileName, format, options) {
  const bitmap = await createImageBitmap(file);
  const scale = Math.min(1, options.maxWidth / bitmap.width);
  const width = Math.max(1, Math.round(bitmap.width * scale));
  const height = Math.max(1, Math.round(bitmap.height * scale));
  const canvas = document.createElement('canvas');
  canvas.width = width;
  canvas.height = height;
  const context = canvas.getContext('2d', { alpha: format.mime !== 'image/jpeg' });

  if (format.mime === 'image/jpeg') {
    context.fillStyle = '#ffffff';
    context.fillRect(0, 0, width, height);
  }

  context.imageSmoothingEnabled = true;
  context.imageSmoothingQuality = options.preset === 'sharp' ? 'high' : 'medium';
  context.drawImage(bitmap, 0, 0, width, height);

  const blob = await new Promise((resolve, reject) => {
    canvas.toBlob(
      (nextBlob) => {
        if (nextBlob) {
          resolve(nextBlob);
        } else {
          reject(new Error(`${format.label} is not supported by this browser.`));
        }
      },
      format.mime,
      options.quality
    );
  });

  return makeOutput({
    blob,
    fileName,
    mime: format.mime,
    previewType: 'image',
    message: `Rendered at ${width} × ${height}.`
  });
}

async function audioToVideo(file, fileName, options) {
  if (!('MediaRecorder' in window)) {
    throw new Error('This browser does not expose MediaRecorder.');
  }

  const audio = new Audio(URL.createObjectURL(file));
  audio.crossOrigin = 'anonymous';
  audio.muted = true;
  audio.playsInline = true;

  await audioLoaded(audio);

  const duration = Number.isFinite(audio.duration) ? Math.min(audio.duration, 30) : 8;
  const canvas = document.createElement('canvas');
  canvas.width = 1280;
  canvas.height = 720;
  const context = canvas.getContext('2d');
  const canvasStream = canvas.captureStream(30);
  const audioContext = new AudioContext();
  const source = audioContext.createMediaElementSource(audio);
  const destination = audioContext.createMediaStreamDestination();
  source.connect(destination);
  source.connect(audioContext.destination);
  destination.stream.getAudioTracks().forEach((track) => canvasStream.addTrack(track));

  const chunks = [];
  const recorder = new MediaRecorder(canvasStream, { mimeType: pickMediaRecorderType() });

  recorder.ondataavailable = (event) => {
    if (event.data.size) {
      chunks.push(event.data);
    }
  };

  const finished = new Promise((resolve) => {
    recorder.onstop = () => resolve(new Blob(chunks, { type: recorder.mimeType || 'video/webm' }));
  });

  recorder.start();
  await audioContext.resume();
  await audio.play();

  const startedAt = performance.now();
  await new Promise((resolve) => {
    const drawFrame = () => {
      const elapsed = (performance.now() - startedAt) / 1000;
      drawAudioPoster(context, canvas, file, elapsed, duration);
      if (elapsed >= duration || audio.ended) {
        resolve();
        return;
      }
      requestAnimationFrame(drawFrame);
    };
    drawFrame();
  });

  audio.pause();
  recorder.stop();
  const blob = await finished;
  URL.revokeObjectURL(audio.src);
  await audioContext.close();

  return makeOutput({
    blob,
    fileName,
    mime: blob.type || 'video/webm',
    previewType: 'video',
    message: 'Created a local WebM video with the audio and a generated cover.'
  });
}

function drawAudioPoster(context, canvas, file, elapsed, duration) {
  const { width, height } = canvas;
  context.fillStyle = '#f9f7f1';
  context.fillRect(0, 0, width, height);
  context.strokeStyle = '#111111';
  context.lineWidth = 8;
  context.strokeRect(34, 34, width - 68, height - 68);

  context.fillStyle = '#111111';
  context.font = '700 58px Arial, sans-serif';
  context.fillText('Audio Video', 92, 138);
  context.font = '500 30px Arial, sans-serif';
  wrapCanvasText(context, file.name, 92, 192, width - 184, 42);

  const progress = Math.max(0, Math.min(1, elapsed / duration));
  const barX = 92;
  const barY = height - 136;
  const barWidth = width - 184;
  context.fillStyle = '#dfded9';
  context.fillRect(barX, barY, barWidth, 16);
  context.fillStyle = '#1463ff';
  context.fillRect(barX, barY, barWidth * progress, 16);

  context.strokeStyle = '#111111';
  context.lineWidth = 3;
  const centerY = 405;
  for (let index = 0; index < 72; index += 1) {
    const x = 92 + index * 15;
    const heightValue = 24 + Math.sin(index * 0.71 + elapsed * 4) * 40 + Math.sin(index * 0.21) * 30;
    context.beginPath();
    context.moveTo(x, centerY - Math.abs(heightValue));
    context.lineTo(x, centerY + Math.abs(heightValue));
    context.stroke();
  }
}

function wrapCanvasText(context, text, x, y, maxWidth, lineHeight) {
  const words = text.split(' ');
  let line = '';
  let nextY = y;
  for (const word of words) {
    const testLine = `${line}${word} `;
    if (context.measureText(testLine).width > maxWidth && line) {
      context.fillText(line, x, nextY);
      line = `${word} `;
      nextY += lineHeight;
    } else {
      line = testLine;
    }
  }
  context.fillText(line, x, nextY);
}

function audioLoaded(audio) {
  return new Promise((resolve, reject) => {
    audio.onloadedmetadata = resolve;
    audio.onerror = () => reject(new Error('Audio could not be decoded in this browser.'));
  });
}

function pickMediaRecorderType() {
  const types = ['video/webm;codecs=vp9,opus', 'video/webm;codecs=vp8,opus', 'video/webm'];
  return types.find((type) => MediaRecorder.isTypeSupported(type)) || '';
}

async function passthrough(file, fileName, mime, message) {
  const blob = file.slice(0, file.size, mime || file.type || 'application/octet-stream');
  return makeOutput({
    blob,
    fileName,
    mime: blob.type || mime,
    previewType: mime?.startsWith('video') ? 'video' : mime?.startsWith('audio') ? 'audio' : 'text',
    message
  });
}

async function textOutput(file, fileName, meta) {
  const text = await fileToBestText(file);
  const content = text.readable
    ? text.value
    : `Source file: ${file.name}\nSize: ${formatBytes(file.size)}\nType: ${file.type || 'unknown'}\n${meta?.dimensions ? `Dimensions: ${meta.dimensions}\n` : ''}\nBinary files cannot be losslessly represented as plain text in the browser.`;
  const blob = new Blob([content], { type: 'text/plain' });
  return makeOutput({
    blob,
    fileName,
    mime: 'text/plain',
    previewType: 'text',
    preview: content.slice(0, 900),
    message: text.readable ? 'Text extracted locally.' : 'Created a readable text summary for a binary file.'
  });
}

async function htmlOutput(file, fileName, meta) {
  const text = await fileToBestText(file);
  const body = text.readable
    ? `<pre>${escapeHtml(text.value)}</pre>`
    : `<p>This file is binary, so the browser created an HTML summary.</p><dl><dt>Name</dt><dd>${escapeHtml(file.name)}</dd><dt>Size</dt><dd>${formatBytes(file.size)}</dd><dt>Type</dt><dd>${escapeHtml(file.type || 'unknown')}</dd>${meta?.dimensions ? `<dt>Dimensions</dt><dd>${escapeHtml(meta.dimensions)}</dd>` : ''}</dl>`;
  const content = `<!doctype html><html><head><meta charset="utf-8"><title>${escapeHtml(file.name)}</title><style>body{font:16px/1.5 system-ui;margin:40px;max-width:900px}pre{white-space:pre-wrap}dt{font-weight:700}</style></head><body><h1>${escapeHtml(file.name)}</h1>${body}</body></html>`;
  const blob = new Blob([content], { type: 'text/html' });
  return makeOutput({
    blob,
    fileName,
    mime: 'text/html',
    previewType: 'text',
    preview: content.slice(0, 900),
    message: 'Built a standalone HTML document locally.'
  });
}

async function jsonPackage(file, fileName, meta, options) {
  const dataUrl = await readAsDataUrl(file);
  const payload = {
    fileName: file.name,
    outputCreatedAt: new Date().toISOString(),
    mimeType: file.type || 'application/octet-stream',
    size: file.size,
    metadata: options.includeMetadata ? meta : undefined,
    dataUrl
  };
  const content = JSON.stringify(payload, null, 2);
  const blob = new Blob([content], { type: 'application/json' });
  return makeOutput({
    blob,
    fileName,
    mime: 'application/json',
    previewType: 'text',
    preview: content.slice(0, 900),
    message: 'Packaged the file as JSON with a data URL.'
  });
}

async function csvOutput(file, fileName, meta) {
  const text = await fileToBestText(file);
  const rows = [
    ['field', 'value'],
    ['name', file.name],
    ['type', file.type || 'unknown'],
    ['size', String(file.size)],
    ['dimensions', meta?.dimensions || '']
  ];

  if (text.readable) {
    const lines = text.value.split(/\r?\n/).slice(0, 40);
    lines.forEach((line, index) => rows.push([`line_${index + 1}`, line]));
  }

  const content = rows.map((row) => row.map(csvEscape).join(',')).join('\n');
  const blob = new Blob([content], { type: 'text/csv' });
  return makeOutput({
    blob,
    fileName,
    mime: 'text/csv',
    previewType: 'text',
    preview: content.slice(0, 900),
    message: 'Created a CSV table from source metadata and readable text.'
  });
}

async function base64Output(file, fileName) {
  const dataUrl = await readAsDataUrl(file);
  const base64 = dataUrl.split(',')[1] || '';
  const blob = new Blob([base64], { type: 'text/plain' });
  return makeOutput({
    blob,
    fileName,
    mime: 'text/plain',
    previewType: 'text',
    preview: base64.slice(0, 900),
    message: 'Encoded the file as Base64 text.'
  });
}

async function dataUrlOutput(file, fileName) {
  const dataUrl = await readAsDataUrl(file);
  const blob = new Blob([dataUrl], { type: 'text/plain' });
  return makeOutput({
    blob,
    fileName,
    mime: 'text/plain',
    previewType: 'text',
    preview: dataUrl.slice(0, 900),
    message: 'Created a portable data URL.'
  });
}

async function zipOutput(file, fileName, options) {
  const zip = new JSZip();
  zip.file(file.name, file);
  if (options.includeMetadata) {
    zip.file(
      'metadata.json',
      JSON.stringify(
        {
          fileName: file.name,
          size: file.size,
          type: file.type || 'application/octet-stream',
          packagedAt: new Date().toISOString()
        },
        null,
        2
      )
    );
  }
  const blob = await zip.generateAsync({ type: 'blob', compression: 'DEFLATE' });
  return makeOutput({
    blob,
    fileName,
    mime: 'application/zip',
    previewType: 'text',
    preview: `${file.name}\nmetadata.json`,
    message: 'Packed the source file into a ZIP archive.'
  });
}

function makeOutput({ blob, fileName, mime, previewType, preview = '', kind = 'converted', message }) {
  return {
    blob,
    fileName,
    mime: mime || blob.type || 'application/octet-stream',
    previewType,
    preview,
    kind,
    message,
    url: URL.createObjectURL(blob)
  };
}

async function readFileMeta(file) {
  const meta = {
    kind: guessSourceKind(file),
    readable: isLikelyText(file)
  };

  if (file.type.startsWith('image/')) {
    try {
      const bitmap = await createImageBitmap(file);
      meta.dimensions = `${bitmap.width} × ${bitmap.height}`;
    } catch {
      meta.dimensions = '';
    }
  }

  return meta;
}

function guessCategory(file) {
  if (file.type.startsWith('image/')) return 'image';
  if (file.type.startsWith('audio/')) return 'audio';
  if (file.type.startsWith('video/')) return 'video';
  if (isLikelyText(file)) return 'document';
  return 'archive';
}

function defaultFormatFor(category, file) {
  if (category === 'image') return file.type === 'image/png' ? 'webp' : 'png';
  if (category === 'audio') return 'mp3';
  if (category === 'video') return 'mp4';
  if (category === 'document') return 'txt';
  if (category === 'archive') return 'zip';
  return OUTPUT_TYPES.find((item) => item.id === category)?.formats[0]?.value || 'png';
}

function guessSourceKind(file) {
  if (file.type.startsWith('image/')) return 'image';
  if (file.type.startsWith('audio/')) return 'audio';
  if (file.type.startsWith('video/')) return 'video';
  if (isLikelyText(file)) return 'text';
  return 'binary';
}

function isLikelyText(file) {
  const extension = getExtension(file.name);
  return (
    file.type.startsWith('text/') ||
    ['txt', 'md', 'csv', 'json', 'html', 'css', 'js', 'ts', 'xml', 'svg', 'yml', 'yaml', 'log'].includes(extension)
  );
}

async function fileToBestText(file) {
  if (!isLikelyText(file)) {
    return { readable: false, value: '' };
  }

  try {
    return { readable: true, value: await file.text() };
  } catch {
    return { readable: false, value: '' };
  }
}

function readAsDataUrl(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = () => reject(new Error('Could not read the file.'));
    reader.readAsDataURL(file);
  });
}

function buildOutputName(sourceName, format) {
  const base = sourceName.replace(/\.[^.]+$/, '') || 'converted-file';
  return `${base}-converted.${extensionForFormat(format)}`;
}

function normalizeFileName(value, extension) {
  const fallback = `converted-file.${extension}`;
  const clean = value.trim().replace(/[\\/:*?"<>|]+/g, '-');
  if (!clean) {
    return fallback;
  }
  return clean.endsWith(`.${extension}`) ? clean : `${clean.replace(/\.[^.]+$/, '')}.${extension}`;
}

function extensionForFormat(format) {
  const map = {
    png: 'png',
    jpg: 'jpg',
    jpeg: 'jpg',
    webp: 'webp',
    avif: 'avif',
    heif: 'heif',
    tiff: 'tiff',
    gif: 'gif',
    mp3: 'mp3',
    wav: 'wav',
    ogg: 'ogg',
    flac: 'flac',
    opus: 'opus',
    aac: 'aac',
    m4a: 'm4a',
    'webm-audio': 'webm',
    mp4: 'mp4',
    'webm-video': 'webm',
    mov: 'mov',
    avi: 'avi',
    mkv: 'mkv',
    'gif-video': 'gif',
    txt: 'txt',
    md: 'md',
    rtf: 'rtf',
    html: 'html',
    pdf: 'pdf',
    doc: 'doc',
    docx: 'docx',
    json: 'json',
    csv: 'csv',
    tsv: 'tsv',
    xlsx: 'xlsx',
    xml: 'xml',
    yaml: 'yaml',
    base64: 'txt',
    zip: 'zip',
    tar: 'tar',
    gzip: 'gz',
    tgz: 'tgz'
  };
  return map[format] || 'bin';
}

function familyLabelForMime(mime) {
  const [family, subtype = 'file'] = String(mime).split('/');
  const labels = {
    application: 'App file',
    audio: 'Audio',
    font: 'Font',
    haptics: 'Haptics',
    image: 'Image',
    message: 'Message',
    model: '3D/model',
    multipart: 'Multipart',
    text: 'Text',
    video: 'Video'
  };
  return labels[family] || subtype.replace(/[.+-]/g, ' ');
}

function suggestOutputForMime(mime, extension) {
  const normalizedMime = String(mime || '').toLowerCase();
  const ext = String(extension || '').toLowerCase();
  const extMap = {
    doc: ['document', 'doc'],
    docm: ['document', 'docx'],
    dot: ['document', 'docx'],
    dotx: ['document', 'docx'],
    odt: ['document', 'docx'],
    rtf: ['document', 'docx'],
    pages: ['document', 'pdf'],
    ppt: ['document', 'pdf'],
    pptx: ['document', 'pdf'],
    odp: ['document', 'pdf'],
    key: ['document', 'pdf'],
    xls: ['data', 'xlsx'],
    xlsm: ['data', 'xlsx'],
    ods: ['data', 'xlsx'],
    numbers: ['data', 'xlsx'],
    sqlite: ['data', 'csv'],
    db: ['data', 'csv'],
    sql: ['document', 'txt'],
    md: ['document', 'html'],
    markdown: ['document', 'html'],
    psd: ['image', 'png'],
    ai: ['image', 'png'],
    eps: ['image', 'png'],
    heic: ['image', 'jpg'],
    heif: ['image', 'jpg'],
    raw: ['image', 'jpg'],
    cr2: ['image', 'jpg'],
    nef: ['image', 'jpg'],
    arw: ['image', 'jpg'],
    rar: ['archive', 'zip'],
    '7z': ['archive', 'zip'],
    gz: ['archive', 'zip'],
    bz2: ['archive', 'zip'],
    xz: ['archive', 'zip']
  };

  let pair = extMap[ext];
  if (!pair && normalizedMime.startsWith('image/')) pair = ['image', 'png'];
  if (!pair && normalizedMime.startsWith('audio/')) pair = ['audio', 'mp3'];
  if (!pair && normalizedMime.startsWith('video/')) pair = ['video', 'mp4'];
  if (!pair && normalizedMime.startsWith('text/')) pair = ['document', 'txt'];
  if (!pair && normalizedMime.includes('spreadsheet')) pair = ['data', 'xlsx'];
  if (!pair && normalizedMime.includes('presentation')) pair = ['document', 'pdf'];
  if (!pair && normalizedMime.includes('wordprocessing')) pair = ['document', 'docx'];
  if (!pair && normalizedMime.includes('zip')) pair = ['archive', 'zip'];
  if (!pair) pair = ['data', 'base64'];

  const [categoryId, formatValue] = pair;
  const entry = FORMAT_INDEX.find((item) => item.categoryId === categoryId && item.formatValue === formatValue);
  return {
    categoryId,
    formatValue,
    label: entry?.formatLabel || formatValue
  };
}

function replaceExtension(fileName, extension) {
  return `${fileName.replace(/\.[^.]+$/, '')}.${extension}`;
}

function normalizeSearch(value) {
  return String(value || '')
    .toLowerCase()
    .replace(/[._-]+/g, ' ')
    .replace(/[^a-z0-9]+/g, ' ')
    .trim();
}

function outputSearchQuery(value) {
  const normalized = normalizeSearch(value);
  const words = normalized.split(' ').filter(Boolean);
  const separators = ['to', 'into', 'as'];
  let splitAt = -1;
  words.forEach((word, index) => {
    if (separators.includes(word)) splitAt = index;
  });
  return splitAt >= 0 && splitAt < words.length - 1 ? words.slice(splitAt + 1).join(' ') : normalized;
}

function shortFormatLabel(label) {
  return String(label).replace(/\s*\([^)]*\)\s*$/, '');
}

function sourceFamily(file) {
  if (!file) return '';
  const extension = getExtension(file.name);
  if (file.type.startsWith('image/')) return 'image';
  if (file.type.startsWith('audio/')) return 'audio';
  if (file.type.startsWith('video/')) return 'video';
  if (['csv', 'tsv', 'xlsx', 'xls', 'json', 'xml', 'yaml', 'yml'].includes(extension)) return 'data';
  if (['doc', 'docx', 'pdf', 'txt', 'md', 'rtf', 'html'].includes(extension) || file.type.startsWith('text/')) return 'document';
  if (['zip', 'tar', 'gz', 'tgz'].includes(extension)) return 'archive';
  return 'file';
}

function isFormatCompatible(file, targetCategory) {
  if (!file) return true;
  if (targetCategory !== 'audio') return true;
  const family = sourceFamily(file);
  return family === 'audio' || family === 'video';
}

function describeConversion(file, targetCategory) {
  if (!file) return { title: 'Ready when you are', detail: 'Choose a file to check this conversion.' };
  const source = sourceFamily(file);
  if (!isFormatCompatible(file, targetCategory)) {
    return { title: 'Audio needs media input', detail: 'Choose an audio or video source for this output.' };
  }
  if (source === targetCategory) return { title: 'Direct conversion', detail: 'Preserves the source content wherever the format allows.' };
  if (targetCategory === 'archive') return { title: 'Packages the original', detail: 'Compresses your source without changing its contents.' };
  if (targetCategory === 'image') return { title: 'Visual rendition', detail: 'Creates an image from a frame, page, or file summary.' };
  if (targetCategory === 'video') return { title: 'Rendered video', detail: 'Uses the source media or creates a short visual clip.' };
  if (targetCategory === 'document') return { title: 'Content extraction', detail: 'Exports readable text; complex layout may be simplified.' };
  if (targetCategory === 'data') return { title: 'Structured export', detail: 'Turns rows or readable content into structured data.' };
  return { title: 'Compatible conversion', detail: 'The local engine can create this output.' };
}

function getExtension(name) {
  return name.includes('.') ? name.split('.').pop().toLowerCase() : '';
}

function formatBytes(bytes) {
  if (!bytes) return '0 B';
  const units = ['B', 'KB', 'MB', 'GB'];
  const index = Math.min(Math.floor(Math.log(bytes) / Math.log(1024)), units.length - 1);
  return `${(bytes / 1024 ** index).toFixed(index === 0 ? 0 : 2)} ${units[index]}`;
}

function csvEscape(value) {
  const text = String(value ?? '');
  return /[",\n]/.test(text) ? `"${text.replace(/"/g, '""')}"` : text;
}

function escapeHtml(value) {
  return String(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

createRoot(document.getElementById('root')).render(<App />);
