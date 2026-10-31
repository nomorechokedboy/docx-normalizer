import { useEffect, useRef, useState } from 'react';
import JSZip from 'jszip';
import { renderAsync } from 'docx-preview';

type Margins = { top: number; right: number; bottom: number; left: number };
type SectionInfo = { index: number; margins: Margins | null };

const WORD_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const presets: Record<string, Margins> = {
  'vn-admin': { top: 25, right: 20, bottom: 25, left: 30 },
  'balanced': { top: 25.4, right: 25.4, bottom: 25.4, left: 25.4 },
  'compact': { top: 15, right: 15, bottom: 15, left: 15 },
};

function twips(mm: number) { return String(Math.round((mm / 25.4) * 1440)); }
function mmFromTwips(value: string | null) {
  if (!value || !/^[-+]?\d+$/.test(value)) return null;
  return Math.round((Number(value) / 1440) * 25.4 * 10) / 10;
}
function getSectPrs(xml: Document) {
  return Array.from(xml.getElementsByTagNameNS(WORD_NS, 'sectPr'));
}
function readSections(xmlText: string): SectionInfo[] {
  const xml = new DOMParser().parseFromString(xmlText, 'application/xml');
  if (xml.getElementsByTagName('parsererror').length) throw new Error('The document XML could not be parsed.');
  return getSectPrs(xml).map((section, i) => {
    const pgMar = Array.from(section.getElementsByTagNameNS(WORD_NS, 'pgMar'))[0];
    if (!pgMar) return { index: i + 1, margins: null };
    const read = (name: string) => mmFromTwips(pgMar.getAttributeNS(WORD_NS, name));
    const values = { top: read('top'), right: read('right'), bottom: read('bottom'), left: read('left') };
    if (Object.values(values).some(v => v === null)) return { index: i + 1, margins: null };
    return { index: i + 1, margins: values as Margins };
  });
}

async function reconstructWithMargins(file: File, margins: Margins) {
  const zip = await JSZip.loadAsync(file);
  const docEntry = zip.file('word/document.xml');
  if (!docEntry) throw new Error('This file does not contain word/document.xml and may not be a valid DOCX.');
  const sourceXml = await docEntry.async('string');
  const xml = new DOMParser().parseFromString(sourceXml, 'application/xml');
  if (xml.getElementsByTagName('parsererror').length) throw new Error('The DOCX main document XML is malformed.');
  const sections = getSectPrs(xml);
  if (!sections.length) throw new Error('No section properties were found in this document.');

  for (const section of sections) {
    let pgMar = Array.from(section.children).find(el => el.namespaceURI === WORD_NS && el.localName === 'pgMar');
    if (!pgMar) {
      pgMar = xml.createElementNS(WORD_NS, 'w:pgMar');
      const pageSize = Array.from(section.children).find(el => el.namespaceURI === WORD_NS && el.localName === 'pgSz');
      if (pageSize?.nextSibling) section.insertBefore(pgMar, pageSize.nextSibling);
      else if (pageSize) section.appendChild(pgMar);
      else section.insertBefore(pgMar, section.firstChild);
    }
    for (const key of ['top', 'right', 'bottom', 'left'] as const) {
      pgMar.setAttributeNS(WORD_NS, `w:${key}`, twips(margins[key]));
    }
  }

  const serialized = new XMLSerializer().serializeToString(xml);
  // Rebuild a new ZIP package from all original entries. Only document.xml is changed;
  // styles, media, relationships, numbering and other OOXML parts are carried forward.
  zip.file('word/document.xml', serialized);
  const blob = await zip.generateAsync({ type: 'blob', compression: 'DEFLATE', compressionOptions: { level: 6 } });
  return { blob, sections: sections.length, originalXml: sourceXml, updatedXml: serialized };
}

function Preview({ blob, label }: { blob: Blob | null; label: string }) {
  const host = useRef<HTMLDivElement>(null);
  const [error, setError] = useState('');
  useEffect(() => {
    let active = true;
    if (!host.current || !blob) return;
    host.current.innerHTML = '';
    setError('');
    blob.arrayBuffer().then(buffer => renderAsync(buffer, host.current!, undefined, {
      className: 'docx-preview', inWrapper: true, ignoreWidth: false, ignoreHeight: false,
      breakPages: true, renderHeaders: true, renderFooters: true, renderFootnotes: true,
    })).catch(err => { if (active) setError(err instanceof Error ? err.message : 'Preview could not be rendered.'); });
    return () => { active = false; if (host.current) host.current.innerHTML = ''; };
  }, [blob]);
  return <section className="preview-card">
    <div className="preview-heading"><span>{label}</span><span className="preview-subtitle">Word layout preview</span></div>
    {error && <div className="preview-error">Preview error: {error}</div>}
    {!blob && <div className="empty-preview"><div className="empty-icon">▤</div><strong>Preview will appear here</strong><span>Choose a DOCX file to get started.</span></div>}
    <div className="preview-host" ref={host} />
  </section>;
}

export default function App() {
  const [file, setFile] = useState<File | null>(null);
  const [originalBlob, setOriginalBlob] = useState<Blob | null>(null);
  const [outputBlob, setOutputBlob] = useState<Blob | null>(null);
  const [preset, setPreset] = useState('vn-admin');
  const [margins, setMargins] = useState<Margins>(presets['vn-admin']);
  const [sections, setSections] = useState<SectionInfo[]>([]);
  const [status, setStatus] = useState('');
  const [busy, setBusy] = useState(false);
  const [stats, setStats] = useState<{ sections: number; preservedParts: number } | null>(null);

  function choosePreset(value: string) {
    setPreset(value);
    if (value in presets) setMargins(presets[value]);
  }
  async function loadFile(next: File | undefined) {
    if (!next) return;
    if (!next.name.toLowerCase().endsWith('.docx')) { setStatus('Please choose a .docx file. Legacy .doc files are not supported in this demo.'); return; }
    setFile(next); setOriginalBlob(next); setOutputBlob(null); setStats(null); setStatus('Inspecting document…');
    try {
      const zip = await JSZip.loadAsync(next);
      const entry = zip.file('word/document.xml');
      if (!entry) throw new Error('Could not find word/document.xml. Check that this is a valid DOCX file.');
      const xml = await entry.async('string');
      const found = readSections(xml);
      setSections(found);
      setStatus(`Loaded ${next.name} · ${found.length} section${found.length === 1 ? '' : 's'} detected.`);
    } catch (err) { setStatus(err instanceof Error ? err.message : 'Unable to inspect this file.'); setSections([]); }
  }
  async function normalize() {
    if (!file) return;
    setBusy(true); setStatus('Rebuilding package and preparing preview…');
    try {
      const result = await reconstructWithMargins(file, margins);
      // Reopen generated archive as a basic integrity check and count retained package parts.
      const check = await JSZip.loadAsync(result.blob);
      if (!check.file('word/document.xml') || !check.file('[Content_Types].xml')) throw new Error('Output package failed integrity checks.');
      const count = Object.keys(check.files).filter(name => !check.files[name].dir).length;
      setOutputBlob(result.blob);
      setStats({ sections: result.sections, preservedParts: count });
      setStatus('Candidate created. Review both previews before using the output; visual safety checks are not automated in this prototype.');
    } catch (err) { setOutputBlob(null); setStatus(err instanceof Error ? err.message : 'Normalization failed.'); }
    finally { setBusy(false); }
  }
  function download() {
    if (!outputBlob || !file) return;
    const url = URL.createObjectURL(outputBlob);
    const a = document.createElement('a'); a.href = url; a.download = `${file.name.replace(/\.docx$/i, '')}-normalized.docx`; a.click();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  }
  const changed = outputBlob !== null;

  return <main className="app-shell">
    <header className="topbar">
      <div className="brand-mark">M</div>
      <div className="brand-copy"><div className="brand-name">Margin Lab</div><div className="brand-caption">DOCX normalization prototype</div></div>
      <div className="local-badge"><span /> Runs locally in your browser</div>
    </header>
    <div className="page-wrap">
      <div className="intro"><div className="eyebrow">PROTOTYPE · 0.1</div><h1>Standardize margins.<br/><span>Keep the document intact.</span></h1><p>Rebuild a DOCX package while retaining its original content XML, media, styles and relationships. Change page margins, then compare the result side by side.</p></div>
      <div className="workspace">
        <aside className="control-panel">
          <div className="panel-title"><span className="step">01</span><div><h2>Source document</h2><p>Choose a Word file to inspect.</p></div></div>
          <label className="upload-zone" htmlFor="docx-file"><span className="upload-icon">↑</span><strong>{file ? file.name : 'Drop a DOCX here'}</strong><span>{file ? `${(file.size / 1024).toFixed(0)} KB · selected` : 'or click to browse · .docx only'}</span><input id="docx-file" type="file" accept=".docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document" onChange={e => void loadFile(e.target.files?.[0])}/></label>
          {sections.length > 0 && <div className="section-summary"><div className="summary-number">{sections.length}</div><div><strong>{sections.length === 1 ? 'Section found' : 'Sections found'}</strong><span>Margins are applied to every section.</span></div></div>}
          <div className="rule" />
          <div className="panel-title"><span className="step">02</span><div><h2>Margin profile</h2><p>All values are in millimeters.</p></div></div>
          <label className="field-label" htmlFor="profile">Preset</label>
          <select id="profile" value={preset} onChange={e => choosePreset(e.target.value)}>
            <option value="vn-admin">Administrative example · 25 / 20 / 25 / 30</option><option value="balanced">Balanced · 25.4 on all sides</option><option value="compact">Compact · 15 on all sides</option><option value="custom">Custom values</option>
          </select>
          <div className="margin-grid">
            {(['top', 'right', 'bottom', 'left'] as const).map(key => <label className="margin-field" key={key}><span>{key[0].toUpperCase() + key.slice(1)}</span><div><input type="number" min="0" max="100" step="0.5" value={margins[key]} onChange={e => { setPreset('custom'); setMargins(old => ({ ...old, [key]: Math.max(0, Math.min(100, Number(e.target.value))) })); setOutputBlob(null); }}/><span>mm</span></div></label>)}
          </div>
          <div className="preservation-note"><span className="note-icon">✓</span><div><strong>Preservation policy</strong><p>Keep original text, tables, images, styles and relationships. Only page margins are changed in this demo.</p></div></div>
          <button className="primary-button" onClick={() => void normalize()} disabled={!file || busy || sections.length === 0}>{busy ? <><span className="spinner"/> Processing…</> : <><span>Normalize margins</span><span>→</span></>}</button>
          {changed && <button className="download-button" onClick={download}><span>↓</span> Download normalized DOCX</button>}
          {status && <div className="status-message" role="status">{status}</div>}
          {stats && <div className="integrity-card"><div className="integrity-head"><span className="check-dot">✓</span><strong>Package integrity check passed</strong></div><div className="integrity-stats"><div><strong>{stats.sections}</strong><span>Sections updated</span></div><div><strong>{stats.preservedParts}</strong><span>Package entries retained</span></div></div><p>This verifies basic package integrity, not visual correctness. Inspect the previews for table or image issues.</p></div>}
          <div className="privacy-foot"><span>⌁</span> Your file stays in this browser tab. Nothing is uploaded.</div>
        </aside>
        <section className="preview-area">
          <div className="preview-toolbar"><div><h2>Visual comparison</h2><p>Original versus normalized candidate</p></div><div className="compare-label"><span className="original-dot"/> Original <span className="candidate-dot"/> Candidate</div></div>
          <div className="preview-grid"><Preview blob={originalBlob} label="Original"/><Preview blob={outputBlob} label="Normalized candidate"/></div>
          <div className="preview-disclaimer"><span>ⓘ</span><p><strong>Prototype limitation:</strong> browser previews are for visual inspection only. This demo does not automatically detect clipping, compare table geometry, or guarantee that Microsoft Word and LibreOffice will render identically.</p></div>
        </section>
      </div>
      <footer><span>Margin Lab · local prototype</span><span>DOCX package reconstruction · not production-ready</span></footer>
    </div>
  </main>;
}
