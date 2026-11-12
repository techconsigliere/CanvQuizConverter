/*
  CanvQuizConverter: a Word quiz in, a Canvas QTI package out.

  Everything happens on this page. The .docx is read with the browser's own
  zip decompression (DecompressionStream) and XML parser (DOMParser), and the
  QTI package is written with a tiny zip writer below. No libraries, no
  network, no upload. Readable top to bottom on purpose.

  MIT licensed. https://github.com/techconsigliere/CanvQuizConverter
*/

const VERSION = '0.2.0';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

/* ---------- reading a .docx (it's a zip of XML files) ------------------- */

async function unzip(buffer) {
  const bytes = new Uint8Array(buffer);
  const view = new DataView(buffer);
  // The central directory's end record sits in the last 64 KB.
  let eocd = -1;
  for (let i = bytes.length - 22; i >= Math.max(0, bytes.length - 65557); i--) {
    if (view.getUint32(i, true) === 0x06054b50) { eocd = i; break; }
  }
  if (eocd < 0) { throw new Error('That file isn’t a Word document (.docx). If it’s an older .doc file, open it in Word and save it as .docx.'); }
  const count = view.getUint16(eocd + 10, true);
  let p = view.getUint32(eocd + 16, true);
  const files = {};
  for (let n = 0; n < count; n++) {
    if (view.getUint32(p, true) !== 0x02014b50) { break; }
    const method = view.getUint16(p + 10, true);
    const csize = view.getUint32(p + 20, true);
    const nlen = view.getUint16(p + 28, true);
    const xlen = view.getUint16(p + 30, true);
    const clen = view.getUint16(p + 32, true);
    const local = view.getUint32(p + 42, true);
    const name = new TextDecoder().decode(bytes.subarray(p + 46, p + 46 + nlen));
    files[name] = { method, csize, local };
    p += 46 + nlen + xlen + clen;
  }
  return async function read(name) {
    const f = files[name];
    if (!f) { return null; }
    const start = f.local + 30 + view.getUint16(f.local + 26, true) + view.getUint16(f.local + 28, true);
    const raw = bytes.subarray(start, start + f.csize);
    if (f.method === 0) { return new TextDecoder().decode(raw); }
    if (f.method !== 8) { throw new Error('This Word file uses a compression this page can’t read. Save it again from Word and retry.'); }
    const stream = new Blob([raw]).stream().pipeThrough(new DecompressionStream('deflate-raw'));
    return new Response(stream).text();
  };
}

function xml(text) {
  const doc = new DOMParser().parseFromString(text, 'application/xml');
  if (doc.getElementsByTagName('parsererror').length) { throw new Error('Part of this Word file is damaged. Open it in Word, save a fresh copy, and try that.'); }
  return doc;
}

const kids = (el, local) => Array.from(el.childNodes).filter(n => n.namespaceURI === W && n.localName === local);
const kid = (el, local) => (el ? kids(el, local)[0] : undefined) || null;
const wval = el => (el ? el.getAttributeNS(W, 'val') : null);

// Word's automatic numbering ("1." and "a." that Word typed for you) is not
// in the paragraph text. It lives in numbering.xml, so rebuild the labels.
function numberingReader(numDoc, styleDoc) {
  const abstract = {}, nums = {}, styleNum = {};
  if (numDoc) {
    for (const a of numDoc.getElementsByTagNameNS(W, 'abstractNum')) {
      const levels = {};
      for (const lvl of kids(a, 'lvl')) {
        levels[lvl.getAttributeNS(W, 'ilvl')] = {
          fmt: wval(kid(lvl, 'numFmt')) || 'decimal',
          start: parseInt(wval(kid(lvl, 'start')) || '1', 10)
        };
      }
      abstract[a.getAttributeNS(W, 'abstractNumId')] = levels;
    }
    for (const n of numDoc.getElementsByTagNameNS(W, 'num')) {
      const starts = {};
      for (const o of kids(n, 'lvlOverride')) {
        const s = kid(o, 'startOverride');
        if (s) { starts[o.getAttributeNS(W, 'ilvl')] = parseInt(wval(s), 10); }
      }
      nums[n.getAttributeNS(W, 'numId')] = { abs: wval(kid(n, 'abstractNumId')), starts };
    }
  }
  if (styleDoc) {
    for (const s of styleDoc.getElementsByTagNameNS(W, 'style')) {
      const np = kid(kid(s, 'pPr'), 'numPr');
      if (np) { styleNum[s.getAttributeNS(W, 'styleId')] = { id: wval(kid(np, 'numId')), ilvl: wval(kid(np, 'ilvl')) || '0' }; }
    }
  }
  const counters = {};   // key `${abstractId}` -> per-level counts
  return function label(p) {
    const pPr = kid(p, 'pPr');
    let np = kid(pPr, 'numPr'), id = null, ilvl = '0';
    if (np) { id = wval(kid(np, 'numId')); ilvl = wval(kid(np, 'ilvl')) || '0'; }
    else {
      const st = styleNum[wval(kid(pPr, 'pStyle'))];
      if (st) { id = st.id; ilvl = st.ilvl; }
    }
    if (!id || id === '0' || !nums[id]) { return ''; }
    const levels = abstract[nums[id].abs] || {};
    const lv = levels[ilvl] || { fmt: 'decimal', start: 1 };
    const key = id;
    const c = counters[key] || (counters[key] = {});
    const start = nums[id].starts[ilvl] != null ? nums[id].starts[ilvl] : lv.start;
    c[ilvl] = c[ilvl] == null ? start : c[ilvl] + 1;
    for (const deeper of Object.keys(c)) { if (+deeper > +ilvl) { delete c[deeper]; } }
    const n = c[ilvl];
    if (lv.fmt === 'decimal') { return n + '. '; }
    if (lv.fmt === 'lowerLetter') { return String.fromCharCode(96 + ((n - 1) % 26) + 1) + '. '; }
    if (lv.fmt === 'upperLetter') { return String.fromCharCode(64 + ((n - 1) % 26) + 1) + '. '; }
    return '';
  };
}

function paragraphText(p) {
  let out = '';
  const walk = node => {
    for (const n of node.childNodes) {
      if (n.namespaceURI !== W) { continue; }
      if (n.localName === 't') { out += n.textContent; }
      else if (n.localName === 'tab' || n.localName === 'br') { out += ' '; }
      else if (n.localName === 'delText' || n.localName === 'instrText') { /* tracked deletions, field codes */ }
      else { walk(n); }
    }
  };
  walk(p);
  return out;
}

async function readDocxLines(buffer) {
  const read = await unzip(buffer);
  const docText = await read('word/document.xml');
  if (!docText) { throw new Error('That file isn’t a Word document (.docx).'); }
  const doc = xml(docText);
  const numText = await read('word/numbering.xml');
  const styleText = await read('word/styles.xml');
  const label = numberingReader(numText && xml(numText), styleText && xml(styleText));
  const body = doc.getElementsByTagNameNS(W, 'body')[0];
  const lines = [];
  let tables = 0;
  for (const el of body.childNodes) {
    if (el.namespaceURI !== W) { continue; }
    if (el.localName === 'tbl') { tables++; continue; }
    if (el.localName !== 'p') { continue; }
    const lab = label(el);
    lines.push(lab + paragraphText(el));
  }
  return { lines, tables };
}

/* ---------- turning lines into questions --------------------------------- */

const Q_RE = /^(\d+)\s*[.)]\s*(.*)$/;
const C_RE = /^(\*?)\s*([a-zA-Z])\s*[.)]\s*(\*?)\s*(.*)$/;

function parseQuiz(lines) {
  const questions = [], notes = [];
  let cur = null;
  lines.forEach((raw, i) => {
    const line = raw.replace(/\s+/g, ' ').trim();
    if (!line) { return; }
    const q = line.match(Q_RE);
    if (q && q[2]) {
      cur = { num: q[1], text: q[2], choices: [] };
      questions.push(cur);
      return;
    }
    const c = line.match(C_RE);
    if (c && cur && c[4]) {
      cur.choices.push({ letter: c[2].toLowerCase(), text: c[4], correct: !!(c[1] || c[3]) });
      return;
    }
    if (cur && !cur.choices.length) { cur.text += ' ' + line; return; }   // a question that runs onto a second line
    notes.push('Skipped a line that isn’t a question or an answer choice: “' + (line.length > 70 ? line.slice(0, 70) + '…' : line) + '”');
  });
  const seen = {};
  for (const q of questions) {
    q.problems = [];
    const right = q.choices.filter(c => c.correct).length;
    if (q.choices.length < 2) { q.problems.push({ level: 'error', text: 'Fewer than two answer choices. Check that each choice starts with a letter and a period, like “b.”' }); }
    if (!right) { q.problems.push({ level: 'error', text: 'No correct answer marked. Put an asterisk before the right letter, like “*b.”' }); }
    if (right > 1) { q.problems.push({ level: 'warn', text: 'More than one correct answer is marked, so this becomes a multiple-answer question: students must pick every marked choice, and only those.' }); }
    if (seen[q.num]) { q.problems.push({ level: 'warn', text: 'Question number ' + q.num + ' appears more than once. That’s fine for Canvas, but check you didn’t paste a question twice.' }); }
    seen[q.num] = true;
  }
  return { questions, notes };
}

/* ---------- writing the QTI package -------------------------------------- */

const QTI = 'http://www.imsglobal.org/xsd/ims_qtiasiv1p2';

function rid() {
  const b = crypto.getRandomValues(new Uint8Array(8));
  return Array.from(b, x => x.toString(16).padStart(2, '0')).join('');
}

// Canvas remembers what it imported by these ids. Fixed ids (the old
// "canvas_export", "q_1") meant a second quiz imported into the same course
// could land on top of the first. Fresh ids per export keep every quiz separate.
function buildQti(questions, title) {
  const run = 'cqc_' + rid();
  const d = document.implementation.createDocument(QTI, 'questestinterop', null);
  const root = d.documentElement;
  const add = (parent, name, attrs, text) => {
    const e = d.createElementNS(QTI, name);
    for (const k in attrs || {}) { e.setAttribute(k, attrs[k]); }
    if (text != null) { e.textContent = text; }
    parent.appendChild(e);
    return e;
  };
  const assessment = add(root, 'assessment', { ident: run, title });
  const section = add(assessment, 'section', { ident: 'root_section' });
  questions.forEach((q, qi) => {
    const qid = run + '_q' + (qi + 1);
    const correct = q.choices.map((c, ci) => c.correct ? qid + '_c' + ci : null).filter(Boolean);
    const multi = correct.length > 1;
    const item = add(section, 'item', { ident: qid, title: 'Question ' + q.num });
    const meta = add(add(item, 'itemmetadata'), 'qtimetadata');
    [['question_type', multi ? 'multiple_answers_question' : 'multiple_choice_question'],
     ['points_possible', '1.0']].forEach(([k, v]) => {
      const f = add(meta, 'qtimetadatafield');
      add(f, 'fieldlabel', null, k);
      add(f, 'fieldentry', null, v);
    });
    const pres = add(item, 'presentation');
    add(add(pres, 'material'), 'mattext', { texttype: 'text/plain' }, q.text);
    const lid = add(pres, 'response_lid', { ident: 'response1', rcardinality: multi ? 'Multiple' : 'Single' });
    const render = add(lid, 'render_choice');
    q.choices.forEach((c, ci) => {
      const lab = add(render, 'response_label', { ident: qid + '_c' + ci });
      add(add(lab, 'material'), 'mattext', { texttype: 'text/plain' }, c.text);
    });
    const rp = add(item, 'resprocessing');
    add(add(rp, 'outcomes'), 'decvar', { maxvalue: '100', minvalue: '0', varname: 'SCORE', vartype: 'Decimal' });
    if (correct.length) {
      const cond = add(rp, 'respcondition', { continue: 'No' });
      const cv = add(cond, 'conditionvar');
      if (multi) {
        const and = add(cv, 'and');
        q.choices.forEach((c, ci) => {
          const id = qid + '_c' + ci;
          if (c.correct) { add(and, 'varequal', { respident: 'response1' }, id); }
          else { add(add(and, 'not'), 'varequal', { respident: 'response1' }, id); }
        });
      } else {
        add(cv, 'varequal', { respident: 'response1' }, correct[0]);
      }
      add(cond, 'setvar', { action: 'Set', varname: 'SCORE' }, '100');
    }
  });
  const quiz = '<?xml version="1.0" encoding="UTF-8"?>\n' + new XMLSerializer().serializeToString(d);
  const manifest = '<?xml version="1.0" encoding="UTF-8"?>\n' +
    '<manifest identifier="' + run + '_manifest" xmlns="http://www.imsglobal.org/xsd/imsccv1p1/imscp_v1p1">' +
    '<resources><resource identifier="' + run + '_res" type="imsqti_xmlv1p2" href="quiz.xml">' +
    '<file href="quiz.xml"/></resource></resources></manifest>';
  return { quiz, manifest };
}

/* ---------- a minimal zip writer (stored, no compression) ---------------- */

const CRC_TABLE = (() => {
  const t = new Uint32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) { c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1; }
    t[n] = c >>> 0;
  }
  return t;
})();
function crc32(bytes) {
  let c = 0xFFFFFFFF;
  for (let i = 0; i < bytes.length; i++) { c = CRC_TABLE[(c ^ bytes[i]) & 0xFF] ^ (c >>> 8); }
  return (c ^ 0xFFFFFFFF) >>> 0;
}

function zip(entries) {   // entries: [{ name, text }]
  const enc = new TextEncoder();
  const parts = [], central = [];
  let offset = 0;
  const now = new Date();
  const dosTime = (now.getHours() << 11) | (now.getMinutes() << 5) | (now.getSeconds() >> 1);
  const dosDate = ((now.getFullYear() - 1980) << 9) | ((now.getMonth() + 1) << 5) | now.getDate();
  for (const e of entries) {
    const name = enc.encode(e.name), data = enc.encode(e.text), crc = crc32(data);
    const h = new DataView(new ArrayBuffer(30));
    h.setUint32(0, 0x04034b50, true); h.setUint16(4, 20, true); h.setUint16(6, 0x0800, true);
    h.setUint16(8, 0, true); h.setUint16(10, dosTime, true); h.setUint16(12, dosDate, true);
    h.setUint32(14, crc, true); h.setUint32(18, data.length, true); h.setUint32(22, data.length, true);
    h.setUint16(26, name.length, true); h.setUint16(28, 0, true);
    parts.push(new Uint8Array(h.buffer), name, data);
    const c = new DataView(new ArrayBuffer(46));
    c.setUint32(0, 0x02014b50, true); c.setUint16(4, 20, true); c.setUint16(6, 20, true); c.setUint16(8, 0x0800, true);
    c.setUint16(10, 0, true); c.setUint16(12, dosTime, true); c.setUint16(14, dosDate, true);
    c.setUint32(16, crc, true); c.setUint32(20, data.length, true); c.setUint32(24, data.length, true);
    c.setUint16(28, name.length, true); c.setUint32(42, offset, true);
    central.push(new Uint8Array(c.buffer), name);
    offset += 30 + name.length + data.length;
  }
  const csize = central.reduce((s, a) => s + a.length, 0);
  const end = new DataView(new ArrayBuffer(22));
  end.setUint32(0, 0x06054b50, true); end.setUint16(8, entries.length, true); end.setUint16(10, entries.length, true);
  end.setUint32(12, csize, true); end.setUint32(16, offset, true);
  return new Blob([...parts, ...central, new Uint8Array(end.buffer)], { type: 'application/zip' });
}

/* ---------- the page ------------------------------------------------------ */

const $ = id => document.getElementById(id);
const make = (tag, cls, text) => { const e = document.createElement(tag); if (cls) { e.className = cls; } if (text != null) { e.textContent = text; } return e; };

let current = null;   // { questions, notes, title }

function status(text, isError) {
  const s = $('qc-status');
  s.textContent = text;
  s.classList.toggle('cs-status--error', !!isError);
}

function render() {
  const { questions, notes } = current;
  const list = $('qc-questions');
  list.textContent = '';
  let errors = 0, warns = 0;
  questions.forEach(q => {
    const li = make('li', 'qc-q');
    li.appendChild(make('p', 'qc-q__text', q.text));
    const ol = make('ul', 'qc-choices');
    q.choices.forEach(c => {
      const ci = make('li', c.correct ? 'qc-choice qc-choice--right' : 'qc-choice');
      ci.appendChild(make('span', 'qc-letter', c.letter + '.'));
      ci.appendChild(document.createTextNode(' ' + c.text));
      if (c.correct) { ci.appendChild(make('span', 'qc-tag', 'Correct')); }
      ol.appendChild(ci);
    });
    li.appendChild(ol);
    q.problems.forEach(pr => {
      if (pr.level === 'error') { errors++; } else { warns++; }
      li.appendChild(make('p', pr.level === 'error' ? 'qc-problem qc-problem--error' : 'qc-problem', pr.text));
    });
    list.appendChild(li);
  });
  const nl = $('qc-notes');
  nl.textContent = '';
  notes.forEach(n => nl.appendChild(make('li', null, n)));
  nl.hidden = !notes.length;
  const n = questions.length;
  let summary = n + (n === 1 ? ' question' : ' questions') + ' found.';
  if (errors) { summary += ' ' + errors + (errors === 1 ? ' needs' : ' need') + ' fixing before Canvas will grade it right; see the red notes below.'; }
  else if (warns) { summary += ' Read the notes below, then download.'; }
  else { summary += ' Everything looks right. Check the correct answers, then download.'; }
  $('qc-summary').textContent = summary;
  $('qc-result').hidden = false;
}

async function load(file) {
  $('qc-result').hidden = true;
  current = null;
  if (!file) { return; }
  if (!/\.docx$/i.test(file.name)) { status('Choose a Word file that ends in .docx. Older .doc files: open in Word and Save As .docx first.', true); return; }
  status('Reading ' + file.name + '…');
  try {
    const { lines, tables } = await readDocxLines(await file.arrayBuffer());
    const parsed = parseQuiz(lines);
    if (tables) { parsed.notes.unshift('This file has ' + tables + (tables === 1 ? ' table' : ' tables') + '. Questions inside a table are skipped; write them as plain lines instead.'); }
    if (!parsed.questions.length) {
      status('No questions found in ' + file.name + '. Each question should start with a number and a period, like “1.”. Try the sample quiz to see the format.', true);
      return;
    }
    current = { ...parsed, title: file.name.replace(/\.docx$/i, '') };
    $('qc-title').value = current.title;
    status('Loaded ' + file.name + '.');
    render();
  } catch (err) {
    status(err.message || String(err), true);
  }
}

function download() {
  if (!current) { return; }
  const title = $('qc-title').value.trim() || current.title;
  const { quiz, manifest } = buildQti(current.questions, title);
  const blob = zip([{ name: 'imsmanifest.xml', text: manifest }, { name: 'quiz.xml', text: quiz }]);
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = title.replace(/[\\/:*?"<>|]+/g, '-').slice(0, 80) + ' (Canvas QTI).zip';
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(a.href), 4000);
  $('qc-done').textContent = 'Downloaded. Next, import it into Canvas (steps below).';
}

if (typeof document !== 'undefined' && $('qc-file')) {
  $('qc-file').addEventListener('change', e => load(e.target.files[0]));
  const drop = $('qc-drop');
  drop.addEventListener('dragover', e => { e.preventDefault(); drop.classList.add('cs-drop--over'); });
  drop.addEventListener('dragleave', () => drop.classList.remove('cs-drop--over'));
  drop.addEventListener('drop', e => { e.preventDefault(); drop.classList.remove('cs-drop--over'); load(e.dataTransfer.files[0]); });
  $('qc-download').addEventListener('click', download);
}

export { readDocxLines, parseQuiz, buildQti, zip, crc32, VERSION };
