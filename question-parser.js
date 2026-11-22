// Word (.docx) বা সাধারণ লেখা -> প্রশ্নসেট। admin.html এটা ব্যবহার করে।
// Word ফাইল ব্রাউজারেই পড়া হয় (JSZip দিয়ে), কোনো server এ যায় না।

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

/* ---------------------------------------------------------------
   1) .docx  ->  লাইনের তালিকা (Word এর auto numbering/lettering সহ)
---------------------------------------------------------------- */
function letters(n) { let s = ''; while (n > 0) { n--; s = String.fromCharCode(97 + (n % 26)) + s; n = Math.floor(n / 26); } return s; }
function roman(n) {
  const m = [[1000, 'M'], [900, 'CM'], [500, 'D'], [400, 'CD'], [100, 'C'], [90, 'XC'], [50, 'L'], [40, 'XL'], [10, 'X'], [9, 'IX'], [5, 'V'], [4, 'IV'], [1, 'I']];
  let s = ''; for (const [v, r] of m) while (n >= v) { s += r; n -= v; } return s;
}
function fmtNum(n, f) {
  switch (f) {
    case 'upperLetter': return letters(n).toUpperCase();
    case 'lowerLetter': return letters(n);
    case 'upperRoman': return roman(n);
    case 'lowerRoman': return roman(n).toLowerCase();
    default: return String(n);
  }
}

export function docxXmlToLines(docXml, numXml, stylesXml) {
  const parse = x => x ? new DOMParser().parseFromString(x, 'application/xml') : null;
  const doc = parse(docXml), num = parse(numXml), sty = parse(stylesXml);
  const attr = (el, n) => {
    if (!el) return null;
    let v = el.getAttributeNS(W, n);
    if (v === null || v === '') v = el.getAttribute('w:' + n);
    return v === '' ? null : v;
  };
  const kids = (el, name) => el ? [...el.children].filter(c => c.localName === name) : [];
  const kid = (el, name) => kids(el, name)[0] || null;

  // numbering.xml
  const abs = {}, nums = {};
  if (num) {
    for (const a of num.getElementsByTagNameNS(W, 'abstractNum')) {
      const id = attr(a, 'abstractNumId'); abs[id] = {};
      for (const l of kids(a, 'lvl')) {
        abs[id][attr(l, 'ilvl')] = {
          start: Number(attr(kid(l, 'start'), 'val') ?? 1),
          fmt: attr(kid(l, 'numFmt'), 'val') || 'decimal',
          text: attr(kid(l, 'lvlText'), 'val') ?? '%1.'
        };
      }
    }
    for (const n of num.getElementsByTagNameNS(W, 'num')) {
      const o = { abs: attr(kid(n, 'abstractNumId'), 'val'), over: {} };
      for (const lo of kids(n, 'lvlOverride')) {
        const s = kid(lo, 'startOverride'); if (s) o.over[attr(lo, 'ilvl')] = Number(attr(s, 'val'));
      }
      nums[attr(n, 'numId')] = o;
    }
  }
  // styles.xml: style এর ভেতরের numbering
  const styleNum = {}, styleBase = {};
  if (sty) {
    for (const s of sty.getElementsByTagNameNS(W, 'style')) {
      const id = attr(s, 'styleId'), ppr = kid(s, 'pPr'), np = ppr && kid(ppr, 'numPr');
      styleBase[id] = attr(kid(s, 'basedOn'), 'val');
      if (np) styleNum[id] = { numId: attr(kid(np, 'numId'), 'val'), ilvl: attr(kid(np, 'ilvl'), 'val') };
    }
  }
  const styleNumOf = id => { for (let i = 0; id && i < 6; i++) { if (styleNum[id]) return styleNum[id]; id = styleBase[id]; } return null; };

  const textOf = el => {
    let out = '';
    for (const c of el.children) {
      const n = c.localName;
      if (n === 'r') {
        for (const x of c.children) {
          const k = x.localName;
          if (k === 't') out += x.textContent;
          else if (k === 'tab') out += ' ';
          else if (k === 'br' || k === 'cr') { if (attr(x, 'type') !== 'page') out += '\n'; }
          else if (k === 'noBreakHyphen') out += '-';
        }
      } else if (['hyperlink', 'ins', 'smartTag', 'sdt', 'sdtContent', 'fldSimple', 'r'].includes(n)) out += textOf(c);
    }
    return out;
  };

  const counters = {}, lines = [];
  for (const p of doc.getElementsByTagNameNS(W, 'p')) {
    const ppr = kid(p, 'pPr');
    const np = ppr && kid(ppr, 'numPr');
    let numId = np ? attr(kid(np, 'numId'), 'val') : null;
    let ilvl = np ? attr(kid(np, 'ilvl'), 'val') : null;
    if (!numId && ppr) {
      const sn = styleNumOf(attr(kid(ppr, 'pStyle'), 'val'));
      if (sn) { numId = sn.numId; ilvl = sn.ilvl; }
    }
    if (ilvl == null) ilvl = '0';
    const text = textOf(p).replace(/[ \t]+$/g, '');
    let label = '';
    if (numId && numId !== '0' && nums[numId]) {
      const def = abs[nums[numId].abs], lv = def && def[ilvl];
      if (lv) {
        const cs = (counters[numId] = counters[numId] || []);
        const L = Number(ilvl);
        const startOf = l => nums[numId].over[l] ?? def[l]?.start ?? 1;
        cs[L] = cs[L] === undefined ? startOf(L) : cs[L] + 1;
        cs.length = L + 1;
        if (lv.fmt === 'bullet') label = '•';
        else label = lv.text.replace(/%(\d)/g, (_, d) => {
          const l = Number(d) - 1;
          return fmtNum(cs[l] === undefined ? startOf(l) : cs[l], def[l]?.fmt || 'decimal');
        });
      }
    }
    if (!text.trim()) continue;
    lines.push((label ? label + ' ' : '') + text);
  }
  return lines;
}

// ব্রাউজারে: ArrayBuffer (.docx) -> লেখা। window.JSZip লাগে।
export async function docxToText(buf) {
  if (!window.JSZip) throw new Error('JSZip লোড হয়নি (internet চেক করো)');
  const zip = await window.JSZip.loadAsync(buf);
  const rd = async n => zip.file(n) ? zip.file(n).async('string') : null;
  const docXml = await rd('word/document.xml');
  if (!docXml) throw new Error('এটা সঠিক .docx ফাইল না');
  return docxXmlToLines(docXml, await rd('word/numbering.xml'), await rd('word/styles.xml')).join('\n');
}

/* ---------------------------------------------------------------
   2) লেখা  ->  প্রশ্নের তালিকা
---------------------------------------------------------------- */
const DIG = '[0-9\\u09E6-\\u09EF]';
const LBL = '[A-Ja-j\\u0995\\u0996\\u0997\\u0998\\u0999\\u099A]';   // A-J, ক খ গ ঘ ঙ চ
const BN_LABEL = { 'ক': 0, 'খ': 1, 'গ': 2, 'ঘ': 3, 'ঙ': 4, 'চ': 5 };
const labelIndex = ch => (ch in BN_LABEL) ? BN_LABEL[ch] : 'ABCDEFGHIJ'.indexOf(ch.toUpperCase());

const RE_Q_PLAIN = new RegExp(`^(${DIG}{1,3})\\s*[.):]\\s*(.*)$`);
const RE_Q_PREFIX = new RegExp(`^(?:Q(?:uestion|ues)?|\\u09AA\\u09CD\\u09B0\\u09B6\\u09CD\\u09A8)\\s*\\.?\\s*(${DIG}{1,3})\\s*[.):\\-\\u2013]?\\s*(.*)$`, 'i');
const RE_OPT = new RegExp(`^\\(?(${LBL})\\s*[.):\\-\\u2013]\\s*(.+)$`);
const RE_BULLET = /^[•▪●○◦]\s*(.+)$/;
const RE_ANS = /^(?:Answer|Ans|Correct\s*Answer|Correct|উত্তর|সঠিক\s*উত্তর)\s*[:：.\-–)]\s*(.*)$/i;
const RE_ANS_LETTER = new RegExp(`^\\(?(${LBL})\\)?(?:[\\s.):,]|$)`);
const RE_MARKS = /^(?:Marks?|Points?|নম্বর)\s*[:：]\s*(\d+(?:\.\d+)?)\s*$/i;
const RE_TYPE = /^Type\s*[:：]\s*(.+)$/i;
const RE_TRAIL = /\s*[\[(（]\s*(\d+(?:\.\d+)?)\s*(?:marks?|pts?|points?|নম্বর)?\s*[\])）]\s*$/i;
const HEADERS = [
  [/^(?:Position|Designation|Post|পদ|পদবি)\s*[:：]\s*(.+)$/i, 'position'],
  [/^Title\s*[:：]\s*(.+)$/i, 'title'],
  [/^(?:Duration|Time)\s*[:：]\s*(\d+)/i, 'duration'],
  [/^(?:Show|Pick|Draw|Random)\s*[:：]\s*(\d+)/i, 'pick']
];

function matchOption(line, expected) {
  let m = line.match(RE_OPT);
  if (m && labelIndex(m[1]) === expected) return m[2].trim();
  m = line.match(RE_BULLET);
  if (m) return m[1].trim();
  return null;
}

export function parseQuestions(raw) {
  const res = { meta: {}, questions: [], errors: [], warnings: [] };
  let cur = null;

  const finish = () => {
    if (!cur) return;
    const q = cur; cur = null;
    const n = res.questions.length + 1, tag = 'Q' + n;
    let text = q.text.trim();
    const tm = text.match(RE_TRAIL);
    if (tm) { if (q.marks == null) q.marks = Number(tm[1]); text = text.slice(0, tm.index).trim(); }
    const marks = q.marks != null ? q.marks : 1;
    if (!text) res.errors.push(`${tag}: প্রশ্নের লেখা নেই`);
    if (!(marks > 0)) res.errors.push(`${tag}: marks ঠিক নেই`);
    const out = { id: 'q' + n, text, marks };
    const wantsText = q.forceType === 'text';
    if (q.options.length >= 2 && !wantsText) {
      out.type = 'mcq'; out.options = q.options;
      const a = (q.ansRaw || '').trim();
      let idx = -1;
      const lm = a.match(RE_ANS_LETTER);
      if (lm) idx = labelIndex(lm[1]);
      else if (a) idx = q.options.findIndex(o => o.toLowerCase() === a.toLowerCase());
      if (!a) res.errors.push(`${tag}: "Answer: B" ধরনের লাইন নেই`);
      else if (idx < 0 || idx >= q.options.length) res.errors.push(`${tag}: Answer "${a.slice(0, 20)}" বোঝা যায়নি (শুধু একটি অক্ষর দাও, যেমন B)`);
      else out.answer = idx;
    } else {
      if (q.options.length === 1) res.errors.push(`${tag}: মাত্র একটি option পাওয়া গেছে`);
      out.type = 'text';
      if (q.ansRaw) out.note = q.ansRaw.trim();
    }
    if (q.extra) res.warnings.push(`${tag}: Answer এর পরের অতিরিক্ত লেখা বাদ দেওয়া হয়েছে`);
    res.questions.push(out);
  };

  for (const rawLine of String(raw || '').split(/\r?\n/)) {
    const line = rawLine.replace(/[​-‍﻿]/g, '').replace(/ /g, ' ').trim();
    if (!line) continue;
    let m;

    if (!cur && !res.questions.length) {
      let hit = false;
      for (const [re, key] of HEADERS) {
        if ((m = line.match(re))) { res.meta[key] = key === 'duration' || key === 'pick' ? Number(m[1]) : m[1].trim(); hit = true; break; }
      }
      if (hit) continue;
    }

    if ((m = line.match(RE_Q_PREFIX)) || (m = line.match(RE_Q_PLAIN))) {
      finish();
      cur = { text: m[2] || '', options: [], ansRaw: null, marks: null, forceType: null, closed: false, extra: false };
      continue;
    }
    if (!cur) { res.warnings.push(`প্রশ্নের আগের লাইন বাদ গেছে: "${line.slice(0, 40)}"`); continue; }

    if ((m = line.match(RE_ANS))) { cur.ansRaw = m[1]; cur.closed = true; continue; }
    if ((m = line.match(RE_MARKS))) { cur.marks = Number(m[1]); continue; }
    if ((m = line.match(RE_TYPE))) { cur.forceType = /writ|text|essay|short|long|descr/i.test(m[1]) ? 'text' : null; continue; }

    if (cur.closed) { cur.extra = true; continue; }
    const opt = matchOption(line, cur.options.length);
    if (opt !== null) { cur.options.push(opt); continue; }
    if (cur.options.length) cur.options[cur.options.length - 1] += '\n' + line;
    else cur.text += '\n' + line;
  }
  finish();

  if (!res.questions.length) res.errors.push('কোনো প্রশ্ন পাওয়া যায়নি। প্রশ্ন "Q1." বা "1." দিয়ে শুরু হতে হবে।');
  return res;
}

/* ---------------------------------------------------------------
   3) Firestore এ রাখার জন্য ভাগ করা (উত্তর আলাদা, শুধু admin দেখে)
---------------------------------------------------------------- */
export function splitForStorage(questions) {
  const pub = [], answers = {}, notes = {}, optCounts = {};
  for (const q of questions) {
    if (q.type === 'mcq') {
      pub.push({ id: q.id, type: 'mcq', text: q.text, options: q.options, marks: q.marks });
      answers[q.id] = q.answer; optCounts[q.id] = q.options.length;
    } else {
      pub.push({ id: q.id, type: 'text', text: q.text, marks: q.marks });
      if (q.note) notes[q.id] = q.note;
    }
  }
  return {
    pub, key: { answers, notes }, qids: pub.map(q => q.id), optCounts,
    totalMarks: pub.reduce((s, q) => s + q.marks, 0)
  };
}
