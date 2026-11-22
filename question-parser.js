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
const DIG = '[0-9০-৯]';
const LBL = '[A-Ja-jকখগঘঙচ]';   // A-J, ক খ গ ঘ ঙ চ
const BN_LABEL = { 'ক': 0, 'খ': 1, 'গ': 2, 'ঘ': 3, 'ঙ': 4, 'চ': 5 };
const labelIndex = ch => (ch in BN_LABEL) ? BN_LABEL[ch] : 'ABCDEFGHIJ'.indexOf(ch.toUpperCase());

// প্রশ্ন নম্বরের পরের চিহ্ন: . ) : । (দাঁড়ি) ইত্যাদি। "1.5 kg" এর মতো দশমিক সংখ্যাকে প্রশ্ন ধরা হয় না
const SEP = '[.):।৷．。․]';
const NOT_DIGIT = '(?![0-9০-৯])';
const RE_Q_PLAIN = new RegExp(`^(${DIG}{1,3})\\s*${SEP}${NOT_DIGIT}\\s*(.*)$`);
const RE_Q_PREFIX = new RegExp(`^(?:Q(?:uestion|ues)?|প্রশ্ন)\\s*\\.?\\s*(${DIG}{1,3})\\s*(?:${SEP}|[\\-–])?${NOT_DIGIT}\\s*(.*)$`, 'i');
const RE_OPT = new RegExp(`^\\(?(${LBL})\\s*[.):\\-–।]\\s*(.+)$`);
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

// অদৃশ্য চিহ্ন বাদ (বাংলার ZWJ/ZWNJ লাইনের ভেতরে থাকতে দেওয়া হয়)
const INVIS = '​‎‏‪-‮⁠-⁤﻿­';
const cleanLine = s => s.replace(new RegExp(`[${INVIS}]`, 'g'), '').replace(/[   ]/g, ' ')
  .replace(/^[\s‌‍]+|[\s‌‍]+$/g, '');

function matchOption(line, expected) {
  let m = line.match(RE_OPT);
  if (m && labelIndex(m[1]) === expected) return m[2].trim();
  m = line.match(RE_BULLET);
  if (m) return m[1].trim();
  return null;
}

// যেভাবেই লেখা হোক, যতটা সম্ভব প্রশ্ন বের করে। MCQ ঠিকঠাক না হলে লিখিত প্রশ্ন ধরে (সতর্কতা সহ)।
export function parseQuestions(raw) {
  const res = { meta: {}, questions: [], errors: [], warnings: [] };
  const lines = String(raw || '').split(/\r?\n/).map(cleanLine).filter(Boolean);

  // ১) শুরুর header (Position/Duration...) আর প্রশ্নের টুকরো ভাগ
  const blocks = []; let cur = null, skipped = 0;
  for (const line of lines) {
    let m;
    if (!cur) {
      let hit = false;
      for (const [re, key] of HEADERS) {
        if ((m = line.match(re))) { res.meta[key] = key === 'duration' || key === 'pick' ? Number(m[1]) : m[1].trim(); hit = true; break; }
      }
      if (hit) continue;
    }
    if ((m = line.match(RE_Q_PREFIX)) || (m = line.match(RE_Q_PLAIN))) {
      cur = { head: m[2] || '', body: [] }; blocks.push(cur); continue;
    }
    if (!cur) { skipped++; res.warnings.push(`প্রশ্নের আগের লাইন বাদ গেছে: "${line.slice(0, 40)}"`); continue; }
    cur.body.push(line);
  }

  // ২) কোনো নম্বর পাওয়া না গেলে প্রতিটি লাইনকে একটি লিখিত প্রশ্ন ধরি
  if (!blocks.length) {
    const rest = lines.filter(l => !HEADERS.some(([re]) => re.test(l)));
    if (rest.length) {
      res.warnings.length = 0;
      res.warnings.push('প্রশ্নের নম্বর (1. 2. 3.) পাওয়া যায়নি, তাই প্রতিটি লাইনকে আলাদা লিখিত প্রশ্ন ধরা হয়েছে। নিচের তালিকা দেখে নাও।');
      for (const l of rest) blocks.push({ head: l, body: [] });
    }
  }

  // ৩) প্রতিটি টুকরো থেকে প্রশ্ন বানাই
  blocks.forEach((blk, i) => {
    const n = i + 1, tag = 'Q' + n;
    let marks = null, forceText = false;
    const answers = [], body = [];
    for (const line of blk.body) {
      let m;
      if ((m = line.match(RE_ANS))) { answers.push(m[1].trim()); continue; }
      if ((m = line.match(RE_MARKS))) { marks = Number(m[1]); continue; }
      if ((m = line.match(RE_TYPE))) { forceText = /writ|text|essay|short|long|descr/i.test(m[1]); continue; }
      body.push(line);
    }
    let head = blk.head.trim();
    const tm = head.match(RE_TRAIL);
    if (tm) { if (marks == null) marks = Number(tm[1]); head = head.slice(0, tm.index).trim(); }
    if (!(marks > 0)) marks = 1;

    // option খোঁজা
    const textLines = head ? [head] : [], options = [];
    for (const line of body) {
      const o = matchOption(line, options.length);
      if (o !== null) options.push(o);
      else if (options.length) options[options.length - 1] += '\n' + line;
      else textLines.push(line);
    }

    const out = { id: 'q' + n, marks };
    let isMcq = false;
    if (!forceText && options.length >= 2) {
      let idx = -1;
      if (answers.length === 1) {
        const a = answers[0];
        const lm = a.match(RE_ANS_LETTER);
        idx = lm ? labelIndex(lm[1]) : options.findIndex(o => o.toLowerCase() === a.toLowerCase());
      }
      if (idx >= 0 && idx < options.length) { isMcq = true; out.type = 'mcq'; out.text = textLines.join('\n').trim(); out.options = options; out.answer = idx; }
      else res.warnings.push(`${tag}: লিখিত প্রশ্ন ধরা হয়েছে (MCQ হলে option এর নিচে "Answer: B" ধরনের একটি লাইন দিতে হবে)`);
    }
    if (!isMcq) {
      out.type = 'text';
      out.text = [head, ...body].filter(Boolean).join('\n').trim();
      if (answers.length) out.note = answers.join('\n');
    }
    if (!out.text) { res.warnings.push(`${tag}: প্রশ্নের লেখা নেই, বাদ দেওয়া হয়েছে`); return; }
    res.questions.push(out);
  });
  // বাদ পড়লে id আবার সাজাই
  res.questions.forEach((q, i) => q.id = 'q' + (i + 1));

  if (!res.questions.length) res.errors.push('ফাইলে কোনো প্রশ্ন পাওয়া যায়নি। Word ফাইলে লেখা আছে কিনা দেখো।');
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
