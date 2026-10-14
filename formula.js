/*
 * Minimal spreadsheet formula engine for the Vonnue Cup workbook.
 *
 * Some tools (openpyxl, for one) save formulas without their calculated results. Excel and
 * Google Sheets fill those in when they open the file, but a browser has to do it itself.
 * This covers what the workbook uses: IF, IFERROR, NOT, N, COUNTIF, COUNTIFS, INDEX, MATCH,
 * LARGE, comparisons, & and arithmetic. Cells that already carry a calculated value keep it.
 */
(function (root) {
  class FErr { constructor(code) { this.code = code; } toString() { return this.code; } }
  const NA = new FErr('#N/A'), VALUE = new FErr('#VALUE!'), REF = new FErr('#REF!'), DIV0 = new FErr('#DIV/0!'), NAME = new FErr('#NAME?');
  const isErr = (v) => v instanceof FErr;

  // ---------- tokenizer
  const REF_RE = /^(?:(?:'((?:[^']|'')+)'|([A-Za-z_][A-Za-z0-9_.]*))!)?(\$?[A-Z]{1,3}\$?\d+)(?::(\$?[A-Z]{1,3}\$?\d+))?/;
  function tokenize(src) {
    const out = []; let i = 0;
    while (i < src.length) {
      const ch = src[i];
      if (ch === ' ' || ch === '\n' || ch === '\t') { i++; continue; }
      if (ch === '"') {
        let j = i + 1, s = '';
        while (j < src.length) { if (src[j] === '"') { if (src[j + 1] === '"') { s += '"'; j += 2; continue; } break; } s += src[j++]; }
        out.push({ t: 'str', v: s }); i = j + 1; continue;
      }
      const num = src.slice(i).match(/^\d+(\.\d+)?([eE][+-]?\d+)?/);
      if (num && !/^[A-Z]/.test(src.slice(i + num[0].length))) { out.push({ t: 'num', v: +num[0] }); i += num[0].length; continue; }
      const rest = src.slice(i);
      const ref = rest.match(REF_RE);
      if (ref && !/^\(/.test(rest.slice(ref[0].length))) {
        const sheet = ref[1] ? ref[1].replace(/''/g, "'") : ref[2] || null;
        out.push({ t: 'ref', sheet, a: ref[3].replace(/\$/g, ''), b: ref[4] ? ref[4].replace(/\$/g, '') : null });
        i += ref[0].length; continue;
      }
      const fn = rest.match(/^([A-Za-z][A-Za-z0-9.]*)\s*\(/);
      if (fn) { out.push({ t: 'fn', v: fn[1].toUpperCase() }); i += fn[0].length; continue; }
      const kw = rest.match(/^(TRUE|FALSE)\b/i);
      if (kw) { out.push({ t: 'bool', v: kw[1].toUpperCase() === 'TRUE' }); i += kw[0].length; continue; }
      const op = rest.match(/^(<=|>=|<>|[=<>&+\-*/(),;])/);
      if (op) { out.push({ t: 'op', v: op[1] === ';' ? ',' : op[1] }); i += op[1].length; continue; }
      throw new Error('Unsupported formula text: ' + rest.slice(0, 20));
    }
    return out;
  }

  // ---------- parser (precedence: comparison < & < + - < * / < unary)
  function parseFormula(src) {
    const toks = tokenize(src.replace(/^=/, '')); let p = 0;
    const peek = () => toks[p], take = () => toks[p++];
    const isOp = (v) => peek() && peek().t === 'op' && peek().v === v;
    function expr() { return comparison(); }
    function comparison() {
      let l = concat();
      while (peek() && peek().t === 'op' && ['=', '<>', '<', '>', '<=', '>='].includes(peek().v)) { const o = take().v; l = { k: 'bin', o, l, r: concat() }; }
      return l;
    }
    function concat() { let l = additive(); while (isOp('&')) { take(); l = { k: 'bin', o: '&', l, r: additive() }; } return l; }
    function additive() { let l = term(); while (isOp('+') || isOp('-')) { const o = take().v; l = { k: 'bin', o, l, r: term() }; } return l; }
    function term() { let l = unary(); while (isOp('*') || isOp('/')) { const o = take().v; l = { k: 'bin', o, l, r: unary() }; } return l; }
    function unary() { if (isOp('-')) { take(); return { k: 'neg', x: unary() }; } if (isOp('+')) { take(); return unary(); } return primary(); }
    function primary() {
      const t = take();
      if (!t) throw new Error('Unexpected end of formula');
      if (t.t === 'num' || t.t === 'str' || t.t === 'bool') return { k: 'lit', v: t.v };
      if (t.t === 'ref') return { k: t.b ? 'range' : 'ref', sheet: t.sheet, a: t.a, b: t.b };
      if (t.t === 'fn') {
        const args = [];
        if (!isOp(')')) { do { if (isOp(',') || isOp(')')) args.push({ k: 'lit', v: null }); else args.push(expr()); } while (isOp(',') && take()); }
        if (!isOp(')')) throw new Error('Expected ) in ' + src);
        take();
        return { k: 'fn', name: t.v, args };
      }
      if (t.t === 'op' && t.v === '(') { const e = expr(); if (!isOp(')')) throw new Error('Expected )'); take(); return e; }
      throw new Error('Unexpected token in ' + src);
    }
    const ast = expr();
    if (p < toks.length) throw new Error('Trailing tokens in ' + src);
    return ast;
  }

  // ---------- values
  const decode = (a) => { const m = a.match(/^([A-Z]+)(\d+)$/); let c = 0; for (const ch of m[1]) c = c * 26 + ch.charCodeAt(0) - 64; return { r: +m[2] - 1, c: c - 1 }; };
  const encode = (r, c) => { let s = ''; c += 1; while (c > 0) { const m = (c - 1) % 26; s = String.fromCharCode(65 + m) + s; c = Math.floor((c - 1) / 26); } return s + (r + 1); };
  const toNum = (v) => {
    if (isErr(v)) return v;
    if (v == null || v === '') return 0;
    if (typeof v === 'number') return v;
    if (typeof v === 'boolean') return v ? 1 : 0;
    const n = Number(v); return isNaN(n) ? VALUE : n;
  };
  const toStr = (v) => (v == null ? '' : typeof v === 'boolean' ? (v ? 'TRUE' : 'FALSE') : typeof v === 'number' ? String(+v.toPrecision(15)) : String(v));
  const truthy = (v) => (typeof v === 'string' ? (/^true$/i.test(v) ? true : /^false$/i.test(v) ? false : VALUE) : !!toNum(v));
  const rank = (v) => (typeof v === 'number' ? 0 : typeof v === 'string' ? 1 : typeof v === 'boolean' ? 2 : 0);
  function cmp(a, b) {
    if (a == null) a = typeof b === 'string' ? '' : typeof b === 'boolean' ? false : 0;
    if (b == null) b = typeof a === 'string' ? '' : typeof a === 'boolean' ? false : 0;
    if (rank(a) !== rank(b)) return rank(a) - rank(b);
    if (typeof a === 'string') { a = a.toLowerCase(); b = b.toLowerCase(); }
    return a < b ? -1 : a > b ? 1 : 0;
  }
  function matchesCriterion(v, crit) {
    if (isErr(v)) return false;
    let op = '=', val = crit;
    if (typeof crit === 'string') {
      const m = crit.match(/^(<=|>=|<>|=|<|>)(.*)$/);
      if (m) { op = m[1]; val = m[2]; }
      if (val !== '' && !isNaN(Number(val))) val = Number(val);
    }
    if (val === '' && (op === '=' || op === '<>')) {
      const empty = v == null || v === '';
      return op === '=' ? empty : !empty;
    }
    if (typeof val === 'number' && typeof v !== 'number') {
      if (typeof v === 'string' && v !== '' && !isNaN(Number(v)) && op === '=') v = Number(v);
      else return op === '<>';
    }
    if (typeof val === 'string' && typeof v !== 'string') return op === '<>';
    const c = cmp(v, val);
    return { '=': c === 0, '<>': c !== 0, '<': c < 0, '>': c > 0, '<=': c <= 0, '>=': c >= 0 }[op];
  }

  // ---------- evaluator bound to a workbook
  function createEvaluator(wb) {
    const memo = new Map(), busy = new Set(), asts = new Map();
    function cellValue(sheet, addr) {
      const ws = wb.Sheets[sheet];
      if (!ws) return REF;
      const key = sheet + '!' + addr;
      if (memo.has(key)) return memo.get(key);
      const c = ws[addr];
      let v;
      if (!c) v = null;
      else if (c.f && (c.t === 'z' || c.v === undefined)) {
        if (busy.has(key)) return REF;
        busy.add(key);
        try { v = evaluate(c.f, sheet); } catch (e) { v = NAME; }
        busy.delete(key);
      } else if (c.t === 'z') v = null;
      else if (c.t === 'e') v = new FErr(c.w || '#N/A');
      else v = c.v;
      memo.set(key, v);
      return v;
    }
    function rangeValues(n, sheet) {
      const s = n.sheet || sheet, a = decode(n.a), b = decode(n.b || n.a);
      const rows = [];
      for (let r = Math.min(a.r, b.r); r <= Math.max(a.r, b.r); r++) {
        const row = [];
        for (let c = Math.min(a.c, b.c); c <= Math.max(a.c, b.c); c++) row.push(cellValue(s, encode(r, c)));
        rows.push(row);
      }
      return rows;
    }
    const flat = (rows) => [].concat(...rows);
    function ev(n, sheet) {
      switch (n.k) {
        case 'lit': return n.v;
        case 'ref': return cellValue(n.sheet || sheet, n.a);
        case 'range': return { range: rangeValues(n, sheet) };
        case 'neg': { const x = toNum(ev(n.x, sheet)); return isErr(x) ? x : -x; }
        case 'bin': {
          const l = scalar(ev(n.l, sheet)), r = scalar(ev(n.r, sheet));
          if (isErr(l)) return l; if (isErr(r)) return r;
          if (n.o === '&') return toStr(l) + toStr(r);
          if (['+', '-', '*', '/'].includes(n.o)) {
            const a = toNum(l), b = toNum(r); if (isErr(a)) return a; if (isErr(b)) return b;
            if (n.o === '/' && b === 0) return DIV0;
            return n.o === '+' ? a + b : n.o === '-' ? a - b : n.o === '*' ? a * b : a / b;
          }
          const c = cmp(l, r);
          return { '=': c === 0, '<>': c !== 0, '<': c < 0, '>': c > 0, '<=': c <= 0, '>=': c >= 0 }[n.o];
        }
        case 'fn': return call(n.name, n.args, sheet);
      }
      return NAME;
    }
    const scalar = (v) => (v && v.range ? (v.range.length === 1 && v.range[0].length === 1 ? v.range[0][0] : VALUE) : v);
    function call(name, args, sheet) {
      const arg = (i) => scalar(ev(args[i], sheet));
      const rng = (i) => { const v = ev(args[i], sheet); return v && v.range ? v.range : [[v]]; };
      switch (name) {
        case 'IF': {
          const c = arg(0); if (isErr(c)) return c;
          const t = truthy(c); if (isErr(t)) return t;
          if (t) return args.length > 1 ? arg(1) : true;
          return args.length > 2 ? arg(2) : false;
        }
        case 'IFERROR': { const v = arg(0); return isErr(v) ? arg(1) : v; }
        case 'NOT': { const v = arg(0); if (isErr(v)) return v; const t = truthy(v); return isErr(t) ? t : !t; }
        case 'N': { const v = arg(0); return isErr(v) ? v : typeof v === 'number' ? v : typeof v === 'boolean' ? (v ? 1 : 0) : 0; }
        case 'COUNTIF': {
          const c = arg(1); if (isErr(c)) return c;
          return flat(rng(0)).filter((v) => matchesCriterion(v, c)).length;
        }
        case 'COUNTIFS': {
          const pairs = [];
          for (let i = 0; i < args.length; i += 2) { const c = arg(i + 1); if (isErr(c)) return c; pairs.push([flat(rng(i)), c]); }
          let n = 0;
          for (let j = 0; j < pairs[0][0].length; j++) if (pairs.every(([vals, c]) => matchesCriterion(vals[j], c))) n++;
          return n;
        }
        case 'INDEX': {
          const r = rng(0), ri = toNum(arg(1)), ci = args.length > 2 ? toNum(arg(2)) : 1;
          if (isErr(ri)) return ri; if (isErr(ci)) return ci;
          const row = r.length === 1 && args.length === 2 ? 1 : ri, col = r.length === 1 && args.length === 2 ? ri : ci;
          const v = r[row - 1] && r[row - 1][col - 1];
          return v === undefined ? REF : v;
        }
        case 'MATCH': {
          const want = arg(0), list = flat(rng(1)), type = args.length > 2 ? toNum(arg(2)) : 1;
          if (isErr(want)) return want;
          if (type === 0) { const i = list.findIndex((v) => !isErr(v) && v != null && cmp(v, want) === 0); return i < 0 ? NA : i + 1; }
          let best = -1;
          list.forEach((v, i) => { if (v != null && !isErr(v) && (type > 0 ? cmp(v, want) <= 0 : cmp(v, want) >= 0)) best = i; });
          return best < 0 ? NA : best + 1;
        }
        case 'LARGE': {
          const k = toNum(arg(1)); if (isErr(k)) return k;
          const nums = flat(rng(0)).filter((v) => typeof v === 'number').sort((a, b) => b - a);
          return k >= 1 && k <= nums.length ? nums[k - 1] : new FErr('#NUM!');
        }
      }
      return NAME;
    }
    function evaluate(f, sheet) {
      if (!asts.has(f)) asts.set(f, parseFormula(f));
      return scalar(ev(asts.get(f), sheet));
    }
    return { cellValue, evaluate };
  }

  // Fill in every formula cell that has no calculated value. Returns how many were computed.
  function recalcMissing(wb) {
    const E = createEvaluator(wb);
    const todo = [];
    wb.SheetNames.forEach((sn) => {
      const ws = wb.Sheets[sn];
      Object.keys(ws).forEach((k) => { if (k[0] !== '!' && ws[k].f && (ws[k].t === 'z' || ws[k].v === undefined)) todo.push([sn, k]); });
    });
    const results = todo.map(([sn, k]) => [sn, k, E.cellValue(sn, k)]);
    results.forEach(([sn, k, v]) => {
      const c = wb.Sheets[sn][k];
      if (isErr(v)) { c.t = 'e'; c.v = undefined; c.w = v.code; }
      else if (v == null) { c.t = 's'; c.v = ''; }
      else if (typeof v === 'number') { c.t = 'n'; c.v = v; }
      else if (typeof v === 'boolean') { c.t = 'b'; c.v = v; }
      else { c.t = 's'; c.v = v; }
    });
    // Remove empty stub cells so the workbook looks like a normally saved one
    wb.SheetNames.forEach((sn) => { const ws = wb.Sheets[sn]; Object.keys(ws).forEach((k) => { if (k[0] !== '!' && ws[k].t === 'z') delete ws[k]; }); });
    return todo.length;
  }

  const api = { parseFormula, createEvaluator, recalcMissing, FErr };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.VonnueFormula = api;
})(typeof window !== 'undefined' ? window : globalThis);
