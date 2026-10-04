/* EXCEL FORMULA QUEST — mini formula engine
 * Excelの式を解析・評価し、X-RAY用の「どの行を使ったか」トレースを返す。
 * ブラウザでは window.EFQFormula、Node では module.exports として使える。
 */
(function (root) {
  'use strict';

  // ---------- errors / helpers ----------
  function Err(code, why) { this.e = code; this.why = why || ''; }
  Err.prototype.toString = function () { return this.e; };
  const isErr = (v) => v instanceof Err;
  const isArr = (v) => v && typeof v === 'object' && Array.isArray(v.arr);
  const isRange = (v) => v && typeof v === 'object' && v.kind === 'range';
  const isEmpty = (v) => v === null || v === undefined || v === '';

  function colToNum(c) { let n = 0; for (const ch of c) n = n * 26 + (ch.charCodeAt(0) - 64); return n; }
  function numToCol(n) { let s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; }

  // 全角ASCII→半角（文字列リテラルの中は触らない）
  function normalizeSource(src) {
    let out = '', inStr = false;
    for (let i = 0; i < src.length; i++) {
      let ch = src[i];
      if (!inStr) {
        if (ch === '“' || ch === '”' || ch === '＂') ch = '"';
        const code = ch.charCodeAt(0);
        if (code >= 0xff01 && code <= 0xff5e && ch !== '＂') ch = String.fromCharCode(code - 0xfee0);
        if (ch === '　') ch = ' ';
        if (ch === '，' || ch === '、') ch = ',';
      } else if (ch === '“' || ch === '”') ch = '"';
      if (ch === '"') inStr = !inStr;
      out += ch;
    }
    return out;
  }

  // ---------- tokenizer ----------
  const SHEET_RE = /^(?:'([^']+)'|([^\s!(),"=<>+\-*/&:'$]+))!/;
  const REF_RE = /^\$?([A-Za-z]{1,3})(\$?(\d+))?(?::\$?([A-Za-z]{1,3})(\$?(\d+))?)?(?![A-Za-z0-9_(])/;

  function tokenize(src) {
    const toks = [];
    let i = 0;
    while (i < src.length) {
      const ch = src[i];
      if (/\s/.test(ch)) { i++; continue; }
      const rest = src.slice(i);
      if (ch === '"') {
        let j = i + 1, s = '';
        while (j < src.length) {
          if (src[j] === '"') { if (src[j + 1] === '"') { s += '"'; j += 2; continue; } break; }
          s += src[j++];
        }
        if (j >= src.length) throw new SyntaxError('文字列の " が閉じていません');
        toks.push({ t: 'str', v: s, p: i, end: j + 1 }); i = j + 1; continue;
      }
      let m = /^\d+(\.\d+)?/.exec(rest);
      if (m && !/^\d+:/.test(rest)) { toks.push({ t: 'num', v: parseFloat(m[0]), p: i, end: i + m[0].length }); i += m[0].length; continue; }
      const two = rest.slice(0, 2);
      if (two === '<=' || two === '>=' || two === '<>') { toks.push({ t: 'op', v: two, p: i, end: i + 2 }); i += 2; continue; }
      if ('+-*/&=<>'.includes(ch)) { toks.push({ t: 'op', v: ch, p: i, end: i + 1 }); i++; continue; }
      if ('(),'.includes(ch)) { toks.push({ t: ch, p: i, end: i + 1 }); i++; continue; }
      // function name: 識別子の直後に ( があれば関数
      m = /^([A-Za-z][A-Za-z0-9.]*)\s*\(/.exec(rest);
      if (m) { toks.push({ t: 'fn', v: m[1].toUpperCase(), p: i, end: i + m[1].length }); i += m[0].length - 1; while (src[i] !== '(') i++; continue; }
      // sheet!ref
      let sheet = null, off = 0;
      const sm = SHEET_RE.exec(rest);
      if (sm) { sheet = sm[1] || sm[2]; off = sm[0].length; }
      const rm = REF_RE.exec(rest.slice(off));
      if (rm) {
        const c1 = rm[1].toUpperCase(), r1 = rm[3] ? +rm[3] : null;
        const c2 = (rm[4] || rm[1]).toUpperCase(), r2 = rm[4] ? (rm[6] ? +rm[6] : null) : r1;
        const len = off + rm[0].length;
        if (!(r1 === null && !rm[4])) { // "A" 単体は参照ではない
          toks.push({ t: 'ref', sheet, c1, r1, c2, r2, p: i, end: i + len }); i += len; continue;
        }
      }
      m = /^(TRUE|FALSE)(?![A-Za-z0-9])/i.exec(rest);
      if (m) { toks.push({ t: 'bool', v: m[1].toUpperCase() === 'TRUE', p: i, end: i + m[1].length }); i += m[1].length; continue; }
      throw new SyntaxError('読めない文字があります: 「' + rest.slice(0, 8) + '」');
    }
    return toks;
  }

  // ---------- parser ----------
  function parse(srcRaw) {
    let src = normalizeSource(String(srcRaw || '').trim());
    src = src.replace(/^[=＝]+\s*/, '');
    if (!src) throw new SyntaxError('式が空です');
    const toks = tokenize(src);
    let k = 0;
    const peek = () => toks[k];
    const next = () => toks[k++];
    const expect = (t) => { const x = next(); if (!x || x.t !== t) throw new SyntaxError(t === ')' ? 'カッコ ) が足りません' : '「' + t + '」が必要です'); return x; };

    function cmp() {
      let l = concat();
      while (peek() && peek().t === 'op' && ['=', '<>', '<', '>', '<=', '>='].includes(peek().v)) { const o = next().v; l = { k: 'bin', o, l, r: concat(), p: l.p, end: 0 }; }
      return l;
    }
    function concat() { let l = add(); while (peek() && peek().t === 'op' && peek().v === '&') { next(); l = { k: 'bin', o: '&', l, r: add(), p: l.p }; } return l; }
    function add() { let l = mul(); while (peek() && peek().t === 'op' && (peek().v === '+' || peek().v === '-')) { const o = next().v; l = { k: 'bin', o, l, r: mul(), p: l.p }; } return l; }
    function mul() { let l = unary(); while (peek() && peek().t === 'op' && (peek().v === '*' || peek().v === '/')) { const o = next().v; l = { k: 'bin', o, l, r: unary(), p: l.p }; } return l; }
    function unary() { if (peek() && peek().t === 'op' && (peek().v === '-' || peek().v === '+')) { const o = next(); const x = unary(); return o.v === '-' ? { k: 'neg', x, p: o.p } : x; } return primary(); }
    function primary() {
      const t = next();
      if (!t) throw new SyntaxError('式が途中で終わっています');
      if (t.t === 'num') return { k: 'num', v: t.v, p: t.p, end: t.end };
      if (t.t === 'str') return { k: 'str', v: t.v, p: t.p, end: t.end };
      if (t.t === 'bool') return { k: 'bool', v: t.v, p: t.p, end: t.end };
      if (t.t === 'ref') return { k: 'ref', sheet: t.sheet, c1: t.c1, r1: t.r1, c2: t.c2, r2: t.r2, p: t.p, end: t.end };
      if (t.t === '(') { const x = cmp(); expect(')'); return x; }
      if (t.t === 'fn') {
        expect('(');
        const args = [];
        if (peek() && peek().t === ')') { const c = next(); return { k: 'call', fn: t.v, args, p: t.p, end: c.end }; }
        for (;;) {
          if (peek() && (peek().t === ',' || peek().t === ')')) args.push({ k: 'missing', p: peek().p, end: peek().p });
          else { const s = peek().p; const a = cmp(); a.p = s; a.end = toks[k - 1].end; args.push(a); }
          const sep = next();
          if (!sep) throw new SyntaxError('カッコ ) が足りません');
          if (sep.t === ')') return { k: 'call', fn: t.v, args, p: t.p, end: sep.end };
          if (sep.t !== ',') throw new SyntaxError('引数の区切りはカンマ , です');
        }
      }
      throw new SyntaxError('「' + (t.v || t.t) + '」の位置がおかしいです');
    }
    const ast = cmp();
    if (k < toks.length) throw new SyntaxError('余分な文字があります: 「' + src.slice(toks[k].p, toks[k].p + 6) + '」');
    ast.src = src;
    return ast;
  }

  // ---------- workbook access ----------
  function getSheet(ctx, name) {
    const s = name ? ctx.wb[name] : ctx.wb[ctx.sheet];
    if (!s) return null;
    return s;
  }
  function maxRow(sheet) { return sheet.rows.length + 1; }
  function cellValue(sheet, col, row) {
    if (row === 1) return sheet.headers[colToNum(col) - 1] ?? null;
    const r = sheet.rows[row - 2];
    if (!r) return null;
    const v = r[colToNum(col) - 1];
    if (v && typeof v === 'object' && v.err) return new Err(v.err);
    return v === undefined ? null : v;
  }
  function rangeOf(ctx, node) {
    const sheet = getSheet(ctx, node.sheet);
    if (!sheet) return new Err('#REF!', 'シート「' + node.sheet + '」がありません');
    const full = node.r1 === null;
    const r1 = full ? 2 : node.r1, r2 = full ? maxRow(sheet) : node.r2;
    return { kind: 'range', sheetName: node.sheet || ctx.sheet, sheet, c1: node.c1, c2: node.c2, r1: Math.min(r1, r2), r2: Math.max(r1, r2), full };
  }
  function rangeCells(rg) {
    const out = [];
    const a = colToNum(rg.c1), b = colToNum(rg.c2);
    for (let r = rg.r1; r <= rg.r2; r++) for (let c = Math.min(a, b); c <= Math.max(a, b); c++) {
      const col = numToCol(c);
      out.push({ col, row: r, v: cellValue(rg.sheet, col, r) });
    }
    return out;
  }
  const rangeSize = (rg) => (rg.r2 - rg.r1 + 1) * (Math.abs(colToNum(rg.c2) - colToNum(rg.c1)) + 1);

  // ---------- coercion ----------
  function toScalar(v) {
    if (isRange(v)) {
      if (rangeSize(v) === 1) return rangeCells(v)[0].v;
      return new Err('#VALUE!', '範囲を1つの値として使っています');
    }
    if (isArr(v)) return v.arr.length ? v.arr[0] : null;
    return v;
  }
  function toNum(v) {
    if (isErr(v)) return v;
    if (typeof v === 'number') return v;
    if (typeof v === 'boolean') return v ? 1 : 0;
    if (isEmpty(v)) return 0;
    const n = Number(String(v).trim());
    if (String(v).trim() !== '' && !isNaN(n)) return n;
    return new Err('#VALUE!', '文字「' + v + '」を数値として計算しています');
  }
  function toStr(v) { if (isEmpty(v)) return ''; if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE'; return String(v); }
  function toBool(v) { if (isErr(v)) return v; if (typeof v === 'boolean') return v; if (typeof v === 'number') return v !== 0; if (isEmpty(v)) return false; const s = String(v).toUpperCase(); if (s === 'TRUE') return true; if (s === 'FALSE') return false; return new Err('#VALUE!'); }
  function toList(v) { // 範囲/配列 → 値の配列
    if (isRange(v)) return rangeCells(v).map((c) => c.v);
    if (isArr(v)) return v.arr;
    return [v];
  }

  // ---------- criteria ----------
  function parseCriterion(c) {
    if (typeof c === 'number') return { op: '=', num: c, raw: String(c) };
    if (typeof c === 'boolean') return { op: '=', bool: c, raw: c ? 'TRUE' : 'FALSE' };
    let s = toStr(c), op = '=';
    const m = /^(<=|>=|<>|=|<|>)/.exec(s);
    if (m) { op = m[1]; s = s.slice(m[1].length); }
    const n = s.trim() !== '' && !isNaN(Number(s)) ? Number(s) : null;
    return { op, str: s, num: n, raw: toStr(c) };
  }
  function wildRe(s) {
    let re = '';
    for (let i = 0; i < s.length; i++) {
      const ch = s[i];
      if (ch === '~' && i + 1 < s.length) { re += s[++i].replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); continue; }
      if (ch === '*') re += '[\\s\\S]*'; else if (ch === '?') re += '[\\s\\S]'; else re += ch.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    }
    return new RegExp('^' + re + '$', 'i');
  }
  function matchCrit(v, cr) {
    if (isErr(v)) return false;
    if (cr.bool !== undefined) return v === cr.bool;
    const cellNum = typeof v === 'number' ? v : (typeof v === 'string' && v.trim() !== '' && !isNaN(Number(v)) ? Number(v) : null);
    if (cr.op === '=' || cr.op === '<>') {
      let eq;
      if (cr.num !== null && cr.num !== undefined) eq = cellNum !== null && cellNum === cr.num;
      else if (cr.str === '') eq = isEmpty(v);
      else eq = typeof v === 'string' && (/[*?~]/.test(cr.str) ? wildRe(cr.str).test(v) : v.toLowerCase() === cr.str.toLowerCase());
      return cr.op === '=' ? eq : !eq;
    }
    if (cr.num !== null && cr.num !== undefined) {
      if (typeof v !== 'number') return false;
      return cr.op === '<' ? v < cr.num : cr.op === '>' ? v > cr.num : cr.op === '<=' ? v <= cr.num : v >= cr.num;
    }
    if (typeof v !== 'string') return false;
    const a = v.toLowerCase(), b = cr.str.toLowerCase();
    return cr.op === '<' ? a < b : cr.op === '>' ? a > b : cr.op === '<=' ? a <= b : a >= b;
  }
  function critLabel(rg, cr) {
    const h = rg.sheet.headers[colToNum(rg.c1) - 1] || rg.c1 + '列';
    const op = cr.op === '=' ? ' = ' : cr.op === '<>' ? ' ≠ ' : ' ' + cr.op + ' ';
    return h + op + (cr.str !== undefined ? (cr.str === '' ? '(空白)' : cr.str) : cr.raw);
  }

  // ---------- evaluator ----------
  function evaluate(src, wb, sheet) {
    const ctx = { wb, sheet, trace: [] };
    let ast;
    try { ast = parse(src); } catch (e) { return { ok: false, syntax: true, message: e.message, value: new Err('#NAME?', e.message), trace: [] }; }
    let value;
    try { value = toScalarFinal(ev(ast, ctx)); } catch (e) { value = new Err('#VALUE!', e.message); }
    return { ok: !isErr(value), value, trace: ctx.trace, ast, fns: listFns(ast), refs: listRefs(ast) };
  }
  function toScalarFinal(v) { if (isRange(v)) return toScalar(v); if (isArr(v)) return v.arr.length === 1 ? v.arr[0] : v.arr[0]; return v; }
  function listFns(ast, out = []) { if (!ast) return out; if (ast.k === 'call') { out.push(ast.fn); ast.args.forEach((a) => listFns(a, out)); } if (ast.k === 'bin') { listFns(ast.l, out); listFns(ast.r, out); } if (ast.k === 'neg') listFns(ast.x, out); return out; }
  function listRefs(ast, out = []) { if (!ast) return out; if (ast.k === 'ref') out.push(ast); if (ast.k === 'call') ast.args.forEach((a) => listRefs(a, out)); if (ast.k === 'bin') { listRefs(ast.l, out); listRefs(ast.r, out); } if (ast.k === 'neg') listRefs(ast.x, out); return out; }

  function binop(o, a, b) {
    if (isErr(a)) return a; if (isErr(b)) return b;
    if (o === '&') return toStr(a) + toStr(b);
    if (['=', '<>', '<', '>', '<=', '>='].includes(o)) {
      let x = a, y = b;
      if (isEmpty(x)) x = typeof y === 'number' ? 0 : ''; if (isEmpty(y)) y = typeof x === 'number' ? 0 : '';
      let c;
      if (typeof x === 'number' && typeof y === 'number') c = x - y;
      else if (typeof x === 'string' && typeof y === 'string') { const p = x.toLowerCase(), q = y.toLowerCase(); c = p < q ? -1 : p > q ? 1 : 0; }
      else if (typeof x === 'boolean' && typeof y === 'boolean') c = (x ? 1 : 0) - (y ? 1 : 0);
      else c = (typeof x === 'number' ? 0 : typeof x === 'string' ? 1 : 2) - (typeof y === 'number' ? 0 : typeof y === 'string' ? 1 : 2);
      return o === '=' ? c === 0 : o === '<>' ? c !== 0 : o === '<' ? c < 0 : o === '>' ? c > 0 : o === '<=' ? c <= 0 : c >= 0;
    }
    const x = toNum(a), y = toNum(b);
    if (isErr(x)) return x; if (isErr(y)) return y;
    if (o === '+') return x + y; if (o === '-') return x - y; if (o === '*') return x * y;
    if (o === '/') return y === 0 ? new Err('#DIV/0!', '0で割っています') : x / y;
    return new Err('#VALUE!');
  }
  function broadcast(o, a, b) {
    const A = isRange(a) && rangeSize(a) > 1 ? { arr: toList(a) } : isRange(a) ? toScalar(a) : a;
    const B = isRange(b) && rangeSize(b) > 1 ? { arr: toList(b) } : isRange(b) ? toScalar(b) : b;
    if (isArr(A) || isArr(B)) {
      const n = Math.max(isArr(A) ? A.arr.length : 1, isArr(B) ? B.arr.length : 1);
      if (isArr(A) && isArr(B) && A.arr.length !== B.arr.length) return { arr: Array(n).fill(new Err('#N/A', '配列の大きさが違います')) };
      const out = [];
      for (let i = 0; i < n; i++) out.push(binop(o, isArr(A) ? A.arr[i] : A, isArr(B) ? B.arr[i] : B));
      return { arr: out };
    }
    return binop(o, A, B);
  }

  function ev(n, ctx) {
    switch (n.k) {
      case 'num': case 'str': case 'bool': return n.v;
      case 'missing': return null;
      case 'ref': return rangeOf(ctx, n);
      case 'neg': { const v = ev(n.x, ctx); if (isArr(v) || (isRange(v) && rangeSize(v) > 1)) return { arr: toList(v).map((x) => { const y = toNum(x); return isErr(y) ? y : -y; }) }; const y = toNum(toScalar(v)); return isErr(y) ? y : -y; }
      case 'bin': return broadcast(n.o, ev(n.l, ctx), ev(n.r, ctx));
      case 'call': {
        const f = FN[n.fn];
        if (!f) return new Err('#NAME?', '「' + n.fn + '」という関数はありません');
        return f(n.args, ctx, n);
      }
    }
    return new Err('#VALUE!');
  }
  const evS = (a, ctx) => { const v = ev(a, ctx); return isRange(v) ? toScalar(v) : v; };
  function needRange(v, name) { if (isErr(v)) return v; if (!isRange(v)) return new Err('#VALUE!', name + 'には範囲（例: B:B）を指定します'); return null; }
  const firstErr = (list) => list.find(isErr);

  function aggNumbers(args, ctx) {
    const nums = []; let err = null;
    for (const a of args) {
      if (a.k === 'missing') continue;
      const v = ev(a, ctx);
      if (isErr(v)) { err = err || v; continue; }
      if (isRange(v) || isArr(v)) { for (const x of toList(v)) { if (isErr(x)) err = err || x; else if (typeof x === 'number') nums.push(x); } }
      else { const x = toNum(v); if (isErr(x)) err = err || x; else nums.push(x); }
    }
    return { nums, err };
  }
  function traceSimple(fnName, args, ctx, result) {
    const rgs = args.filter((a) => a.k === 'ref').map((a) => rangeOf(ctx, a)).filter(isRange);
    if (!rgs.length) return;
    const rg = rgs[0];
    const rows = [], values = [];
    for (const c of rangeCells(rg)) { if (typeof c.v === 'number') { rows.push(c.row); values.push(c.v); } }
    ctx.trace.push({ fn: fnName, sheet: rg.sheetName, valueCol: rg.c1, scanRows: rangeCells(rg).map((c) => c.row), steps: [], rows, values, skipped: rangeCells(rg).filter((c) => !isEmpty(c.v) && typeof c.v !== 'number').map((c) => ({ row: c.row, v: c.v })), result });
  }

  function ifsCore(ctx, sumArg, pairs, fnName, mode) {
    let sumRg = null;
    if (sumArg) { sumRg = ev(sumArg, ctx); const e = needRange(sumRg, '合計範囲'); if (e) return e; }
    const conds = [];
    for (const [ra, ca] of pairs) {
      if (!ra || !ca) return new Err('#VALUE!', '条件範囲と条件はペアで指定します');
      const rg = ev(ra, ctx); const e = needRange(rg, '条件範囲'); if (e) return e;
      let c = evS(ca, ctx); if (isErr(c)) return c;
      conds.push({ rg, cr: parseCriterion(c) });
    }
    const n = conds.length ? rangeSize(conds[0].rg) : 0;
    for (const c of conds) if (rangeSize(c.rg) !== n) return new Err('#VALUE!', '範囲の大きさ（行数）がそろっていません');
    if (sumRg && rangeSize(sumRg) !== n) return new Err('#VALUE!', '合計範囲と条件範囲の行数がそろっていません');
    const cellsList = conds.map((c) => rangeCells(c.rg));
    const sumCells = sumRg ? rangeCells(sumRg) : null;
    let alive = Array.from({ length: n }, (_, i) => i);
    const steps = [];
    conds.forEach((c, ci) => {
      alive = alive.filter((i) => matchCrit(cellsList[ci][i].v, c.cr));
      steps.push({ col: c.rg.c1, sheet: c.rg.sheetName, label: critLabel(c.rg, c.cr), rows: alive.map((i) => cellsList[0][i].row) });
    });
    const rows = [], values = [], skipped = [];
    let total = 0, cnt = 0, err = null;
    for (const i of alive) {
      if (mode === 'count') { cnt++; rows.push(cellsList[0][i].row); continue; }
      const v = sumCells[i].v;
      if (isErr(v)) { err = err || v; continue; }
      if (typeof v === 'number') { total += v; cnt++; rows.push(sumCells[i].row); values.push(v); }
      else if (!isEmpty(v)) skipped.push({ row: sumCells[i].row, v });
    }
    let result = mode === 'count' ? cnt : mode === 'avg' ? (cnt ? total / cnt : new Err('#DIV/0!', '条件に合う行がありません')) : total;
    if (err) result = err;
    ctx.trace.push({ fn: fnName, sheet: (sumRg || conds[0].rg).sheetName, valueCol: sumRg ? sumRg.c1 : null, rowOffset: sumRg ? sumRg.r1 - conds[0].rg.r1 : 0, scanRows: cellsList[0].map((c) => c.row), steps, rows, values, skipped, result });
    return result;
  }

  function lookupCore(ctx, val, look, ret, ifnf, fnName, retColIdx) {
    if (isErr(val)) return val;
    const lc = rangeCells(look);
    let idx = -1;
    for (let i = 0; i < lc.length; i++) {
      const v = lc[i].v;
      if (typeof val === 'number' ? v === val : (typeof v === 'string' && typeof val === 'string' && v.toLowerCase() === val.toLowerCase())) { idx = i; break; }
    }
    const tr = { fn: fnName, sheet: look.sheetName, lookupCol: look.c1, value: val, scanRows: lc.map((c) => c.row), matchRow: idx >= 0 ? lc[idx].row : null, steps: [], rows: [], values: [] };
    let result;
    if (idx < 0) {
      result = ifnf !== undefined ? ifnf : new Err('#N/A', '「' + toStr(val) + '」が' + (look.sheet.headers[colToNum(look.c1) - 1] || '') + '列に見つかりません');
      tr.notFound = true;
    } else {
      const retCol = retColIdx ? numToCol(colToNum(look.c1) + retColIdx - 1) : ret.c1;
      const row = retColIdx ? lc[idx].row : ret.r1 + (lc[idx].row - look.r1);
      result = retColIdx && colToNum(retCol) > colToNum(ret.c2) ? new Err('#REF!', '列番号が表の範囲を超えています') : cellValue(retColIdx ? look.sheet : ret.sheet, retCol, row);
      if (isEmpty(result)) result = 0;
      tr.returnCol = retCol; tr.returnRow = row; tr.returnSheet = retColIdx ? look.sheetName : ret.sheetName;
    }
    tr.result = result;
    ctx.trace.push(tr);
    return result;
  }

  const FN = {
    SUM(args, ctx) { const { nums, err } = aggNumbers(args, ctx); const r = err || nums.reduce((s, x) => s + x, 0); traceSimple('SUM', args, ctx, r); return r; },
    AVERAGE(args, ctx) { const { nums, err } = aggNumbers(args, ctx); const r = err || (nums.length ? nums.reduce((s, x) => s + x, 0) / nums.length : new Err('#DIV/0!')); traceSimple('AVERAGE', args, ctx, r); return r; },
    MAX(args, ctx) { const { nums, err } = aggNumbers(args, ctx); return err || (nums.length ? Math.max(...nums) : 0); },
    MIN(args, ctx) { const { nums, err } = aggNumbers(args, ctx); return err || (nums.length ? Math.min(...nums) : 0); },
    COUNT(args, ctx) { const { nums } = aggNumbers(args, ctx); traceSimple('COUNT', args, ctx, nums.length); return nums.length; },
    COUNTA(args, ctx) { let n = 0; for (const a of args) { const v = ev(a, ctx); for (const x of toList(v)) if (!isEmpty(x)) n++; } return n; },
    COUNTBLANK(args, ctx) { const v = ev(args[0], ctx); return toList(v).filter(isEmpty).length; },
    SUMIF(args, ctx) {
      if (args.length < 2) return new Err('#VALUE!', 'SUMIFには少なくとも2つの引数が必要です');
      const rg = ev(args[0], ctx); const e = needRange(rg, '条件範囲'); if (e) return e;
      let sumNode = args[2] && args[2].k !== 'missing' ? args[2] : null;
      if (sumNode) { // Excel同様：合計範囲は条件範囲と同じ大きさに読み替える
        const sr = ev(sumNode, ctx); const e2 = needRange(sr, '合計範囲'); if (e2) return e2;
        const h = rg.r2 - rg.r1;
        sumNode = { k: 'ref', sheet: sr.sheetName === ctx.sheet ? null : sr.sheetName, c1: sr.c1, c2: sr.c1, r1: sr.r1, r2: sr.r1 + h };
        if (rg.full && sr.full) { sumNode.r1 = null; sumNode.r2 = null; }
      }
      return ifsCore(ctx, sumNode || args[0], [[args[0], args[1]]], 'SUMIF', 'sum');
    },
    SUMIFS(args, ctx) { if (args.length < 3 || args.length % 2 === 0) return new Err('#VALUE!', 'SUMIFSは (合計範囲, 条件範囲1, 条件1, …) の形です'); const pairs = []; for (let i = 1; i < args.length; i += 2) pairs.push([args[i], args[i + 1]]); return ifsCore(ctx, args[0], pairs, 'SUMIFS', 'sum'); },
    AVERAGEIFS(args, ctx) { const pairs = []; for (let i = 1; i < args.length; i += 2) pairs.push([args[i], args[i + 1]]); return ifsCore(ctx, args[0], pairs, 'AVERAGEIFS', 'avg'); },
    COUNTIF(args, ctx) { if (args.length !== 2) return new Err('#VALUE!', 'COUNTIFは (範囲, 条件) の形です'); return ifsCore(ctx, null, [[args[0], args[1]]], 'COUNTIF', 'count'); },
    COUNTIFS(args, ctx) { if (args.length < 2 || args.length % 2) return new Err('#VALUE!', 'COUNTIFSは (条件範囲1, 条件1, …) の形です'); const pairs = []; for (let i = 0; i < args.length; i += 2) pairs.push([args[i], args[i + 1]]); return ifsCore(ctx, null, pairs, 'COUNTIFS', 'count'); },
    XLOOKUP(args, ctx) {
      if (args.length < 3) return new Err('#VALUE!', 'XLOOKUPは (探す値, 探す範囲, 戻す範囲) の形です');
      const val = evS(args[0], ctx), look = ev(args[1], ctx), ret = ev(args[2], ctx);
      const e = needRange(look, '探す範囲') || needRange(ret, '戻す範囲'); if (e) return e;
      if (look.r2 - look.r1 !== ret.r2 - ret.r1) return new Err('#VALUE!', '探す範囲と戻す範囲の行数が違います');
      const ifnf = args[3] && args[3].k !== 'missing' ? evS(args[3], ctx) : undefined;
      return lookupCore(ctx, val, look, ret, ifnf, 'XLOOKUP');
    },
    VLOOKUP(args, ctx) {
      if (args.length < 3) return new Err('#VALUE!', 'VLOOKUPは (探す値, 表, 列番号, FALSE) の形です');
      const val = evS(args[0], ctx), tbl = ev(args[1], ctx), col = toNum(evS(args[2], ctx));
      const e = needRange(tbl, '表'); if (e) return e; if (isErr(col)) return col;
      const exact = args[3] ? !toBool(evS(args[3], ctx)) : false;
      const look = Object.assign({}, tbl, { c2: tbl.c1 });
      if (!exact) { // 近似一致（並び替え前提）
        const lc = rangeCells(look); let idx = -1;
        for (let i = 0; i < lc.length; i++) { const v = lc[i].v; if (isEmpty(v)) continue; if (binop('<=', v, val) === true) idx = i; }
        const tr = { fn: 'VLOOKUP', sheet: tbl.sheetName, lookupCol: tbl.c1, value: val, scanRows: lc.map((c) => c.row), matchRow: idx >= 0 ? lc[idx].row : null, approx: true, steps: [], rows: [], values: [] };
        if (idx < 0) { tr.result = new Err('#N/A'); ctx.trace.push(tr); return tr.result; }
        const rc = numToCol(colToNum(tbl.c1) + col - 1); tr.returnCol = rc; tr.returnRow = lc[idx].row; tr.returnSheet = tbl.sheetName;
        tr.result = cellValue(tbl.sheet, rc, lc[idx].row); ctx.trace.push(tr); return tr.result;
      }
      return lookupCore(ctx, val, look, tbl, undefined, 'VLOOKUP', col);
    },
    IF(args, ctx) {
      const c = ev(args[0], ctx);
      if (isArr(c) || (isRange(c) && rangeSize(c) > 1)) {
        const cl = toList(c);
        const A = args[1] ? ev(args[1], ctx) : true, B = args[2] ? ev(args[2], ctx) : false;
        const al = isArr(A) || (isRange(A) && rangeSize(A) > 1) ? toList(A) : null, bl = isArr(B) || (isRange(B) && rangeSize(B) > 1) ? toList(B) : null;
        return { arr: cl.map((x, i) => { const b = toBool(x); if (isErr(b)) return b; return b ? (al ? al[i] : toScalar(A)) : (bl ? bl[i] : toScalar(B)); }) };
      }
      const b = toBool(toScalar(c)); if (isErr(b)) return b;
      if (b) return args[1] && args[1].k !== 'missing' ? evS(args[1], ctx) : true;
      return args[2] && args[2].k !== 'missing' ? evS(args[2], ctx) : false;
    },
    IFS(args, ctx) { for (let i = 0; i + 1 < args.length; i += 2) { const b = toBool(evS(args[i], ctx)); if (isErr(b)) return b; if (b) return evS(args[i + 1], ctx); } return new Err('#N/A', 'どの条件にも当てはまりません'); },
    AND(args, ctx) { let r = true; for (const a of args) for (const x of toList(ev(a, ctx))) { const b = toBool(x); if (isErr(b)) return b; r = r && b; } return r; },
    OR(args, ctx) { let r = false; for (const a of args) for (const x of toList(ev(a, ctx))) { const b = toBool(x); if (isErr(b)) return b; r = r || b; } return r; },
    NOT(args, ctx) { const b = toBool(evS(args[0], ctx)); return isErr(b) ? b : !b; },
    IFERROR(args, ctx) {
      const v = evS(args[0], ctx);
      if (isErr(v)) { const t = ctx.trace[ctx.trace.length - 1]; if (t) t.hiddenError = v.e; ctx.hidden = v; return args[1] ? evS(args[1], ctx) : ''; }
      return v;
    },
    ISERROR(args, ctx) { return isErr(evS(args[0], ctx)); },
    ISBLANK(args, ctx) { return isEmpty(evS(args[0], ctx)); },
    TRIM(args, ctx) { const v = evS(args[0], ctx); return isErr(v) ? v : toStr(v).replace(/ +/g, ' ').trim(); },
    LEN(args, ctx) { const v = evS(args[0], ctx); return isErr(v) ? v : toStr(v).length; },
    LEFT(args, ctx) { const v = evS(args[0], ctx); const n = args[1] ? toNum(evS(args[1], ctx)) : 1; return isErr(v) ? v : isErr(n) ? n : toStr(v).slice(0, n); },
    RIGHT(args, ctx) { const v = evS(args[0], ctx); const n = args[1] ? toNum(evS(args[1], ctx)) : 1; return isErr(v) ? v : isErr(n) ? n : (n ? toStr(v).slice(-n) : ''); },
    MID(args, ctx) { const v = evS(args[0], ctx); const s = toNum(evS(args[1], ctx)), n = toNum(evS(args[2], ctx)); return firstErr([v, s, n]) || toStr(v).substr(s - 1, n); },
    SUBSTITUTE(args, ctx) { const v = evS(args[0], ctx), a = evS(args[1], ctx), b = evS(args[2], ctx); return firstErr([v, a, b]) || (toStr(a) === '' ? toStr(v) : toStr(v).split(toStr(a)).join(toStr(b))); },
    UPPER(args, ctx) { const v = evS(args[0], ctx); return isErr(v) ? v : toStr(v).toUpperCase(); },
    VALUE(args, ctx) { return toNum(evS(args[0], ctx)); },
    ROUND(args, ctx) { const v = toNum(evS(args[0], ctx)), d = args[1] ? toNum(evS(args[1], ctx)) : 0; if (isErr(v)) return v; const p = Math.pow(10, d); return Math.round(v * p) / p; },
    ABS(args, ctx) { const v = toNum(evS(args[0], ctx)); return isErr(v) ? v : Math.abs(v); },
  };

  // ---------- formatting / explanation ----------
  function fmt(v) {
    if (isErr(v)) return v.e;
    if (typeof v === 'number') return Number.isInteger(v) ? v.toLocaleString('en-US') : (Math.round(v * 100) / 100).toLocaleString('en-US');
    if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
    if (isEmpty(v)) return '(空白)';
    return String(v);
  }
  function sameValue(a, b) {
    if (isErr(a) || isErr(b)) return isErr(a) && isErr(b) && a.e === b.e;
    if (typeof a === 'number' && typeof b === 'number') return Math.abs(a - b) < 1e-9;
    return toStr(a) === toStr(b);
  }

  const ARG_LABELS = {
    SUM: ['合計するもの'], AVERAGE: ['平均するもの'], COUNT: ['数える範囲'], COUNTA: ['数える範囲'],
    SUMIF: ['条件範囲', '条件', '合計するもの'],
    SUMIFS: (i) => (i === 0 ? '合計するもの' : (i % 2 ? '条件範囲' : '条件') + '①②③④⑤'[Math.floor((i - 1) / 2)]),
    COUNTIF: ['条件範囲', '条件'],
    COUNTIFS: (i) => (i % 2 ? '条件' : '条件範囲') + '①②③④⑤'[Math.floor(i / 2)],
    XLOOKUP: ['探す値', '探す場所', '持ってくる列', '見つからない時'],
    VLOOKUP: ['探す値', '表', '何列目', '一致の方法'],
    IF: ['判定条件', '正しいとき', '違うとき'], IFERROR: ['本来の計算', 'エラーの時に出す値'],
    TRIM: ['整える文字'], LEFT: ['文字', '左から何文字'], RIGHT: ['文字', '右から何文字'], MID: ['文字', '開始位置', '文字数'], LEN: ['文字'],
    SUBSTITUTE: ['文字', '探す文字', '置き換える文字'],
  };
  function headerFor(wb, sheet, node) {
    if (!node || node.k !== 'ref') return null;
    const s = wb[node.sheet || sheet]; if (!s) return null;
    const h = s.headers[colToNum(node.c1) - 1];
    if (node.r1 !== null && node.r1 === node.r2 && node.c1 === node.c2) {
      const v = cellValue(s, node.c1, node.r1);
      return (h || node.c1 + '列') + '（' + node.c1 + node.r1 + ' = ' + fmt(v) + '）';
    }
    return (node.sheet ? node.sheet + 'の' : '') + (h || node.c1 + '列') + '（' + node.c1 + (node.r1 !== null ? node.r1 + ':' + node.c2 + node.r2 : '列') + '）';
  }
  // 式 → 意味ブロック
  function explain(src, wb, sheet) {
    let ast; try { ast = parse(src); } catch (e) { return null; }
    if (ast.k !== 'call') return { fn: null, blocks: [{ label: '計算', text: ast.src, meaning: ast.src }] };
    const top = ast;
    const labels = ARG_LABELS[top.fn];
    const blocks = top.args.map((a, i) => {
      const text = ast.src.slice(a.p, a.end);
      const label = typeof labels === 'function' ? labels(i) : (labels && labels[i]) || '引数' + (i + 1);
      let meaning = headerFor(wb, sheet, a);
      if (!meaning) meaning = a.k === 'str' ? '「' + a.v + '」' : a.k === 'num' ? String(a.v) : text;
      return { label, text, meaning, col: a.k === 'ref' ? a.c1 : null, sheet: a.k === 'ref' ? a.sheet || sheet : null, node: a };
    });
    // 条件ペアを「列 = 値」にまとめる
    const pairs = [];
    const pairStart = top.fn === 'SUMIFS' ? 1 : top.fn === 'COUNTIFS' ? 0 : top.fn === 'SUMIF' || top.fn === 'COUNTIF' ? 0 : -1;
    if (pairStart >= 0) {
      const step = 2;
      const end = top.fn === 'SUMIF' ? 2 : blocks.length;
      for (let i = pairStart; i + 1 < end; i += step) {
        const rg = blocks[i], c = blocks[i + 1];
        const s = wb[(rg.node.k === 'ref' && rg.node.sheet) || sheet];
        const h = rg.node.k === 'ref' && s ? s.headers[colToNum(rg.node.c1) - 1] || rg.node.c1 + '列' : rg.text;
        const cv = c.node.k === 'str' ? c.node.v : c.text;
        const m = /^(<=|>=|<>|=|<|>)/.exec(cv);
        pairs.push(h + (m ? ' ' + m[1].replace('<>', '≠') + ' ' + cv.slice(m[1].length) : ' = ' + cv));
      }
    }
    return { fn: top.fn, blocks, pairs };
  }

  const api = { parse, evaluate, explain, fmt, sameValue, isErr, Err, colToNum, numToCol, cellValue, normalizeSource };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.EFQFormula = api;
})(typeof window !== 'undefined' ? window : globalThis);
