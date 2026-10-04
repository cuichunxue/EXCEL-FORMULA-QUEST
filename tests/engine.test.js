// node tests/engine.test.js — 式エンジンとミッションデータの整合性チェック
const assert = require('assert');
const F = require('../js/formula.js');
require('../js/data.js');
const D = globalThis.EFQData;
const wb = D.workbook;
let n = 0;
const ev = (f, s) => F.evaluate(f, wb, s).value;
function eq(f, s, exp) { const v = ev(f, s); assert.ok(F.sameValue(v, exp), `${s} ${f} => ${F.fmt(v)} (expected ${F.fmt(exp)})`); n++; }

eq('=SUM(C:C)', '本日', 1248);
eq('=SUMIF(B:B,"B",D:D)', '生産実績', 1428);
eq('=SUMIFS(D:D,B:B,"B",C:C,"10月")', '生産実績', 1248);
eq('=sumifs(d:d, b:b, "b", c:c, "10月")', '生産実績', 1248);
eq('==SUMIFS(D:D,B:B,"B",C:C,"10月")', '生産実績', 1248);
eq('＝ＳＵＭＩＦＳ（Ｄ：Ｄ，Ｂ：Ｂ，"B"，Ｃ：Ｃ，"10月"）', '生産実績', 1248);
eq('=SUMIFS(D:D,B:B,"C",C:C,"10月")', '生産実績', 980);
eq('=SUMIFS(D:D,B:B,"A",C:C,"10月")', '生産実績', 1252);
eq('=COUNTIFS(B:B,"A",C:C,"10月")', '生産実績', 3);
eq('=SUMIFS(D2:D12,B2:B12,"B",C2:C12,"10月")', '生産実績', 738);
eq('=SUM(IF((B2:B15="B")*(C2:C15="10月"),D2:D15))', '生産実績', 1248);
eq('=SUMIF(B:B,"C",D:D)', '生産実績', 1480);
eq('=SUMIF(B2:B15,"B",D3:D16)', '生産実績', 260 + 275 + 305 + 400); // 範囲ずれ
eq('=XLOOKUP(A2,マスター!A:A,マスター!B:B)', '出荷指示', 'ブラケット');
eq('=XLOOKUP("A-003",マスター!A:A,マスター!C:C)', 'マスター', '加工');
eq('=VLOOKUP("A-003",マスター!A:E,2,FALSE)', '生産実績', 'ハウジング');
eq('=IFERROR(XLOOKUP("B-003",マスター!A:A,マスター!E:E),0)', 'マスター', 0);
assert.strictEqual(ev('=XLOOKUP(A5,マスター!A:A,マスター!B:B)', '出荷指示').e, '#N/A'); n++;
assert.strictEqual(ev('=XLOOKUP(A6,マスター!A:A,マスター!B:B)', '出荷指示').e, '#N/A'); n++;
eq('=XLOOKUP("B-003",マスター!A:A,マスター!E:E,"マスター未登録")', 'マスター', 'マスター未登録');
eq('=COUNTIFS(生産実績!B:B,"A")', '月次レポート', 5);
eq('=COUNTIFS(生産実績!B:B,"A",生産実績!C:C,"10月")', '月次レポート', 3);
eq('=IF(B2>=400,"OK","NG")', '検査結果', 'OK');
eq('=IF(B4>400,"OK","NG")', '検査結果', 'NG');
eq('=COUNTIF(D:D,"外観")', '不良実績', 3);
eq('=COUNT(D:D)', '不良実績', 0);
eq('=COUNTA(D:D)', '不良実績', 8);
eq('=COUNTIFS(B:B,"B",D:D,"外観")', '不良実績', 1);
eq('=COUNTIF(A:A,"Bライン")', '受入データ', 1);
eq('=TRIM(A4)', '受入データ', 'Bライン');
eq('=LEFT(C2,1)', '受入データ', 'A');
eq('=SUMIFS(F:F,C:C,"Bライン",E:E,"10月")', '10月実績(最終)', 320);
eq('=COUNTIF(B:B,">=400")', '検査結果', 3);
eq('=SUMIF(B:B,"<>B",D:D)', '生産実績', 4890 - 1428);
assert.ok(F.evaluate('=SUMIFS(D:D,B:B,"B"', wb, '生産実績').syntax); n++;
assert.strictEqual(ev('=FOO(1)', '生産実績').e, '#NAME?'); n++;
assert.strictEqual(ev('=SUMIFS(D:D,B2:B5,"B")', '生産実績').e, '#VALUE!'); n++;
// X-RAY trace
const t = F.evaluate('=SUMIFS(D:D,B:B,"B",C:C,"10月")', wb, '生産実績').trace[0];
assert.deepStrictEqual(t.steps[0].rows, [5, 8, 11, 14]); assert.deepStrictEqual(t.rows, [8, 11, 14]); n++;
const ex = F.explain('=SUMIFS(D:D,B:B,"B",C:C,"10月")', wb, '生産実績');
assert.deepStrictEqual(ex.pairs, ['ライン = B', '月 = 10月']); n++;

// mission data: build answers & template answers must produce expected values
for (const [id, m] of Object.entries(D.missions)) m.steps.forEach((s, i) => {
  const chk = (f, exp, sh) => { const v = ev(f, sh); assert.ok(F.sameValue(v, exp), `${id}#${i} ${f} => ${F.fmt(v)} expected ${F.fmt(exp)}`); n++; };
  if (s.type === 'build' || s.type === 'repair') {
    chk(s.hints[4], s.expect, s.sheet);
    if (s.template) chk('=' + s.template.fn + '(' + s.template.slots.map((x) => x.ans[0]).join(',') + ')', s.expect, s.sheet);
  }
  if (s.type === 'detect') { assert.ok(!F.sameValue(ev(s.formula, s.sheet), s.correct), `${id}#${i} detect formula should be wrong`); n++; if (s.repairTo) chk(s.repairTo.hints[4], s.repairTo.expect, s.sheet); }
  if (s.type === 'ai') s.rounds.forEach((r) => { const v = ev(r.formula, r.sheet || s.sheet); assert.ok(!F.isErr(v), `${id}#${i} ai formula errors`); n++; if (r.fix) assert.ok(!F.isErr(ev(r.fix, r.sheet || s.sheet))); });
  if (s.type === 'verify') { const tr = F.evaluate(s.formula, wb, s.sheet).trace[0]; assert.deepStrictEqual(tr.rows, s.rowsOk, `${id}#${i} verify rows`); n++; }
});
console.log('OK', n, 'assertions');
