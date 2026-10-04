/* EXCEL FORMULA QUEST — 架空製造企業「みらい精機」のデータとミッション定義 */
(function (root) {
  'use strict';

  // ===================== DATASET =====================
  // 生産実績：ラインB・10月 = 320 + 418 + 510 = 1,248 / ラインB(月条件なし) = 1,428
  const production = {
    headers: ['日付', 'ライン', '月', '生産数', '品番', '稼働時間'],
    rows: [
      ['2024/09/27', 'A', '9月', 380, 'A-002', 8.0],
      ['2024/09/27', 'C', '9月', 240, 'C-002', 8.0],
      ['2024/09/30', 'A', '9月', 350, 'A-001', 8.0],
      ['2024/09/30', 'B', '9月', 180, 'A-002', 4.0],
      ['2024/09/30', 'C', '9月', 260, 'C-001', 8.0],
      ['2024/10/01', 'A', '10月', 410, 'A-001', 8.0],
      ['2024/10/01', 'B', '10月', 320, 'A-002', 8.0],
      ['2024/10/01', 'C', '10月', 275, 'C-001', 8.0],
      ['2024/10/02', 'A', '10月', 390, 'A-003', 7.5],
      ['2024/10/02', 'B', '10月', 418, 'B-002', 8.0],
      ['2024/10/02', 'C', '10月', 305, 'C-002', 8.0],
      ['2024/10/03', 'A', '10月', 452, 'A-001', 8.0],
      ['2024/10/03', 'B', '10月', 510, 'A-002', 8.0],
      ['2024/10/03', 'C', '10月', 400, 'C-001', 8.0],
    ],
  };
  const today = {
    headers: ['ライン', '品番', '生産数'],
    rows: [['A', 'A-001', 300], ['B', 'A-002', 410], ['C', 'C-001', 250], ['A', 'A-003', 288]],
  };
  const master = {
    headers: ['品番', '品名', '工程', '基準値', '単価'],
    rows: [
      ['A-001', 'シャフト', '切削', 500, 120],
      ['A-002', 'ブラケット', '組立', 300, 80],
      ['A-003', 'ハウジング', '加工', 400, 150],
      ['B-002', 'カバー', '成形', 200, 60],
      ['C-001', 'ギア', '熱処理', 100, 300],
      ['C-002', 'スプリング', '組立', 200, 50],
    ],
  };
  const defect = {
    headers: ['日付', 'ライン', '品番', '不良分類', '不良数'],
    rows: [
      ['2024/10/01', 'A', 'A-001', '外観', 5],
      ['2024/10/01', 'B', 'A-002', '寸法', 12],
      ['2024/10/02', 'A', 'A-001', '機能', 3],
      ['2024/10/02', 'B', 'B-002', '外観', 8],
      ['2024/10/03', 'A', 'A-003', '寸法', 6],
      ['2024/10/03', 'C', 'C-001', '材料', 4],
      ['2024/10/03', 'B', 'A-002', '機能', 9],
      ['2024/10/04', 'C', 'C-002', '外観', 7],
    ],
  };
  const order = {
    headers: ['品番', '品名', '数量'],
    rows: [['A-002', null, 120], ['C-001', null, 40], ['A-003', null, 60], ['Ａ-001', null, 80], ['B-003', null, 30]],
  };
  const inspect = {
    headers: ['ロット', '測定値', '判定'],
    rows: [['L-01', 420, null], ['L-02', 380, null], ['L-03', 400, null], ['L-04', 455, null]],
  };
  const intake = {
    headers: ['入力ライン', '生産数', '品番'],
    rows: [['Bライン', 320, 'A-002'], ['Aライン', 410, 'A-001'], ['Bライン ', 418, 'B-002'], [' Bライン', 510, 'A-002'], ['Cライン', 275, 'C-001']],
  };
  const report = {
    headers: ['項目', '値', '式'],
    rows: [
      ['ラインB 10月 生産数', null, '=SUMIF(生産実績!B:B,"B",生産実績!D:D)'],
      ['A-003 品名', null, '=XLOOKUP("A-003",マスター!A:A,マスター!C:C)'],
      ['B-003 単価', null, '=IFERROR(XLOOKUP("B-003",マスター!A:A,マスター!E:E),0)'],
      ['ラインC 10月 生産数', null, '=SUMIFS(生産実績!D:D,生産実績!B:B,"C",生産実績!C:C,"10月")'],
    ],
  };
  // BRIDGE：実務に近い「汚れた」ファイル
  const bridge = {
    headers: ['記入日', '製品ｺｰﾄﾞ', 'ﾗｲﾝ名', '区分', '月', '生産数(個)', '前回比', '備考'],
    rows: [
      ['2024/10/01', '00123', 'Aライン', '量産', '10月', 410, { err: '#REF!' }, ''],
      ['2024/10/01', '00124', 'Bライン', '量産', '10月', 320, { err: '#REF!' }, ''],
      ['2024/10/01', '00125', 'Cライン', '量産', '10月', 275, { err: '#REF!' }, ''],
      [null, null, null, null, null, null, null, null],
      ['2024/10/02', '00123', 'Aライン', '量産', '10月', 390, { err: '#REF!' }, ''],
      ['2024/10/02', '00126', 'Bライン ', '量産', '10月', 418, { err: '#REF!' }, '夜勤分 手入力'],
      ['2024/10/02', '00127', 'Cライン', '試作', '10月', 305, { err: '#REF!' }, ''],
      ['2024/10/03', '00123', 'Aライン', '量産', '10月', 452, { err: '#REF!' }, ''],
      ['2024/10/03', '00124', 'Bライン', '量産', '10月', '510', { err: '#REF!' }, 'CSV貼付'],
      ['2024/10/03', '00125', 'Cライン', '量産', '10月', 400, { err: '#REF!' }, ''],
      ['2024/09/30', '00124', 'Bライン', '量産', '9月', 180, { err: '#REF!' }, '9月分'],
    ],
  };
  const bridgeOld = {
    headers: ['日付', 'ライン', '数量', '担当'],
    rows: [['2024/09/30', 'B', 180, '佐藤'], ['2024/09/27', 'A', 380, '佐藤'], ['2024/09/27', 'C', 240, '田中']],
  };

  const workbook = {
    '生産実績': production, '本日': today, 'マスター': master, '不良実績': defect,
    '出荷指示': order, '検査結果': inspect, '受入データ': intake, '月次レポート': report,
    '10月実績(最終)': bridge, '10月実績(旧)': bridgeOld,
  };

  // ===================== CHARACTERS / VOICE =====================
  const chars = {
    navi: { name: 'ミライ', role: 'ナビゲーター', img: 'assets/img/chars/navi.jpg' },
    support: { name: 'ソウタ', role: '現場サポート', img: 'assets/img/chars/support.jpg' },
    ai: { name: 'AIアシスタント', role: 'AI', img: 'assets/img/chars/ai.jpg' },
    boss: { name: '工藤課長', role: '上司', img: 'assets/img/chars/boss.jpg' },
  };
  const voices = {
    start: { text: 'ミッション開始。関数を覚えなくて大丈夫です。まず、何を求めたいかを確認しましょう。', url: 'https://www.aidocmaker.com/g0/audio?name=6409993e6f0c427aa58a18d10b869f64' },
    discovery: { text: 'フォーミュラ・ディスカバリー。条件に合うデータだけが残りました。式が、どのデータを使って計算しているか確認してみましょう。', url: 'https://www.aidocmaker.com/g0/audio?name=c9f6fd6f1a124516bdf09e8f147ccf53' },
    silent: { text: 'サイレントエラーを発見しました。Excelはエラーを出していません。しかし、十月という条件が抜けています。エラーがないことと、正しいことは同じではありません。', url: 'https://www.aidocmaker.com/g0/audio?name=f85ae00b7b724738a4448d6dc4096dc0' },
    verified: { text: '検証完了。対象、条件、元データの確認ができました。結果は検証済みです。', url: 'https://www.aidocmaker.com/g0/audio?name=425eb56f78a3485baa081a60f5eed435' },
    final: { text: 'ファイナルミッション。十六時の会議までに、数字が合わない月次レポートを確認してください。正しい式だけでなく、正常に見える間違いにも注意してください。', url: 'https://www.aidocmaker.com/g0/audio?name=079896ecb9f64733b81521bdedb088e6' },
    clear: { text: 'レポート、検証完了。今日のクリアはここでは終わりません。次は、自分のExcelで一度使ってみましょう。', url: 'https://www.aidocmaker.com/g0/audio?name=e26dd183bdb543a8880c0ba37a87f04' },
  };

  // ===================== JOBS (ENTRY) =====================
  const jobs = [
    { id: 'sum', emoji: '🎯', icon: 'condition.svg', title: '条件に合う数字だけ合計したい', sub: '例：ラインBの10月の生産数', mission: 'c1', skill: '条件集計' },
    { id: 'lookup', emoji: '🔍', icon: 'lookup.svg', title: '品番から品名を自動表示したい', sub: '例：出荷指示に品名を出す', mission: 'lookup', skill: 'マスター検索' },
    { id: 'judge', emoji: '🚦', icon: 'decision.svg', title: 'OK/NGを自動判定したい', sub: '例：測定値400以上ならOK', mission: 'judge', skill: '自動判定' },
    { id: 'count', emoji: '🔢', icon: 'sum.svg', title: '件数を数えたい', sub: '例：外観不良は何件？', mission: 'count', skill: '件数集計' },
    { id: 'clean', emoji: '🧹', icon: 'clean.svg', title: '文字を整えたい', sub: '例：見えない空白を消す', mission: 'clean', skill: 'データ整形' },
    { id: 'error', emoji: '🚨', icon: 'error.svg', title: 'エラーの原因を調べたい', sub: '例：#N/A が出た', mission: 'error', skill: 'エラー確認' },
  ];

  // ===================== RESCUE =====================
  const rescue = [
    { id: 'sumifs', job: '条件付き合計', fn: 'SUMIFS', syntax: '=SUMIFS(合計範囲, 条件範囲1, 条件1, 条件範囲2, 条件2)', example: '=SUMIFS(D:D,B:B,"B",C:C,"10月")', sheet: '生産実績',
      meaning: '複数の条件に合う行だけを選んで、その行の数値を合計する。', trap: '条件が1つ抜けても、エラーは出ずに「それらしい数字」が出る。', verify: 'フィルターで同じ条件に絞り込み、ステータスバーの合計と一致するか見る。' },
    { id: 'countifs', job: '件数を数える', fn: 'COUNTIFS', syntax: '=COUNTIFS(条件範囲1, 条件1, 条件範囲2, 条件2)', example: '=COUNTIFS(B:B,"A",C:C,"10月")', sheet: '生産実績',
      meaning: '条件に合う行が何行あるかを数える。', trap: '「Bライン」と「Bライン␣」は別の文字として数えられる。', verify: 'フィルター後の件数表示（○件中△件）と比べる。' },
    { id: 'xlookup', job: '別表から取得', fn: 'XLOOKUP', syntax: '=XLOOKUP(探す値, 探す範囲, 持ってくる範囲, "見つからない時")', example: '=XLOOKUP("A-003",マスター!A:A,マスター!B:B)', sheet: '生産実績',
      meaning: '探す範囲から値を見つけ、同じ行の別の列を持ってくる。', trap: '持ってくる列を間違えても、別の値が正常に表示される。', verify: '1件だけマスターを目で引き、同じ値か確かめる。' },
    { id: 'if', job: 'OK/NG判定', fn: 'IF', syntax: '=IF(条件, 正しいとき, 違うとき)', example: '=IF(B2>=400,"OK","NG")', sheet: '検査結果',
      meaning: '条件を満たすかどうかで、表示する内容を切り替える。', trap: '「以上」は >=、「より大きい」は >。境界値（ちょうど400）で結果が変わる。', verify: '境界値ちょうどの行を1つ用意して判定を確認する。' },
    { id: 'trim', job: '文字を整える', fn: 'TRIM', syntax: '=TRIM(文字)', example: '=TRIM(A4)', sheet: '受入データ',
      meaning: '前後の余分な半角スペースを取り除く。', trap: '見えない空白があると、検索や条件集計で一致しない。', verify: '=LEN(セル) で文字数を比べると空白に気づける。' },
    { id: 'iferror', job: 'エラーを処理', fn: 'IFERROR', syntax: '=IFERROR(計算, エラー時の値)', example: '=IFERROR(XLOOKUP(A2,マスター!A:A,マスター!B:B),"未登録")', sheet: '出荷指示',
      meaning: 'エラーが出たときだけ、指定した値を表示する。', trap: '0 や 空白 で隠すと、マスター漏れなどの本当の問題が見えなくなる。', verify: 'IFERRORを一度外して、何のエラーが隠れていたか確認する。' },
  ];

  // ===================== MISSIONS =====================
  // step.type: talk / choice / build / xray / detect / repair / verify / debug / ai / bridge / chapterEnd / card / mywork
  const H = (a, b, c, d, e) => [a, b, c, d, e];

  const missions = {
    // ---------- 第1章 FIRST SUCCESS ----------
    c1: {
      mins: 5, title: 'FIRST SUCCESS', chapter: 1, job: 'sum', next: 'c2',
      steps: [
        { type: 'talk', who: 'navi', voice: 'start', text: 'ようこそ、みらい精機へ。関数を覚えなくて大丈夫。まずは今日の日報から、かんたんな仕事を1つ片づけよう。' },
        { type: 'choice', tag: 'WHAT', skill: 'WHAT', sheet: '本日', title: 'REQUEST', request: '今日の生産数を全部合計してください。',
          prompt: 'まず、何を求める？', options: [
            { t: '生産数の合計', ok: true, fb: 'そのとおり。「生産数」列の数字を全部足す仕事です。' },
            { t: '何行あるか（件数）', fb: '件数ではなく「数を足す」仕事です。リクエストの「合計」に注目。' },
            { t: 'ラインごとの平均', fb: 'ラインで分ける指示はまだありません。まずは全部です。' },
          ] },
        { type: 'choice', tag: 'CHOOSE', skill: 'CHOOSE', sheet: '本日', title: 'CHOOSE', request: '今日の生産数を全部合計してください。',
          prompt: '「全部足す」ができる道具はどれ？', options: [
            { t: '全部たす', code: 'SUM', ok: true, fb: '正解。SUM は「範囲の数値を全部足す」関数です。' },
            { t: '数える', code: 'COUNT', fb: 'COUNT は「数値が何個あるか」を数えます。足し算ではありません。' },
            { t: '探して持ってくる', code: 'XLOOKUP', fb: 'XLOOKUP は別の表から値を探す道具です。' },
          ] },
        { type: 'build', skill: 'BUILD', sheet: '本日', title: 'BUILD', request: '今日の生産数を全部合計してください。', forceLevel: 0, job: 'sum',
          expect: 1248, fns: ['SUM'],
          template: { fn: 'SUM', slots: [{ label: '合計するもの', ans: ['C:C', 'C2:C5'], chips: ['C:C', 'A:A', 'B:B'] }] },
          hints: H('「生産数」はどの列？ 列の上の文字（A・B・C）を見てみよう。', '「全部足す」系の関数です。', '関数名は SUM。', '=SUM(合計したい列)', '=SUM(C:C)'),
          mistakes: [{ v: 0, msg: '0 になりました。文字の列（ラインや品番）を選んでいない？ 数字が入っているのは「生産数」の列です。' }],
          success: { title: 'FORMULA WORKED!', meaning: 'SUM = 範囲の数値を全部足す' } },
        { type: 'talk', who: 'navi', text: '最初の1式、完成！ では次の依頼。「ラインBだけ」ならどうする？ 今度は月のデータがある「生産実績」シートを使うよ。' },
        { type: 'build', skill: 'BUILD', sheet: '生産実績', title: 'SMALL CHALLENGE', request: 'ラインBの生産数だけ合計してください。', forceLevel: 0, job: 'sum',
          expect: 1428, fns: ['SUMIF', 'SUMIFS'],
          template: { fn: 'SUMIF', slots: [
            { label: '条件範囲（どの列で判定？）', ans: ['B:B'], chips: ['B:B', 'C:C', 'D:D'] },
            { label: '条件（何と一致？）', ans: ['"B"'], chips: ['"B"', '"10月"', '"A"'] },
            { label: '合計するもの', ans: ['D:D'], chips: ['D:D', 'F:F', 'B:B'] }] },
          hints: H('「ライン = B」の行だけ残して、生産数を足すイメージ。', '条件つきで合計する仕事です。', '関数名は SUMIF。', '=SUMIF(条件の列, 条件, 合計する列)', '=SUMIF(B:B,"B",D:D)'),
          success: { title: 'FORMULA WORKED!', meaning: 'SUMIF = 1つの条件に合う行だけを合計する' } },
        { type: 'ai', skill: 'AI REVIEW', title: 'AI CHECK（早期体験）', sheet: '生産実績', short: true,
          rounds: [{ request: 'ラインCの10月の生産数を知りたい', say: 'この式でできます！', formula: '=SUMIF(B:B,"C",D:D)', good: ['REPAIR', 'REJECT'],
            checks: ['何を計算？ → 生産数の合計', '条件は全部ある？ → ラインCだけ。10月がない', '数件確認 → 9月の240と260まで入っている'],
            explain: '条件が「ラインC」だけで、10月の条件がありません。AIの式もエラーは出ませんが、9月分まで合計されています。', fix: '=SUMIFS(D:D,B:B,"C",C:C,"10月")' }] },
        { type: 'chapterEnd', title: 'FIRST SUCCESS CLEAR', gained: ['全部合計（SUM）', '1条件の合計（SUMIF）', 'AIの式を一度疑ってみる'] },
      ],
    },
    // ---------- 第2章 MISSION ENGINE & X-RAY ----------
    c2: {
      mins: 5, title: 'X-RAY MISSION', chapter: 2, job: 'sum', next: 'c3',
      steps: [
        { type: 'talk', who: 'boss', text: 'ラインBの「10月」の生産数を出してくれ。9月分は入れないように。' },
        { type: 'choice', tag: 'WHAT', skill: 'WHAT', sheet: '生産実績', title: 'WHAT', request: 'ラインB・10月の生産数を求めて。',
          prompt: '何を求める？', options: [
            { t: '生産数（の合計）', ok: true, fb: 'OK。合計するのは「生産数」列です。' },
            { t: '稼働時間（の合計）', fb: '依頼は「生産数」です。稼働時間ではありません。' },
            { t: 'ラインBの行数', fb: '件数ではなく、生産数を足す仕事です。' },
          ] },
        { type: 'choice', tag: 'LOGIC', skill: 'LOGIC', sheet: '生産実績', title: 'LOGIC', request: 'ラインB・10月の生産数を求めて。', multi: true,
          prompt: '条件はどれ？（当てはまるものを全部選ぶ）', options: [
            { t: 'ライン = B', ok: true }, { t: '月 = 10月', ok: true }, { t: '品番 = A-002' }, { t: '稼働時間 = 8.0' },
          ], fbOk: '条件は2つ。「ライン = B」と「月 = 10月」。', fbNg: '依頼文に出てくる条件だけを選ぼう。品番や稼働時間の指定はありません。' },
        { type: 'choice', tag: 'CHOOSE', skill: 'CHOOSE', sheet: '生産実績', title: 'CHOOSE', request: 'ラインB・10月の生産数を求めて。',
          prompt: '条件2つで合計する道具は？', options: [
            { t: '全部たす', code: 'SUM', fb: 'SUM には条件を付けられません。全行が足されます。' },
            { t: '条件1つで合計', code: 'SUMIF', fb: 'SUMIF は条件1つまで。今回は条件が2つあります。' },
            { t: '複数条件で合計', code: 'SUMIFS', ok: true, fb: '正解。SUMIFS は条件をいくつでも追加できます。' },
            { t: '複数条件で数える', code: 'COUNTIFS', fb: 'COUNTIFS は「何件あるか」。今回は生産数を足す仕事です。' },
          ] },
        { type: 'build', skill: 'BUILD', sheet: '生産実績', title: 'BUILD', request: 'ラインB・10月の生産数を求めて。', job: 'sum',
          expect: 1248, fns: ['SUMIFS'],
          template: { fn: 'SUMIFS', slots: [
            { label: '合計するもの', ans: ['D:D'], chips: ['D:D', 'F:F'] },
            { label: '条件範囲①', ans: ['B:B'], chips: ['B:B', 'C:C', 'E:E'] },
            { label: '条件①', ans: ['"B"'], chips: ['"B"', '"10月"'] },
            { label: '条件範囲②', ans: ['C:C'], chips: ['C:C', 'B:B', 'A:A'] },
            { label: '条件②', ans: ['"10月"'], chips: ['"10月"', '"B"', '"9月"'] }] },
          hints: H('「生産数」を合計。ただし「ライン = B」かつ「月 = 10月」の行だけ。', '条件つきで合計する仕事（条件が2つ）。', '関数名は SUMIFS。', '=SUMIFS(合計する列, 条件の列1, 条件1, 条件の列2, 条件2)', '=SUMIFS(D:D,B:B,"B",C:C,"10月")'),
          mistakes: [{ v: 1428, msg: '1,428 は「ラインB」の全期間です。10月の条件は入っている？' }, { v: 2615, msg: '10月の全ラインになっていない？ ラインの条件を確認しよう。' }],
          success: { title: 'FORMULA WORKED!', meaning: 'SUMIFS = 複数条件に合う行だけを集計する' } },
        { type: 'xray', skill: 'READ', sheet: '生産実績', title: 'FORMULA X-RAY', voice: 'discovery', formula: '=SUMIFS(D:D,B:B,"B",C:C,"10月")', meaning: 'SUMIFS = 複数条件に合う行だけを集計する' },
        { type: 'choice', tag: 'READ', skill: 'READ', sheet: '生産実績', title: 'READ', formula: '=COUNTIFS(B:B,"A",C:C,"10月")',
          prompt: 'この式は何を求めている？', options: [
            { t: 'ラインAの10月の生産数の合計', fb: 'COUNTIFS は「足す」ではなく「数える」。合計する列がありません。' },
            { t: 'ラインAで10月の行が何件あるか', ok: true, fb: '正解。条件に合う行を数えます。答えは 3 件。' },
            { t: 'ラインAまたは10月の行数', fb: 'COUNTIFS の条件は「かつ（AND）」です。「または」ではありません。' },
          ] },
        { type: 'choice', tag: 'PREDICT', skill: 'PREDICT', sheet: '生産実績', title: 'PREDICT', formula: '=SUMIFS(D:D,B:B,"C",C:C,"10月")',
          prompt: '実行する前に予想。結果はどれくらい？（表を見て概算でOK）', options: [
            { t: '約500', fb: 'ラインCの10月は3行あります。3行とも300前後…。' },
            { t: '約1,000', ok: true, fb: '正解！ 275 + 305 + 400 = 980。概算できると検算にも使えます。' },
            { t: '約1,500', fb: 'それだと9月の分（240, 260）まで入っています。' },
          ], xrayAfter: true },
        { type: 'chapterEnd', title: 'X-RAY CLEAR', gained: ['WHAT → LOGIC → 関数の順で考える', '複数条件の合計（SUMIFS）', '式が使ったデータを X-RAY で確認'] },
      ],
    },
    // ---------- 第3章 SILENT ERROR ----------
    c3: {
      mins: 5, title: 'SILENT ERROR', chapter: 3, job: 'sum', next: 'c4',
      steps: [
        { type: 'talk', who: 'support', text: '先輩が作った集計表があるんだ。ラインB・10月の数字、このまま提出していいかな？' },
        { type: 'detect', skill: 'DETECT', sheet: '生産実績', title: 'FORMULA DETECTIVE', voice: 'silent',
          request: 'ラインB＋10月の生産数を合計。', formula: '=SUMIF(B:B,"B",D:D)', correct: 1248,
          causes: [
            { t: '条件が足りない（10月がない）', ok: true },
            { t: '合計する列が違う', fb: '合計しているのは D列（生産数）なので、列は合っています。' },
            { t: '範囲がずれている', fb: '範囲は B:B と D:D で、ずれはありません。' },
            { t: '問題なし', fb: 'エラーは出ていないけれど…X-RAYで使った行を見てみよう。' },
          ],
          explain: 'Excelはエラーを出していません。でも「10月」の条件が抜けていて、9月の180まで合計されています。' },
        { type: 'repair', skill: 'REPAIR', sheet: '生産実績', title: 'REPAIR', request: 'ラインB＋10月の生産数を合計。', formula: '=SUMIF(B:B,"B",D:D)', expect: 1248, fns: ['SUMIFS'], job: 'sum',
          hints: H('足りない条件を足すには？', '条件を2つ以上つけられる関数に変える。', 'SUMIF → SUMIFS。引数の順番が変わるので注意。', '=SUMIFS(合計する列, 条件の列1, 条件1, 条件の列2, 条件2)', '=SUMIFS(D:D,B:B,"B",C:C,"10月")') },
        { type: 'verify', skill: 'VERIFY', sheet: '生産実績', title: '10 SECOND VERIFY', voice: 'verified',
          request: 'ラインB＋10月の生産数を合計。', formula: '=SUMIFS(D:D,B:B,"B",C:C,"10月")',
          target: { col: 'D', header: '生産数' }, conds: ['ライン = B', '月 = 10月'], decoys: ['品番 = A-002', '稼働時間 = 8.0'],
          rowsOk: [8, 11, 14] },
        { type: 'chapterEnd', title: 'DETECT & REPAIR CLEAR', gained: ['エラーが出ない間違い（サイレントエラー）を見抜く', '式を直す', '10秒で結果を確かめる'] },
      ],
    },
    // ---------- 第4章 DEBUG & RECOVERY ----------
    c4: {
      mins: 4, title: 'DEBUG MISSION', chapter: 4, job: 'lookup', next: 'c5',
      steps: [
        { type: 'talk', who: 'support', text: '出荷指示に品名を自動で出したら、1行だけ #N/A が出た…。原因を調べよう。' },
        { type: 'debug', skill: 'DEBUG', sheet: '出荷指示', title: 'DEBUG MISSION', cell: 'B5', formula: '=XLOOKUP(A5,マスター!A:A,マスター!B:B)', variant: 'fullwidth' },
        { type: 'build', skill: 'BUILD', sheet: '生産実績', title: 'RECOVERY MISSION', recovery: true, request: 'ラインAの10月の生産数を求めて。', job: 'sum', expect: 1252, fns: ['SUMIFS'],
          template: { fn: 'SUMIFS', slots: [
            { label: '合計するもの', ans: ['D:D'], chips: ['D:D', 'F:F'] },
            { label: '条件範囲①', ans: ['B:B'], chips: ['B:B', 'C:C'] },
            { label: '条件①', ans: ['"A"'], chips: ['"A"', '"B"'] },
            { label: '条件範囲②', ans: ['C:C'], chips: ['C:C', 'B:B'] },
            { label: '条件②', ans: ['"10月"'], chips: ['"10月"', '"9月"'] }] },
          hints: H('さっきと同じ形。ラインだけ A に変える。', '複数条件の合計。', 'SUMIFS。', '=SUMIFS(D:D, B:B, ライン, C:C, 月)', '=SUMIFS(D:D,B:B,"A",C:C,"10月")'),
          success: { title: 'RECOVERED!', meaning: '同じ形の式は、条件を変えるだけで使い回せる' } },
        { type: 'chapterEnd', title: 'DEBUG CLEAR', gained: ['#N/A の原因を特定する', 'エラーを隠さずに直す', '同じ形の式を使い回す'] },
      ],
    },
    // ---------- 第5章 AI REVIEW ----------
    c5: {
      mins: 3, title: 'AI REVIEW', chapter: 5, job: 'sum', next: 'c6',
      steps: [
        { type: 'talk', who: 'ai', text: 'Excelの式なら任せてください！ いくつか提案します。ただし…ちゃんと確認してくださいね。' },
        { type: 'ai', skill: 'AI REVIEW', title: 'AI REVIEW', sheet: '生産実績',
          rounds: [
            { request: 'ラインBの10月の生産数', say: 'この式でできます。', formula: '=SUMIFS(D:D,B:B,"B",C:C,"10月")', good: ['ACCEPT'],
              checks: ['何を計算？ → D列（生産数）の合計', '条件は全部ある？ → ラインB・10月の2つ', '数件確認 → 320 + 418 + 510 = 1,248'], explain: 'この式は正しいです。確認したうえでの ACCEPT が正解。' },
            { request: 'ラインBの10月の生産数', say: '範囲を絞ったほうが軽くなります。', formula: '=SUMIFS(D2:D12,B2:B12,"B",C2:C12,"10月")', good: ['REPAIR', 'REJECT'],
              checks: ['範囲は正しい？ → 12行目までで止まっている', '数件確認 → 10/03 の 510 が入っていない'], explain: '範囲が12行目で切れていて、最終行（510）が入っていません。エラーは出ません。', fix: '=SUMIFS(D:D,B:B,"B",C:C,"10月")' },
            { request: '出荷指示のB-003の単価を表示', say: 'エラーが出ないようにIFERRORで囲みました。', formula: '=IFERROR(XLOOKUP("B-003",マスター!A:A,マスター!E:E),0)', good: ['REPAIR', 'REJECT'], sheet: 'マスター',
              checks: ['データ型は？ → 単価 0 円はありえる？', 'IFERRORを外すと？ → #N/A（マスターに B-003 がない）'], explain: 'マスターに B-003 が登録されていません。0 で隠すと「単価0円」として計算が進んでしまいます。', fix: '=XLOOKUP("B-003",マスター!A:A,マスター!E:E,"マスター未登録")' },
            { request: 'ラインBの10月の生産数', say: '配列数式で書くとプロっぽいですよ。', formula: '=SUM(IF((B2:B15="B")*(C2:C15="10月"),D2:D15))', good: ['REPAIR', 'ACCEPT'], best: 'REPAIR',
              checks: ['結果は？ → 1,248 で正しい', '読める？ → 複雑で、後任者が直せない'], explain: '結果は正しいけれど、不要に複雑です。同じことが SUMIFS で読みやすく書けます。動く式より「読める式」。', fix: '=SUMIFS(D:D,B:B,"B",C:C,"10月")' },
          ] },
        { type: 'chapterEnd', title: 'AI REVIEW CLEAR', gained: ['AIの式を ACCEPT / CHECK / REPAIR / REJECT で判断', 'IFERRORによる問題の隠蔽を見抜く', '範囲の切れ・不要な複雑さに気づく'] },
      ],
    },
    // ---------- 第6章 THE BROKEN REPORT ----------
    c6: {
      mins: 8, title: 'THE BROKEN REPORT', chapter: 6, job: 'sum', next: 'c7', final: true,
      steps: [
        { type: 'talk', who: 'boss', clock: '15:40', voice: 'final', text: '16時の会議資料、数字が合わない。確認して。' , finalIntro: true },
        { type: 'detect', skill: 'DETECT', clock: '15:43', sheet: '生産実績', title: 'REPORT ① ラインB 10月 生産数', report: 0,
          request: 'ラインBの10月の生産数', formula: '=SUMIF(生産実績!B:B,"B",生産実績!D:D)', correct: 1248,
          causes: [{ t: '条件が足りない', ok: true }, { t: '列が違う', fb: 'D列（生産数）を合計していて列は正しいです。' }, { t: '問題なし', fb: '9月の行も入っていないか、X-RAYで確認を。' }],
          explain: '月の条件が抜けています。', repairTo: { expect: 1248, fns: ['SUMIFS'], hints: H('月の条件を足す。', '複数条件の合計。', 'SUMIFS', '=SUMIFS(合計, 条件列1, 条件1, 条件列2, 条件2)', '=SUMIFS(生産実績!D:D,生産実績!B:B,"B",生産実績!C:C,"10月")') } },
        { type: 'detect', skill: 'DETECT', clock: '15:47', sheet: 'マスター', title: 'REPORT ② A-003 品名', report: 1,
          request: '品番 A-003 の品名', formula: '=XLOOKUP("A-003",マスター!A:A,マスター!C:C)', correct: 'ハウジング',
          causes: [{ t: '持ってくる列が違う', ok: true }, { t: '探す値が違う', fb: '探しているのは A-003 で合っています。' }, { t: '問題なし', fb: '「加工」は品名？ マスターの見出しを確認しよう。' }],
          explain: '「加工」は工程（C列）です。品名は B列にあります。', repairTo: { expect: 'ハウジング', fns: ['XLOOKUP', 'VLOOKUP'], hints: H('品名はマスターの何列目？', '持ってくる範囲を変える。', 'XLOOKUP のまま。', '=XLOOKUP("A-003",マスター!A:A,マスター!品名の列)', '=XLOOKUP("A-003",マスター!A:A,マスター!B:B)') } },
        { type: 'choice', tag: 'DETECT', skill: 'DETECT', clock: '15:51', sheet: 'マスター', title: 'REPORT ③ B-003 単価', report: 2, formula: '=IFERROR(XLOOKUP("B-003",マスター!A:A,マスター!E:E),0)',
          prompt: '単価が「0」と表示されている。どう対応する？', options: [
            { t: 'エラーが出ていないのでそのまま', fb: 'IFERRORの裏で #N/A が隠れています。0円で計算が進むのは危険。' },
            { t: 'IFERRORを外して原因を確認 → マスター未登録と表示し、担当へ連絡', ok: true, fb: '正解。IFERRORの裏に #N/A。B-003 がマスターに未登録でした。問題を見える形にします。' },
            { t: 'IFERRORの0を空白""に変える', fb: '見た目が変わるだけで、問題は隠れたままです。' },
          ], reveal: 'IFERRORを外すと → #N/A（「B-003」がマスターに見つかりません）' },
        { type: 'ai', skill: 'AI REVIEW', clock: '15:55', title: 'AIが修正版を提案', sheet: '生産実績',
          rounds: [
            { clock: '15:55', request: 'ラインBの10月の生産数（レポート①の修正版）', say: 'レポート①の修正版です。', formula: '=SUMIFS(生産実績!D:D,生産実績!B:B,"B",生産実績!C:C,"10月")', good: ['ACCEPT'], checks: ['条件：ラインB・10月 → OK', '結果 1,248 → 元データと一致'], explain: '確認済みの正しい式です。' },
            { clock: '15:57', request: 'ラインAの10月の生産件数（追加の項目）', say: '追加の項目も作っておきました。', formula: '=COUNTIFS(生産実績!B:B,"A")', good: ['REPAIR', 'REJECT'], checks: ['条件は全部ある？ → 月の条件がない', '数件確認 → 9月の2件も数えて 5件'], explain: 'AI式にも1つ誤り。10月の条件が抜けて、9月分も数えています。正しくは 3件。', fix: '=COUNTIFS(生産実績!B:B,"A",生産実績!C:C,"10月")' },
          ] },
        { type: 'verify', skill: 'VERIFY', clock: '15:59', sheet: '生産実績', title: 'FINAL VERIFY', voice: 'verified',
          request: 'ラインBの10月の生産数（レポート①）', formula: '=SUMIFS(生産実績!D:D,生産実績!B:B,"B",生産実績!C:C,"10月")',
          target: { col: 'D', header: '生産数' }, conds: ['ライン = B', '月 = 10月'], decoys: ['品番 = A-002'], rowsOk: [8, 11, 14], finalClear: true },
        { type: 'chapterEnd', clock: '16:00', title: 'REPORT VERIFIED ✓', gained: ['正常に見える間違いを3つ発見', 'AIの修正案も検証', '会議に間に合った'] },
      ],
    },
    // ---------- 第7章 BRIDGE → CARD → MY WORK ----------
    c7: {
      mins: 5, title: 'BRIDGE MISSION', chapter: 7, job: 'sum', next: null,
      steps: [
        { type: 'talk', who: 'support', text: '本物の現場ファイルはこんな感じ。列名もバラバラ、空白行、古い式のエラー…。でも「JobとLogicは同じ」だよ。' },
        { type: 'bridge', skill: 'TRANSFER', sheet: '10月実績(最終)', title: 'BRIDGE MISSION', request: 'Bラインの10月の生産数を出して。', correct: 1248 },
        { type: 'card' },
        { type: 'mywork' },
      ],
    },

    // ===================== JOB MINI MISSIONS =====================
    lookup: {
      mins: 7, title: 'マスター検索', job: 'lookup', next: null, mini: true,
      steps: [
        { type: 'talk', who: 'navi', voice: 'start', text: '出荷指示に品名を自動で出したい。マスター表から「探して持ってくる」仕事だよ。' },
        { type: 'choice', tag: 'WHAT', skill: 'WHAT', sheet: '出荷指示', title: 'WHAT', request: 'A2の品番（A-002）の品名を、マスターから表示して。',
          prompt: '何をする仕事？', options: [
            { t: '品番をマスターで探して、同じ行の品名を持ってくる', ok: true, fb: 'そのとおり。「探す値」「探す場所」「持ってくる列」の3つがそろえばOK。' },
            { t: '品番の数を数える', fb: '数える仕事ではありません。' },
            { t: '品名を手で入力する', fb: '手入力だとマスター変更に追従できません。' },
          ] },
        { type: 'build', skill: 'BUILD', sheet: '出荷指示', title: 'BUILD', request: 'A2の品番（A-002）の品名を、マスターから表示して。', job: 'lookup', expect: 'ブラケット', fns: ['XLOOKUP', 'VLOOKUP'],
          template: { fn: 'XLOOKUP', slots: [
            { label: '探す値', ans: ['A2'], chips: ['A2', 'C2', '"品名"'] },
            { label: '探す場所', ans: ['マスター!A:A'], chips: ['マスター!A:A', 'マスター!B:B'] },
            { label: '持ってくる列', ans: ['マスター!B:B'], chips: ['マスター!B:B', 'マスター!C:C', 'マスター!E:E'] }] },
          hints: H('「A-002」をマスターの品番列で探して、同じ行の品名を持ってくる。', '探して持ってくる系（Lookup）。', 'XLOOKUP。', '=XLOOKUP(探す値, マスター!品番の列, マスター!品名の列)', '=XLOOKUP(A2,マスター!A:A,マスター!B:B)'),
          mistakes: [{ v: '組立', msg: '「組立」は工程（C列）です。品名はどの列？' }],
          success: { title: 'FORMULA WORKED!', meaning: 'XLOOKUP = 探す値を見つけて、同じ行の別の列を持ってくる' } },
        { type: 'xray', skill: 'READ', sheet: '出荷指示', title: 'FORMULA X-RAY', formula: '=XLOOKUP(A2,マスター!A:A,マスター!B:B)', meaning: 'XLOOKUP = 探す値を見つけて、同じ行の別の列を持ってくる' },
        { type: 'detect', skill: 'DETECT', sheet: 'マスター', title: 'FORMULA DETECTIVE', request: '品番 A-003 の品名', formula: '=XLOOKUP("A-003",マスター!A:A,マスター!C:C)', correct: 'ハウジング',
          causes: [{ t: '持ってくる列が違う', ok: true }, { t: '探す値が違う', fb: 'A-003 で合っています。' }, { t: '問題なし', fb: 'マスターの見出しを見てみよう。' }], explain: '「加工」は工程の列。品名はB列です。エラーは出ません。' },
        { type: 'debug', skill: 'DEBUG', sheet: '出荷指示', title: 'DEBUG MISSION', cell: 'B5', formula: '=XLOOKUP(A5,マスター!A:A,マスター!B:B)', variant: 'fullwidth' },
        { type: 'card', mini: true },
      ],
    },
    judge: {
      mins: 5, title: '自動判定', job: 'judge', next: null, mini: true,
      steps: [
        { type: 'talk', who: 'navi', voice: 'start', text: '検査結果をOK/NGで自動判定しよう。基準は「測定値が400以上ならOK」。' },
        { type: 'choice', tag: 'LOGIC', skill: 'LOGIC', sheet: '検査結果', title: 'LOGIC', request: '測定値が400以上ならOK、それ以外はNG。',
          prompt: '「400以上」を式の記号で書くと？', options: [
            { t: 'B2>400', fb: '「>」は「より大きい」。400ちょうどが NG になってしまいます。' },
            { t: 'B2>=400', ok: true, fb: '正解。「以上」は >= 。400ちょうどもOKになります。' },
            { t: 'B2=400', fb: '400ちょうどだけがOKになります。' },
          ] },
        { type: 'build', skill: 'BUILD', sheet: '検査結果', title: 'BUILD', request: 'L-01（B2）の判定を出して。400以上ならOK、それ以外はNG。', job: 'judge', expect: 'OK', fns: ['IF'],
          template: { fn: 'IF', slots: [
            { label: '判定条件', ans: ['B2>=400'], chips: ['B2>=400', 'B2>400', 'B2<400'] },
            { label: '正しいとき', ans: ['"OK"'], chips: ['"OK"', '"NG"'] },
            { label: '違うとき', ans: ['"NG"'], chips: ['"NG"', '"OK"'] }] },
          hints: H('条件を満たす？ → OK、満たさない → NG。', '判定する系。', 'IF。', '=IF(条件, 正しいとき, 違うとき)', '=IF(B2>=400,"OK","NG")'),
          success: { title: 'FORMULA WORKED!', meaning: 'IF = 条件で表示を切り替える' } },
        { type: 'choice', tag: 'PREDICT', skill: 'PREDICT', sheet: '検査結果', title: 'PREDICT', formula: '=IF(B4>=400,"OK","NG")',
          prompt: 'L-03（測定値 400 ちょうど）の結果を予想しよう。', options: [{ t: 'OK', ok: true, fb: '正解。400 >= 400 は「正しい」。' }, { t: 'NG', fb: '400は「400以上」に含まれます。' }] },
        { type: 'detect', skill: 'DETECT', sheet: '検査結果', title: 'FORMULA DETECTIVE', request: 'L-03（B4）を判定。400以上ならOK。', formula: '=IF(B4>400,"OK","NG")', correct: 'OK',
          causes: [{ t: '境界値（ちょうど400）の扱いが違う', ok: true }, { t: 'OKとNGが逆', fb: '400より大きい値はOKになっていて、逆ではありません。' }, { t: '問題なし', fb: 'L-03 は 400 ちょうど。基準は「以上」です。' }], explain: '「>」だと400ちょうどがNGになります。合格品が不合格扱いに。' },
        { type: 'card', mini: true },
      ],
    },
    count: {
      mins: 5, title: '件数集計', job: 'count', next: null, mini: true,
      steps: [
        { type: 'talk', who: 'navi', voice: 'start', text: '不良実績から「外観不良」が何件あるか数えよう。' },
        { type: 'choice', tag: 'WHAT', skill: 'WHAT', sheet: '不良実績', title: 'WHAT', request: '外観不良は何件？',
          prompt: '何を求める？', options: [
            { t: '不良分類が「外観」の行の数', ok: true, fb: 'OK。不良数を足すのではなく「行の数」を数えます。' },
            { t: '外観不良の不良数の合計', fb: '依頼は「何件」。件数です。' },
          ] },
        { type: 'build', skill: 'BUILD', sheet: '不良実績', title: 'BUILD', request: '外観不良は何件？', job: 'count', expect: 3, fns: ['COUNTIF', 'COUNTIFS'],
          template: { fn: 'COUNTIF', slots: [
            { label: '条件範囲', ans: ['D:D'], chips: ['D:D', 'B:B', 'E:E'] },
            { label: '条件', ans: ['"外観"'], chips: ['"外観"', '"寸法"'] }] },
          hints: H('不良分類の列で「外観」の行を数える。', '条件つきで数える系。', 'COUNTIF。', '=COUNTIF(条件の列, 条件)', '=COUNTIF(D:D,"外観")'),
          success: { title: 'FORMULA WORKED!', meaning: 'COUNTIF = 条件に合う行を数える' } },
        { type: 'detect', skill: 'DETECT', sheet: '不良実績', title: 'FORMULA DETECTIVE', request: '不良の記録は全部で何件？', formula: '=COUNT(D:D)', correct: 8,
          causes: [{ t: '文字の列をCOUNTで数えている', ok: true }, { t: '範囲が違う', fb: '範囲はD列全体で問題ありません。' }, { t: '問題なし', fb: '結果 0 件…ほんとうに？' }], explain: 'COUNT は「数値」だけを数えます。文字の列は COUNTA か、数値の列（不良数）を使います。' },
        { type: 'build', skill: 'BUILD', sheet: '不良実績', title: 'RECOVERY', recovery: true, request: 'ラインBの外観不良は何件？', job: 'count', expect: 1, fns: ['COUNTIFS'],
          template: { fn: 'COUNTIFS', slots: [
            { label: '条件範囲①', ans: ['B:B'], chips: ['B:B', 'D:D'] }, { label: '条件①', ans: ['"B"'], chips: ['"B"', '"外観"'] },
            { label: '条件範囲②', ans: ['D:D'], chips: ['D:D', 'B:B'] }, { label: '条件②', ans: ['"外観"'], chips: ['"外観"', '"B"'] }] },
          hints: H('条件が2つ。ライン=B、不良分類=外観。', '複数条件で数える。', 'COUNTIFS。', '=COUNTIFS(列1, 条件1, 列2, 条件2)', '=COUNTIFS(B:B,"B",D:D,"外観")'),
          success: { title: 'RECOVERED!', meaning: 'COUNTIFS = 複数条件に合う行を数える' } },
        { type: 'card', mini: true },
      ],
    },
    clean: {
      mins: 5, title: 'データ整形', job: 'clean', next: null, mini: true,
      steps: [
        { type: 'talk', who: 'support', text: '受入データの「Bライン」を数えたら、3件あるはずが1件になった…。見えない何かがいる。' },
        { type: 'detect', skill: 'DETECT', sheet: '受入データ', title: 'FORMULA DETECTIVE', request: 'Bラインの行は何件？', formula: '=COUNTIF(A:A,"Bライン")', correct: 3, showSpaces: true,
          causes: [{ t: '前後に見えない空白がある', ok: true }, { t: '数える列が違う', fb: 'A列（入力ライン）で合っています。' }, { t: '問題なし', fb: '表を見ると Bライン は3行ありそうです。' }], explain: '「Bライン␣」「␣Bライン」は別の文字。Excelはエラーを出しません。' },
        { type: 'build', skill: 'BUILD', sheet: '受入データ', title: 'BUILD', request: 'A4（ Bライン）の前後の空白を取り除いて。', job: 'clean', expect: 'Bライン', fns: ['TRIM'],
          template: { fn: 'TRIM', slots: [{ label: '整える文字', ans: ['A4'], chips: ['A4', 'B4', 'C4'] }] },
          hints: H('前後の空白を消す道具を使う。', '文字を整える系。', 'TRIM。', '=TRIM(セル)', '=TRIM(A4)'),
          success: { title: 'CLEANED!', meaning: 'TRIM = 前後の余分なスペースを取る' } },
        { type: 'build', skill: 'BUILD', sheet: '受入データ', title: 'SMALL CHALLENGE', request: 'C2の品番（A-002）の先頭1文字（ライン記号）を取り出して。', job: 'clean', expect: 'A', fns: ['LEFT'],
          template: { fn: 'LEFT', slots: [{ label: '文字', ans: ['C2'], chips: ['C2', 'A2'] }, { label: '左から何文字', ans: ['1'], chips: ['1', '2', '5'] }] },
          hints: H('左から1文字だけ取り出す。', '文字を切り出す系。', 'LEFT。', '=LEFT(セル, 文字数)', '=LEFT(C2,1)'),
          success: { title: 'FORMULA WORKED!', meaning: 'LEFT = 左から指定した文字数を取り出す' } },
        { type: 'card', mini: true },
      ],
    },
    error: {
      mins: 6, title: 'エラー確認', job: 'error', next: null, mini: true,
      steps: [
        { type: 'talk', who: 'navi', voice: 'start', text: 'エラーは失敗じゃなくて「手がかり」。消すのではなく、原因を見つけよう。' },
        { type: 'debug', skill: 'DEBUG', sheet: '出荷指示', title: 'DEBUG MISSION', cell: 'B5', formula: '=XLOOKUP(A5,マスター!A:A,マスター!B:B)', variant: 'fullwidth' },
        { type: 'debug', skill: 'DEBUG', sheet: '出荷指示', title: 'DEBUG MISSION 2', cell: 'B6', formula: '=XLOOKUP(A6,マスター!A:A,マスター!B:B)', variant: 'missing' },
        { type: 'choice', tag: 'DETECT', skill: 'DETECT', sheet: '出荷指示', title: 'IFERROR TRAP', formula: '=IFERROR(XLOOKUP(A6,マスター!A:A,マスター!B:B),"")',
          prompt: '同僚が「エラーが消えた！」と喜んでいる。この式の問題は？', options: [
            { t: '問題なし。表がきれいになった', fb: '空白で隠しただけ。B-003 がマスターにない問題は残っています。' },
            { t: 'マスター未登録が隠れて、誰も気づけなくなる', ok: true, fb: '正解。隠すなら「マスター未登録」など、気づける表示に。' },
          ] },
        { type: 'card', mini: true },
      ],
    },
  };

  // 60 SEC WORK CHALLENGE
  const challenge = [
    { skill: 'WHAT', q: '「10月にラインBで作った数を知りたい」。求めるものは？', options: ['生産数の合計', '行の件数', '稼働時間の平均'], a: 0 },
    { skill: 'LOGIC', q: '上の依頼の条件は？', options: ['ライン=B だけ', 'ライン=B と 月=10月', '品番=A-002'], a: 1 },
    { skill: 'BUILD', q: '「条件2つで合計」に向くのは？', options: ['SUM', 'SUMIFS', 'XLOOKUP', 'IF'], a: 1 },
    { skill: 'READ', q: '=COUNTIF(B:B,"A") の意味は？', options: ['B列で「A」の行数', 'B列の合計', 'A列の件数'], a: 0 },
    { skill: 'VERIFY', q: '依頼「ラインB＋10月の合計」に =SUMIF(B:B,"B",D:D) → 1,428（エラーなし）。この結果は？', options: ['正しい', '怪しい（条件が足りない）'], a: 1 },
    { skill: 'DEBUG', q: '=XLOOKUP(A5,…) が #N/A。A5 は「Ａ-001」。原因は？', options: ['マスターが壊れている', '全角/半角の表記違い', 'Excelのバグ'], a: 1 },
    { skill: 'AI REVIEW', q: 'AIが「IFERRORで囲んでエラーを0にしました」。どうする？', options: ['そのまま使う', '何のエラーが隠れているか確認する'], a: 1 },
  ];
  const challengeMap = { WHAT: 'c2', LOGIC: 'c2', BUILD: 'c1', READ: 'c2', VERIFY: 'c3', DEBUG: 'c4', 'AI REVIEW': 'c5' };

  const myWork = [
    { id: 'sum', t: '条件集計', mission: '自分のExcelで、条件つき合計（SUMIFS）を1回使ってみる' },
    { id: 'lookup', t: 'マスター検索', mission: '自分のExcelで、品番などから名前を引く（XLOOKUP）を1回使ってみる' },
    { id: 'judge', t: '自動判定', mission: '自分のExcelで、OK/NG判定（IF）を1回使ってみる' },
    { id: 'error', t: 'エラー確認', mission: '自分のExcelで、IFERRORの裏に隠れたエラーがないか1回確認する' },
    { id: 'clean', t: 'データ整形', mission: '自分のExcelで、見えない空白（TRIM / LEN）を1回確認する' },
  ];

  const chapters = ['c1', 'c2', 'c3', 'c4', 'c5', 'c6', 'c7'];

  root.EFQData = { workbook, chars, voices, jobs, rescue, missions, challenge, challengeMap, myWork, chapters };
})(typeof window !== 'undefined' ? window : globalThis);
