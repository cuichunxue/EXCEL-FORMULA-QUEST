# EXCEL FORMULA QUEST

**関数を覚えるな。仕事を解決せよ。**

実務Jobを入口に、WHAT → LOGIC → CHOOSE → BUILD → UNDERSTAND → VERIFY → DETECT → REPAIR → APPLY を体験する、
Just-in-Time型 Excel Formula Performance Support System（要件定義 FINAL v3.0 準拠）。

## 起動

ビルド不要の静的Webアプリです。

- `index.html` をブラウザ（Edge / Chrome）で開くだけで動きます（`file://` でもOK）。
- 社内サーバー等に置く場合は、フォルダごとアップロードしてください。
- 直接リンク：`index.html#c1`〜`#c7`（各章）、`#lookup` `#judge` `#count` `#clean` `#error`（Jobミッション）、`#rescue`（10秒RESCUE）

学習データ（進捗・FORMULA CARD・1週間ログ・TOMORROW MISSION）はブラウザの `localStorage` に保存されます。サーバーへの送信はありません。

## 体験の流れ（COMPLETE JOURNEY）

| 章 | 内容 | 主な要件 |
|---|---|---|
| 1 FIRST SUCCESS | 今日の生産数を全部合計（SUM）→「ラインBだけなら？」（SUMIF）→ 最初の AI CHECK | §7 Easy Success / §23 AI早期体験 |
| 2 X-RAY MISSION | WHAT → LOGIC → CHOOSE → BUILD（SUMIFS）→ FORMULA X-RAY → READ → PREDICT | §10 §11 §14 |
| 3 SILENT ERROR | エラーの出ない誤答 1,428 を見抜く → REPAIR → 10 SECOND VERIFY | §15 §18 |
| 4 DEBUG | #N/A の原因（全角「Ａ-001」）を調査ツールで特定 → 正しく直す → RECOVERY | §19 §21 |
| 5 AI REVIEW | ACCEPT / CHECK / REPAIR / REJECT（正解・範囲切れ・IFERROR隠蔽・不要な複雑化） | §22 |
| 6 THE BROKEN REPORT | 15:40→16:00。関数名の指定なしで月次レポートの3つの誤りを修正・AI案も検証 | §24（時間制限は設定でON/OFF） |
| 7 BRIDGE | 実務風の汚れたファイル（半角ｶﾅ列名・空白行・#REF!・末尾スペース・文字の数字）→ FORMULA CARD → MY WORK BRIDGE | §25 §26 §27 |

トップの6つのJob（品名表示 / OK・NG判定 / 件数 / 文字整形 / エラー原因）からは短いミッションに直接入れます。
ほかに **60秒 WORK CHALLENGE**（§6）、**10秒 RESCUE＋30秒 UNDERSTAND**（§28）、**YOUR WEEK**（§30）、**TOMORROW MISSION → REAL-WORK SUCCESS**（§29 §39）があります。

### 仕組み

- **本物の式エンジン**（`js/formula.js`）：SUM / COUNT(A) / AVERAGE / SUMIF(S) / COUNTIF(S) / XLOOKUP / VLOOKUP / IF / IFS / AND / OR / IFERROR / TRIM / LEFT / RIGHT / MID / LEN / SUBSTITUTE ほか。
  ユーザーが自由に書いた式をその場で評価するので、別解も正しく判定でき、間違えた式も X-RAY で「どの行を使ったか」を見せられます。全角入力・小文字・`==` も受け付けます。
- **FORMULA X-RAY**：条件ごとに対象外の行が暗くなり、対象値だけが発光 → `320 + 418 + 510 = 1,248` → 意味表示。行をクリックすると「なぜ対象/対象外か」が出ます（式 → データ → 結果の双方向追跡）。
- **FADING SUPPORT**：A FULL GUIDE（意味ブロック＋候補）→ B PARTIAL → C FUNCTION ONLY → D JOB ONLY。ヒントなし成功で自動的に減り、初めての仕事では A/B に留まります。
- **PROGRESSIVE HINT**：考え方 → カテゴリー → 関数名 → 構文 → 完成式。完成式を見た成功は「習得」扱いにしません。
- **EMOTION ADAPTIVE**：3回連続でつまずくとナビが声をかけ、支援をAに戻してヒントを開きます。誤操作で止まる画面はありません（✕でいつでも終了でき、その時点で持ち帰れる技能を表示）。
- **MASTERY MODEL**：NEW → DISCOVERED → PRACTICED → INDEPENDENT → VERIFIED → TRANSFERRED（Job単位）。
- **サウンド**：初期設定は OFF（職場利用のため）。ナレーションは字幕を常時表示し、音声はパック記載のリンク→ブラウザ読み上げの順で再生します。

## ファイル構成

```
index.html
css/style.css
js/formula.js     式エンジン（ブラウザ / Node 両対応）
js/data.js        架空企業「みらい精機」のデータセット・ミッション・RESCUE定義
js/app.js         画面・ミッションエンジン
assets/img/chars  キャラクター（素材パックからトリミング）
assets/img/bg     背景
assets/img/icons  Jobアイコン（SVG）
assets/sfx        効果音（素材パック）
tests/            エンジン単体テスト / E2E通しプレイ
```

素材パックについて：UIの文字は要件どおりHTML/CSSで描画し、画像はキャラクター・雰囲気用に使用しています。
パック内の個別画像はカタログからの切り出し位置がずれて隣の素材やラベルが写り込んでいたため、キャラクター・背景をトリミングし直しました。
`clean_icons` は2点（lookup の文字化け、decision と verify が同一）に問題があったため、同じデザイン（Navy＋Safety Yellow枠）でSVGを作り直しています。

## テスト

```bash
node tests/engine.test.js                       # 式エンジンと全ミッションの答え・数値の整合性
NODE_PATH=$(npm root -g) node tests/e2e.js      # Playwrightで全章＋全Jobを通しプレイ（要 playwright）
```

## ミッションの追加

`js/data.js` の `missions` にステップを並べるだけです。ステップ種別：
`talk` / `choice`（WHAT・LOGIC・CHOOSE・READ・PREDICT）/ `build` / `repair` / `xray` / `detect` / `verify` / `debug` / `ai` / `bridge` / `chapterEnd` / `card` / `mywork`。
`tests/engine.test.js` が、各ミッションの完成式・テンプレート解答が期待値になるかを自動で検証します。
