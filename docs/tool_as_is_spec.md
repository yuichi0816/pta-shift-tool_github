# tool.html 現行仕様書（As-Is）

- 作成日: 2026-02-07
- 対象画面: `tool.html`
- 関連実装: `js/script.js`, `css/style.css`, `manual.html`
- 目的: 再構築時に機能欠落を防ぐため、現行挙動を仕様として固定する

## 1. スコープ

本仕様書は `tool.html` の以下を対象とする。
- 画面UI（表示要素、タブ、ボタン、文言）
- ファイル読込、プレビュー、割当処理、結果表示、Excel出力
- `js/script.js` が提供する割当ロジックと制約判定

対象外:
- `validator.html` / `js/validator.js` の検証UI
- `gas-config.html` や `region.html` の別機能

## 2. 画面構成（固定すべきUI）

### 2.1 ヘッダー
- タイトル: シフト自動割当て
- ナビリンク:
  - `index.html`
  - `validator.html`
  - `manual.html`
  - `survey-sample.html`
  - `gas-config.html`

### 2.2 STEP 1: ファイル読込
- セクション: `#uploadSection`
- アップロードボックス:
  - `#surveyUploadBox`
  - `#shiftUploadBox`
- `#surveyFile`（アンケート）
- `#shiftFile`（旗振りシフト日と場所）
- 状態表示:
  - `#surveyStatus`
  - `#shiftStatus`
- 両ファイル読込後に `STEP 1.5`（設問タグマッピング）を表示

### 2.3 STEP 1.5: 設問タグマッピング
- セクション: `#mappingSection`
- マッピング操作:
  - `#autoMapBtn`
  - `#resetMapBtn`
  - `#saveMapBtn`
  - `#loadMapBtn`
  - `#confirmMappingBtn`
- 状態表示:
  - `#mappingConfirmStatus`
  - `#mappingSummary`
  - `#mappingIssueNav`
  - `#mappingRestoreNotice`
- 表示先:
  - `#mappingMeta`
  - `#mappingTableBody`
- マッピング確定後に STEP 2 / STEP 3 を表示

### 2.4 STEP 2: データプレビュー
- タブ:
  - `data-tab="shiftPreview"`
  - `data-tab="surveyPreview"`
- 表示先:
  - `#shiftStats`, `#shiftTable`
  - `#surveyStats`, `#surveyTable`

### 2.5 STEP 3: 割当設定・実行
- 前後NG期間入力: `#ngPeriodDays`（0-30、初期値7）
- ロジック選択: `#logicSelect`
  - `minimum_guarantee`（推奨）
  - `participant_priority`
  - `standard`
  - `date_order`
- 実行ボタン: `#assignBtn`
- 実行可否メッセージ: `#assignBlockedReason`
- 進捗表示:
  - `#progressContainer`
  - `#progressFill`
  - `#progressText`

### 2.6 STEP 4: 結果表示
- サマリー: `#resultSummary`
- タブ:
  - `shiftResultTab`（`#resultTable`）
  - `assignmentSummaryTab`（`#assignmentSummaryTable`）
  - `preferredDatesDensityTab`（`#preferredDatesDensityTable`）
  - `detailAnalysisTab`（`#detailAnalysisTable`）
- 出力ボタン: `#exportBtn`

### 2.7 未割当セクション
- `#unassignedSection`, `#unassignedList`
- 現行実装では結果表示時に非表示化（タブ内集計へ統合済み）

## 3. 入力仕様

### 3.1 共通
- 読込対象: `.xlsx`, `.xls`, `.csv`
- 読込先: 1シート目
- 読込方式: SheetJS `sheet_to_json(..., { header: 1, defval: '' })`

### 3.2 旗振りシフト日と場所
- ヘッダー行判定:
  - 先頭5行以内で1列目に「日」を含む行
  - 見つからない場合は2行目（index 1）
- 1列目: 日付
- 2列目: 曜日
- 3列目以降: 場所
- 除外列: ヘッダー名に「更新」「変更」「削除」を含む列
- 複数人体制:
  - 例: `正門前（2人体制）`
  - 正規表現で定員抽出（全角数字対応）
  - `baseName`（定員表記除去後）を保持

### 3.3 アンケート回答
- ヘッダー名のキーワードから列認識（固定列番号に依存しない）
- 主な認識対象:
  - 参加可能回数
  - 参加/免除判定
  - 希望月 / 希望曜日 / 特定参加可能日
  - NG日 / NG曜日 / NG月
  - 希望場所（第1〜第5）
  - 追加対応 / 追加回数
  - 学年 / 氏名（単一・姓名分離）/ クラス / 自由記述

## 4. 前処理仕様

### 4.1 場所一致バリデーション
- 比較対象:
  - シフト側: `baseName` 集合
  - アンケート側: 第1〜第3希望（先頭100行）
- 不一致がある場合:
  - STEP2領域にエラーカードを表示
  - 割当処理へ進ませない

### 4.2 参加者生成
- 参加者除外条件:
  - 空行（メールと第1希望ともに空）
  - 参加可能回数が未入力/0/数値抽出不可
  - 参加可能回数に「免除」を含む
- 参加可能回数:
  - `何回でも/無制限/週/以上` を `999` として扱う
  - 文字列から数値抽出（全角数字対応）
- 表示名:
  - 基本: `氏名_学年`
  - 重複時: `(2)` などを付与
- 既存割当照合キー:
  - 表示名
  - 氏名
  - 姓名結合
  - メールローカル部
  - 学年サフィックス除去、重複サフィックス除去を含む正規化

### 4.3 スロット生成
- 土日祝を除外
- 場所ごとの定員分だけスロット展開
- 既にセルに値がある場合はプレ割当 (`isPreAssigned=true`)
- 1セル内の複数名は `,` `、` 改行区切りで解釈

## 5. 共通割当制約

以下は全ロジックで適用（ドライラン除く）。
- 同一日重複禁止（同一参加者の同日複数枠禁止）
- 前後NG期間制約
  - `#ngPeriodDays=0` で無効
  - 日付差分が `ngPeriod` 未満なら不可
- NG条件（絶対除外）
  - NG日
  - NG曜日
  - NG月
- 希望場所一致
- 特定参加可能日がある場合はその日のみ可
- 特定参加可能日がない場合は希望月/希望曜日で判定

## 6. 割当ロジック仕様

### 6.1 `minimum_guarantee`（全員最低1回保証）
- 第1パス: 0回参加者へ1回を優先配分
- 第2パス: 希望回数まで追加配分
- 第3パス: 柔軟配分（追加対応可を利用）
- 参加者並び順:
  - 入れる日数少ない順
  - 制約レベル高い順
  - 希望回数少ない順
  - 高学年優先
  - 同条件はランダム

### 6.2 `participant_priority`（参加者優先）
- 制約の強い参加者を先に処理
- 参加者ごとに入れるスロットへ割当
- 第2パスで柔軟配分

### 6.3 `standard`（標準）
- 候補者が少ないスロットから優先して埋める
- `selectBestCandidate` スコアで候補者選抜
- 第2パスで柔軟配分

### 6.4 `date_order`（日付順）
- スロット順（通常は日付順）に単純割当
- 候補選定は共通ルールを利用

## 7. 結果表示仕様

### 7.1 サマリー
表示項目:
- 有効回答者数
- 参加可能回数合計（`999`は除外）
- 総枠数
- 実割当て合計（充足率表示）
- 未割当て数

### 7.2 シフト結果タブ
- ヘッダー1行目: 場所
- ヘッダー2行目: 充足/総枠
- ヘッダー3行目: 未割当て率
- セル色:
  - `cell-assigned`
  - `cell-partial`
  - `cell-empty`
- 複数人体制は同日同場所を集約し `、` 連結表示

### 7.3 割当て集計タブ
- 参加者全員表示（0回含む、免除除外）
- 列: No/名前/参加可能回数/既存/新規/合計/割当日
- 上限超過行の強調
- 追加対応による超過は別色で表示

### 7.4 特定希望日集中度タブ
- 日付×場所のクロス表
- 特定希望日の人数カウント
- しきい値着色:
  - 3人以上: 赤
  - 2人: 橙
  - 1人: 緑
- ツールチップに氏名一覧

### 7.5 詳細分析タブ
- 希望条件と実割当の突合
- 判定: `OK / 0回 / 問題`
- 条件不一致の内訳（場・日・月・曜）を可視化

## 8. Excel出力仕様

出力ファイル名:
- `旗振りシフト表_YYYYMMDD.xlsx`

シート:
- `シフト表`
- `割り当て集計`
- `詳細分析`

出力ルール:
- 複数人体制は `、` 連結
- 割当て集計は0回参加者を含む
- 詳細分析は希望条件・割当詳細・判定を出力

## 9. デザイン仕様（踏襲対象）

- カラーテーマ: 緑系（`css/style.css` の CSS Variables）
- ステップカード構成（STEP1〜4）
- タブUIとテーブルの固定ヘッダー/固定左2列
- レスポンシブ対応（768px以下）
- 進捗バー演出、カードのフェードイン

## 10. 削除禁止機能リスト

再構築時に削除不可（同等機能必須）:
- 2ファイル読込と読込状態表示
- 場所不一致エラー表示
- プレビュー2タブ
- 前後NG期間入力
- ロジック4種類の選択
- 既存入力セルの維持（再割当）
- 複数人体制（N人体制）
- 結果4タブ
- 0回参加者の集計表示
- Excel3シート出力

## 11. 現行実装上の注意（再構築時に要整理）

- `manual.html` 記述と `js/script.js` 挙動に一部差分がある（例: ヘッダー認識、空欄時の月/曜日制約の扱い）。
- `md/ASSIGNMENT_INPUT_AUDIT_2026-0405.md` で、特定データに対する列誤認識リスクが確認されている。
- 再構築では「現行互換を維持しつつ、誤認識リスクを下げる設計」が必要。
