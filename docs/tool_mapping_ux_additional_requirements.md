# tool.html UX改善 追加要件定義書（マッピング導線強化・実行導線改善・運用性改善）

- 作成日: 2026-02-07
- 対象画面: `tool.html`（STEP 1.5 設問タグマッピング / STEP 3 割当実行 / STEP 1 ファイル読込）
- 関連仕様: `docs/tool_as_is_spec.md`, `docs/tool_rebuild_requirements.md`
- 目的: 着手前に、UX改善7点を実装可能レベルで具体化する

## 1. 背景

STEP 1.5 のタグマッピング導入により列認識の確実性は向上したが、初回利用時に以下の課題が残る。

- `mappingSummary` に不足/重複が表示されても、該当列を表から探しにくい
- マッピング完了状態が暗黙で、ユーザーが「次に進んでよいか」を判断しづらい
- 列数が多いと `UNMAPPED` の探索に時間がかかる
- 割当実行不可の理由が明示されず、次アクションが分かりにくい
- ファイル読み込みがクリック選択中心で、PCでの操作コストが高い
- 前回と同形式のアンケートでも毎回タグマッピングを手作業でやり直す必要がある
- `manual.html` と `tool.html` の用語差で、入力や確認時に混乱が起きやすい

本追加要件は、上記7課題をUIの最小差分で改善し、作業時間と誤操作を減らすことを目的とする。

## 2. スコープ

### 2.1 対象
- 不足タグナビ機能
- マッピング確定アクション
- 未対応列のみ表示フィルタ
- 割当実行ボタンの無効理由表示
- ファイル読込のドラッグ＆ドロップ対応
- マッピングJSONの自動復元（前回設定）
- 用語の統一

### 2.2 対象外
- 割り当てアルゴリズム自体の変更
- `validator.html` の改修
- 全体デザインリニューアル

## 3. 要件サマリー

| ID | 追加機能 | 期待効果 |
|---|---|---|
| UX-001 | 不足タグナビ | 修正対象列の探索時間を短縮 |
| UX-002 | マッピング確定アクション | フローの明確化、誤実行防止 |
| UX-003 | 未対応列のみ表示 | 多列データでの編集効率向上 |
| UX-004 | 割当実行ボタンの無効理由表示 | 次アクションの明確化 |
| UX-005 | ファイル読込のドラッグ＆ドロップ | アップロード操作コスト削減 |
| UX-006 | マッピングJSONの自動復元（前回設定） | 繰り返し作業の削減 |
| UX-007 | 用語の統一 | 認知負荷・入力ミスの低減 |

## 4. 機能要件（FR）

## 4.1 UX-001 不足タグナビ

### FR-UX-001
- `mappingSummary` で検出した問題（必須不足/重複）を、クリック可能なナビとして表示する。
- ナビ項目クリック時、該当列行へ自動スクロールし、視覚的にハイライトする。

### 表示要件
- 新規領域: `#mappingIssueNav`（`mappingSummary` 直下）
- 表示条件:
  - 問題あり時のみ表示
  - 問題なし時は非表示
- 表示内容:
  - 必須不足: `COUNT 未設定` など
  - 重複: `LOC_1 重複（3, 7列）` など

### 挙動要件
- ナビを押すと最初の対象列へ `scrollIntoView({ behavior: 'smooth', block: 'center' })`
- スクロール先行に 1.5 秒程度の一時ハイライトを付与
- ハイライトクラス（例: `mapping-row-highlight`）は複数回押しても再点灯可能

### 補足要件
- 対象列が複数ある重複ケースでは、
  - 方式A: 1項目内に「次へ」を持つ
  - 方式B: 列ごとに別ナビ表示
  - 本要件では方式Bを推奨（実装容易）

## 4.2 UX-002 マッピング確定アクション

### FR-UX-002
- `STEP 1.5` に `マッピング確定` ボタンを追加する。
- 有効なマッピング状態でのみ確定可能とする。
- 確定後にのみ `STEP 2/3` を開放する。

### 追加UI
- 新規ボタン: `#confirmMappingBtn`
- 新規状態表示: `#mappingConfirmStatus`（例: `未確定`, `確定済み`）

### 状態遷移
- `draft`（初期）: マッピング未完了
- `valid`（必須/重複クリア）: 確定可能
- `confirmed`: 確定済み、STEP 2/3開放
- `dirtyAfterConfirm`: 確定後にタグ変更が発生、未確定へ戻す

### 制御要件
- `confirmMappingBtn` 活性条件:
  - `surveyMapping.isValid === true`
- `checkBothFilesLoaded()` の開放条件:
  - `shiftData` 読込済み
  - `surveyData` 読込済み
  - `surveyMapping.isValid === true`
  - `surveyMapping.isConfirmed === true`
- `runAssignment()` の実行条件も上記と同条件を要求

### 文言要件
- 未確定時: `マッピングを確定すると STEP 2 へ進めます`
- 確定時: `マッピング確定済み（タグ変更で再確定が必要）`

## 4.3 UX-003 未対応列のみ表示

### FR-UX-003
- マッピング表に `未対応列のみ表示` トグルを追加する。
- ON時は `UNMAPPED` 行のみ表示する。
- OFF時は全列表示に戻す。

### 追加UI
- 新規トグル: `#showUnmappedOnly`（checkbox推奨）
- 任意で補助表示: `#unmappedCountBadge`（例: `未対応 5件`）

### 挙動要件
- `renderSurveyMappingTable()` は `state.surveyMapping.filter.unmappedOnly` を参照して描画行を絞り込む
- フィルタON中に行がすべて解消された場合、空状態メッセージを表示
  - 例: `未対応列はありません（すべて設定済み）`

### 非機能要件
- フィルタ切替時の再描画は 200ms 以内を目標
- スクロール位置は可能な範囲で保持（難しければ先頭復帰可）

## 4.4 UX-004 割当実行ボタンの無効理由表示

### FR-UX-004
- `assignBtn` が無効の間、無効理由を常時表示する。
- 無効理由はユーザーの次操作が分かる文言で表示する。

### 追加UI
- 新規領域: `#assignBlockedReason`（`assignBtn` 近傍）

### 制御要件
- `assignBtn` の活性状態を単一関数で判定する（例: `evaluateAssignmentReadiness()`）。
- 判定で `isReady=false` の場合、理由文字列を `#assignBlockedReason` に表示する。
- 判定で `isReady=true` の場合、`#assignBlockedReason` は「実行可能」表示または非表示とする。

### 理由の優先表示（推奨）
1. シフトファイル未読込
2. アンケートファイル未読込
3. タグマッピング未確定（必須不足/重複含む）
4. 場所不一致エラーあり
5. 処理中

### 文言例
- `シフトファイルを読み込んでください`
- `アンケートファイルを読み込んでください`
- `設問タグマッピングを確定してください`
- `場所名不一致を解消してください`

## 4.5 UX-005 ファイル読込のドラッグ＆ドロップ対応

### FR-UX-005
- `surveyUploadBox` / `shiftUploadBox` にドラッグ＆ドロップでのファイル読込を追加する。
- 既存のクリック選択（`input[type=file]`）は維持する。

### 挙動要件
- 対応イベント:
  - `dragenter`
  - `dragover`
  - `dragleave`
  - `drop`
- `drop` 時に対象ファイルを取得し、既存の読込処理に委譲する。
- 対象拡張子以外は受け付けず、エラーメッセージを表示する。

### 表示要件
- ドラッグ中はアップロードボックスの視覚状態を変更する（境界色/背景色）。
- ドロップ完了後は既存の `uploaded` 表示ルールに従う。

### 実装要件
- 読込ロジックをイベント依存で重複実装しない。
- `change`（ファイル選択）と `drop`（D&D）から同一処理関数を呼ぶ。

## 4.6 UX-006 マッピングJSONの自動復元（前回設定）

### FR-UX-006
- 最新のマッピング設定を `localStorage` に保存し、同形式ファイル読込時に自動復元候補として提示する。

### 保存要件
- 保存タイミング:
  - マッピング確定時
  - マッピングJSON手動読込成功時
- 保存キー（例）:
  - `pta_shift:last_mapping`
- 保存内容:
  - 列順・ヘッダー・タグの対応
  - 生成日時
  - ソース識別情報（ファイル名、列数、ヘッダー署名）

### 復元要件
- 復元条件:
  - 列数一致
  - ヘッダー署名一致（厳密一致 or 正規化一致）
- 条件一致時:
  - `前回設定を復元` を表示
  - ユーザー同意で適用（自動適用は任意設定）
- 条件不一致時:
  - 自動復元しない
  - 「前回設定はこのファイル形式に適合しません」を表示

### セキュリティ/運用要件
- `localStorage` が利用不可の環境では静かに無効化し、機能全体は継続可能とする。
- 個人情報値（回答データ本文）は保存しない。タグ対応情報のみ保存する。

## 4.7 UX-007 用語の統一

### FR-UX-007
- `tool.html` / `manual.html` / 関連UI文言で、同一概念の用語を統一する。

### 統一対象（初期セット）
- `参加可能回数`（主用語）
  - 置換対象候補: `参加回数`, `参加確認`（文脈別に再定義）
- `割当て回数`
- `未割当て`
- `特定参加可能日`

### 実装要件
- 用語辞書を定義（`docs/terminology_glossary.md` 新規推奨）。
- UI文言変更時は辞書を参照し、差分レビュー項目に「用語整合」を追加する。

### 表示要件
- 同一画面内で同義語を混在させない。
- `manual.html` の説明文と `tool.html` 表示ラベルが一致すること。

## 5. データ/状態要件

`state.surveyMapping` に以下を追加する。

- `isConfirmed: boolean`
- `issues: Array<{ type: 'missing'|'duplicate', tagId: string, colIndices: number[] }>`
- `filter: { unmappedOnly: boolean }`

`state` に以下を追加する。

- `assignmentBlockReason: string`
- `uploadDragState: { survey: boolean, shift: boolean }`
- `lastMappingSnapshot: { signature: string, updatedAt: string } | null`

更新ルール:
- タグ変更時:
  - `isConfirmed` を `false` に戻す
  - `issues` を再計算
- `confirmMappingBtn` 押下時:
  - `isValid === true` の場合のみ `isConfirmed = true`
  - `localStorage` へ保存

## 6. UI/スタイル要件

- 既存の緑系デザインを維持
- 追加部品のスタイルは既存クラスを優先利用
  - ボタン: `.btn .btn-outline .btn-primary`
  - 情報帯: `analysis-legend` または `stat-item` の組み合わせ
- 新規クラス最小化:
  - `.mapping-issue-nav`
  - `.mapping-row-highlight`
  - `.upload-box.drag-active`

## 7. 受け入れ基準（UAT）

### UAT-UX-001 不足タグナビ
- 必須不足時にナビが表示される
- ナビクリックで該当行へ移動し、ハイライトされる
- 重複時に対象列が明示される

### UAT-UX-002 マッピング確定
- 有効マッピング前は確定不可
- 確定後は STEP 2/3 が開放される
- 確定後にタグ変更すると再確定が必要になる
- 未確定状態では `assignBtn` 実行不可

### UAT-UX-003 未対応列フィルタ
- トグルONで `UNMAPPED` のみ表示される
- `UNMAPPED` がゼロの場合、空状態メッセージが表示される
- トグルOFFで全列に戻る

### UAT-UX-004 割当ボタン無効理由
- `assignBtn` 無効時に理由が常時表示される
- 条件の解消に応じて理由が即時更新される
- 実行可能時に無効理由表示が消える、または `実行可能` に切り替わる

### UAT-UX-005 ドラッグ＆ドロップ
- `surveyUploadBox` へドロップでアンケートが読込できる
- `shiftUploadBox` へドロップでシフトが読込できる
- 非対応拡張子ドロップ時にエラー表示される
- 既存のクリック選択読込が引き続き動作する

### UAT-UX-006 自動復元
- 同形式ファイル読込時に前回マッピングを復元できる
- 不一致ファイルでは復元しない
- 個人情報データ本文を `localStorage` に保存しない

### UAT-UX-007 用語統一
- `tool.html` と `manual.html` で同一概念の表示語が一致する
- 用語辞書にない新語を追加せず運用できる

## 8. 実装指針（推奨）

### 実装順
1. `UX-002`（確定状態導入）
2. `UX-001`（不足タグナビ）
3. `UX-003`（未対応列フィルタ）
4. `UX-004`（割当実行ボタンの無効理由表示）
5. `UX-005`（ドラッグ＆ドロップ）
6. `UX-006`（自動復元）
7. `UX-007`（用語統一）

理由:
- 先に状態管理を固めることで、他機能の条件分岐が明確になるため。

### 影響ファイル
- `tool.html`
  - ボタン/ナビ/トグルの追加
  - `assignBlockedReason` 領域追加
- `js/script.js`
  - `surveyMapping` 状態拡張
  - `validateSurveyTagMapping()` の戻り情報拡張
  - テーブル描画フィルタ
  - 開放条件（STEP2/3, assign）更新
  - `assignBtn` 活性判定と理由表示の統合
  - D&Dイベント処理
  - `localStorage` 保存/復元処理
- `manual.html`
  - 用語統一反映
- `css/style.css`（必要時のみ）
  - ハイライトアニメーション
  - D&D中の視覚スタイル
- `docs/terminology_glossary.md`（新規推奨）

## 9. リスクと対策

- リスク: 状態増加により分岐が複雑化
  - 対策: `computeMappingIssues()`, `updateMappingConfirmState()`, `renderMappingIssueNav()` を分離
- リスク: 確定状態の戻し忘れ
  - 対策: タグ変更ハンドラで `isConfirmed=false` を一元実施
- リスク: 行ジャンプの対象計算ミス
  - 対策: `data-col-index` を唯一キーとして利用
- リスク: D&Dのイベント競合でブラウザ既定動作（ファイル表示）が発生
  - 対策: `dragover/drop` で `preventDefault()` を徹底
- リスク: 無効理由の分岐追加で条件矛盾が発生
  - 対策: 判定ロジックを `evaluateAssignmentReadiness()` に一元化
- リスク: `localStorage` の旧データ適用で誤マッピングが起きる
  - 対策: ヘッダー署名一致チェックと復元前確認ダイアログを必須化
- リスク: 用語統一で既存ユーザーが一時的に戸惑う
  - 対策: 主要変更点を `manual.html` の更新履歴に明記

## 10. 完了定義（Definition of Done）

- 本書の UAT-UX-001〜007 を満たす
- `tool.html` の既存機能（アップロード、割り当て、結果表示、Excel出力）を毀損しない
- `node --check js/script.js` で構文エラーなし
- 手動確認で以下が通る
  - アンケート読込 -> マッピング -> 確定 -> STEP2表示
  - 既存の割り当て実行フロー完走

## 11. 参考

- 既存仕様: `docs/tool_as_is_spec.md`
- 再構築要件: `docs/tool_rebuild_requirements.md`
