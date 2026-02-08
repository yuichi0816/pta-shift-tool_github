# tool.html 再構築向け詳細要件定義書（To-Be）

- 作成日: 2026-02-07
- 対象: `tool.html` + 新規ロジック実装（現 `js/script.js` の置換）
- 方針: UI/デザインを踏襲しつつ、保守可能な構造へ全面再設計

## 1. 背景と目的

現行 `js/script.js` は単一ファイルに処理が集中し、以下の課題がある。
- ロジックとDOM操作が密結合
- 列認識・候補抽出が複数箇所に重複
- 仕様変更時の影響範囲が読みにくい
- 特定入力での列誤認識リスク（監査記録あり）

本再構築では、以下を同時達成する。
- 既存機能を落とさない（機能同等）
- UI/見た目/操作フローの互換維持
- 実装をモジュール化し改修容易性を確保

## 2. 開発原則

- 機能削除禁止: 現行のユーザー操作・出力機能を削らない
- 互換優先: `tool.html` のID、主要クラス、操作導線は維持
- 仕様明確化: manual記載と実装差分は要件として明示し、意図を固定
- 変更容易性: 単一責務・純粋関数中心・テスト可能な構造に分割

## 3. スコープ

### 3.1 対象
- `tool.html`（必要最小限のマークアップ調整）
- 新規JS実装（例: `tool-app.js` + 機能別モジュール）
- 既存`css/style.css`のデザイン維持（必要最小限の差分のみ）

### 3.2 対象外
- `validator.html` / `js/validator.js` の全面改修
- `manual.html` の全面改稿（必要な整合修正は別タスク化）
- サーバーサイド化

## 4. 機能要件（FR）

### FR-001 画面互換
- STEP1〜STEP4構成を維持する。
- 既存IDを維持する。
  - `uploadSection`, `surveyUploadBox`, `shiftUploadBox`
  - `surveyFile`, `shiftFile`, `surveyStatus`, `shiftStatus`
  - `previewSection`, `assignSection`, `resultSection`
  - `advancedSettings`, `ngPeriodDays`, `logicSelect`, `assignBtn`, `exportBtn`
  - `resultTable`, `assignmentSummaryTable`, `preferredDatesDensityTable`, `detailAnalysisTable`
- 既存タブ構成と文言を維持する（軽微な文言修正は可）。

### FR-002 ファイル読込
- Excel/CSVの1シート目を読み込む。
- 文字コード混在（全角数字・日本語）を扱えること。
- 読込成功/失敗をステータス表示すること。

### FR-003 シフトデータ解析
- ヘッダー行探索（先頭5行優先、fallbackあり）。
- 場所列抽出時にメタ列を除外（更新/変更/削除）。
- `（N人体制）` 等の定員表記を解析し、`capacity` と `baseName` を生成する。
- 日付・曜日を保持し、土日祝判定を行う。

### FR-004 アンケート列認識
- ヘッダーキーワードで論理列へマッピングする。
- 対象列:
  - 回数、参加判定
  - 希望日/希望月/希望曜日
  - NG日/NG月/NG曜日
  - 希望場所（第1〜第5）
  - 追加対応/追加回数
  - 氏名/学年/クラス/自由記述
- 誤認識が起きやすい項目は優先順位ルールを明示し実装する。

### FR-005 場所整合性検証
- アンケート希望場所とシフト場所（`baseName`）を突合する。
- 不一致時は割当実行を禁止し、修正可能なエラーUIを表示する。

### FR-006 参加者生成
- 除外条件:
  - 空行
  - 回数未入力/0/解析不可
  - 免除希望
- 回数解析:
  - 数値抽出
  - 無制限語彙（何回でも/無制限/以上/週〜）対応
- 氏名生成:
  - 氏名単一列優先
  - 旧形式（姓/名分離）後方互換
  - 重複名サフィックス付与

### FR-007 スロット生成
- 登校日（非土日祝）のみ対象。
- 場所×定員でスロット展開。
- 既存セル値がある枠はプレ割当として保持し、再割当対象から除外する。

### FR-008 共通制約判定
- 同日重複禁止
- 前後NG期間制約（0で無効）
- NG日/NG曜日/NG月の絶対除外
- 希望場所一致
- 特定希望日優先（指定時はその日限定）
- 特定希望日未指定時は希望月/曜日で判定

### FR-009 ロジック選択
- 以下4ロジックを提供する。
  - `minimum_guarantee`
  - `participant_priority`
  - `standard`
  - `date_order`
- すべて同じ共通制約関数を利用する（ロジック間で条件差が出ない構造）。

### FR-010 結果表示
- サマリーカードを表示する。
  - 有効回答者数
  - 参加可能回数合計
  - 総枠数
  - 実割当て合計（充足率）
  - 未割当て数
- 結果4タブを表示する。
  - シフト表
  - 割当て集計
  - 特定希望日集中度
  - 詳細分析
- 0回参加者を集計・表示対象に含める（免除は除外）。

### FR-011 Excel出力
- 出力ファイル名: `旗振りシフト表_YYYYMMDD.xlsx`
- シート:
  - `シフト表`
  - `割り当て集計`
  - `詳細分析`
- 画面表示と整合する内容を出力する。

### FR-012 フロントエンド機能維持（削除禁止）
- 既存機能一覧（`docs/tool_as_is_spec.md` 第10章）を1つも削除しない。
- UI簡素化・機能統合による実質削除は禁止。

## 5. 非機能要件（NFR）

### NFR-001 保守性
- ロジック層をUI層から分離する。
- 1ファイル1責務を基本とする。
- 主要アルゴリズムは副作用の少ない純粋関数として実装する。

### NFR-002 可読性
- JavaScript 4スペース、`const`/`let`、セミコロンを統一。
- 関数名は責務を表す命名に統一。
- コメントは「意図」と「判断理由」を最小限記載。

### NFR-003 性能
- ブラウザ内完結を維持。
- 参加者数200〜500、枠数300〜1000程度で操作不能にならないこと。
- 不要な全探索重複を減らす（候補算出の共通化・キャッシュ利用）。

### NFR-004 信頼性
- 例外発生時にUIが固まらない。
- エラーメッセージは原因と対応が分かる文言にする。
- プレ割当保持・再割当の安全性を担保する。

### NFR-005 デザイン互換
- 緑系テーマ、カード、タブ、テーブルの見た目を維持。
- モバイル（<=768px）で現行同等の表示崩れ抑制。

## 6. アーキテクチャ要件

推奨モジュール構成（例）:
- `ui/dom.js`: DOM取得、イベント配線
- `ui/render-*.js`: プレビュー/結果/エラー描画
- `core/state.js`: 画面状態管理
- `core/parser/shift-parser.js`: シフト解析
- `core/parser/survey-parser.js`: アンケート解析
- `core/mapping/column-mapper.js`: 列認識
- `core/assignment/constraints.js`: 共通制約判定
- `core/assignment/strategies/*.js`: 4ロジック
- `core/export/excel-exporter.js`: Excel出力
- `core/utils/*.js`: 日付・文字列ユーティリティ

必須ルール:
- ロジック内で直接DOM操作しない。
- 候補者判定を単一関数に集約し、各ロジックから再利用する。
- 名前照合・正規化処理を共通化する。

## 7. データモデル要件

### Participant
- `id`, `displayName`, `email`
- `maxAssignments`, `currentAssignments`
- `preferredLocations`, `preferredMonths`, `preferredDays`, `preferredDates`
- `ngDates`, `ngDays`, `ngMonths`
- `canSupportAdditional`, `maxAdditionalAssignments`
- `assignedDates`
- `matchKeySet`

### Slot
- `date`, `dayOfWeek`, `month`
- `location`, `baseName`
- `rowIndex`, `locationIndex`
- `capacity`, `slotIndex`
- `assignedTo`, `isPreAssigned`

### AssignmentResult
- `slots`, `participants`
- `assignedCount`, `unassignedSlots`, `totalSlots`, `preAssignedCount`

## 8. 既知課題に対する改善要件

### RSK-001 列誤認識
- `count` と `additionalCount` の衝突を防ぐ優先順位を明文化。
- `pref3`/`pref4` 上書き誤認識を防ぐ。
- `preferredDates` と `ngDates` の取り違えを防ぐ。
- 監査ログを容易に出せるデバッグ出力を維持する。

### RSK-002 ロジック間の判定差
- 場所比較は `baseName` を正とし、全ロジックで統一。
- 共通制約は単一実装を使用して分岐差分をなくす。

### RSK-003 manualとの差分
- 現行互換を維持する仕様と、将来是正する仕様を分離して記録する。
- 差分項目は `manual` 改訂候補として一覧化する。

## 9. 受け入れ基準（UAT）

### UAT-01 機能同等
- `docs/tool_as_is_spec.md` の削除禁止機能がすべて動作する。

### UAT-02 主要フロー
- ファイル読込 -> プレビュー -> 割当 -> 4タブ確認 -> Excel出力 が完走する。

### UAT-03 再割当
- 既存入力済みセルを保持し、空欄のみ補完できる。

### UAT-04 複数人体制
- `（N人体制）` で定員展開・表示・出力が一致する。

### UAT-05 NG期間
- `ngPeriodDays=0` で無効化、`>0` で間隔制約が機能する。

### UAT-06 列認識
- 監査対象サンプルで `count/pref1-5/preferredDates/ngDates` が意図通りに解釈される。

### UAT-07 出力整合
- 画面集計値とExcel集計値が仕様どおり一致する（0回表示含む）。

## 10. 開発ステップ（提案）

1. 現行互換テストケース作成（手動 + 最小自動）
2. データモデル/パーサ/制約判定の分離実装
3. ロジック4種をStrategy化して移植
4. レンダラと既存DOMを接続
5. Excel出力の差分吸収
6. 互換テスト実施とmanual差分整理

## 11. 成果物

- 新規仕様書: `docs/tool_as_is_spec.md`
- 詳細要件定義: `docs/tool_rebuild_requirements.md`
- 追加要件定義（UX改善）: `docs/tool_mapping_ux_additional_requirements.md`
- 用語辞書: `docs/terminology_glossary.md`
- 実装確認チェックリスト: `docs/tool_implementation_checklist.md`
- 実装フェーズで追加予定:
  - 設計ノート（モジュール依存図）
  - 移行チェックリスト（機能同等確認表）

## 12. 追加UX要件の統合（旧: `tool_mapping_ux_additional_requirements.md`）

本章は、`docs/tool_mapping_ux_additional_requirements.md` の内容を本仕様へ統合したものである。
以降、再構築の実装判定は本書（第4章 + 本章）を正とする。

### FR-013 不足タグナビ（UX-001）
- `mappingSummary` の必須不足/重複を、クリック可能なナビとして表示する。
- ナビクリックで該当列へスクロールし、一時ハイライトする。
- 表示領域: `#mappingIssueNav`

### FR-014 マッピング確定アクション（UX-002）
- `STEP 1.5` に `#confirmMappingBtn` と `#mappingConfirmStatus` を追加する。
- `surveyMapping.isValid === true` のときのみ確定可能とする。
- `STEP 2/3` と `runAssignment()` の開放条件に `surveyMapping.isConfirmed === true` を含める。
- タグ変更時は `isConfirmed = false` に戻す。

### FR-015 未対応列のみ表示（UX-003）
- `#showUnmappedOnly` で `UNMAPPED` 行のみ表示/全表示を切替可能にする。
- `state.surveyMapping.filter.unmappedOnly` を描画条件として利用する。
- 未対応列が0件のときは空状態メッセージを表示する。

### FR-016 割当実行ボタンの無効理由表示（UX-004）
- `#assignBlockedReason` に無効理由を常時表示する。
- 判定ロジックは単一関数（例: `evaluateAssignmentReadiness()`）に集約する。
- 優先理由は以下順を推奨する:
  1. シフトファイル未読込
  2. アンケートファイル未読込
  3. タグマッピング未確定
  4. 場所不一致エラー
  5. 処理中

### FR-017 ファイル読込のドラッグ＆ドロップ（UX-005）
- `#surveyUploadBox` / `#shiftUploadBox` に `dragenter/dragover/dragleave/drop` を実装する。
- `change`（ファイル選択）と `drop`（D&D）は同一読込関数へ委譲する。
- 非対応拡張子はエラー表示し、ブラウザ既定のファイル表示遷移を抑止する。

### FR-018 マッピングJSON自動復元（UX-006）
- `localStorage` に最新マッピングを保存し、同形式ファイル読込時に復元候補を提示する。
- 保存キー例: `pta_shift:last_mapping`
- 復元条件:
  - 列数一致
  - ヘッダー署名一致（正規化一致可）
- `localStorage` 利用不可時は機能を静かに無効化し、主機能は継続可能とする。
- 保存対象はタグ対応情報のみとし、回答本文（個人情報）は保存しない。

### FR-019 用語統一（UX-007）
- `tool.html` / `manual.html` / 関連UI文言の同一概念は同一語彙で表示する。
- 初期統一語:
  - `参加可能回数`
  - `割当て回数`
  - `未割当て`
  - `特定参加可能日`
- 用語運用は `docs/terminology_glossary.md` を正とする。

### 12.1 追加状態要件

`state.surveyMapping`:
- `isConfirmed: boolean`
- `issues: Array<{ type: 'missing'|'duplicate', tagId: string, colIndices: number[] }>`
- `filter: { unmappedOnly: boolean }`

`state`:
- `assignmentBlockReason: string`
- `uploadDragState: { survey: boolean, shift: boolean }`
- `lastMappingSnapshot: { signature: string, updatedAt: string } | null`

### 12.2 統合後UAT

#### UAT-08（UX-001）
- 必須不足/重複時にナビ表示、クリックで対象行へ移動・ハイライトされる。

#### UAT-09（UX-002）
- マッピング有効前は確定不可、確定後のみ STEP2/3 開放、タグ変更で再確定が必要。

#### UAT-10（UX-003）
- `UNMAPPED` フィルタON/OFFが正しく動作し、0件時は空状態メッセージが表示される。

#### UAT-11（UX-004）
- `assignBtn` 無効時に理由が表示され、状態変化に即時追従する。

#### UAT-12（UX-005）
- D&D読込が動作し、非対応拡張子でエラー表示される。クリック読込も維持される。

#### UAT-13（UX-006）
- 同形式ファイルで前回マッピングを復元できる。不一致形式では復元しない。
- `localStorage` に回答本文が保存されない。

#### UAT-14（UX-007）
- `tool.html` と `manual.html` の同一概念が統一用語で表示される。
