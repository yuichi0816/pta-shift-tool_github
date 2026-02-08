# 引継書（次チャット用）

- 作成日: 2026-02-07
- 対象プロジェクト: `pta-shift-tool_github_ver2`
- 主対象: `tool.html` / `js/script.js`

## 1. 依頼背景（ユーザー要望）

- `tool.html` を中心に、プロジェクトを実質的に作り直したい。
- UI/デザインは現行を可能な限り維持（踏襲）。
- 既存機能は削除しない（機能同等を維持）。
- 保守しやすい構造で改修可能性を高めたい。
- マニュアルと実装を分析して、仕様書・要件定義を整備したい。
- 設問タグマッピングは `tool.html`（STEP 1.5）へ統合済みとし、独立ページは廃止する。

## 2. このチャットで作成・更新したドキュメント

## 2.1 作成済み
- `docs/tool_as_is_spec.md`
  - `tool.html` 現行仕様（As-Is）
  - UI、ロジック、出力、削除禁止機能を固定
- `docs/tool_rebuild_requirements.md`
  - 再構築向けの詳細要件（To-Be）
  - 保守性・互換性・受け入れ基準
- `docs/tool_mapping_ux_additional_requirements.md`
  - UX追加要件（現時点でUX-001〜UX-007）
  - マッピング導線強化、実行導線改善、運用性改善

## 2.2 重要な要件トピック（追加要件定義書）
- UX-001: 不足タグナビ
- UX-002: マッピング確定アクション
- UX-003: 未対応列のみ表示
- UX-004: 割当実行ボタンの無効理由表示
- UX-005: ファイル読込のドラッグ＆ドロップ
- UX-006: マッピングJSONの自動復元（localStorage）
- UX-007: 用語の統一（`tool.html` と `manual.html`）

## 3. 実装の現状（コード）

- `tool.html` と `js/script.js` に、`survey-mapper` 統合実装が入っている。
- 方向性:
  - `STEP 1.5` としてタグマッピングUIを `tool.html` に追加
  - `js/script.js` でタグマッピング結果を列解決ソースに使用
  - キーワード直接判定の中心ロジックを縮退

## 3.1 変更状態（未コミット）
- 変更あり: `tool.html`
- 変更あり: `js/script.js`
- `docs/` は新規作成状態（未追跡）

## 3.2 注意
- この時点の `tool.html` / `js/script.js` は、手動動作確認と整合調整を実施すること。
- 追加UX（UX-004〜UX-007）は要件定義済みだが、実装は未完了。

## 4. 次チャットで優先して進めるべきこと

1. 現在の `tool.html` / `js/script.js` の動作確認（読み込み〜割当〜出力）。
2. `docs/tool_mapping_ux_additional_requirements.md` の要件順に実装。
   - まず UX-004（無効理由表示）
   - 次に UX-005（D&D）
   - 続いて UX-006（自動復元）
   - 最後に UX-007（用語統一）
3. 実装ごとに UAT 観点を満たすか確認。
4. `manual.html` の用語整合を実施（UX-007）。

## 5. 非交渉の制約（引継先で必ず守る）

- UIデザインは原則踏襲（大幅変更しない）。
- 既存機能は削除禁止。
- `tool.html` 主体で進める。
- Excel入出力（既存ワークフロー）を壊さない。

## 6. 次チャットに貼るための依頼文（コピペ用）

以下を新しいチャットに貼る:

```text
この引継書に沿って作業を継続してください。

対象リポジトリ: pta-shift-tool_github_ver2

必読:
- docs/tool_as_is_spec.md
- docs/tool_rebuild_requirements.md
- docs/tool_mapping_ux_additional_requirements.md
- docs/handover_next_chat.md

前提:
- tool.html のUI/デザインはなるべく維持
- 既存機能は削除しない

まずやること:
1) 現在の tool.html / js/script.js の動作確認
2) UX-004 と UX-005 を実装
3) 手動確認結果を報告

その後:
- UX-006（localStorage自動復元）
- UX-007（manual.html と tool.html の用語統一）

実装時は、追加要件定義書の UAT を満たしているか明示してください。
```

## 7. 参照ファイル

- `docs/tool_as_is_spec.md`
- `docs/tool_rebuild_requirements.md`
- `docs/tool_mapping_ux_additional_requirements.md`
