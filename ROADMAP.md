# `openpyxl`との機能差分を埋める

現在、このライブラリは`openpyxl`の基本的な機能（ブックの読み込み、シートの選択、セルの値の読み書き）を実装していますが、`openpyxl`が提供する豊富な機能の多くがまだ実装されていません。

このドキュメントは、`openpyxl`との機能差分を追跡し、実装を管理するための中央ハブとして機能します。

## 未実装機能リスト

*   **データ型:**
    *   [x] 日付・時刻型 (`datetime`)
    *   [x] ブール型 (`bool`)
    *   [x] 数式 (`formula`)
*   **スタイル:**
    *   [x] フォント（名前、サイズ、太字、イタリック、色など）
    *   [x] 塗りつぶし（パターン、背景色、前景色）
    *   [x] 罫線（スタイル、色）
    *   [x] 配置（水平・垂直方向の配置、折り返し、インデント）
    *   [x] 数値フォーマット
    *   [x] 保護
    *   [x] 名前付きスタイル
*   **ワークシートの操作:**
    *   [x] 行・列の挿入・削除
    *   [x] 行の高さ・列の幅の設定
    *   [x] セルのマージ・アンマージ
    *   [x] ウィンドウ枠の固定
    *   [x] 印刷設定（タイトル行、印刷範囲など）
    *   [x] シートのコピー
*   **ブックの操作:**
    *   [x] アクティブシートの設定
    *   [x] 名前付き範囲の操作
*   **オブジェクトの挿入:**
    *   [ ] 画像
    *   [ ] グラフ（棒、線、円、散布図など）
    *   [ ] 図形
*   **その他:**
    *   [x] コメントの追加
    *   [x] 条件付き書式
    *   [x] データバリデーション
    *   [x] テーブル
    *   [x] ピボットテーブル
    *   [x] VBAマクロの保持

## 貢献

このリストにある機能の実装に貢献したい方は、このリポジトリのissueやプルリクエストで議論を開始してください。
