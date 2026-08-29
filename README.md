# excel_to_backlog

Excel ファイルから Backlog の課題を一括登録・更新する CLI ツールです。

---

## 概要

Excel で管理しているタスク・問い合わせ一覧などを Backlog 課題として自動登録できます。

- **複数の Excel ファイル**を1つの設定ファイルにまとめて処理できる
- **ドライラン（デフォルト）** で登録内容を事前確認できる。既存課題の照合まで行うため、作成されるのか更新されるのかまで分かる
- **upsert 対応**：既存課題は更新、新規は作成（issueKey または件名で重複判定）
- **カスタム属性**・ステータス・担当者・開始日・期限日に対応
- フィルタリング・テンプレート・value_map など柔軟なマッピング設定が可能

---

## 必要環境・インストール

**Python 3.10 以上**

```bash
pip install -e .          # 依存関係（openpyxl / pyyaml）をまとめて導入
```

```bash
pip install -e ".[dev]"   # 開発時（テスト実行に必要）
```

---

## クイックスタート

**1. 設定ファイルを作成する**

```bash
cp config.sample.yaml config.yaml
```

**2. `config.yaml` を編集する**

```yaml
backlog:
  space_host: "yourcompany.backlog.com"
  api_key: "YOUR_API_KEY_HERE"
  project_key: "MYPROJ"
```

**3. ドライランで登録内容を確認する**

```bash
python excel_to_backlog.py
```

**4. Markdown プレビューで詳細を確認する（任意）**

```bash
python excel_to_backlog.py --preview
```

**5. 実際に登録・更新する**

```bash
python excel_to_backlog.py --execute
```

---

## 実行コマンド一覧

```
python excel_to_backlog.py [オプション]
```

| オプション | 説明 |
|---|---|
| （なし） | ドライラン。変換結果を確認するのみで Backlog には書き込まない |
| `--execute` | Backlog に実際に課題を作成・更新する |
| `--preview` | 登録予定の課題内容（本文全文）を Markdown ファイルに出力する |
| `--source "名前"` | 指定した name のソースのみ処理する |
| `--config path` | 設定ファイルのパスを指定する（デフォルト: スクリプトと同じディレクトリの `config.yaml`） |
| `-y` / `--yes` | 実行前の確認を省略する（非対話環境ではこの指定が必要） |
| `--resume CSV` | 過去の実行ログを読み、作成・更新済みの行を飛ばして再開する |
| `--no-log` | 実行ログ（`run_*.csv`）を出力しない |
| `--debug` | API リクエスト・レスポンスの詳細を表示する（カスタム属性の反映確認に有用） |

> `--preview` と `--execute` は同時に指定できません。

`--execute` を指定すると、書き込みを始める前に一度だけ確認を求めます。`-y` を付けると省略できます。cron やパイプ経由など**入力を受け取れない環境では `-y` が必須**で、指定がないと終了コード 1 で停止します。

### 元の Excel ファイルは変更しません

このツールは読み込んだ Excel を一切書き換えません。生成するのは以下の新規ファイルのみで、いずれも設定ファイル（`config.yaml`）と同じディレクトリに出力されます。

| ファイル | 内容 |
|---|---|
| `run_YYYYMMDD_HHMMSS.csv` | 実行ログ |
| `preview_YYYYMMDD_HHMMSS_<ソース名>.md` | `--preview` の出力 |

そのため、作成した issueKey が Excel の `key_col` に自動で書き戻されることはありません。対応表が必要な場合は実行ログを参照してください。

> この方針は `tests/test_source_excel_untouched.py` でテストとして固定されています。openpyxl でファイルを開き直して保存すると、数式が失われるほか Excel 側で作成したグラフ・ピボットテーブル・画像などが失われる可能性があるためです。

### 終了コード

| コード | 意味 |
|---|---|
| `0` | すべて正常に処理された |
| `1` | 1件以上のエラー、中断、または設定不備 |
| `2` | コマンドライン引数の誤り |

### 実行ログと再開

`--execute` で実行すると、処理した行を `run_YYYYMMDD_HHMMSS.csv` に1件ずつ記録します。1行ごとに書き込むため、途中で強制終了しても直前までの内容が残ります。

通信エラーなどで中断した場合は、続きから再開できます。

```bash
python excel_to_backlog.py --execute --resume run_20250828_101500.csv
```

作成・更新まで完了した行だけを飛ばすため、失敗した行は再実行されます。

---

## 設定ファイル（config.yaml）

### Backlog 接続設定

```yaml
backlog:
  space_host: "yourcompany.backlog.com"  # スペースのホスト名
  api_key: "YOUR_API_KEY_HERE"           # Backlog API キー
  project_key: "YOUR_PROJECT_KEY"        # プロジェクトキー（例: MYPROJ）
  ssl_verify: true                       # SSL 証明書検証（オンプレ版で false にする場合あり）
  base_path: ""                          # オンプレ版でパスプレフィックスがある場合（例: "/backlog"）
```

> `config.yaml` は `.gitignore` で除外されています（API キーを含むため）。

### sources の基本構造

```yaml
sources:
  - name: "ソースの識別名"  # --source オプションで指定できる名前
    excel:      # Excel 読み込み設定
      ...
    filters:    # 行の絞り込み条件（任意）
      ...
    issue_mapping:  # Backlog 課題へのマッピング
      ...
    upsert:     # 重複チェック・更新設定（任意）
      ...
```

---

## 機能別リファレンス

### Excel 読み込み設定

```yaml
excel:
  path: 'C:\Users\username\Documents\task_list.xlsx'
  sheet: "Sheet1"        # シート名（省略時: 最初のシート）
  header_start_row: 1    # ヘッダー開始行（1始まり）
  header_end_row: 1      # ヘッダー終了行（複数行ヘッダーの場合に end > start）
  data_start_row: 2      # データ開始行
  col_start: "A"         # 読み込み開始列
  col_end: "H"           # 読み込み終了列
```

**Windows でのパス記述について**

バックスラッシュはYAMLのエスケープ文字のため、以下のいずれかで記述します。

```yaml
path: 'C:\Users\username\Documents\task.xlsx'   # シングルクォート（推奨）
path: "C:/Users/username/Documents/task.xlsx"   # スラッシュ区切り
```

**複数行ヘッダーについて**

`header_start_row: 2` / `header_end_row: 3` のように指定すると、2〜3行目をヘッダーとして読み込みます。複数行のセルは `" / "` で結合された1つの列名になります。

```
行2: "大分類"  → 結合後の列名: "大分類 / 小分類"
行3: "小分類"
```

---

### フィルタリング

#### filters（AND 評価）

複数条件をすべて満たす行のみを処理します。

```yaml
filters:
  - col_name: "ステータス"
    value: "未着手"              # 完全一致（デフォルト）
    # match: "contains"         # 部分一致
    # match: "startswith"       # 前方一致
  - col_name: "種別"
    values: ["タスク", "バグ"]   # いずれかに一致（OR）
```

#### filter_groups（グループ間 OR、グループ内 AND）

複合キー（例: 項番＋枝番のペア）で特定行を指定したいときに使います。

```yaml
filter_groups:
  - filters:                  # グループ1: 項番=1 かつ 枝番=A
      - col_name: "項番"
        value: "1"
      - col_name: "枝番"
        value: "A"
  - filters:                  # グループ2: 項番=3 かつ 枝番=B
      - col_name: "項番"
        value: "3"
      - col_name: "枝番"
        value: "B"
```

> `filters` と `filter_groups` を両方指定した場合は `filter_groups` が優先されます。

---

### 必須の設定

`issue_type` と `priority` はすべての課題に共通で適用されます。**どちらも省略できません。**

```yaml
issue_mapping:
  issue_type: "タスク"   # Backlog の種別名
  priority: "中"         # Backlog の優先度名（高 / 中 / 低）
```

名前は Backlog のプロジェクト設定にあるものと完全に一致させてください。起動時にターミナルへ利用可能な一覧が表示されます。

> 設定を間違えた場合、そのソースは処理されずエラーになります（行ごとのスキップにはなりません）。

---

### 件名の設定

#### summary_col：列の値をそのまま件名にする

```yaml
issue_mapping:
  summary_col: "タスク名"
```

#### summary_template：テンプレートで件名を組み立てる

`{{列名}}` でセルの値を埋め込みます。`summary_col` より優先されます。

```yaml
summary_template: "【{{_source_name}}】{{タスク名}}（{{担当者}}）"
```

**条件ブロック `{{#列名}}...{{/列名}}`**

指定列の値が空でなければブロック内を出力し、空なら出力しません。セパレーターやプレフィックスを値がある場合のみ表示したいときに使います。

```yaml
# 枝番="A" → "項番1-A"、枝番="" → "項番1"
summary_template: "項番{{項番}}{{#枝番}}-{{枝番}}{{/枝番}}"
```

**利用できる特殊キー**

| キー | 内容 |
|---|---|
| `{{_source_name}}` | そのソースの `name` の値 |
| `{{_excel_path}}` | `excel.path` の値 |
| `{{_excel_sheet}}` | `excel.sheet` の値 |

---

### 本文（description）の設定

#### template 方式（デフォルト）

`description_template` に Markdown を記述します。`{{列名}}` の部分がセルの値に置換されます。

```yaml
description_format: "template"
description_template: |
  ## 概要
  {{概要}}

  ## 対応内容
  {{対応内容}}

  ## 備考
  {{備考}}
```

#### auto 方式

列名を見出し（`#`）、セルの値を本文として自動生成します。出力形式は `excel_md_tool` と同じです。

```yaml
description_format: "auto"
description_cols:      # 出力する列を絞る場合に指定（省略時は全列）
  - "概要"
  - "対応内容"
```

- 複数行ヘッダー（`"大分類 / 小分類"`）は階層見出し（`#` `##`）に変換
- セル内改行は `<br>` に変換
- 空セルは「（値なし）」と出力

#### {{auto}} プレースホルダー

`template` 方式のテンプレート内で `{{auto}}` を使うと、その位置に auto 方式の出力を展開できます。ヘッダー・フッターを追加したい場合に便利です。

```yaml
description_format: "template"
description_cols:
  - "概要"
  - "対応内容"
description_template: |
  担当: {{担当者}} / 期限: {{期限日}}

  {{auto}}

  ---
  ※ このチケットは自動生成されました。
```

---

### 開始日・期限日

列名を指定すると、その列の値を日付として設定します。

```yaml
start_date_col: "開始予定日"
due_date_col: "期限"
```

`{{列名}}` を含めるとテンプレートとして展開されます。複数の列を組み合わせたい場合に使います。

```yaml
due_date_col: "{{年}}/{{月}}/{{日}}"
```

**受理する形式**

| 形式 | 例 |
|---|---|
| `YYYY-MM-DD` | `2025-01-05` |
| `YYYY/M/D` | `2025/1/5`（ゼロ埋め不要） |
| `YYYY年M月D日` | `2025年1月5日` |

Excel の日付型セルはそのまま使えます。解釈できない値（和暦 `R7/1/5`、年のない `9/1` など）は警告を出し、その項目は未設定のまま登録されます。

---

### 取り消し線を Markdown に変換する（rich_text）

`rich_text: true` にすると、Excel のセルに引かれた取り消し線を Markdown の `~~text~~` に変換して本文に反映します。

```yaml
issue_mapping:
  rich_text: true
  description_format: "auto"
```

反映されるのは `description_format: "auto"` の出力と、テンプレート内の `{{auto}}` の部分だけです。`{{列名}}` は取り消し線を含まないプレーンテキストのままになります。

> openpyxl 3.1 以上が必要です。古いバージョンでは警告を出してプレーンテキストのまま処理を続けます。

---

### 担当者設定

#### assignee_col：列からユーザーを設定する

Backlog の表示名またはログイン ID と一致する文字列が列に入っている必要があります。

```yaml
assignee_col: "担当者"
```

#### default_assignee：担当者のデフォルト値

`assignee_col` が未設定、またはセルが空の場合に適用されるデフォルトの担当者を指定します。セルに値がある場合はセルの値が優先されます。

```yaml
default_assignee: "yamada"   # Backlog の表示名 or ログインID
```

---

### ステータス制御

Excel のステータス列の値を Backlog のステータスに対応付けます。Backlog のステータス名はプロジェクト設定で確認してください（例: 未対応 / 処理中 / 処理済み / 完了）。

```yaml
status_col: "ステータス"
status_map:
  "未着手": "未対応"
  "対応中": "処理中"
  "確認待ち": "処理済み"
  "完了": "完了"
```

`status_map` に存在しない値はスキップされ（警告を出力）、ステータスは変更されません。

---

### カスタム属性

```yaml
custom_fields:
  - field_name: "カテゴリ"    # Backlog のカスタム属性名
    col_name: "分類"          # Excel の列名
```

`value_map` を省略した場合は Excel のセルの値をそのまま Backlog に渡します。

#### value_map：値の変換テーブル

Excel の値と Backlog の値が異なる場合に変換テーブルを定義します。テーブルに存在しない値はスキップされます（警告を出力）。

```yaml
value_map:
  "A": "カテゴリA"    # 完全一致
  "B": "カテゴリB"    # 完全一致
```

**正規表現パターン**

キーに正規表現（`re.fullmatch`）を使えます。セル値に改行が含まれる場合も正しくマッチします（`re.DOTALL` 適用済み）。マッチング順序は完全一致が先で、その後定義順に正規表現を評価します。

```yaml
value_map:
  "設計.*":      "設計"    # 「設計A」「設計B」など前方一致
  ".*テスト":    "QA"      # 「単体テスト」「結合テスト」など後方一致
  "(?!.*ABC).*": "未分類"  # 「ABC」を含まない場合（否定先読み）
  "その他":      "未分類"  # 完全一致
```

#### value_separator：複数選択型カスタム属性

typeId 6（複数リスト）・typeId 7（チェックボックス）のカスタム属性に複数の選択肢を渡す場合、`value_separator` でセルの値を分割します。

```yaml
- field_name: "タグ"
  col_name: "タグ"
  value_separator: ","     # 「設計,開発,QA」→ ["設計", "開発", "QA"] に分割
  # value_map:             # 分割後の各値に適用（任意）
  #   "設計": "Design"
```

> typeId 5（単一リスト）・typeId 8（ラジオ）では複数値を渡しても先頭の1件のみ使用します。

#### 必須列チェック

リスト内のいずれか1列でも空の行はスキップされます。

```yaml
required_cols:
  - "タスク名"
  - "対応内容"
```

---

### 重複チェック・更新（upsert）

```yaml
upsert:
  enabled: true
```

`enabled: true` にすると既存の課題を更新し、見つからない場合は新規作成します。重複の判定方法を以下のどちらかで指定します。

#### key_col：issueKey 列で判定する

Excel に issueKey（例: `PROJ-123`）を書いた列がある場合に使います。

```yaml
upsert:
  enabled: true
  key_col: "Backlog課題番号"
```

#### match_summary：件名で検索して判定する

件名が一致する既存課題を検索して更新します。

```yaml
upsert:
  enabled: true
  match_summary: true
```

> `key_col` と `match_summary` を両方指定した場合は `key_col` が優先されます。`key_col` の列に値がある行は issueKey で検索し、値がない行のみ `match_summary` による件名検索にフォールバックします。

---

## トラブルシューティング

**「カスタム属性が更新されない」**

`--debug` オプションをつけて実行すると、送信パラメータとBacklogからのレスポンスに含まれるカスタム属性の値が出力されます。Backlog が値を受け取ったかどうかを確認してください。

```bash
python excel_to_backlog.py --execute --debug
```

**「— 変更なし」と表示されて更新されない**

Backlog 側でフィールドに変更がなかった場合にこのメッセージが出ます。サマリーでは「変更なし」として「スキップ」とは別に集計されます。

更新対象のフィールドに実際の変更が含まれているか、ドライランで確認してください。ドライランは既存課題の照合まで行うため、その行が作成されるのか更新されるのかを事前に確認できます。

**「種別/優先度/ステータスが見つかりません」**

ツール起動時に取得できるマスターデータの名称と、設定ファイルの値が一致しているか確認してください。起動時にターミナルへ一覧が表示されます。

**「カスタム属性が見つかりません」**

カスタム属性名は起動時には表示されません。`--debug` を付けて実行するか、Backlog のプロジェクト設定で正確な名称を確認してください。

**「認識できないキーです」**

設定ファイルのキー名が間違っています。近い名前があれば提案されます。キー名を間違えると値が読まれず、静かに既定値で動作してしまうため、実行前に停止します。

**「担当者が見つかりません」**

Backlog の表示名またはログイン ID と完全一致する文字列を `assignee_col` の列または `default_assignee` に設定してください。使用できる名称はエラーメッセージに一覧表示されます。

**「設定が参照する列名が Excel のヘッダーに存在しません」**

設定に書いた列名が Excel のヘッダーと一致していません。列名が一致しないとフィルター条件が無視されて全行が対象になるなどの誤動作につながるため、実行前に停止します。エラーメッセージに設定箇所・列名・実際のヘッダー一覧が表示されます。前後の空白による不一致は画面上見分けがつかないため、その場合はメッセージで明示されます。

**「Strict Open XML 形式のため読み込めません」**

Excel の「名前を付けて保存」でファイルの種類が **Strict Open XML スプレッドシート** になっています。openpyxl はこの形式に対応していません。ファイルの種類を **Excel ブック (*.xlsx)** にして保存し直してください。

**「ヘッダー名が重複しています」**

同じ列名が複数あります。行データは列名をキーにするため、そのままでは後ろの列が前の列を上書きして内容が失われます。2つ目以降に `備考 (2)` のような連番を付けて区別します。設定から `備考` を参照すると左端の列が使われます。

**「日付として解釈できません」**

`YYYY-MM-DD` / `YYYY/M/D` / `YYYY年M月D日` のいずれかの形式にしてください。和暦（`R7/1/5`）や年のない表記（`9/1`）は解釈できません。この場合、期限日・開始日は未設定のまま登録されます。

**「SSL 証明書エラーが出る」（オンプレ版）**

```yaml
backlog:
  ssl_verify: false
```

**「Backlog のパスが異なる」（オンプレ版）**

```yaml
backlog:
  base_path: "/backlog"
```

---

## 開発

### テスト

```bash
pytest
```

`push` とプルリクエストで GitHub Actions が自動実行します（Python 3.10 / 3.12）。

カバレッジを見る場合:

```bash
pytest --cov=. --cov-report=term-missing
```

主要な観点は以下のファイルにまとまっています。

| ファイル | 観点 |
|---|---|
| `test_upsert_behavior.py` | key_col と match_summary の違い（運用方法の選択に直結） |
| `test_source_excel_untouched.py` | 元 Excel を書き換えないこと |
| `test_cli.py` | 引数・確認フロー・実行ログ・中断時の扱い |
| `test_http_layer.py` | リクエストボディの組み立て（過去に2回バグが出た箇所） |
| `test_excel_reader.py` | 取り消し線・重複ヘッダー・フィルタ |
| `test_mapper.py` | 日付・テンプレート・担当者・カスタム属性 |
