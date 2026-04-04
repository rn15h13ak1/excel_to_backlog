# excel_to_backlog

Excel ファイルから Backlog の課題を一括登録・更新する CLI ツールです。

---

## 概要

Excel で管理しているタスク・問い合わせ一覧などを Backlog 課題として自動登録できます。

- **複数の Excel ファイル**を1つの設定ファイルにまとめて処理できる
- **ドライラン（デフォルト）** で登録内容を事前確認してから実行できる
- **upsert 対応**：既存課題は更新、新規は作成（issueKey または件名で重複判定）
- **カスタム属性**・ステータス・担当者・開始日・期限日に対応
- フィルタリング・テンプレート・value_map など柔軟なマッピング設定が可能

---

## 必要環境・インストール

**Python 3.10 以上**

```bash
pip install openpyxl pyyaml
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
| `--debug` | API リクエスト・レスポンスの詳細を表示する（カスタム属性の反映確認に有用） |

> `--preview` と `--execute` は同時に指定できません。

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

**「スキップ（変更なし）」と表示されて更新されない**

Backlog 側でフィールドに変更がなかった場合にこのメッセージが出ます。更新対象のフィールドに実際の変更が含まれているかドライランで確認してください。

**「種別/優先度/ステータスが見つかりません」**

ツール起動時に取得できるマスターデータの名称と、設定ファイルの値が一致しているか確認してください。起動時にターミナルへ一覧が表示されます。

**「担当者が見つかりません」**

Backlog の表示名またはログイン ID と完全一致する文字列を `assignee_col` の列またはd `default_assignee` に設定してください。使用できる名称はエラーメッセージに一覧表示されます。

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
