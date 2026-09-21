# excel_to_backlog 固有の取り決め

共通規約は [`../ws-conventions/README.md`](../ws-conventions/README.md) に従う。
本ファイルには、本リポジトリだけの事情を書く。

## 規約からの逸脱: commit / push を自動で行う

共通規約 A は「commit / push は、利用者が明示的に指示したときだけ実行する」としている。
**本リポジトリでは、修正のたびに自動で commit / push する。** 利用者の指示による。

本リポジトリは利用者 1 人の開発環境にあり、レビューの相手がいない。区切りごとに
確認を挟んでも判断材料は増えず、往復だけが増える。テストが通ることを commit の
前提にしているため、壊れたまま push されることはない。

| | |
|---|---|
| 対象 | 本リポジトリのみ |
| 前提 | 全テストが通ること。落ちていれば commit しない |
| 単位 | 修正ごと。1 コミットに大量の修正を詰めない |
| 他リポジトリ | 共通規約 B のとおり、変更系 git は行わない |

## 作業のたびに実行する

Markdown を編集したら、commit 前に検査する。

```bash
../ws-conventions/bin/check-markdown.sh .
../ws-conventions/bin/check-privacy.sh .
```

検査やスクリプトが 3 つを超えたら、手順そのものも検査する。

```bash
../ws-conventions/bin/check-commands.sh .
```

ADR は使っていないため、語の検査と索引の生成は対象外。`docs/term-rules.md` を
置いて決定を記録するようになったら、次も実行する。

```bash
../ws-conventions/bin/check-terms.sh .
../ws-conventions/bin/gen-decision-index.py .
```

テストは、依存（openpyxl / PyYAML / pytest）を入れた Python で実行する。

```bash
python -m pytest -q -p no:warnings
```

## 元の Excel は変更しない

読み込んだ Excel を書き換える機能は持たない。作成した issueKey の書き戻しも行わない。
openpyxl で開き直して保存すると、数式が失われるほか、Excel 側で作成したグラフ・
ピボットテーブル・画像などが失われる可能性があるため。

`tests/test_source_excel_untouched.py` がこの方針をテストとして固定している。

## 設定ファイル

`config.yaml` は Backlog の API キーを含むため追跡しない。サンプルは
`config.sample.yaml`（全項目）と `config.minimal.yaml`（最小構成）の 2 つ。
