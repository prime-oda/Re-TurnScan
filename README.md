# Re-TurnScan
片面スキャンしかできないADF用マージツール

## 概要

ADF（自動原稿送り装置）で片面ずつスキャンした「表面のみのPDF」と「裏面のみのPDF」を、正しいページ順（1, 2, 3 …）にマージするツールです。

監視フォルダを定期チェックし、以下の条件を満たすPDFペアを自動検出してマージします。

- 最新ファイルの更新時刻から過去5分以内に別のPDFが存在する
- 両ファイルのページ数が等しい

マージ前に Tesseract OSD（Orientation and Script Detection）を使って各ページの向きを検出し、90/180/270 度回転しているページを自動補正します。

## 依存関係

### Python パッケージ

```
pip install -r requirements.txt
```

- **pypdf** ≥ 4.0.0 — PDF の読み書き
- **pdf2image** ≥ 1.16.0 — PDF ページを画像に変換（poppler が必要）

### 外部コマンド

| ツール | 用途 | インストール例 |
|--------|------|---------------|
| **Poppler** (`pdftoppm`) | PDF → 画像変換（pdf2image の依存） | `brew install poppler` / `apt install poppler-utils` |
| **Tesseract OCR** | ページ向き検出（OSD） | `brew install tesseract` / `apt install tesseract-ocr` |
| **Tesseract OSD 言語データ** | OSD に必須 | `apt install tesseract-ocr-osd` （通常 tesseract に同梱） |

> **注意:** Tesseract または Poppler が見つからない場合、OSD はスキップされ回転補正なしでマージを続行します（警告ログを出力）。

## 使い方

```bash
python merge_scan.py <監視フォルダパス>
```

## 回転補正の仕組み

1. マージ直前に、各 PDF ページを画像（150 dpi）に変換します（`pdf2image` / `pdftoppm`）。
2. `tesseract <image> stdout --psm 0 -l osd` を実行し、出力の `Rotate:` 行から回転角度（0/90/180/270）を取得します。
3. 回転が必要なページは `page.rotate(angle)` で補正してからマージします。
4. OSD 判定に失敗した場合（コマンド不在・タイムアウト・出力パース失敗など）は補正なしで続行し、警告ログを出力します。

