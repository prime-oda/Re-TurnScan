#!/usr/bin/env python3
"""
Re-TurnScan: ADF片面スキャンPDFマージツール

片面ずつスキャンした「表面のみのPDF」と「裏面のみのPDF」を、
正しいページ順（1, 2, 3...）にマージするツール。

監視フォルダへ毎分確認を行い、
  - 最新PDFの更新時刻から過去5分以内に別のPDFが存在する
  - 両ファイルのページ数が等しい
という条件を満たすペアを自動的に検出してマージする。

ADFで裏面をまとめてスキャンする際、用紙スタックを反転して再投入するため
裏面PDFのページ順は逆（最終ページの裏→先頭ページの裏）になっている。
そのため裏面ページを逆順にしてから表面と交互に結合する。
"""

import logging
import os
import subprocess
import sys
import tempfile
import time
from datetime import datetime, timedelta
from pathlib import Path

from pypdf import PdfReader, PdfWriter

# ── 設定 ──────────────────────────────────────────────────────────────────────
WATCH_INTERVAL_SECONDS = 60   # 監視間隔（秒）
RECENT_FILE_WINDOW_SECONDS = 120  # 最新ファイルとみなす時間窓（監視間隔の2倍）
PAIR_WINDOW_SECONDS = 300     # ペアとみなす時間窓（5分）
MERGED_PREFIX = "merged_"     # マージ済みファイルの接頭辞

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
)
logger = logging.getLogger(__name__)


# ── コア関数 ──────────────────────────────────────────────────────────────────

def detect_rotation_osd(image_path: Path) -> int:
    """
    Tesseract OSD (--psm 0) を使って画像の回転角度を検出する。

    tesseract コマンドが存在しない場合、OSD に失敗した場合などは
    警告ログを出して 0 を返す（回転補正なし）。

    Args:
        image_path: 検査対象の画像ファイルパス

    Returns:
        回転角度 (0, 90, 180, 270)。検出失敗時は 0。
    """
    try:
        result = subprocess.run(
            ["tesseract", str(image_path), "stdout", "--psm", "0", "-l", "osd"],
            capture_output=True,
            text=True,
            timeout=30,
        )
        for line in result.stdout.splitlines():
            if line.startswith("Rotate:"):
                try:
                    angle = int(line.split(":")[1].strip())
                except ValueError:
                    break
                if angle in (0, 90, 180, 270):
                    return angle
                logger.warning("OSD から予期しない回転角度を受け取りました: %d", angle)
                return 0
    except FileNotFoundError:
        logger.warning("tesseract が見つかりません。OSD をスキップします。")
    except subprocess.TimeoutExpired:
        logger.warning("OSD 検出がタイムアウトしました: %s", image_path)
    except Exception as e:
        logger.warning("OSD 検出に失敗しました: %s", e)
    return 0


def correct_pdf_pages(pdf_path: Path) -> list:
    """
    PDF の各ページを Tesseract OSD で解析し、回転補正したページのリストを返す。

    pdf2image (poppler) および tesseract が利用できない場合は
    補正なしで元のページを返す（フォールバック）。

    Args:
        pdf_path: 処理対象の PDF ファイルパス

    Returns:
        回転補正済みの pypdf ページオブジェクトのリスト
    """
    reader = PdfReader(str(pdf_path))
    pages = list(reader.pages)

    try:
        from pdf2image import convert_from_path  # type: ignore[import]
    except ImportError:
        logger.warning(
            "pdf2image が見つかりません。OSD をスキップします。"
            " `pip install pdf2image` でインストールしてください。"
        )
        return pages

    try:
        images = convert_from_path(str(pdf_path), dpi=150)
    except Exception as e:
        logger.warning("PDF→画像変換に失敗しました: %s", e)
        return pages

    with tempfile.TemporaryDirectory() as tmpdir:
        corrected = []
        for i, (page, image) in enumerate(zip(pages, images)):
            img_path = Path(tmpdir) / f"page_{i:04d}.png"
            image.save(str(img_path), "PNG")
            angle = detect_rotation_osd(img_path)
            if angle != 0:
                logger.info("ページ %d を %d 度回転補正します。", i + 1, angle)
                page.rotate(angle)
            corrected.append(page)
    return corrected


def get_pdf_page_count(pdf_path: Path) -> int:
    """PDFのページ数を返す。読み取れない場合は -1 を返す。"""
    try:
        reader = PdfReader(str(pdf_path))
        return len(reader.pages)
    except Exception as e:
        logger.warning("PDFの読み取りに失敗しました: %s: %s", pdf_path, e)
        return -1


def merge_front_back(front_path: Path, back_path: Path, output_path: Path) -> None:
    """
    表面PDFと裏面PDFを正しいページ順にマージする。

    表面は順番通り（ページ 1, 3, 5 … の内容）、裏面はADFの仕様により
    逆順（最後のページの裏から先頭ページの裏の順）となっているため、
    裏面を逆順に並べ直してから 表面[i] → 裏面[n-1-i] の順で交互に結合する。

    Args:
        front_path:  表面のみのPDFファイルパス（スキャン順）
        back_path:   裏面のみのPDFファイルパス（逆スキャン順）
        output_path: 出力先ファイルパス
    Raises:
        ValueError: 両PDFのページ数が異なる場合
    """
    front_pages = correct_pdf_pages(front_path)
    back_pages = correct_pdf_pages(back_path)

    if len(front_pages) != len(back_pages):
        raise ValueError(
            f"ページ数が一致しません: 表面={len(front_pages)}, 裏面={len(back_pages)}"
        )

    writer = PdfWriter()
    # 裏面を逆順にして表面と交互に追加
    for front_page, back_page in zip(front_pages, reversed(back_pages)):
        writer.add_page(front_page)
        writer.add_page(back_page)

    with open(output_path, "wb") as f:
        writer.write(f)

    logger.info("マージ完了: %s (%d ページ)", output_path, len(front_pages) * 2)


def find_and_merge_pairs(folder: Path, already_merged: set) -> set:
    """
    フォルダ内のPDFを検索し、条件に合うペアをマージする。

    条件:
      - 最新ファイルが現在時刻から2分以内に更新されている（新規スキャンの有無を判定）
      - 最新ファイルの更新時刻から過去5分以内に更新された別のPDFがある
      - 両ファイルのページ数が同じ

    マージ済みファイル（MERGED_PREFIX 始まり）は対象外。
    already_merged に登録済みのペアは重複マージしない。

    Args:
        folder:         監視対象フォルダ
        already_merged: 処理済みペアの frozenset 集合

    Returns:
        更新された already_merged セット
    """
    pdf_files = sorted(
        [
            f for f in folder.glob("*.pdf")
            if not f.name.startswith(".")
            and not f.name.startswith(MERGED_PREFIX)
        ],
        key=lambda f: f.stat().st_mtime,
        reverse=True,
    )

    if len(pdf_files) < 2:
        return already_merged

    latest = pdf_files[0]
    latest_mtime = datetime.fromtimestamp(latest.stat().st_mtime)

    # 最新ファイルが一定時間以上前なら新規スキャンなしとみなしてスキップ
    if datetime.now() - latest_mtime > timedelta(seconds=RECENT_FILE_WINDOW_SECONDS):
        return already_merged

    threshold = latest_mtime - timedelta(seconds=PAIR_WINDOW_SECONDS)

    for candidate in pdf_files[1:]:
        candidate_mtime = datetime.fromtimestamp(candidate.stat().st_mtime)
        if candidate_mtime < threshold:
            # ファイルはmtime降順なので以降は全て期限外
            break

        pair_key = frozenset([str(latest), str(candidate)])
        if pair_key in already_merged:
            continue

        latest_pages = get_pdf_page_count(latest)
        candidate_pages = get_pdf_page_count(candidate)

        if latest_pages <= 0 or candidate_pages <= 0:
            continue

        if latest_pages != candidate_pages:
            logger.debug(
                "ページ数不一致のためスキップ: %s(%dp) vs %s(%dp)",
                latest.name, latest_pages,
                candidate.name, candidate_pages,
            )
            continue

        # 古い方が表面、新しい方が裏面
        if candidate.stat().st_mtime < latest.stat().st_mtime:
            front, back = candidate, latest
        else:
            front, back = latest, candidate

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_name = f"{MERGED_PREFIX}{timestamp}_{front.stem}_and_{back.stem}.pdf"
        output_path = folder / output_name

        logger.info(
            "ペアを検出しました: 表面=%s, 裏面=%s (%dページ×2)",
            front.name, back.name, latest_pages,
        )

        try:
            merge_front_back(front, back, output_path)
            already_merged.add(pair_key)
        except Exception as e:
            logger.error("マージに失敗しました: %s", e)

    return already_merged


# ── 監視ループ ─────────────────────────────────────────────────────────────────

def monitor(folder_path: str) -> None:
    """
    指定フォルダを定期的に監視し、条件に合うPDFペアをマージし続ける。

    Args:
        folder_path: 監視対象フォルダのパス
    Raises:
        FileNotFoundError: フォルダが存在しない場合
    """
    folder = Path(folder_path)
    if not folder.is_dir():
        raise FileNotFoundError(f"フォルダが見つかりません: {folder_path}")

    logger.info("監視開始: %s (間隔: %d秒)", folder, WATCH_INTERVAL_SECONDS)

    already_merged: set = set()
    while True:
        try:
            already_merged = find_and_merge_pairs(folder, already_merged)
        except Exception as e:
            logger.error("監視中にエラーが発生しました: %s", e)
        time.sleep(WATCH_INTERVAL_SECONDS)


# ── エントリーポイント ──────────────────────────────────────────────────────────

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print(f"使用方法: python {os.path.basename(sys.argv[0])} <監視フォルダパス>")
        sys.exit(1)

    monitor(sys.argv[1])
