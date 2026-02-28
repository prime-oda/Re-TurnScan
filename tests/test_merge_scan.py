"""
Re-TurnScan ユニットテスト

merge_scan モジュールのコア関数をテストする。
"""

import os
import subprocess
import time
from pathlib import Path
from unittest.mock import MagicMock, patch

import pytest
from pypdf import PdfReader, PdfWriter

import merge_scan


# ── ヘルパー ──────────────────────────────────────────────────────────────────

def _make_pdf(path: Path, num_pages: int) -> None:
    """指定ページ数の空白PDFを作成する。"""
    writer = PdfWriter()
    for _ in range(num_pages):
        writer.add_blank_page(width=595, height=842)
    with open(path, "wb") as f:
        writer.write(f)


def _page_count(path: Path) -> int:
    return len(PdfReader(str(path)).pages)


# ── get_pdf_page_count ────────────────────────────────────────────────────────

class TestGetPdfPageCount:
    def test_valid_pdf(self, tmp_path):
        pdf = tmp_path / "test.pdf"
        _make_pdf(pdf, 3)
        assert merge_scan.get_pdf_page_count(pdf) == 3

    def test_invalid_file_returns_minus_one(self, tmp_path):
        bad = tmp_path / "bad.pdf"
        bad.write_bytes(b"not a pdf")
        assert merge_scan.get_pdf_page_count(bad) == -1

    def test_missing_file_returns_minus_one(self, tmp_path):
        missing = tmp_path / "missing.pdf"
        assert merge_scan.get_pdf_page_count(missing) == -1


# ── detect_rotation_osd ───────────────────────────────────────────────────────

class TestDetectRotationOsd:
    def _make_tesseract_result(self, stdout: str, returncode: int = 0):
        result = MagicMock()
        result.stdout = stdout
        result.returncode = returncode
        return result

    def test_returns_0_when_no_rotation(self):
        output = "Orientation in degrees: 0\nRotate: 0\nOrientation confidence: 9.50\n"
        with patch("subprocess.run", return_value=self._make_tesseract_result(output)):
            assert merge_scan.detect_rotation_osd(Path("dummy.png")) == 0

    def test_returns_90(self):
        output = "Orientation in degrees: 270\nRotate: 90\nOrientation confidence: 8.00\n"
        with patch("subprocess.run", return_value=self._make_tesseract_result(output)):
            assert merge_scan.detect_rotation_osd(Path("dummy.png")) == 90

    def test_returns_180(self):
        output = "Orientation in degrees: 180\nRotate: 180\nOrientation confidence: 7.00\n"
        with patch("subprocess.run", return_value=self._make_tesseract_result(output)):
            assert merge_scan.detect_rotation_osd(Path("dummy.png")) == 180

    def test_returns_270(self):
        output = "Orientation in degrees: 90\nRotate: 270\nOrientation confidence: 9.00\n"
        with patch("subprocess.run", return_value=self._make_tesseract_result(output)):
            assert merge_scan.detect_rotation_osd(Path("dummy.png")) == 270

    def test_returns_0_when_tesseract_not_found(self):
        with patch("subprocess.run", side_effect=FileNotFoundError):
            assert merge_scan.detect_rotation_osd(Path("dummy.png")) == 0

    def test_returns_0_when_timeout(self):
        with patch("subprocess.run", side_effect=subprocess.TimeoutExpired(cmd="tesseract", timeout=30)):
            assert merge_scan.detect_rotation_osd(Path("dummy.png")) == 0

    def test_returns_0_when_rotate_line_missing(self):
        output = "Orientation in degrees: 90\nOrientation confidence: 9.50\n"
        with patch("subprocess.run", return_value=self._make_tesseract_result(output)):
            assert merge_scan.detect_rotation_osd(Path("dummy.png")) == 0

    def test_returns_0_for_unexpected_angle(self):
        output = "Rotate: 45\n"
        with patch("subprocess.run", return_value=self._make_tesseract_result(output)):
            assert merge_scan.detect_rotation_osd(Path("dummy.png")) == 0


# ── correct_pdf_pages ─────────────────────────────────────────────────────────

class TestCorrectPdfPages:
    def test_returns_pages_when_pdf2image_unavailable(self, tmp_path):
        """pdf2image がない場合は元ページをそのまま返す。"""
        pdf = tmp_path / "test.pdf"
        _make_pdf(pdf, 2)
        with patch.dict("sys.modules", {"pdf2image": None}):
            pages = merge_scan.correct_pdf_pages(pdf)
        assert len(pages) == 2

    def test_returns_pages_when_convert_fails(self, tmp_path):
        """pdf2image の変換が例外を起こしても元ページを返す。"""
        pdf = tmp_path / "test.pdf"
        _make_pdf(pdf, 3)
        mock_pdf2image = MagicMock()
        mock_pdf2image.convert_from_path.side_effect = Exception("変換失敗")
        with patch.dict("sys.modules", {"pdf2image": mock_pdf2image}):
            pages = merge_scan.correct_pdf_pages(pdf)
        assert len(pages) == 3

    def test_rotates_pages_based_on_osd(self, tmp_path):
        """OSD が 180 度を返したページが rotate() される。"""
        pdf = tmp_path / "test.pdf"
        _make_pdf(pdf, 1)

        mock_image = MagicMock()
        mock_pdf2image = MagicMock()
        mock_pdf2image.convert_from_path.return_value = [mock_image]

        osd_output = "Rotate: 180\n"
        with patch.dict("sys.modules", {"pdf2image": mock_pdf2image}):
            with patch("subprocess.run", return_value=MagicMock(stdout=osd_output)):
                pages = merge_scan.correct_pdf_pages(pdf)

        assert len(pages) == 1
        # rotate() が適用されたページの /Rotate エントリを確認
        assert pages[0].get("/Rotate", 0) == 180

    def test_no_rotation_when_osd_returns_0(self, tmp_path):
        """OSD が 0 を返した場合はページを変更しない。"""
        pdf = tmp_path / "test.pdf"
        _make_pdf(pdf, 1)

        mock_image = MagicMock()
        mock_pdf2image = MagicMock()
        mock_pdf2image.convert_from_path.return_value = [mock_image]

        osd_output = "Rotate: 0\n"
        with patch.dict("sys.modules", {"pdf2image": mock_pdf2image}):
            with patch("subprocess.run", return_value=MagicMock(stdout=osd_output)):
                pages = merge_scan.correct_pdf_pages(pdf)

        assert len(pages) == 1
        assert pages[0].get("/Rotate", 0) == 0




class TestMergeFrontBack:
    def test_merge_produces_correct_page_count(self, tmp_path):
        front = tmp_path / "front.pdf"
        back = tmp_path / "back.pdf"
        output = tmp_path / "merged.pdf"
        _make_pdf(front, 3)
        _make_pdf(back, 3)
        merge_scan.merge_front_back(front, back, output)
        assert _page_count(output) == 6

    def test_merge_page_order(self, tmp_path):
        """
        表面ページは順番通り、裏面ページは逆順で交互に結合されることを確認する。

        各ページに一意の幅を割り当てることでページの同一性を検証する。

        表面PDF (幅): [100, 200, 300]  （スキャン順）
        裏面PDF (幅): [600, 500, 400]  （ADFにより逆順 = 最終ページの裏から）
        マージ後の期待幅: [100, 400, 200, 500, 300, 600]
        """
        def _make_sized_pdf(path: Path, widths: list[int]) -> None:
            """各ページに指定幅を持つPDFを作成する（高さは固定）。"""
            writer = PdfWriter()
            for w in widths:
                writer.add_blank_page(width=w, height=842)
            with open(path, "wb") as f:
                writer.write(f)

        front = tmp_path / "front.pdf"
        back = tmp_path / "back.pdf"
        output = tmp_path / "merged.pdf"

        _make_sized_pdf(front, [100, 200, 300])
        _make_sized_pdf(back, [600, 500, 400])

        merge_scan.merge_front_back(front, back, output)

        reader = PdfReader(str(output))
        widths = [int(float(page.mediabox.width)) for page in reader.pages]
        assert widths == [100, 400, 200, 500, 300, 600]

    def test_merge_raises_on_page_count_mismatch(self, tmp_path):
        front = tmp_path / "front.pdf"
        back = tmp_path / "back.pdf"
        output = tmp_path / "merged.pdf"
        _make_pdf(front, 2)
        _make_pdf(back, 3)
        with pytest.raises(ValueError, match="ページ数が一致しません"):
            merge_scan.merge_front_back(front, back, output)

    def test_output_file_is_created(self, tmp_path):
        front = tmp_path / "front.pdf"
        back = tmp_path / "back.pdf"
        output = tmp_path / "out.pdf"
        _make_pdf(front, 1)
        _make_pdf(back, 1)
        assert not output.exists()
        merge_scan.merge_front_back(front, back, output)
        assert output.exists()


# ── find_and_merge_pairs ──────────────────────────────────────────────────────

class TestFindAndMergePairs:
    def test_pairs_same_page_count_within_window(self, tmp_path):
        front = tmp_path / "front.pdf"
        back = tmp_path / "back.pdf"
        _make_pdf(front, 2)
        _make_pdf(back, 2)

        now = time.time()
        os.utime(front, (now - 120, now - 120))  # 2分前
        os.utime(back, (now, now))               # 現在

        merged = merge_scan.find_and_merge_pairs(tmp_path, set())

        output_files = list(tmp_path.glob(f"{merge_scan.MERGED_PREFIX}*.pdf"))
        assert len(output_files) == 1
        assert len(merged) == 1

    def test_no_pair_when_only_one_pdf(self, tmp_path):
        only = tmp_path / "only.pdf"
        _make_pdf(only, 2)
        merged = merge_scan.find_and_merge_pairs(tmp_path, set())
        assert merged == set()

    def test_no_pair_when_outside_time_window(self, tmp_path):
        front = tmp_path / "front.pdf"
        back = tmp_path / "back.pdf"
        _make_pdf(front, 2)
        _make_pdf(back, 2)

        now = time.time()
        os.utime(front, (now - 400, now - 400))  # 6分40秒前（窓外）
        os.utime(back, (now, now))

        merged = merge_scan.find_and_merge_pairs(tmp_path, set())
        output_files = list(tmp_path.glob(f"{merge_scan.MERGED_PREFIX}*.pdf"))
        assert len(output_files) == 0
        assert merged == set()

    def test_no_pair_when_page_counts_differ(self, tmp_path):
        front = tmp_path / "front.pdf"
        back = tmp_path / "back.pdf"
        _make_pdf(front, 2)
        _make_pdf(back, 3)

        now = time.time()
        os.utime(front, (now - 60, now - 60))
        os.utime(back, (now, now))

        merged = merge_scan.find_and_merge_pairs(tmp_path, set())
        output_files = list(tmp_path.glob(f"{merge_scan.MERGED_PREFIX}*.pdf"))
        assert len(output_files) == 0

    def test_already_merged_pair_skipped(self, tmp_path):
        front = tmp_path / "front.pdf"
        back = tmp_path / "back.pdf"
        _make_pdf(front, 2)
        _make_pdf(back, 2)

        now = time.time()
        os.utime(front, (now - 60, now - 60))
        os.utime(back, (now, now))

        already = merge_scan.find_and_merge_pairs(tmp_path, set())
        assert len(already) == 1

        # 再度呼び出してもマージファイルは増えない
        already = merge_scan.find_and_merge_pairs(tmp_path, already)
        output_files = [
            f for f in tmp_path.glob(f"{merge_scan.MERGED_PREFIX}*.pdf")
        ]
        assert len(output_files) == 1

    def test_merged_files_excluded_from_candidates(self, tmp_path):
        """MERGED_PREFIX で始まるファイルはペア候補にならない。"""
        normal = tmp_path / "scan.pdf"
        merged_existing = tmp_path / f"{merge_scan.MERGED_PREFIX}old.pdf"
        _make_pdf(normal, 2)
        _make_pdf(merged_existing, 2)

        now = time.time()
        os.utime(normal, (now - 60, now - 60))
        os.utime(merged_existing, (now, now))

        result = merge_scan.find_and_merge_pairs(tmp_path, set())
        output_files = [
            f for f in tmp_path.glob(f"{merge_scan.MERGED_PREFIX}*.pdf")
            if f.name != merged_existing.name
        ]
        assert len(output_files) == 0

    def test_dot_files_excluded_from_candidates(self, tmp_path):
        """「.」で始まるファイル（隠しファイル・AppleDoubleファイル等）はペア候補にならない。"""
        normal = tmp_path / "scan.pdf"
        dot_file = tmp_path / "._scan.pdf"
        _make_pdf(normal, 2)
        dot_file.write_bytes(b"\x00\x05\x16\x07\x00")  # 不正なPDFバイト列

        now = time.time()
        os.utime(normal, (now - 60, now - 60))
        os.utime(dot_file, (now, now))

        result = merge_scan.find_and_merge_pairs(tmp_path, set())
        output_files = list(tmp_path.glob(f"{merge_scan.MERGED_PREFIX}*.pdf"))
        assert len(output_files) == 0
        assert result == set()

    def test_older_file_is_front(self, tmp_path):
        """更新時刻が古い方が表面として使われることを確認する。"""
        older = tmp_path / "older.pdf"
        newer = tmp_path / "newer.pdf"
        _make_pdf(older, 1)
        _make_pdf(newer, 1)

        now = time.time()
        os.utime(older, (now - 60, now - 60))
        os.utime(newer, (now, now))

        merge_scan.find_and_merge_pairs(tmp_path, set())

        output_files = list(tmp_path.glob(f"{merge_scan.MERGED_PREFIX}*.pdf"))
        assert len(output_files) == 1
        # 出力ファイル名に古い方（表面）のステムが先に現れる
        assert f"_{older.stem}_and_" in output_files[0].name
