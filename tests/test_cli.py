"""Unit tests for decx.cli and decx.config — no COM required."""

import os
import tempfile

import pytest

from decx.cli import parse_pair, resolve_output_path
from decx.config import apply_overrides, DEFAULT_CONFIG


class TestParsePair:
    def test_simple_pair(self):
        pptx, xlsx = parse_pair("report.pptx:data.xlsx")
        assert pptx.endswith("report.pptx")
        assert xlsx.endswith("data.xlsx")

    def test_windows_drive_pptx(self):
        pptx, xlsx = parse_pair(r"C:\docs\report.pptx:data.xlsx")
        assert pptx == os.path.abspath(r"C:\docs\report.pptx")
        assert xlsx.endswith("data.xlsx")

    def test_windows_drive_xlsx(self):
        pptx, xlsx = parse_pair(r"report.pptx:C:\data\file.xlsx")
        assert pptx.endswith("report.pptx")
        assert xlsx == os.path.abspath(r"C:\data\file.xlsx")

    def test_both_windows_drives(self):
        pptx, xlsx = parse_pair(r"C:\docs\report.pptx:C:\data\file.xlsx")
        assert pptx == os.path.abspath(r"C:\docs\report.pptx")
        assert xlsx == os.path.abspath(r"C:\data\file.xlsx")

    def test_returns_absolute_paths(self):
        pptx, xlsx = parse_pair("report.pptx:data.xlsx")
        assert os.path.isabs(pptx)
        assert os.path.isabs(xlsx)

    def test_no_colon_exits(self):
        with pytest.raises(SystemExit):
            parse_pair("no_colon_here")


class TestResolveOutputPath:
    def test_no_output_returns_original(self):
        result = resolve_output_path("/some/file.pptx", None, False, 1)
        assert result == "/some/file.pptx"

    def test_pptx_output_copies_file(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            src = os.path.join(tmpdir, "source.pptx")
            with open(src, "w") as f:
                f.write("test")
            out = os.path.join(tmpdir, "output.pptx")
            result = resolve_output_path(src, out, False, 1)
            assert result == os.path.abspath(out)
            assert os.path.exists(out)

    def test_directory_output_copies_file(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            src = os.path.join(tmpdir, "source.pptx")
            with open(src, "w") as f:
                f.write("test")
            out_dir = os.path.join(tmpdir, "output_dir")
            result = resolve_output_path(src, out_dir, False, 1)
            expected = os.path.join(os.path.abspath(out_dir), "source.pptx")
            assert result == expected
            assert os.path.exists(expected)

    def test_pptx_output_batch_multiple_exits(self):
        with pytest.raises(SystemExit):
            resolve_output_path("/some/file.pptx", "out.pptx", True, 3)

    def test_directory_output_batch_ok(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            src = os.path.join(tmpdir, "source.pptx")
            with open(src, "w") as f:
                f.write("test")
            out_dir = os.path.join(tmpdir, "batch_out")
            result = resolve_output_path(src, out_dir, True, 3)
            expected = os.path.join(os.path.abspath(out_dir), "source.pptx")
            assert result == expected


class TestApplyOverrides:
    def _base_config(self):
        import copy

        return copy.deepcopy(DEFAULT_CONFIG)

    def test_string_value(self):
        config = apply_overrides(self._base_config(), ["ccst.positive_prefix=+"])
        assert config["ccst"]["positive_prefix"] == "+"

    def test_empty_string(self):
        config = apply_overrides(self._base_config(), ["ccst.positive_prefix="])
        assert config["ccst"]["positive_prefix"] == ""

    def test_int_conversion(self):
        config = apply_overrides(self._base_config(), ["delta.template_slide=2"])
        assert config["delta"]["template_slide"] == 2
        assert isinstance(config["delta"]["template_slide"], int)

    def test_bool_true(self):
        config = apply_overrides(self._base_config(), ["links.set_manual=true"])
        assert config["links"]["set_manual"] is True

    def test_bool_false(self):
        config = apply_overrides(self._base_config(), ["links.set_manual=false"])
        assert config["links"]["set_manual"] is False

    def test_hex_color(self):
        config = apply_overrides(self._base_config(), ["heatmap.color_minimum=#FF0000"])
        assert config["heatmap"]["color_minimum"] == "#FF0000"

    def test_multiple_overrides(self):
        config = apply_overrides(
            self._base_config(),
            ["ccst.positive_prefix=", "ccst.symbol_removal=", "delta.template_slide=3"],
        )
        assert config["ccst"]["positive_prefix"] == ""
        assert config["ccst"]["symbol_removal"] == ""
        assert config["delta"]["template_slide"] == 3

    def test_invalid_key_raises(self):
        with pytest.raises(ValueError, match="Unknown config key"):
            apply_overrides(self._base_config(), ["nonexistent.key=value"])

    def test_no_equals_raises(self):
        with pytest.raises(ValueError, match="Invalid override format"):
            apply_overrides(self._base_config(), ["no_equals_sign"])

    def test_override_wins_over_default(self):
        config = self._base_config()
        assert config["ccst"]["positive_prefix"] == "+"
        config = apply_overrides(config, ["ccst.positive_prefix=-"])
        assert config["ccst"]["positive_prefix"] == "-"
