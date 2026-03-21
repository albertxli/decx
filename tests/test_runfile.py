"""Unit tests for decx.runfile — no COM required."""

import os
import tempfile
import textwrap

import pytest

from decx.runfile import load_runfile, _validate_default_output, _resolve_output


def _write_runfile(tmpdir: str, content: str, name: str = "run.py") -> str:
    """Write a runfile to tmpdir and return its path."""
    path = os.path.join(tmpdir, name)
    with open(path, "w") as f:
        f.write(textwrap.dedent(content))
    return path


class TestValidateDefaultOutput:
    def test_directory_slash(self):
        _validate_default_output("output/")

    def test_directory_backslash(self):
        _validate_default_output("output\\")

    def test_format_string(self):
        _validate_default_output("output/rpm_{name}.pptx")

    def test_no_name_no_slash_raises(self):
        with pytest.raises(ValueError, match="Invalid default_output"):
            _validate_default_output("output/report.pptx")

    def test_no_pptx_raises(self):
        with pytest.raises(ValueError, match="Invalid default_output"):
            _validate_default_output("output/{name}")

    def test_bare_string_raises(self):
        with pytest.raises(ValueError, match="Invalid default_output"):
            _validate_default_output("report")

    def test_name_without_pptx_raises(self):
        with pytest.raises(ValueError, match="Invalid default_output"):
            _validate_default_output("{name}.txt")


class TestResolveOutput:
    def test_directory(self):
        result = _resolve_output("australia", "output/")
        assert result == os.path.join("output/", "australia.pptx")

    def test_format_string(self):
        result = _resolve_output("australia", "output/rpm_{name}.pptx")
        assert result == "output/rpm_australia.pptx"


class TestLoadRunfile:
    def test_load_minimal(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            path = _write_runfile(
                tmpdir,
                """\
                jobs = {
                    "template.pptx": {
                        "argentina": "data.xlsx",
                    },
                }
                """,
            )
            spec = load_runfile(path)
            assert len(spec.jobs) == 1
            assert spec.jobs[0].name == "argentina"
            assert spec.jobs[0].template.endswith("template.pptx")
            assert spec.jobs[0].excel.endswith("data.xlsx")
            assert spec.jobs[0].output.endswith("argentina.pptx")
            assert spec.steps is None
            assert spec.config == {}

    def test_load_full(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            path = _write_runfile(
                tmpdir,
                """\
                jobs = {
                    "template.pptx": {
                        "argentina": "arg.xlsx",
                        "mexico": "mex.xlsx",
                    },
                }
                default_output = "out/rpm_{name}.pptx"
                steps = ["links", "tables"]
                config = {"ccst.positive_prefix": ""}
                """,
            )
            spec = load_runfile(path)
            assert len(spec.jobs) == 2
            assert spec.steps == ["links", "tables"]
            assert spec.config == {"ccst.positive_prefix": ""}
            names = {j.name for j in spec.jobs}
            assert names == {"argentina", "mexico"}
            for job in spec.jobs:
                assert "rpm_" in job.output
                assert job.output.endswith(".pptx")

    def test_missing_jobs_raises(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            path = _write_runfile(tmpdir, "steps = ['links']\n")
            with pytest.raises(ValueError, match="must define a 'jobs'"):
                load_runfile(path)

    def test_empty_jobs_raises(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            path = _write_runfile(tmpdir, "jobs = {}\n")
            with pytest.raises(ValueError, match="non-empty dict"):
                load_runfile(path)

    def test_invalid_step_raises(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            path = _write_runfile(
                tmpdir,
                """\
                jobs = {"t.pptx": {"a": "d.xlsx"}}
                steps = ["bogus"]
                """,
            )
            with pytest.raises(ValueError, match="Unknown step"):
                load_runfile(path)

    def test_invalid_config_key_raises(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            path = _write_runfile(
                tmpdir,
                """\
                jobs = {"t.pptx": {"a": "d.xlsx"}}
                config = {"nonexistent.key": "value"}
                """,
            )
            with pytest.raises(ValueError, match="Unknown config key"):
                load_runfile(path)

    def test_paths_relative_to_runfile(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            path = _write_runfile(
                tmpdir,
                """\
                jobs = {
                    "templates/region1.pptx": {
                        "aus": "data/aus.xlsx",
                    },
                }
                default_output = "output/"
                """,
            )
            spec = load_runfile(path)
            job = spec.jobs[0]
            # All paths should be absolute and rooted in tmpdir
            assert os.path.isabs(job.template)
            assert os.path.isabs(job.excel)
            assert os.path.isabs(job.output)
            assert tmpdir in job.template
            assert tmpdir in job.excel
            assert tmpdir in job.output

    def test_default_output_format_string(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            path = _write_runfile(
                tmpdir,
                """\
                jobs = {"t.pptx": {"australia": "d.xlsx"}}
                default_output = "out/rpm_{name}.pptx"
                """,
            )
            spec = load_runfile(path)
            assert spec.jobs[0].output.endswith("rpm_australia.pptx")

    def test_default_output_directory(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            path = _write_runfile(
                tmpdir,
                """\
                jobs = {"t.pptx": {"australia": "d.xlsx"}}
                default_output = "results/"
                """,
            )
            spec = load_runfile(path)
            assert spec.jobs[0].output.endswith("australia.pptx")
            assert "results" in spec.jobs[0].output

    def test_default_output_omitted(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            path = _write_runfile(
                tmpdir,
                """\
                jobs = {"t.pptx": {"australia": "d.xlsx"}}
                """,
            )
            spec = load_runfile(path)
            assert spec.jobs[0].output.endswith("australia.pptx")
            assert "output" in spec.jobs[0].output

    def test_default_output_no_name_no_slash_raises(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            path = _write_runfile(
                tmpdir,
                """\
                jobs = {"t.pptx": {"a": "d.xlsx"}}
                default_output = "output/report.pptx"
                """,
            )
            with pytest.raises(ValueError, match="Invalid default_output"):
                load_runfile(path)

    def test_default_output_no_pptx_raises(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            path = _write_runfile(
                tmpdir,
                """\
                jobs = {"t.pptx": {"a": "d.xlsx"}}
                default_output = "output/{name}"
                """,
            )
            with pytest.raises(ValueError, match="Invalid default_output"):
                load_runfile(path)

    def test_default_output_bare_string_raises(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            path = _write_runfile(
                tmpdir,
                """\
                jobs = {"t.pptx": {"a": "d.xlsx"}}
                default_output = "report"
                """,
            )
            with pytest.raises(ValueError, match="Invalid default_output"):
                load_runfile(path)

    def test_per_job_output_override(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            path = _write_runfile(
                tmpdir,
                """\
                jobs = {
                    "t.pptx": {
                        "australia": "aus.xlsx",
                        "japan": {"data": "jpn.xlsx", "output": "special/japan.pptx"},
                    },
                }
                default_output = "output/"
                """,
            )
            spec = load_runfile(path)
            jobs_by_name = {j.name: j for j in spec.jobs}
            assert jobs_by_name["australia"].output.endswith("australia.pptx")
            assert "output" in jobs_by_name["australia"].output
            assert jobs_by_name["japan"].output.endswith("japan.pptx")
            assert "special" in jobs_by_name["japan"].output

    def test_multi_template(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            path = _write_runfile(
                tmpdir,
                """\
                jobs = {
                    "region1.pptx": {
                        "australia": "aus.xlsx",
                        "japan": "jpn.xlsx",
                    },
                    "region4.pptx": {
                        "argentina": "arg.xlsx",
                    },
                }
                default_output = "output/"
                """,
            )
            spec = load_runfile(path)
            assert len(spec.jobs) == 3
            templates = {j.template for j in spec.jobs}
            assert len(templates) == 2
            names = {j.name for j in spec.jobs}
            assert names == {"australia", "japan", "argentina"}

    def test_job_dict_missing_data_key_raises(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            path = _write_runfile(
                tmpdir,
                """\
                jobs = {
                    "t.pptx": {
                        "japan": {"output": "out.pptx"},
                    },
                }
                """,
            )
            with pytest.raises(ValueError, match="must have a 'data' key"):
                load_runfile(path)

    def test_runfile_not_found(self):
        with pytest.raises(FileNotFoundError):
            load_runfile("/nonexistent/path/run.py")
