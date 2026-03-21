"""Runfile loader for decx — parse Python runfiles into JobSpec/RunSpec."""

import importlib.util
import os
import sys
from dataclasses import dataclass, field

from decx.cli import VALID_STEPS
from decx.config import DEFAULT_CONFIG


@dataclass
class JobSpec:
    """A single job to execute."""

    name: str  # output name (dict key, e.g. "australia")
    template: str  # absolute path to template pptx
    excel: str  # absolute path to excel data file
    output: str  # absolute path to output pptx


@dataclass
class RunSpec:
    """Parsed runfile specification."""

    jobs: list[JobSpec]
    steps: list[str] | None = None  # None = all steps
    config: dict[str, str] = field(default_factory=dict)  # empty = use defaults


def _validate_default_output(default_output: str) -> None:
    """Validate default_output format.

    Must be either:
    - A directory path ending with '/'
    - A .pptx path containing '{name}'

    Raises ValueError otherwise.
    """
    if default_output.endswith("/") or default_output.endswith("\\"):
        return
    if "{name}" in default_output and default_output.endswith(".pptx"):
        return
    raise ValueError(
        f"Invalid default_output: '{default_output}'. "
        "Must end with '/' (directory) or contain '{{name}}' and end with '.pptx'."
    )


def _resolve_output(name: str, default_output: str) -> str:
    """Resolve the output path for a single job."""
    if default_output.endswith("/") or default_output.endswith("\\"):
        return os.path.join(default_output, f"{name}.pptx")
    return default_output.format(name=name)


def _validate_steps(steps: list) -> None:
    """Validate step names."""
    invalid = set(steps) - VALID_STEPS
    if invalid:
        raise ValueError(
            f"Unknown step(s): {', '.join(sorted(invalid))}. "
            f"Valid steps: {', '.join(sorted(VALID_STEPS))}"
        )


def _validate_config_keys(config: dict) -> None:
    """Validate config keys against DEFAULT_CONFIG."""
    valid_keys = set()
    for section, values in DEFAULT_CONFIG.items():
        for key in values:
            valid_keys.add(f"{section}.{key}")

    invalid = set(config.keys()) - valid_keys
    if invalid:
        raise ValueError(
            f"Unknown config key(s): {', '.join(sorted(invalid))}. "
            f"Valid keys: {', '.join(sorted(valid_keys))}"
        )


def load_runfile(path: str) -> RunSpec:
    """Load a Python runfile and return a validated RunSpec.

    The runfile is a .py file with module-level variables:
    - jobs (required): dict[template_path, dict[name, excel_path | {"data": path, "output": path}]]
    - default_output (optional): str — format string with {name} or directory ending in /
    - steps (optional): list[str] — subset of valid step names
    - config (optional): dict[str, str] — config overrides (dot notation keys)

    All paths resolve relative to the runfile's parent directory.
    """
    path = os.path.abspath(path)
    if not os.path.exists(path):
        raise FileNotFoundError(f"Runfile not found: {path}")

    runfile_dir = os.path.dirname(path)

    # Load the module
    module_name = os.path.splitext(os.path.basename(path))[0]
    spec = importlib.util.spec_from_file_location(module_name, path)
    if spec is None or spec.loader is None:
        raise ValueError(f"Cannot load runfile: {path}")
    module = importlib.util.module_from_spec(spec)

    # Temporarily add runfile dir to sys.path for imports
    old_path = sys.path.copy()
    sys.path.insert(0, runfile_dir)
    try:
        spec.loader.exec_module(module)
    finally:
        sys.path = old_path

    # Extract variables
    jobs_raw = getattr(module, "jobs", None)
    default_output = getattr(module, "default_output", "output/{name}.pptx")
    steps = getattr(module, "steps", None)
    config = getattr(module, "config", {})

    # Validate jobs
    if jobs_raw is None:
        raise ValueError("Runfile must define a 'jobs' variable.")
    if not isinstance(jobs_raw, dict) or not jobs_raw:
        raise ValueError("'jobs' must be a non-empty dict.")

    # Validate default_output
    _validate_default_output(default_output)

    # Validate steps
    if steps is not None:
        if not isinstance(steps, list) or not steps:
            raise ValueError("'steps' must be a non-empty list of step names.")
        _validate_steps(steps)

    # Validate config
    if config:
        _validate_config_keys(config)

    # Flatten jobs into JobSpec list
    job_specs: list[JobSpec] = []

    for template_rel, markets in jobs_raw.items():
        if not isinstance(markets, dict) or not markets:
            raise ValueError(
                f"Jobs for template '{template_rel}' must be a non-empty dict."
            )

        template_abs = os.path.normpath(os.path.join(runfile_dir, template_rel))

        for name, value in markets.items():
            if isinstance(value, str):
                # Simple: name → excel_path
                excel_abs = os.path.normpath(os.path.join(runfile_dir, value))
                output_abs = os.path.normpath(
                    os.path.join(runfile_dir, _resolve_output(name, default_output))
                )
            elif isinstance(value, dict):
                # Dict: {"data": excel_path, "output": output_path}
                if "data" not in value:
                    raise ValueError(
                        f"Job '{name}': dict value must have a 'data' key."
                    )
                excel_abs = os.path.normpath(os.path.join(runfile_dir, value["data"]))
                if "output" in value:
                    output_abs = os.path.normpath(
                        os.path.join(runfile_dir, value["output"])
                    )
                else:
                    output_abs = os.path.normpath(
                        os.path.join(runfile_dir, _resolve_output(name, default_output))
                    )
            else:
                raise ValueError(
                    f"Job '{name}': value must be a string (excel path) or dict."
                )

            job_specs.append(
                JobSpec(
                    name=name,
                    template=template_abs,
                    excel=excel_abs,
                    output=output_abs,
                )
            )

    return RunSpec(
        jobs=job_specs,
        steps=steps,
        config=config,
    )
