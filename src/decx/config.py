"""Configuration defaults and --set override support."""

import copy

DEFAULT_CONFIG = {
    "heatmap": {
        "color_minimum": "#F8696B",
        "color_midpoint": "#FFEB84",
        "color_maximum": "#63BE7B",
        "dark_font": "#000000",
        "light_font": "#FFFFFF",
    },
    "ccst": {
        "positive_color": "#33CC33",
        "negative_color": "#ED0590",
        "neutral_color": "#595959",
        "positive_prefix": "+",
        "symbol_removal": "%",
    },
    "delta": {
        "template_positive": "tmpl_delta_pos",
        "template_negative": "tmpl_delta_neg",
        "template_none": "tmpl_delta_none",
        "template_slide": 1,
    },
    "links": {
        "set_manual": True,
    },
}


def _coerce_value(value: str):
    """Auto-convert string value to appropriate Python type."""
    if value.lower() == "true":
        return True
    if value.lower() == "false":
        return False
    try:
        return int(value)
    except ValueError:
        pass
    try:
        return float(value)
    except ValueError:
        pass
    return value


def apply_overrides(config: dict, overrides: list[str]) -> dict:
    """Apply --set key=value overrides to a config dict.

    Supports dot notation: "ccst.positive_prefix=+"
    Auto-converts types: "true"->True, "1"->1, etc.
    Empty string is preserved as "".

    Raises ValueError for invalid keys.
    """
    valid_keys = set()
    for section, values in DEFAULT_CONFIG.items():
        for key in values:
            valid_keys.add(f"{section}.{key}")

    for override in overrides:
        eq_pos = override.find("=")
        if eq_pos < 0:
            raise ValueError(
                f"Invalid override format: '{override}'. Expected key=value"
            )

        key = override[:eq_pos]
        value = override[eq_pos + 1 :]

        if key not in valid_keys:
            raise ValueError(
                f"Unknown config key: '{key}'. "
                f"Valid keys: {', '.join(sorted(valid_keys))}"
            )

        parts = key.split(".")
        section, name = parts[0], parts[1]

        if section not in config:
            config[section] = {}

        config[section][name] = _coerce_value(value)

    return config


def get_config(overrides: list[str] | None = None) -> dict:
    """Get config with defaults, optionally applying --set overrides."""
    config = copy.deepcopy(DEFAULT_CONFIG)
    if overrides:
        apply_overrides(config, overrides)
    return config
