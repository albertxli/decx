import os
import yaml

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


def load_config(config_path: str | None = None) -> dict:
    if config_path and os.path.exists(config_path):
        with open(config_path, "r", encoding="utf-8") as f:
            return yaml.safe_load(f)
    return DEFAULT_CONFIG.copy()
