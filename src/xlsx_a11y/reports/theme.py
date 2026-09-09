"""Theme loader and SMACSS layer assembler with user theme support.

Delegates to engine_a11y.reports.theme.
"""
from pathlib import Path
from typing import Any, Dict, List, Optional, Union
from engine_a11y.reports.theme import (
    BUNDLED_THEMES,
    available_themes as _engine_available_themes,
    theme_css as _engine_theme_css,
)

DEFAULT_USER_CONFIG_DIR = Path.home() / ".config" / "xlsx-a11y"


def available_themes(config_dir: Optional[Union[str, Path]] = None) -> List[Dict[str, Any]]:
    target_config = config_dir if config_dir else DEFAULT_USER_CONFIG_DIR
    return _engine_available_themes(config_dir=target_config)


def theme_css(name: str = "light", config_dir: Optional[Union[str, Path]] = None) -> str:
    target_config = config_dir if config_dir else DEFAULT_USER_CONFIG_DIR
    css = _engine_theme_css(name=name, config_dir=target_config)
    return css.replace("/* docx-a11y theme:", "/* xlsx-a11y theme:")


__all__ = [
    "BUNDLED_THEMES",
    "available_themes",
    "theme_css",
    "DEFAULT_USER_CONFIG_DIR",
]
