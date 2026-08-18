"""
The Style Spec data model (technical plan, Section 6).

This is the contract between all three stages: Stage 1 produces it, Stage 2
and Stage 3 consume it. It's plain JSON on disk so it isn't tied to any one
library or language.

Only the fields needed for Phases 1-2 are modeled here. chart_style,
table_style, and icon_palette are Phase 3 scope (chart/table/icon
automation) and are left as empty placeholders so the schema is
forward-compatible without pretending those phases are implemented.
"""
from __future__ import annotations

import json
from dataclasses import dataclass, field, asdict
from pathlib import Path
from typing import Optional


# Theme color roles, in the order OOXML defines them in <a:clrScheme>.
THEME_COLOR_ROLES = [
    "dk1", "lt1", "dk2", "lt2",
    "accent1", "accent2", "accent3", "accent4", "accent5", "accent6",
    "hlink", "folHlink",
]

# Roles that are safe to treat as "accent" colors when nearest-color-matching
# an arbitrary RGB value found on a content slide (Stage 3). dk1/lt1/dk2/lt2
# are background/text roles, not decorative accents, so they're excluded.
ACCENT_ROLES = ["accent1", "accent2", "accent3", "accent4", "accent5", "accent6"]


@dataclass
class ThemeFonts:
    major_latin: str = "Calibri"
    minor_latin: str = "Calibri"


@dataclass
class BrandLogo:
    # Position/size are in EMU (English Metric Units, the OOXML length unit),
    # matching python-pptx's Emu type. None means "no logo detected".
    image_path: Optional[str] = None
    left: Optional[int] = None
    top: Optional[int] = None
    width: Optional[int] = None
    height: Optional[int] = None


@dataclass
class BrandFooter:
    text: Optional[str] = None
    show_slide_number: bool = True


@dataclass
class Grid:
    # Simple margin/column model used by Stage 3's grid-alignment pass
    # (Section 5.3.3). EMU units, matching slide width/height.
    margin_left: int = 457200    # 0.5"
    margin_top: int = 457200
    margin_right: int = 457200
    margin_bottom: int = 457200
    columns: int = 12
    gutter: int = 91440           # 0.1"


@dataclass
class StyleSpec:
    """See technical plan Section 6 for the full field reference."""
    theme_colors: dict = field(default_factory=dict)     # role -> "RRGGBB"
    theme_fonts: ThemeFonts = field(default_factory=ThemeFonts)
    brand_logo: BrandLogo = field(default_factory=BrandLogo)
    brand_footer: BrandFooter = field(default_factory=BrandFooter)
    grid: Grid = field(default_factory=Grid)
    slide_width: int = 9144000    # EMU, default 10" (4:3); pipeline overwrites from source
    slide_height: int = 6858000

    # Phase 3 scope - intentionally left empty here.
    chart_style: dict = field(default_factory=dict)
    table_style: dict = field(default_factory=dict)
    icon_palette: list = field(default_factory=list)

    meta: dict = field(default_factory=dict)             # source file, version, timestamp, etc.

    def to_json(self) -> str:
        return json.dumps(_to_plain(self), indent=2)

    def save(self, path: str | Path) -> None:
        Path(path).write_text(self.to_json(), encoding="utf-8")

    @classmethod
    def load(cls, path: str | Path) -> "StyleSpec":
        data = json.loads(Path(path).read_text(encoding="utf-8"))
        return cls.from_dict(data)

    @classmethod
    def from_dict(cls, data: dict) -> "StyleSpec":
        data = dict(data)
        data["theme_fonts"] = ThemeFonts(**data.get("theme_fonts", {}))
        data["brand_logo"] = BrandLogo(**data.get("brand_logo", {}))
        data["brand_footer"] = BrandFooter(**data.get("brand_footer", {}))
        data["grid"] = Grid(**data.get("grid", {}))
        known = {f for f in cls.__dataclass_fields__}
        return cls(**{k: v for k, v in data.items() if k in known})

    def accent_hex_list(self) -> list[str]:
        return [self.theme_colors[r] for r in ACCENT_ROLES if r in self.theme_colors]


def _to_plain(obj):
    if hasattr(obj, "__dataclass_fields__"):
        return {k: _to_plain(v) for k, v in asdict(obj).items()}
    if isinstance(obj, dict):
        return {k: _to_plain(v) for k, v in obj.items()}
    if isinstance(obj, list):
        return [_to_plain(v) for v in obj]
    return obj
