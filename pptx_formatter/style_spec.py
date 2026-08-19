"""
The Style Spec - a versioned JSON contract between all three stages.

Extraction produces it, layout generation and formatting consume it, and
the Template Bank archives it. Nothing here depends on python-pptx, so a
future consumer (a PowerPoint add-in, a rendering service) can read the
same document without inheriting this library.

Document shape (v1.0):

    theme.colors      role -> hex, for dk1/lt1/dk2/lt2/accent1-6/hlink/folHlink
    theme.fonts       major/minor family per script: Latin, East Asian, complex
    layouts[]         per layout: archetype, placeholder geometry, source
    brand.logo        image reference, geometry, and which layouts it appears on
    brand.footer      footer text/field rules and page-number behavior
    grid              margins, gutters, and column/row guide positions
    chart_style       series color rotation, font, gridline/axis styling
    table_style       header shading, borders, cell padding, fonts
    icon_palette      accent colors approved for monochrome icon recoloring
    meta              client/project ids, source master, spec version, timestamp

Two conventions run through the whole document:

1. **Geometry is fractions of the slide, never EMU.** A layout banked from
   a 4:3 master has to be reusable on a 16:9 submission, and absolute
   offsets don't survive that. Everything is 0.0-1.0; `to_emu()` resolves
   against the target deck's dimensions.

2. **`None` means "inherits".** A placeholder whose font size is None
   inherits from the master, which is the desired state. Recording that
   distinction is what lets the inheritance check tell a real override
   apart from an inherited value.
"""
from __future__ import annotations

import json
from dataclasses import dataclass, field, asdict
from pathlib import Path
from typing import Optional

from . import archetypes

# Bumped when the document shape changes in a way older readers can't handle.
SPEC_VERSION = "1.0"

# Theme color roles, in the order OOXML defines them in <a:clrScheme>.
THEME_COLOR_ROLES = [
    "dk1", "lt1", "dk2", "lt2",
    "accent1", "accent2", "accent3", "accent4", "accent5", "accent6",
    "hlink", "folHlink",
]

# Roles safe to treat as "accent" colors when nearest-color-matching an
# arbitrary RGB found on a content slide. dk/lt roles are background and
# text roles, not decorative accents, so they're excluded.
ACCENT_ROLES = ["accent1", "accent2", "accent3", "accent4", "accent5", "accent6"]


# --- theme ---------------------------------------------------------------

@dataclass
class ThemeFonts:
    """
    Major (heading) and minor (body) family per script.

    OOXML's `<a:fontScheme>` declares three scripts per family: `latin`,
    `ea` (East Asian) and `cs` (complex script). Arabic resolves through
    `cs`, which is why Prezlab's bilingual decks need it carried explicitly
    rather than collapsed into the Latin face.
    """
    major_latin: str = "Calibri"
    minor_latin: str = "Calibri"
    major_ea: str = ""       # East Asian; empty means "none declared"
    minor_ea: str = ""
    major_cs: str = ""       # complex script (Arabic, Hebrew, Thai...)
    minor_cs: str = ""

    def latin_for(self, is_major: bool) -> str:
        return self.major_latin if is_major else self.minor_latin

    def cs_for(self, is_major: bool) -> str:
        """The Arabic typeface, falling back to the Latin face when the
        theme declares no separate complex-script family."""
        return (self.major_cs if is_major else self.minor_cs) or self.latin_for(is_major)

    def ea_for(self, is_major: bool) -> str:
        return (self.major_ea if is_major else self.minor_ea) or self.latin_for(is_major)


@dataclass
class Theme:
    colors: dict = field(default_factory=dict)      # role -> "RRGGBB"
    fonts: ThemeFonts = field(default_factory=ThemeFonts)


# --- brand ---------------------------------------------------------------

# Which layouts the logo appears on.
LOGO_ALL = "all"
LOGO_EXCEPT_TITLE = "except_title"      # the common case: not on covers
LOGO_TITLE_ONLY = "title_only"
LOGO_NONE = "none"


@dataclass
class BrandLogo:
    """
    Logo geometry as slide fractions, the image bytes on disk, and the rule
    governing which layouts it gets stamped onto.

    `asset_path` is what makes re-insertion possible - geometry alone only
    records where a logo used to be.
    """
    asset_path: Optional[str] = None
    left_frac: Optional[float] = None
    top_frac: Optional[float] = None
    width_frac: Optional[float] = None
    height_frac: Optional[float] = None
    layout_rule: str = LOGO_EXCEPT_TITLE
    exclude_archetypes: list = field(default_factory=list)

    @property
    def present(self) -> bool:
        return bool(self.asset_path) and self.left_frac is not None

    def appears_on(self, archetype: str) -> bool:
        """Whether the logo belongs on a layout of this archetype."""
        if not self.present or self.layout_rule == LOGO_NONE:
            return False
        if archetype in self.exclude_archetypes:
            return False
        cover_like = (archetypes.TITLE_SLIDE, archetypes.CLOSING, archetypes.PICTURE_FULL)
        if self.layout_rule == LOGO_EXCEPT_TITLE:
            return archetype not in cover_like
        if self.layout_rule == LOGO_TITLE_ONLY:
            return archetype in cover_like
        return True


@dataclass
class BrandFooter:
    """Footer text/field rules plus page-number placeholder behavior."""
    text: Optional[str] = None
    show_footer: bool = True
    show_slide_number: bool = True
    show_date: bool = False
    date_format: Optional[str] = None       # OOXML field format, e.g. "datetime1"
    # Same vocabulary as the logo rule; covers usually carry no page number.
    footer_rule: str = LOGO_EXCEPT_TITLE
    slide_number_rule: str = LOGO_EXCEPT_TITLE

    left_frac: Optional[float] = None
    top_frac: Optional[float] = None
    width_frac: Optional[float] = None
    height_frac: Optional[float] = None

    slide_number_left_frac: Optional[float] = None
    slide_number_top_frac: Optional[float] = None
    slide_number_width_frac: Optional[float] = None
    slide_number_height_frac: Optional[float] = None

    def _rule_allows(self, rule: str, archetype: str) -> bool:
        cover_like = (archetypes.TITLE_SLIDE, archetypes.CLOSING, archetypes.PICTURE_FULL)
        if rule == LOGO_NONE:
            return False
        if rule == LOGO_EXCEPT_TITLE:
            return archetype not in cover_like
        if rule == LOGO_TITLE_ONLY:
            return archetype in cover_like
        return True

    def footer_on(self, archetype: str) -> bool:
        return self.show_footer and self._rule_allows(self.footer_rule, archetype)

    def slide_number_on(self, archetype: str) -> bool:
        return self.show_slide_number and self._rule_allows(self.slide_number_rule, archetype)


@dataclass
class Brand:
    logo: BrandLogo = field(default_factory=BrandLogo)
    footer: BrandFooter = field(default_factory=BrandFooter)


# --- backgrounds & layouts ----------------------------------------------

@dataclass
class BackgroundSpec:
    """
    A layout or master background.

    `kind` is one of "inherit" (no explicit background - use the master's),
    "solid" (literal RGB), "theme" (a theme color role), "image" (picture
    fill, bytes at `asset_path`), or "gradient".
    """
    kind: str = "inherit"
    color_hex: Optional[str] = None
    theme_role: Optional[str] = None
    asset_path: Optional[str] = None
    gradient_stops: list = field(default_factory=list)   # list of "RRGGBB"


@dataclass
class PlaceholderSpec:
    """One placeholder on a layout, geometry normalized to the slide."""
    ph_type: str = "body"          # OOXML p:ph/@type
    idx: Optional[int] = None      # OOXML p:ph/@idx; None for title
    name: str = ""
    left_frac: float = 0.0
    top_frac: float = 0.0
    width_frac: float = 1.0
    height_frac: float = 1.0

    # Typography explicitly declared on the layout. None means "inherit".
    font_size_pt: Optional[float] = None
    bold: Optional[bool] = None
    alignment: Optional[str] = None    # OOXML: "l" | "ctr" | "r" | "just"
    anchor: Optional[str] = None       # OOXML: "t" | "ctr" | "b"
    rtl: Optional[bool] = None

    def to_emu(self, slide_width: int, slide_height: int) -> tuple[int, int, int, int]:
        """Resolve fractional geometry against a concrete slide size."""
        return (
            int(round(self.left_frac * slide_width)),
            int(round(self.top_frac * slide_height)),
            int(round(self.width_frac * slide_width)),
            int(round(self.height_frac * slide_height)),
        )

    @property
    def is_title(self) -> bool:
        return self.ph_type in ("title", "ctrTitle")


# LayoutSpec.source values.
SOURCE_DESIGNER = "designer"        # authored in the submitted master
SOURCE_GENERATED = "generated"      # synthesized from the submission's own grid


@dataclass
class LayoutSpec:
    """A single slide layout, as extracted or as banked."""
    name: str = ""
    archetype: str = archetypes.TITLE_AND_CONTENT
    ooxml_type: str = "obj"        # p:sldLayout/@type
    placeholders: list = field(default_factory=list)   # list[PlaceholderSpec]
    background: BackgroundSpec = field(default_factory=BackgroundSpec)
    # "designer" | "generated" | "bank:<entry_id>"
    source: str = SOURCE_DESIGNER

    def signature(self) -> archetypes.Signature:
        """Structural fingerprint, used for closest-match selection."""
        n_title = n_body = n_pic = n_tbl = n_cht = n_other = 0
        lefts = []
        for ph in self.placeholders:
            t = ph.ph_type
            if t in ("title", "ctrTitle"):
                n_title += 1
            elif t in ("body", "subTitle", "obj"):
                n_body += 1
                lefts.append(round(ph.left_frac, 2))
            elif t == "pic":
                n_pic += 1
                lefts.append(round(ph.left_frac, 2))
            elif t == "tbl":
                n_tbl += 1
            elif t in ("chart", "dgm"):
                n_cht += 1
            elif t in ("dt", "ftr", "sldNum"):
                continue      # furniture, not structure
            else:
                n_other += 1
        return archetypes.Signature(
            n_title=n_title, n_body=n_body, n_picture=n_pic,
            n_table=n_tbl, n_chart=n_cht, n_other=n_other,
            columns=max(1, len(set(lefts))),
        )

    def content_placeholders(self) -> list:
        """Placeholders holding real content, excluding date/footer/number."""
        return [p for p in self.placeholders if p.ph_type not in ("dt", "ftr", "sldNum")]

    @property
    def is_banked(self) -> bool:
        return self.source.startswith("bank:")


# --- grid ----------------------------------------------------------------

@dataclass
class Grid:
    """
    Margins, gutters and guide positions inferred from the master.

    Guides are stored as explicit fraction lists rather than being recomputed
    by every consumer, so a designer reviewing the spec can see the actual
    guide positions the tool will snap to.
    """
    margin_left_frac: float = 0.05
    margin_top_frac: float = 0.07
    margin_right_frac: float = 0.05
    margin_bottom_frac: float = 0.07
    columns: int = 12
    gutter_frac: float = 0.01
    rows: int = 12
    row_gutter_frac: float = 0.0
    column_guides: list = field(default_factory=list)   # x fractions
    row_guides: list = field(default_factory=list)      # y fractions

    def margins_emu(self, slide_width: int, slide_height: int) -> tuple[int, int, int, int]:
        return (
            int(self.margin_left_frac * slide_width),
            int(self.margin_top_frac * slide_height),
            int(self.margin_right_frac * slide_width),
            int(self.margin_bottom_frac * slide_height),
        )

    def compute_guides(self) -> None:
        """(Re)derive column/row guide fractions from margins and counts."""
        self.column_guides = _guide_positions(
            self.margin_left_frac, self.margin_right_frac, self.columns, self.gutter_frac
        )
        self.row_guides = _guide_positions(
            self.margin_top_frac, self.margin_bottom_frac, self.rows, self.row_gutter_frac
        )

    def column_width_frac(self) -> float:
        usable = 1.0 - self.margin_left_frac - self.margin_right_frac
        if self.columns <= 0:
            return usable
        return (usable - self.gutter_frac * (self.columns - 1)) / self.columns


def _guide_positions(start_margin: float, end_margin: float, count: int, gutter: float) -> list:
    """Guide line positions (fractions) for `count` tracks between margins."""
    if count <= 0:
        return []
    usable = 1.0 - start_margin - end_margin
    track = (usable - gutter * (count - 1)) / count
    if track <= 0:
        return []
    stride = track + gutter
    return [round(start_margin + i * stride, 5) for i in range(count)]


# --- chart / table / icon styling ---------------------------------------

@dataclass
class ChartStyle:
    """Native-chart styling derived from the brand's theme."""
    series_colors: list = field(default_factory=list)   # rotation, hex
    font: str = ""
    font_size_pt: float = 12.0
    title_font_size_pt: float = 14.0
    gridline_color: Optional[str] = None
    gridline_width_pt: float = 0.75
    axis_color: Optional[str] = None
    axis_font_size_pt: float = 10.0
    legend_position: str = "b"        # b | t | l | r | none
    show_major_gridlines: bool = True
    show_minor_gridlines: bool = False


@dataclass
class TableStyle:
    """Table styling derived from the brand's theme."""
    header_fill: Optional[str] = None
    header_font_color: Optional[str] = None
    header_bold: bool = True
    header_font_size_pt: float = 12.0
    body_font_size_pt: float = 11.0
    body_font_color: Optional[str] = None
    banded_fill: Optional[str] = None      # alternating row shade, None = off
    border_color: Optional[str] = None
    border_width_pt: float = 0.75
    cell_padding_emu: int = 45720          # 0.05"
    font: str = ""


# --- meta ----------------------------------------------------------------

@dataclass
class Meta:
    """Provenance. What makes an archived spec re-runnable months later."""
    spec_version: str = SPEC_VERSION
    client: Optional[str] = None
    project: Optional[str] = None
    source_master: Optional[str] = None
    source_name: Optional[str] = None
    extracted_at: Optional[str] = None      # ISO 8601 UTC
    layouts_found: int = 0
    content_slides_ignored: int = 0
    notes: list = field(default_factory=list)


# --- the document --------------------------------------------------------

@dataclass
class StyleSpec:
    """The full extracted identity of a submitted master."""
    theme: Theme = field(default_factory=Theme)
    brand: Brand = field(default_factory=Brand)
    layouts: list = field(default_factory=list)          # list[LayoutSpec]
    grid: Grid = field(default_factory=Grid)
    master_background: BackgroundSpec = field(default_factory=BackgroundSpec)

    chart_style: ChartStyle = field(default_factory=ChartStyle)
    table_style: TableStyle = field(default_factory=TableStyle)
    icon_palette: list = field(default_factory=list)     # approved accent hexes

    slide_width: int = 12192000    # EMU, 13.333" (16:9)
    slide_height: int = 6858000

    meta: Meta = field(default_factory=Meta)

    # -- ergonomic accessors ----------------------------------------------
    # The JSON nests theme/brand per the contract; these keep call sites in
    # the pipeline short without flattening the document itself.

    @property
    def theme_colors(self) -> dict:
        return self.theme.colors

    @property
    def theme_fonts(self) -> ThemeFonts:
        return self.theme.fonts

    @property
    def brand_logo(self) -> BrandLogo:
        return self.brand.logo

    @property
    def brand_footer(self) -> BrandFooter:
        return self.brand.footer

    # -- serialization -----------------------------------------------------

    def to_dict(self) -> dict:
        return _to_plain(self)

    def to_json(self) -> str:
        return json.dumps(self.to_dict(), indent=2, ensure_ascii=False)

    def save(self, path: str | Path) -> None:
        Path(path).write_text(self.to_json(), encoding="utf-8")

    @classmethod
    def load(cls, path: str | Path) -> "StyleSpec":
        return cls.from_dict(json.loads(Path(path).read_text(encoding="utf-8")))

    @classmethod
    def from_dict(cls, data: dict) -> "StyleSpec":
        data = dict(data)

        theme = dict(data.get("theme", {}))
        theme["fonts"] = ThemeFonts(**theme.get("fonts", {}))
        data["theme"] = Theme(**{k: v for k, v in theme.items() if k in Theme.__dataclass_fields__})

        brand = dict(data.get("brand", {}))
        brand["logo"] = BrandLogo(**brand.get("logo", {}))
        brand["footer"] = BrandFooter(**brand.get("footer", {}))
        data["brand"] = Brand(**{k: v for k, v in brand.items() if k in Brand.__dataclass_fields__})

        data["grid"] = Grid(**data.get("grid", {}))
        data["master_background"] = BackgroundSpec(**data.get("master_background", {}))
        data["chart_style"] = ChartStyle(**data.get("chart_style", {}))
        data["table_style"] = TableStyle(**data.get("table_style", {}))
        data["meta"] = Meta(**data.get("meta", {}))
        data["layouts"] = [_layout_from_dict(d) for d in data.get("layouts", [])]

        known = set(cls.__dataclass_fields__)
        return cls(**{k: v for k, v in data.items() if k in known})

    # -- convenience -------------------------------------------------------

    def accent_hex_list(self) -> list:
        return [self.theme.colors[r] for r in ACCENT_ROLES if r in self.theme.colors]

    def archetypes_present(self) -> set:
        return {l.archetype for l in self.layouts}

    def layout_for(self, archetype: str):
        for layout in self.layouts:
            if layout.archetype == archetype:
                return layout
        return None

    @property
    def aspect_ratio(self) -> float:
        return self.slide_width / self.slide_height if self.slide_height else 1.7778


def _layout_from_dict(d: dict) -> LayoutSpec:
    d = dict(d)
    d["placeholders"] = [
        PlaceholderSpec(**{k: v for k, v in p.items() if k in PlaceholderSpec.__dataclass_fields__})
        for p in d.get("placeholders", [])
    ]
    d["background"] = BackgroundSpec(**d.get("background", {}))
    known = set(LayoutSpec.__dataclass_fields__)
    return LayoutSpec(**{k: v for k, v in d.items() if k in known})


def _to_plain(obj):
    if hasattr(obj, "__dataclass_fields__"):
        return {k: _to_plain(v) for k, v in asdict(obj).items()}
    if isinstance(obj, dict):
        return {k: _to_plain(v) for k, v in obj.items()}
    if isinstance(obj, (list, tuple)):
        return [_to_plain(v) for v in obj]
    return obj
