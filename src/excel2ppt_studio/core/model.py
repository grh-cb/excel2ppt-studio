"""数据模型：模板元素、槽位、映射与 Profile。

本模块只描述结构，不依赖 python-pptx / openpyxl，便于单测与跨层复用。
"""

from __future__ import annotations

from dataclasses import asdict, dataclass, field
from typing import Any

EMU_PER_CM = 360_000
EMU_PER_PT = 12_700


@dataclass(frozen=True)
class Rect:
    """绝对坐标与尺寸，单位 EMU（1 cm = 360000 EMU，1 pt = 12700 EMU）。"""

    x: int
    y: int
    w: int
    h: int

    @property
    def cm(self) -> tuple[float, float, float, float]:
        return (
            round(self.x / EMU_PER_CM, 2),
            round(self.y / EMU_PER_CM, 2),
            round(self.w / EMU_PER_CM, 2),
            round(self.h / EMU_PER_CM, 2),
        )

    def to_dict(self) -> dict[str, Any]:
        return {"x": self.x, "y": self.y, "w": self.w, "h": self.h, "cm": list(self.cm)}


@dataclass
class Run:
    """一段连续同格式文本（对应 OOXML ``a:r``）。"""

    text: str
    bold: bool | None = None
    italic: bool | None = None
    size_pt: float | None = None
    color: str | None = None
    font: str | None = None


@dataclass
class Paragraph:
    index: int
    align: str | None = None
    level: int = 0
    runs: list[Run] = field(default_factory=list)

    @property
    def text(self) -> str:
        return "".join(r.text for r in self.runs)


@dataclass
class Slot:
    """Field 内一个可填充位置。

    ``kind``：
    - ``explicit``：``{{x}}`` / ``${x}`` 占位符
    - ``named``：形状名 ``fld_x``
    - ``alt``：替代文字 ``{{x}}``
    - ``label``：``标签：`` 后跟空白/下划线
    - ``blank``：整块空白/下划线
    """

    name: str
    raw: str
    field_id: str
    para: int
    run: int
    kind: str


@dataclass
class TableMeta:
    rows: int
    cols: int
    cells: list[list[str]]
    col_widths: list[int] = field(default_factory=list)
    row_heights: list[int] = field(default_factory=list)
    first_row_header: bool = True


@dataclass
class ImageMeta:
    filename: str = ""
    ext: str = ""
    width_px: int | None = None
    height_px: int | None = None


@dataclass
class Field:
    """模板中的一个可映射元素。"""

    id: str
    kind: str  # text | table | image | chart | group | shape
    name: str
    role: str  # title | body | label | image | table | chart | decoration | unknown
    rect: Rect
    z: int
    text: str
    paragraphs: list[Paragraph] = field(default_factory=list)
    slots: list[Slot] = field(default_factory=list)
    table: TableMeta | None = None
    image: ImageMeta | None = None
    chart: dict[str, Any] | None = None
    meta: dict[str, Any] = field(default_factory=dict)

    @property
    def slot_names(self) -> list[str]:
        return [s.name for s in self.slots]

    def to_dict(self) -> dict[str, Any]:
        d = asdict(self)
        d["rect"] = self.rect.to_dict()
        d["slot_names"] = self.slot_names
        return d


@dataclass
class Column:
    index: int
    name: str
    dtype: str  # text | number | date | bool | image
    samples: list[str] = field(default_factory=list)

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


@dataclass
class Binding:
    """一个 Field 与一个数据来源的绑定。"""

    field_id: str
    kind: str  # text | table | image | constant | page_var
    source: dict[str, Any] = field(default_factory=dict)
    options: dict[str, Any] = field(default_factory=dict)

    @classmethod
    def from_dict(cls, d: dict[str, Any]) -> Binding:
        return cls(
            field_id=d["field_id"],
            kind=d["kind"],
            source=dict(d.get("source") or {}),
            options=dict(d.get("options") or {}),
        )


@dataclass
class Profile:
    """可保存、可复用的完整映射方案。"""

    version: int = 1
    template: str = ""
    slide_index: int = 0
    excel: str = ""
    sheet: str = ""
    header_row: int = 1
    bindings: list[Binding] = field(default_factory=list)
    pagination: dict[str, Any] = field(default_factory=lambda: {"mode": "one_per_row"})
    output: dict[str, Any] = field(default_factory=dict)
    # 可选编排：{"cover": {...}, "section": {...}, "end": {...}}
    # 缺省为空表示单页模式（只生成正文页，与旧版行为一致）
    layout: dict[str, Any] = field(default_factory=dict)

    def to_dict(self) -> dict[str, Any]:
        data = {
            "version": self.version,
            "template": self.template,
            "slide_index": self.slide_index,
            "excel": self.excel,
            "sheet": self.sheet,
            "header_row": self.header_row,
            "bindings": [asdict(b) for b in self.bindings],
            "pagination": self.pagination,
            "output": self.output,
        }
        if self.layout:
            data["layout"] = self.layout
        return data

    @classmethod
    def from_dict(cls, d: dict[str, Any]) -> Profile:
        return cls(
            version=int(d.get("version", 1)),
            template=d.get("template", ""),
            slide_index=int(d.get("slide_index", 0)),
            excel=d.get("excel", ""),
            sheet=d.get("sheet", ""),
            header_row=int(d.get("header_row", 1)),
            bindings=[Binding.from_dict(b) for b in d.get("bindings", [])],
            pagination=dict(d.get("pagination") or {"mode": "one_per_row"}),
            output=dict(d.get("output") or {}),
            layout=dict(d.get("layout") or {}),
        )
