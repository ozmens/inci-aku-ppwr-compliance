"""Phase O6 — strict shared grid front-end for HOME / NAVIGATION / SEARCH.

All shapes placed via grid helpers; ivory canvas = CELL FILL only (A1:P40).
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any

MSO_RECT = 1
MSO_ROUNDED = 5
MSO_TRUE = -1
MSO_FALSE = 0
XL_FREE_FLOATING = 3
MSO_GRADIENT_HORIZ = 1

NAVY = "0B2341"
MIDNIGHT = "07182B"
GOLD = "C9A24A"
DARK_GOLD = "A9842E"
IVORY = "F5F1E8"
CARD = "FCFBF8"
PALE = "E8EEF4"
OK_BG = "E5F0E7"
OK_FG = "325D3E"
INK = "1E2C3A"
MUTED = "6A7785"
WHITE = "FFFFFF"
FONT = "Tahoma"

# Grid: columns A..P = 1..16, visual groups of ~5 cols for 3-up cards
CANVAS_COLS = 16  # A:P
CANVAS_ROWS = 40
COL_WIDTH = 8.6
GAP_X = 10.0
GAP_Y = 12.0
MARGIN = 14.0


def _rgb(h: str) -> int:
    h = h.lstrip("#")
    return int(h[4:6] + h[2:4] + h[0:2], 16)


@dataclass
class Bounds:
    left: float
    top: float
    right: float
    bottom: float

    @property
    def width(self) -> float:
        return self.right - self.left

    @property
    def height(self) -> float:
        return self.bottom - self.top


class GridCanvas:
    """Shared layout engine for Class A sheets."""

    def __init__(self, excel, wb, logo_path: Path) -> None:
        self.excel = excel
        self.wb = wb
        self.logo_path = logo_path
        self.shapes_created = 0
        self.hyperlinks_added = 0
        self.locked_shapes = 0
        self._bounds: Bounds | None = None
        self._placed: list[tuple[str, float, float, float, float]] = []  # name,L,T,W,H

    def _focus(self, name: str):
        ws = self.wb.Worksheets(name)
        ws.Select()
        ws.Activate()
        self.excel.CutCopyMode = False
        return ws

    def _clear_shapes(self, ws) -> None:
        for i in range(int(ws.Shapes.Count), 0, -1):
            try:
                ws.Shapes(i).Delete()
            except Exception:
                pass

    def prepare_canvas(self, ws, *, zoom: int = 92, rows: int = CANVAS_ROWS) -> Bounds:
        """Fill A1:P{rows} with ivory CELL fill; set widths; hide chrome."""
        self._placed = []
        win = self.excel.ActiveWindow
        win.DisplayGridlines = False
        try:
            win.DisplayHeadings = False
        except Exception:
            pass
        win.ScrollRow = 1
        win.ScrollColumn = 1
        try:
            win.Zoom = zoom
        except Exception:
            pass

        ws.Cells.Clear()
        try:
            ws.Cells.UnMerge()
        except Exception:
            pass

        for c in range(1, CANVAS_COLS + 1):
            ws.Columns(c).ColumnWidth = COL_WIDTH
        # Compact row heights for predictable layout
        for r in range(1, rows + 1):
            ws.Rows(r).RowHeight = 15

        rng = ws.Range(f"A1:P{rows}")
        rng.Interior.Color = _rgb(IVORY)
        rng.Font.Name = FONT
        rng.Borders.LineStyle = 0

        left = float(ws.Range("A1").Left)
        top = float(ws.Range("A1").Top)
        right = float(ws.Range("P1").Left) + float(ws.Range("P1").Width)
        bottom = float(ws.Range(f"A{rows}").Top) + float(ws.Range(f"A{rows}").Height)
        self._bounds = Bounds(left + 0, top + 0, right, bottom)
        return self._bounds

    def cell_box(self, ws, c1: int, r1: int, c2: int, r2: int) -> tuple[float, float, float, float]:
        """Return Left, Top, Width, Height for cell range (1-based cols/rows)."""
        a = ws.Cells(r1, c1)
        b = ws.Cells(r2, c2)
        left = float(a.Left)
        top = float(a.Top)
        right = float(b.Left) + float(b.Width)
        bottom = float(b.Top) + float(b.Height)
        # inset margins within cells
        return left + 2, top + 2, max(right - left - 4, 8), max(bottom - top - 4, 8)

    def _free(self, shp) -> None:
        try:
            shp.Placement = XL_FREE_FLOATING
        except Exception:
            pass

    def _shadow(self, shp) -> None:
        try:
            sh = shp.Shadow
            sh.Visible = MSO_TRUE
            sh.Style = 1
            sh.OffsetX = 0.6
            sh.OffsetY = 1.2
            sh.Transparency = 0.85
            try:
                sh.Blur = 4
            except Exception:
                pass
        except Exception:
            pass

    def _no_shadow(self, shp) -> None:
        try:
            shp.Shadow.Visible = MSO_FALSE
        except Exception:
            pass

    def _track(self, shp) -> None:
        self._placed.append(
            (str(shp.Name), float(shp.Left), float(shp.Top), float(shp.Width), float(shp.Height))
        )

    def place_round(
        self,
        ws,
        name: str,
        left: float,
        top: float,
        w: float,
        h: float,
        *,
        fill: str = CARD,
        line: str | None = PALE,
        shadow: bool = True,
        text: str = "",
        size: float = 9,
        bold: bool = False,
        color: str = INK,
        align: str = "left",
    ):
        shp = ws.Shapes.AddShape(MSO_ROUNDED, left, top, w, h)
        shp.Name = name
        shp.Fill.Solid()
        shp.Fill.ForeColor.RGB = _rgb(fill)
        if line:
            shp.Line.Visible = MSO_TRUE
            shp.Line.ForeColor.RGB = _rgb(line)
            shp.Line.Weight = 0.75
        else:
            shp.Line.Visible = MSO_FALSE
        self._free(shp)
        if shadow:
            self._shadow(shp)
        else:
            self._no_shadow(shp)
        if text:
            self._set_text(shp, text, size=size, bold=bold, color=color, align=align)
        self.shapes_created += 1
        self._track(shp)
        return shp

    def place_rect(
        self,
        ws,
        name: str,
        left: float,
        top: float,
        w: float,
        h: float,
        *,
        fill: str,
        line: str | None = None,
        text: str = "",
        size: float = 9,
        bold: bool = False,
        color: str = INK,
        align: str = "left",
        gradient: bool = False,
    ):
        shp = ws.Shapes.AddShape(MSO_RECT, left, top, w, h)
        shp.Name = name
        if gradient:
            try:
                shp.Fill.TwoColorGradient(MSO_GRADIENT_HORIZ, 1)
                shp.Fill.ForeColor.RGB = _rgb(MIDNIGHT)
                shp.Fill.BackColor.RGB = _rgb(NAVY)
            except Exception:
                shp.Fill.Solid()
                shp.Fill.ForeColor.RGB = _rgb(fill)
        else:
            shp.Fill.Solid()
            shp.Fill.ForeColor.RGB = _rgb(fill)
        if line:
            shp.Line.Visible = MSO_TRUE
            shp.Line.ForeColor.RGB = _rgb(line)
        else:
            shp.Line.Visible = MSO_FALSE
        self._free(shp)
        self._no_shadow(shp)
        if text:
            self._set_text(shp, text, size=size, bold=bold, color=color, align=align)
        self.shapes_created += 1
        self._track(shp)
        return shp

    def _set_text(self, shp, text, size=9, bold=False, color=INK, align="left") -> None:
        tf = shp.TextFrame
        tf.Characters().Text = text
        f = tf.Characters().Font
        f.Name = FONT
        f.Size = size
        f.Bold = bold
        f.Color = _rgb(color)
        amap = {"left": 1, "center": 2, "right": 3}
        try:
            tf.HorizontalAlignment = amap.get(align, 1)
            tf.VerticalAlignment = 2
            tf.MarginLeft = 10
            tf.MarginRight = 10
            tf.MarginTop = 6
            tf.MarginBottom = 6
            tf.WordWrap = MSO_TRUE
        except Exception:
            pass

    def link(self, ws, shp, sheet: str) -> None:
        try:
            ws.Hyperlinks.Add(Anchor=shp, Address="", SubAddress=f"'{sheet}'!A1")
            self.hyperlinks_added += 1
        except Exception:
            pass

    def lock_all(self, ws) -> int:
        n = 0
        for i in range(1, int(ws.Shapes.Count) + 1):
            try:
                ws.Shapes(i).Locked = True
                n += 1
            except Exception:
                pass
        self.locked_shapes += n
        return n

    def protect(self, name: str, unlock: list[str] | None = None) -> dict[str, Any]:
        ws = self._focus(name)
        locked = self.lock_all(ws)
        try:
            ws.Cells.Locked = False
        except Exception:
            pass
        if unlock:
            for a in unlock:
                try:
                    ws.Range(a).Locked = False
                except Exception:
                    pass
        try:
            ws.Unprotect(Password="")
        except Exception:
            pass
        try:
            ws.Protect(
                Password="",
                DrawingObjects=True,
                Contents=True,
                Scenarios=True,
                UserInterfaceOnly=True,
            )
            return {"sheet": name, "ok": True, "locked_shapes": locked}
        except Exception as exc:
            return {"sheet": name, "ok": False, "error": str(exc)}

    def place_logo(self, ws, name: str) -> None:
        if not self.logo_path.exists() or not self._bounds:
            return
        # Right-aligned inside canvas with padding
        w, h = 118.0, 52.0
        left = self._bounds.right - MARGIN - w
        top = float(ws.Range("A3").Top) + 10
        try:
            pic = ws.Shapes.AddPicture(
                str(self.logo_path.resolve()), False, True, left, top, w, h
            )
            pic.Name = name
            self._free(pic)
            self.shapes_created += 1
            self._track(pic)
        except Exception:
            pass

    def top_nav(self, ws, active: str) -> None:
        """Shared top nav — pills in rows 1–2."""
        L, T, W, H = self.cell_box(ws, 1, 1, 16, 2)
        self.place_rect(ws, "NavBar", L, T, W, H, fill=NAVY, gradient=True)
        items = [
            (1, 3, "HOME", "00_HOME"),
            (4, 6, "NAVIGATION", "NAVIGATION"),
            (7, 9, "SEARCH", "SEARCH"),
            (10, 12, "DOCUMENTS", "DOCUMENT_CENTER"),
        ]
        for c1, c2, label, target in items:
            l, t, w, h = self.cell_box(ws, c1, 1, c2, 2)
            # inset inside bar
            t += 4
            h = max(h - 8, 16)
            selected = label == active
            fill = GOLD if selected else WHITE
            fg = MIDNIGHT if selected else NAVY
            p = self.place_round(
                ws,
                f"NavPill_{target}",
                l + 4,
                t,
                w - 8,
                h,
                fill=fill,
                line=None,
                shadow=False,
                text=label,
                size=8,
                bold=True,
                color=fg,
                align="center",
            )
            self.link(ws, p, target)
        # status pills right
        for c1, c2, label, fill, fg in (
            (13, 14, "EXCEL VALIDATED", OK_BG, OK_FG),
            (15, 15, "REV.00", GOLD, MIDNIGHT),
            (16, 16, "CTRL", PALE, NAVY),
        ):
            l, t, w, h = self.cell_box(ws, c1, 1, c2, 2)
            self.place_round(
                ws,
                f"Status_{label[:6]}",
                l + 2,
                t + 4,
                max(w - 4, 40),
                max(h - 8, 16),
                fill=fill,
                shadow=False,
                text=label if label != "CTRL" else "CONTROLLED",
                size=6.5,
                bold=True,
                color=fg,
                align="center",
            )

    def hero(self, ws, title: str, subtitle: str, logo_name: str) -> None:
        L, T, W, H = self.cell_box(ws, 1, 3, 16, 7)
        hero = self.place_rect(
            ws,
            "Hero",
            L,
            T,
            W,
            H - 4,
            fill=NAVY,
            gradient=True,
            text=f"İNCİ AKÜ PPWR\n{title}\n{subtitle}",
            size=11,
            bold=True,
            color=WHITE,
            align="left",
        )
        # Refine text sizes via Characters is hard for multi-style; keep readable block
        self._set_text(
            hero,
            f"İNCİ AKÜ PPWR\r{title}\r{subtitle}",
            size=12,
            bold=True,
            color=WHITE,
            align="left",
        )
        # gold hairline under hero
        gl, gt, gw, gh = self.cell_box(ws, 1, 7, 16, 7)
        self.place_rect(
            ws,
            "HeroGold",
            gl,
            gt + gh - 5,
            gw,
            3.5,
            fill=GOLD,
        )
        self.place_logo(ws, logo_name)

    def clear_shapes_over_cells(self, ws, addr: str) -> int:
        cell = ws.Range(addr)
        cx = float(cell.Left) + float(cell.Width) / 2
        cy = float(cell.Top) + float(cell.Height) / 2
        removed = 0
        for i in range(int(ws.Shapes.Count), 0, -1):
            shp = ws.Shapes(i)
            name = str(shp.Name)
            if name.startswith(("NavBar", "NavPill", "Status_", "Hero", "Inci", "SQ_", "SearchBtn")):
                continue
            try:
                sl, st = float(shp.Left), float(shp.Top)
                sw, sh = float(shp.Width), float(shp.Height)
            except Exception:
                continue
            if sl <= cx <= sl + sw and st <= cy <= st + sh:
                try:
                    shp.Delete()
                    removed += 1
                except Exception:
                    pass
        return removed

    # ─── HOME ───────────────────────────────────────────────
    def design_home(self) -> dict[str, Any]:
        ws = self._focus("00_HOME")
        self._clear_shapes(ws)
        self.prepare_canvas(ws, zoom=92, rows=40)
        self.top_nav(ws, "HOME")
        self.hero(
            ws,
            "Packaging Information Management System",
            "Controlled Packaging Data  •  Compliance  •  Document Registry",
            "InciAkuLogo",
        )

        # 4 KPI cards — rows 9–14, equal quarters of A:P
        kpis = [
            (1, 4, "247\nPACKAGING\nCONFIGURATIONS"),
            (5, 8, "2,046\nPRODUCTS"),
            (9, 12, "988\nCONTROLLED\nDOCUMENTS"),
            (13, 16, "0\nBLOCKING\nERRORS"),
        ]
        for i, (c1, c2, text) in enumerate(kpis):
            l, t, w, h = self.cell_box(ws, c1, 9, c2, 14)
            self.place_round(
                ws,
                f"KPI_{i}",
                l,
                t,
                w,
                h,
                text=text,
                size=14,
                bold=True,
                color=NAVY,
                align="left",
            )

        # Secondary strip rows 16–18 — ONE shape
        l, t, w, h = self.cell_box(ws, 1, 16, 16, 18)
        self.place_round(
            ws,
            "SecStrip",
            l,
            t,
            w,
            h,
            text="112 Components   ·   1,690 BOM Lines   ·   247 Technical Files   ·   247 EU DoCs   ·   247 Labels   ·   247 Statements",
            size=9,
            bold=True,
            color=NAVY,
            align="center",
        )

        # Three panels rows 20–27
        panels = [
            (
                1,
                5,
                "PanelHealth",
                "SYSTEM HEALTH\n\n"
                "Master Data              READY\n"
                "Golden Register        247/247\n"
                "Document Registry      LINKED\n"
                "Excel Validation         PASS\n"
                "Blocking Errors              0",
            ),
            (
                6,
                10,
                "PanelPort",
                "PACKAGING PORTFOLIO\n\n"
                "Starter ████████████████  240\n"
                "Industrial ██                      3\n"
                "Container ███                     4",
            ),
            (
                11,
                16,
                "PanelComp",
                "DOCUMENT COMPLETENESS\n\n"
                "Technical Files          100%\n"
                "EU DoCs                  100%\n"
                "Labels                   100%\n"
                "Statements               100%",
            ),
        ]
        for c1, c2, name, text in panels:
            l, t, w, h = self.cell_box(ws, c1, 20, c2, 27)
            self.place_round(
                ws, name, l, t, w, h, text=text, size=8.5, bold=False, color=INK, align="left"
            )

        # Quick actions 5×2 rows 29–39
        actions = [
            ("Document Center", "988 linked documents", "DOCUMENT_CENTER"),
            ("Packaging Configurations", "247 configurations", "PACKAGING_CONFIGURATIONS"),
            ("Technical Files", "247 Rev.00 files", "TECHNICAL_FILES"),
            ("EU Declarations", "247 EU DoCs", "DECLARATIONS_OF_CONFORMITY"),
            ("Labels", "247 labels", "LABELS"),
            ("Shipment Statements", "247 statements", "SHIPMENT_STATEMENTS"),
            ("Product Master", "2,046 products", "PRODUCT_MASTER"),
            ("Component Master", "112 components", "COMPONENT_MASTER"),
            ("Global Search", "Set / config inquiry", "SEARCH"),
            ("Navigation", "Module launchpad", "NAVIGATION"),
        ]
        # row1: cols 1-3,4-6,7-9,10-12,13-16
        col_spans = [(1, 3), (4, 6), (7, 9), (10, 12), (13, 16)]
        for i, (title, desc, target) in enumerate(actions):
            r = 0 if i < 5 else 1
            c1, c2 = col_spans[i % 5]
            r1 = 29 + r * 5
            r2 = r1 + 4
            l, t, w, h = self.cell_box(ws, c1, r1, c2, r2)
            card = self.place_round(
                ws,
                f"Act_{target}",
                l,
                t,
                w,
                h,
                text=f"{title}\n{desc}\nOpen →",
                size=8.5,
                bold=True,
                color=NAVY,
                align="left",
            )
            self.link(ws, card, target)

        self.lock_all(ws)
        ws.Range("A1").Select()
        return {"sheet": "00_HOME", "shapes": int(ws.Shapes.Count)}

    # ─── NAVIGATION ─────────────────────────────────────────
    def design_navigation(self) -> dict[str, Any]:
        ws = self._focus("NAVIGATION")
        self._clear_shapes(ws)
        self.prepare_canvas(ws, zoom=92, rows=40)
        self.top_nav(ws, "NAVIGATION")
        self.hero(
            ws,
            "NAVIGATION",
            "Access all controlled PPWR management modules",
            "InciAkuLogo_Nav",
        )

        # Section headers as cell text (no shapes) to save count
        def section_label(row: int, text: str) -> None:
            ws.Cells(row, 1).Value = text
            ws.Cells(row, 1).Font.Name = FONT
            ws.Cells(row, 1).Font.Size = 9
            ws.Cells(row, 1).Font.Bold = True
            ws.Cells(row, 1).Font.Color = _rgb(NAVY)

        def card_row(r1: int, r2: int, cards: list[tuple[str, str, str, str]]) -> None:
            # 3 equal columns: A:E, F:J, K:O  (P = margin)
            spans = [(1, 5), (6, 10), (11, 15)]
            for i, (title, desc, target, meta) in enumerate(cards):
                if i >= 3:
                    break
                c1, c2 = spans[i]
                l, t, w, h = self.cell_box(ws, c1, r1, c2, r2)
                shp = self.place_round(
                    ws,
                    f"NavCard_{target}",
                    l,
                    t,
                    w,
                    h,
                    text=f"{title}\n{desc}\n{meta}  ·  Open →",
                    size=9,
                    bold=True,
                    color=NAVY,
                    align="left",
                )
                self.link(ws, shp, target)

        section_label(9, "WORKSPACE")
        card_row(
            10,
            13,
            [
                ("HOME", "Executive cockpit", "00_HOME", "Cockpit"),
                ("Global Search", "Packaging set inquiry", "SEARCH", "Lookup"),
            ],
        )

        section_label(15, "MASTER DATA")
        card_row(
            16,
            19,
            [
                ("Packaging Configurations", "247 controlled configurations", "PACKAGING_CONFIGURATIONS", "247"),
                ("Product Master", "2,046 products", "PRODUCT_MASTER", "2,046"),
                ("Component Master", "112 packaging components", "COMPONENT_MASTER", "112"),
            ],
        )

        section_label(21, "DOCUMENT CONTROL")
        card_row(
            22,
            25,
            [
                ("Document Center", "988 linked documents", "DOCUMENT_CENTER", "988"),
                ("Technical Files", "247 controlled Rev.00 files", "TECHNICAL_FILES", "247"),
                ("EU Declarations", "247 EU declarations", "DECLARATIONS_OF_CONFORMITY", "247"),
            ],
        )
        card_row(
            26,
            29,
            [
                ("Labels", "247 labels", "LABELS", "247"),
                ("Shipment Statements", "247 statements", "SHIPMENT_STATEMENTS", "247"),
            ],
        )

        section_label(31, "OPERATIONS & SYSTEM")
        card_row(
            32,
            35,
            [
                ("Shipments", "Transactional register", "SHIPMENTS", "Empty"),
                ("Document Engine Map", "Read-only mapping", "DOC_ENGINE_MAP", "Map"),
            ],
        )

        self.lock_all(ws)
        ws.Range("A1").Select()
        return {"sheet": "NAVIGATION", "shapes": int(ws.Shapes.Count)}

    # ─── SEARCH ─────────────────────────────────────────────
    def design_search(self) -> dict[str, Any]:
        ws = self._focus("SEARCH")
        self._clear_shapes(ws)
        self.prepare_canvas(ws, zoom=92, rows=40)
        self.top_nav(ws, "SEARCH")
        self.hero(
            ws,
            "GLOBAL SEARCH",
            "Find a packaging set, configuration, source BOM or document.",
            "InciAkuLogo_Search",
        )

        # Search input card area rows 9–14 — CELL chrome only
        for r in range(9, 15):
            for c in range(1, 17):
                ws.Cells(r, c).Interior.Color = _rgb(CARD)
        # gold left accent cells
        for r in range(9, 15):
            ws.Cells(r, 1).Interior.Color = _rgb(GOLD)

        ws.Range("B10").Value = "SEARCH TERM"
        ws.Range("B10").Font.Name = FONT
        ws.Range("B10").Font.Size = 9
        ws.Range("B10").Font.Bold = True
        ws.Range("B10").Font.Color = _rgb(MUTED)

        # REAL input B11:H12 — no covering shape
        try:
            ws.Range("B11:H12").UnMerge()
        except Exception:
            pass
        ws.Range("B11:H12").Merge()
        inp = ws.Range("B11")
        inp.Value = ""
        inp.Interior.Color = _rgb(WHITE)
        inp.Font.Name = FONT
        inp.Font.Size = 12
        inp.Font.Bold = True
        inp.Font.Color = _rgb(NAVY)
        inp.HorizontalAlignment = -4131
        inp.VerticalAlignment = -4108
        for edge in (7, 8, 9, 10):
            try:
                b = ws.Range("B11:H12").Borders(edge)
                b.LineStyle = 1
                b.Weight = 3
                b.Color = _rgb(GOLD)
            except Exception:
                pass
        ws.Range("B11:H12").Locked = False

        # SEARCH button J11:L12
        l, t, w, h = self.cell_box(ws, 10, 11, 12, 12)
        self.place_round(
            ws,
            "SearchBtn",
            l,
            t,
            w,
            h,
            fill=NAVY,
            line=None,
            shadow=True,
            text="SEARCH",
            size=11,
            bold=True,
            color=WHITE,
            align="center",
        )

        ws.Range("B13").Value = (
            "Results update automatically after entering a valid ID.   ·   "
            "Examples: ST-051-STD-01  •  CNT-20-STD-01  •  IND-24V-01"
        )
        ws.Range("B13").Font.Name = FONT
        ws.Range("B13").Font.Size = 8
        ws.Range("B13").Font.Color = _rgb(MUTED)

        # Results card rows 16–27 — cell surface
        for r in range(16, 28):
            for c in range(1, 17):
                ws.Cells(r, c).Interior.Color = _rgb(CARD)
        for r in range(16, 28):
            ws.Cells(r, 1).Interior.Color = _rgb(GOLD)

        ws.Range("B16").Value = "RESULTS"
        ws.Range("B16").Font.Name = FONT
        ws.Range("B16").Font.Size = 9
        ws.Range("B16").Font.Bold = True
        ws.Range("B16").Font.Color = _rgb(NAVY)

        # Empty state — use B11 as search key (same as O5 logic but new cell)
        # Keep formulas referencing $B$11
        ws.Range("B17").Formula = (
            '=IF($B$11="","Enter a Packaging Set or Configuration ID'
            ' to view controlled packaging data.","")'
        )
        ws.Range("B17").Font.Name = FONT
        ws.Range("B17").Font.Size = 9
        ws.Range("B17").Font.Italic = True
        ws.Range("B17").Font.Color = _rgb(MUTED)

        fields = [
            (18, "Packaging Set Code", 'IF($B$11="","",$B$11)'),
            (19, "Configuration ID", "PACKAGING_CONFIGURATIONS!B:B"),
            (20, "Family", "PACKAGING_CONFIGURATIONS!D:D"),
            (21, "Source Configuration ID", "PACKAGING_CONFIGURATIONS!C:C"),
            (22, "Packaging Tare", "PACKAGING_CONFIGURATIONS!H:H"),
            (23, "Technical File ID", "PACKAGING_CONFIGURATIONS!K:K"),
            (24, "EU DoC ID", "PACKAGING_CONFIGURATIONS!L:L"),
            (25, "Label ID", "PACKAGING_CONFIGURATIONS!M:M"),
            (26, "Statement ID", "PACKAGING_CONFIGURATIONS!N:N"),
        ]
        for r, label, colref in fields:
            ws.Range(f"B{r}").Value = label
            ws.Range(f"B{r}").Font.Name = FONT
            ws.Range(f"B{r}").Font.Size = 8.5
            ws.Range(f"B{r}").Font.Color = _rgb(MUTED)
            if colref.startswith("IF("):
                ws.Range(f"E{r}").Formula = f"={colref}"
            else:
                ws.Range(f"E{r}").Formula = (
                    f'=IF($B$11="","",IFERROR(XLOOKUP($B$11,PACKAGING_CONFIGURATIONS!A:A,{colref}),'
                    f'"Not found"))'
                )
            ws.Range(f"E{r}").Font.Name = FONT
            ws.Range(f"E{r}").Font.Size = 9.5
            ws.Range(f"E{r}").Font.Color = _rgb(INK)
            ws.Range(f"E{r}").Interior.Color = _rgb(PALE)

        # Right status panel K18:O25
        ws.Range("K18").Value = "DOCUMENT PACK STATUS"
        ws.Range("K18").Font.Name = FONT
        ws.Range("K18").Font.Size = 8
        ws.Range("K18").Font.Bold = True
        ws.Range("K18").Font.Color = _rgb(NAVY)
        ws.Range("K20").Value = "988 / 988 LINKED"
        ws.Range("K20").Font.Name = FONT
        ws.Range("K20").Font.Size = 12
        ws.Range("K20").Font.Bold = True
        ws.Range("K20").Font.Color = _rgb(OK_FG)
        ws.Range("K20").Interior.Color = _rgb(OK_BG)
        ws.Range("K22").Value = "Rev.00"
        ws.Range("K22").Font.Name = FONT
        ws.Range("K22").Font.Size = 10
        ws.Range("K22").Font.Color = _rgb(MUTED)

        # Quick links rows 29–33 — four equal cards A:D E:H I:L M:P
        links = [
            (1, 4, "Document Center", "DOCUMENT_CENTER"),
            (5, 8, "Packaging Configurations", "PACKAGING_CONFIGURATIONS"),
            (9, 12, "Technical Files", "TECHNICAL_FILES"),
            (13, 16, "Product Master", "PRODUCT_MASTER"),
        ]
        for c1, c2, title, target in links:
            l, t, w, h = self.cell_box(ws, c1, 29, c2, 33)
            shp = self.place_round(
                ws,
                f"SQ_{target}",
                l,
                t,
                w,
                h,
                text=f"{title}\nOpen →",
                size=9,
                bold=True,
                color=NAVY,
                align="left",
            )
            self.link(ws, shp, target)

        self.lock_all(ws)
        ws.Range("B11:H12").Locked = False
        self.clear_shapes_over_cells(ws, "B11")
        # ensure SearchBtn still exists and is to the right
        try:
            btn = ws.Shapes("SearchBtn")
            # if somehow over input, nudge right
            if float(btn.Left) < float(ws.Range("B11").Left) + float(ws.Range("B11:H12").Width):
                l, t, w, h = self.cell_box(ws, 10, 11, 12, 12)
                btn.Left = l
                btn.Top = t
        except Exception:
            pass
        ws.Range("B11").Select()
        return {"sheet": "SEARCH", "shapes": int(ws.Shapes.Count)}
