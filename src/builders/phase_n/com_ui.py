"""Phase N — native Excel COM executive UI designer.

Critical: always ungroup sheets (SelectedSheets=1) before AddShape / Add.
"""

from __future__ import annotations

from pathlib import Path
from typing import Any

import pythoncom
import win32com.client

# Excel / MSO constants
MSO_SHAPE_RECT = 1
MSO_SHAPE_ROUNDED_RECT = 5
MSO_TRUE = -1
MSO_FALSE = 0
MSO_LINKED = 0
MSO_CFALSE = 0
XL_SCREEN = 1
XL_BITMAP = 2
MSO_GRADIENT_HORIZ = 1
MSO_GRADIENT_VERT = 2

# Palette
NAVY = "0E2A47"
NAVY_DEEP = "0A1F35"
STEEL = "315E87"
GOLD = "C8A24A"
GOLD_SOFT = "D2B15B"
IVORY = "F7F5F0"
CARD = "FFFFFF"
STONE = "F3F1EB"
TRACK = "E4E8EE"
SUCCESS_BG = "DDEBDD"
SUCCESS_FG = "2F5D3A"
INK = "1C2430"
MUTED = "5C6B7A"
WHITE = "FFFFFF"

FONT = "Tahoma"

UI_SHEETS = [
    "00_HOME",
    "NAVIGATION",
    "SEARCH",
    "PACKAGING_CONFIGURATIONS",
    "PRODUCT_MASTER",
    "COMPONENT_MASTER",
    "DOCUMENT_CENTER",
    "TECHNICAL_FILES",
    "DECLARATIONS_OF_CONFORMITY",
    "LABELS",
    "SHIPMENT_STATEMENTS",
    "SHIPMENTS",
    "DOC_ENGINE_MAP",
]

REGISTER_META = {
    "PACKAGING_CONFIGURATIONS": (
        "PACKAGING CONFIGURATIONS",
        "Final packaging set register — controlled configurations",
        [("CONFIGS", "247"), ("STARTER", "240"), ("INDUSTRIAL", "3"), ("CONTAINER", "4")],
    ),
    "PRODUCT_MASTER": (
        "PRODUCT MASTER",
        "Products linked to packaging configurations",
        [("PRODUCTS", "2,046"), ("CONFIGS", "247")],
    ),
    "COMPONENT_MASTER": (
        "COMPONENT MASTER",
        "Packaging component catalogue",
        [("COMPONENTS", "112")],
    ),
    "DOCUMENT_CENTER": (
        "DOCUMENT CENTER",
        "Per-configuration document pack — open controlled Word files",
        [("CONFIGS", "247"), ("DOCUMENTS", "988"), ("LINKED", "988"), ("REV", "00")],
    ),
    "TECHNICAL_FILES": (
        "TECHNICAL FILES",
        "PPWR technical file index",
        [("FILES", "247"), ("LINKED", "247/247")],
    ),
    "DECLARATIONS_OF_CONFORMITY": (
        "DECLARATIONS OF CONFORMITY",
        "EU Declaration of Conformity index",
        [("EU DOCS", "247"), ("LINKED", "247/247")],
    ),
    "LABELS": (
        "LABELS",
        "Packaging identification label index",
        [("LABELS", "247"), ("LINKED", "247/247")],
    ),
    "SHIPMENT_STATEMENTS": (
        "SHIPMENT STATEMENTS",
        "Shipment statement index",
        [("STATEMENTS", "247"), ("LINKED", "247/247")],
    ),
    "SHIPMENTS": (
        "SHIPMENTS",
        "Transactional shipment register",
        [("MODULE", "OPS")],
    ),
    "DOC_ENGINE_MAP": (
        "DOCUMENT ENGINE MAP",
        "Read-only mapping — Python remains document authority",
        [("MODE", "READ-ONLY")],
    ),
}


def _rgb(hex6: str) -> int:
    h = hex6.lstrip("#")
    return int(h[4:6] + h[2:4] + h[0:2], 16)


class ExcelComUI:
    def __init__(self, workbook_path: Path, logo_path: Path) -> None:
        self.workbook_path = workbook_path
        self.logo_path = logo_path
        self.excel = None
        self.wb = None
        self.stats: dict[str, Any] = {
            "shapes_created": 0,
            "home_buttons": 0,
            "sheets_designed": [],
        }

    # ── lifecycle ──────────────────────────────────────────────────────
    def run(self) -> dict[str, Any]:
        pythoncom.CoInitialize()
        try:
            self.excel = win32com.client.DispatchEx("Excel.Application")
            self.excel.Visible = False
            self.excel.DisplayAlerts = False
            self.excel.AskToUpdateLinks = False
            self.excel.ScreenUpdating = False
            self.excel.EnableEvents = False
            self.wb = self.excel.Workbooks.Open(
                str(self.workbook_path.resolve()),
                UpdateLinks=0,
                ReadOnly=False,
                CorruptLoad=0,
            )
            self._ungroup()

            self.design_home()
            self.design_navigation()
            self.design_search()
            for name in REGISTER_META:
                if self._has(name):
                    self.design_register(name)

            # Ensure home shape buttons on all UI sheets
            for name in UI_SHEETS:
                if self._has(name):
                    self._ensure_home_button(name)

            self._reorder_ui_first()
            self.excel.ScreenUpdating = True
            self.wb.Save()
            path = str(self.workbook_path.resolve())
            self.wb.Close(True)
            self.wb = None
            self.excel.Quit()
            self.excel = None
            # Native reopen validation
            reopen = self._reopen_check(path)
            self.stats["native_reopen"] = reopen
            return self.stats
        finally:
            try:
                if self.wb is not None:
                    self.wb.Close(False)
            except Exception:
                pass
            try:
                if self.excel is not None:
                    self.excel.Quit()
            except Exception:
                pass
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass

    def _reopen_check(self, path: str) -> dict:
        pythoncom.CoInitialize()
        excel = None
        out = {"ok": False, "error": None, "sheets": None}
        try:
            excel = win32com.client.DispatchEx("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            excel.AskToUpdateLinks = False
            wb = excel.Workbooks.Open(path, UpdateLinks=0, ReadOnly=True, CorruptLoad=0)
            out["ok"] = True
            out["sheets"] = int(wb.Worksheets.Count)
            wb.Close(False)
        except Exception as exc:  # noqa: BLE001
            out["error"] = str(exc)
        finally:
            if excel is not None:
                try:
                    excel.Quit()
                except Exception:
                    pass
            pythoncom.CoUninitialize()
        return out

    def _has(self, name: str) -> bool:
        try:
            self.wb.Worksheets(name)
            return True
        except Exception:
            return False

    def _ungroup(self) -> None:
        # Critical: grouped sheets block AddShape ("won't work on multiple selections")
        ws = self.wb.Worksheets(1)
        ws.Select()
        ws.Activate()
        ws.Range("A1").Select()
        self.excel.CutCopyMode = False

    def _focus(self, name: str):
        ws = self.wb.Worksheets(name)
        ws.Select()
        ws.Activate()
        ws.Range("A1").Select()
        self.excel.CutCopyMode = False
        return ws

    def _delete_all_shapes(self, ws) -> None:
        # Delete from end to start
        n = int(ws.Shapes.Count)
        for i in range(n, 0, -1):
            try:
                ws.Shapes(i).Delete()
            except Exception:
                pass

    def _canvas(self, ws, hide_headings: bool = True) -> None:
        ws.Activate()
        win = self.excel.ActiveWindow
        win.DisplayGridlines = False
        if hide_headings:
            try:
                win.DisplayHeadings = False
            except Exception:
                pass
        win.ScrollRow = 1
        win.ScrollColumn = 1
        try:
            win.Zoom = 95
        except Exception:
            pass
        # Ivory canvas fill for visible area
        rng = ws.Range("A1:N50")
        rng.Interior.Color = _rgb(IVORY)
        rng.Borders.LineStyle = 0
        rng.Font.Name = FONT
        rng.Font.Color = _rgb(INK)
        # Clear prior values in chrome area carefully done by callers

    def _soft_shadow(self, shp) -> None:
        try:
            sh = shp.Shadow
            sh.Visible = MSO_TRUE
            sh.Style = 1  # outer
            sh.OffsetX = 1.5
            sh.OffsetY = 2.0
            sh.Transparency = 0.72
            try:
                sh.Blur = 5
            except Exception:
                pass
            sh.ForeColor.RGB = _rgb("000000")
        except Exception:
            pass

    def _round_card(
        self,
        ws,
        name: str,
        left: float,
        top: float,
        width: float,
        height: float,
        fill: str = CARD,
        line: str | None = None,
    ):
        shp = ws.Shapes.AddShape(MSO_SHAPE_ROUNDED_RECT, left, top, width, height)
        shp.Name = name
        shp.Fill.Solid()
        shp.Fill.ForeColor.RGB = _rgb(fill)
        if line:
            shp.Line.Visible = MSO_TRUE
            shp.Line.ForeColor.RGB = _rgb(line)
            shp.Line.Weight = 0.75
        else:
            shp.Line.Visible = MSO_FALSE
        self._soft_shadow(shp)
        self.stats["shapes_created"] += 1
        return shp

    def _rect(
        self,
        ws,
        name: str,
        left: float,
        top: float,
        width: float,
        height: float,
        fill: str,
        line: str | None = None,
    ):
        shp = ws.Shapes.AddShape(MSO_SHAPE_RECT, left, top, width, height)
        shp.Name = name
        shp.Fill.Solid()
        shp.Fill.ForeColor.RGB = _rgb(fill)
        if line:
            shp.Line.Visible = MSO_TRUE
            shp.Line.ForeColor.RGB = _rgb(line)
            shp.Line.Weight = 0.75
        else:
            shp.Line.Visible = MSO_FALSE
        self.stats["shapes_created"] += 1
        return shp

    def _set_text(
        self,
        shp,
        text: str,
        size: float = 10,
        bold: bool = False,
        color: str = INK,
        align: str = "left",
    ) -> None:
        tf = shp.TextFrame
        tf.Characters().Text = text
        font = tf.Characters().Font
        font.Name = FONT
        font.Size = size
        font.Bold = bold
        font.Color = _rgb(color)
        # Alignment: 1 left, 2 center, 3 right
        amap = {"left": 1, "center": 2, "right": 3}
        try:
            tf.HorizontalAlignment = amap.get(align, 1)
            tf.VerticalAlignment = 2  # center
            tf.MarginLeft = 8
            tf.MarginRight = 8
            tf.MarginTop = 4
            tf.MarginBottom = 4
            tf.WordWrap = MSO_TRUE
        except Exception:
            pass

    def _link_shape(self, ws, shp, sheet: str) -> None:
        try:
            ws.Hyperlinks.Add(
                Anchor=shp,
                Address="",
                SubAddress=f"'{sheet}'!A1",
                TextToDisplay="",
            )
        except Exception:
            try:
                shp.Select()
                self.excel.ActiveSheet.Hyperlinks.Add(
                    Anchor=self.excel.Selection,
                    Address="",
                    SubAddress=f"'{sheet}'!A1",
                )
            except Exception:
                pass

    def _badge(self, ws, name: str, left, top, width, height, text: str, bg: str, fg: str):
        shp = self._round_card(ws, name, left, top, width, height, fill=bg, line=None)
        self._set_text(shp, text, size=8, bold=True, color=fg, align="center")
        return shp

    def _kpi_card(
        self,
        ws,
        name: str,
        left: float,
        top: float,
        width: float,
        height: float,
        label: str,
        value: str,
        footer: str = "",
    ):
        card = self._round_card(ws, name, left, top, width, height, fill=CARD)
        # gold accent line at top
        self._rect(ws, name + "_accent", left + 10, top + 8, width - 20, 2.5, GOLD)
        # label
        lab = self._rect(ws, name + "_lab", left + 8, top + 14, width - 16, 16, CARD)
        lab.Line.Visible = MSO_FALSE
        self._set_text(lab, label.upper(), size=7.5, bold=True, color=MUTED, align="left")
        # value
        val = self._rect(ws, name + "_val", left + 8, top + 30, width - 16, 28, CARD)
        val.Line.Visible = MSO_FALSE
        self._set_text(val, value, size=20, bold=True, color=NAVY, align="left")
        if footer:
            ft = self._rect(ws, name + "_ft", left + 8, top + height - 20, width - 16, 14, CARD)
            ft.Line.Visible = MSO_FALSE
            self._set_text(ft, footer, size=7.5, bold=False, color=MUTED, align="left")
        return card

    def _progress_row(
        self,
        ws,
        name: str,
        left: float,
        top: float,
        width: float,
        label: str,
        pct: float,
        right_text: str,
    ):
        lab = self._rect(ws, name + "_l", left, top, 150, 16, CARD)
        lab.Line.Visible = MSO_FALSE
        self._set_text(lab, label, size=8.5, bold=False, color=INK, align="left")
        track_l = left + 155
        track_w = width - 230
        self._round_card(ws, name + "_track", track_l, top + 3, track_w, 10, fill=TRACK)
        fill_w = max(track_w * max(0.0, min(pct, 1.0)), 4)
        bar = self._round_card(ws, name + "_bar", track_l, top + 3, fill_w, 10, fill=GOLD)
        try:
            bar.Shadow.Visible = MSO_FALSE
        except Exception:
            pass
        rt = self._rect(ws, name + "_r", left + width - 70, top, 70, 16, CARD)
        rt.Line.Visible = MSO_FALSE
        self._set_text(rt, right_text, size=8, bold=True, color=NAVY, align="right")

    def _ensure_home_button(self, sheet_name: str) -> None:
        ws = self._focus(sheet_name)
        # Remove prior home buttons
        for i in range(int(ws.Shapes.Count), 0, -1):
            try:
                if str(ws.Shapes(i).Name).startswith("HomeBtn"):
                    ws.Shapes(i).Delete()
            except Exception:
                pass
        btn = self._round_card(ws, "HomeBtn", 12, 8, 72, 22, fill=NAVY)
        self._set_text(btn, "← HOME", size=8, bold=True, color=WHITE, align="center")
        self._link_shape(ws, btn, "00_HOME")
        self.stats["home_buttons"] += 1

    def _top_nav_strip(self, ws, sheet_name: str) -> None:
        # Minimal app header strip under home button area
        bar = self._rect(ws, f"TopNav_{sheet_name}", 90, 8, 620, 22, STONE)
        bar.Line.Visible = MSO_FALSE
        # nav pills
        items = [
            (96, "Navigation", "NAVIGATION"),
            (190, "Search", "SEARCH"),
            (270, "Document Center", "DOCUMENT_CENTER"),
        ]
        for left, label, target in items:
            if target == sheet_name:
                continue
            pill = self._round_card(
                ws, f"NavPill_{sheet_name}_{target}", left, 10, 88 if len(label) < 12 else 110, 18, fill=CARD, line=STEEL
            )
            try:
                pill.Shadow.Visible = MSO_FALSE
            except Exception:
                pass
            self._set_text(pill, label, size=7.5, bold=True, color=STEEL, align="center")
            self._link_shape(ws, pill, target)
        rev = self._badge(ws, f"RevBadge_{sheet_name}", 720, 10, 54, 18, "REV.00", GOLD, NAVY)

    # ── HOME ───────────────────────────────────────────────────────────
    def design_home(self) -> None:
        ws = self._focus("00_HOME")
        self._delete_all_shapes(ws)
        ws.Cells.Clear()
        self._canvas(ws, hide_headings=True)
        # Widen columns for positioning reference
        for col in range(1, 15):
            ws.Columns(col).ColumnWidth = 8.5

        # Full-width hero with subtle gradient via overlapping rects
        hero = self._rect(ws, "HeroBar", 10, 36, 760, 78, NAVY)
        try:
            hero.Fill.TwoColorGradient(MSO_GRADIENT_HORIZ, 1)
            hero.Fill.ForeColor.RGB = _rgb(NAVY)
            hero.Fill.BackColor.RGB = _rgb(NAVY_DEEP)
        except Exception:
            pass

        brand = self._rect(ws, "HeroBrand", 22, 42, 420, 14, NAVY)
        brand.Line.Visible = MSO_FALSE
        self._set_text(
            brand,
            "İNCİ AKÜ  •  PPWR PIMS",
            size=8,
            bold=True,
            color=GOLD_SOFT,
            align="left",
        )

        title = self._rect(ws, "HeroTitle", 22, 56, 520, 28, NAVY)
        title.Line.Visible = MSO_FALSE
        self._set_text(
            title,
            "İnci Akü PPWR\nPackaging Information Management System",
            size=14,
            bold=True,
            color=WHITE,
            align="left",
        )
        try:
            title.TextFrame.Characters().Font.Size = 14
        except Exception:
            pass

        sub = self._rect(ws, "HeroSub", 22, 92, 480, 14, NAVY)
        sub.Line.Visible = MSO_FALSE
        self._set_text(
            sub,
            "Controlled Packaging Data & Compliance Workspace",
            size=8,
            bold=False,
            color="B8C7D6",
            align="left",
        )

        # Logo top-right
        if self.logo_path.exists():
            try:
                pic = ws.Shapes.AddPicture(
                    str(self.logo_path.resolve()),
                    LinkToFile=False,
                    SaveWithDocument=True,
                    Left=620,
                    Top=48,
                    Width=140,
                    Height=48,
                )
                pic.Name = "InciAkuLogo"
                self.stats["shapes_created"] += 1
            except Exception:
                pass

        # Status badges
        self._badge(ws, "BadgeRev", 620, 12, 50, 18, "REV.00", GOLD, NAVY)
        self._badge(ws, "BadgeCtrl", 676, 12, 78, 18, "CONTROLLED", SUCCESS_BG, SUCCESS_FG)
        self._badge(ws, "BadgeVal", 760, 12, 10, 18, "", IVORY, NAVY)  # spacer placeholder
        # fix third badge position within canvas - place left of logo area at top
        try:
            ws.Shapes("BadgeVal").Delete()
            self.stats["shapes_created"] -= 1
        except Exception:
            pass
        self._badge(ws, "BadgeExcel", 500, 12, 110, 18, "EXCEL VALIDATED", SUCCESS_BG, SUCCESS_FG)

        # HOME button
        btn = self._round_card(ws, "HomeBtn", 12, 8, 72, 22, fill=GOLD)
        self._set_text(btn, "← HOME", size=8, bold=True, color=NAVY, align="center")
        self._link_shape(ws, btn, "00_HOME")
        self.stats["home_buttons"] += 1

        # Top nav
        for left, label, target in (
            (96, "Navigation", "NAVIGATION"),
            (196, "Search", "SEARCH"),
            (276, "Documents", "DOCUMENT_CENTER"),
        ):
            pill = self._round_card(ws, f"HomeNav_{target}", left, 10, 88, 18, fill=CARD, line=STEEL)
            try:
                pill.Shadow.Visible = MSO_FALSE
            except Exception:
                pass
            self._set_text(pill, label, size=7.5, bold=True, color=STEEL, align="center")
            self._link_shape(ws, pill, target)

        # Section: KPI
        sec = self._rect(ws, "SecKPI", 12, 126, 760, 18, IVORY)
        sec.Line.Visible = MSO_FALSE
        self._set_text(sec, "EXECUTIVE OVERVIEW", size=9, bold=True, color=NAVY, align="left")

        kpis_r1 = [
            ("Configurations", "247"),
            ("Products", "2,046"),
            ("Components", "112"),
            ("BOM Lines", "1,690"),
        ]
        kpis_r2 = [
            ("Technical Files", "247"),
            ("EU DoCs", "247"),
            ("Labels", "247"),
            ("Statements", "247"),
        ]
        kpis_r3 = [
            ("Documents", "988", "COMPLETE"),
            ("Linked", "988 / 988", "PASS"),
            ("Configs Valid", "247 / 247", "PASS"),
            ("Blocking Errors", "0", "CLEAR"),
        ]
        card_w, card_h, gap = 180, 78, 12
        base_left, base_top = 12, 148
        for i, (lab, val) in enumerate(kpis_r1):
            self._kpi_card(
                ws, f"KPI1_{i}", base_left + i * (card_w + gap), base_top, card_w, card_h, lab, val
            )
        for i, (lab, val) in enumerate(kpis_r2):
            self._kpi_card(
                ws,
                f"KPI2_{i}",
                base_left + i * (card_w + gap),
                base_top + card_h + gap,
                card_w,
                card_h,
                lab,
                val,
            )
        for i, item in enumerate(kpis_r3):
            lab, val = item[0], item[1]
            ft = item[2] if len(item) > 2 else ""
            self._kpi_card(
                ws,
                f"KPI3_{i}",
                base_left + i * (card_w + gap),
                base_top + 2 * (card_h + gap),
                card_w,
                card_h,
                lab,
                val,
                ft,
            )

        # System status panel (right of lower area - place below KPIs left, and mix panels)
        panel_top = base_top + 3 * (card_h + gap) + 8
        status_panel = self._round_card(ws, "StatusPanel", 12, panel_top, 300, 168, fill=CARD)
        hdr = self._rect(ws, "StatusHdr", 24, panel_top + 10, 270, 18, CARD)
        hdr.Line.Visible = MSO_FALSE
        self._set_text(hdr, "SYSTEM STATUS", size=9, bold=True, color=NAVY, align="left")
        self._rect(ws, "StatusGold", 24, panel_top + 30, 40, 2.5, GOLD)

        status_lines = [
            ("Production Master Data", "READY"),
            ("Golden Variant Register", "247 / 247"),
            ("Document Pack", "988 / 988"),
            ("Document Registry", "LINKED"),
            ("Native Excel Validation", "PASS"),
            ("Blocking QA Errors", "0"),
        ]
        for i, (k, v) in enumerate(status_lines):
            y = panel_top + 42 + i * 20
            row = self._rect(ws, f"StRow_{i}", 24, y, 270, 18, CARD)
            row.Line.Visible = MSO_FALSE
            self._set_text(row, f"●  {k}", size=8, bold=False, color=MUTED, align="left")
            pill = self._badge(
                ws, f"StPill_{i}", 210, y, 80, 16, v, SUCCESS_BG, SUCCESS_FG
            )

        # Configuration mix — composition bar (not skewed column chart)
        mix = self._round_card(ws, "MixPanel", 326, panel_top, 446, 78, fill=CARD)
        mh = self._rect(ws, "MixHdr", 338, panel_top + 8, 400, 16, CARD)
        mh.Line.Visible = MSO_FALSE
        self._set_text(mh, "CONFIGURATION MIX", size=9, bold=True, color=NAVY, align="left")
        # 100% bar: 240/247, 3/247, 4/247
        track_l, track_t, track_w, track_h = 338, panel_top + 36, 410, 16
        self._round_card(ws, "MixTrack", track_l, track_t, track_w, track_h, fill=TRACK)
        w1 = track_w * (240 / 247)
        w2 = track_w * (3 / 247)
        w3 = track_w - w1 - w2
        b1 = self._rect(ws, "MixS", track_l, track_t, w1, track_h, NAVY)
        b2 = self._rect(ws, "MixI", track_l + w1, track_t, max(w2, 3), track_h, STEEL)
        b3 = self._rect(ws, "MixC", track_l + w1 + w2, track_t, max(w3, 4), track_h, GOLD)
        legend = self._rect(ws, "MixLeg", 338, panel_top + 56, 410, 14, CARD)
        legend.Line.Visible = MSO_FALSE
        self._set_text(
            legend,
            "Starter 240   ·   Industrial 3   ·   Container 4",
            size=8,
            bold=False,
            color=MUTED,
            align="left",
        )

        # Document completeness progress panel
        docp = self._round_card(ws, "DocPanel", 326, panel_top + 90, 446, 120, fill=CARD)
        dh = self._rect(ws, "DocHdr", 338, panel_top + 98, 400, 16, CARD)
        dh.Line.Visible = MSO_FALSE
        self._set_text(
            dh, "DOCUMENT PACK COMPLETENESS", size=9, bold=True, color=NAVY, align="left"
        )
        rows = [
            ("Technical Files", 1.0, "247/247  100%"),
            ("EU DoCs", 1.0, "247/247  100%"),
            ("Labels", 1.0, "247/247  100%"),
            ("Shipment Statements", 1.0, "247/247  100%"),
        ]
        for i, (lab, pct, rt) in enumerate(rows):
            self._progress_row(
                ws, f"DocProg_{i}", 338, panel_top + 120 + i * 20, 420, lab, pct, rt
            )

        # Quick actions
        qa_top = panel_top + 230
        qh = self._rect(ws, "QAHdr", 12, qa_top, 760, 16, IVORY)
        qh.Line.Visible = MSO_FALSE
        self._set_text(qh, "QUICK ACTIONS", size=9, bold=True, color=NAVY, align="left")

        actions = [
            ("PACKAGING", "Packaging Configurations", "PACKAGING_CONFIGURATIONS"),
            ("PRODUCTS", "Product Master", "PRODUCT_MASTER"),
            ("COMPONENTS", "Component Master", "COMPONENT_MASTER"),
            ("DOCUMENTS", "Document Center", "DOCUMENT_CENTER"),
            ("TECHNICAL", "Technical Files", "TECHNICAL_FILES"),
            ("DECLARATIONS", "EU DoCs", "DECLARATIONS_OF_CONFORMITY"),
            ("LABELS", "Identification Labels", "LABELS"),
            ("SHIPMENTS", "Shipment Statements", "SHIPMENT_STATEMENTS"),
            ("SEARCH", "Global Search", "SEARCH"),
            ("SYSTEM", "Navigation", "NAVIGATION"),
        ]
        aw, ah, ag = 146, 52, 10
        for i, (eyebrow, title, target) in enumerate(actions):
            r, c = divmod(i, 5)
            left = 12 + c * (aw + ag)
            top = qa_top + 22 + r * (ah + ag)
            card = self._round_card(ws, f"Act_{target}", left, top, aw, ah, fill=CARD)
            self._rect(ws, f"ActGold_{target}", left, top + 8, 3, ah - 16, GOLD)
            eb = self._rect(ws, f"ActEb_{target}", left + 10, top + 8, aw - 18, 12, CARD)
            eb.Line.Visible = MSO_FALSE
            self._set_text(eb, eyebrow, size=7, bold=True, color=GOLD, align="left")
            tt = self._rect(ws, f"ActTt_{target}", left + 10, top + 22, aw - 18, 22, CARD)
            tt.Line.Visible = MSO_FALSE
            self._set_text(tt, title, size=9, bold=True, color=NAVY, align="left")
            self._link_shape(ws, card, target)

        # Footer
        foot = self._rect(ws, "Footer", 12, qa_top + 22 + 2 * (ah + ag) + 8, 760, 28, IVORY)
        foot.Line.Visible = MSO_FALSE
        self._set_text(
            foot,
            "İnci Akü Sanayi ve Ticaret A.Ş.   ·   PPWR Packaging Information Management System   ·   Rev.00 Controlled Baseline",
            size=7.5,
            bold=False,
            color=MUTED,
            align="left",
        )

        ws.Range("A1").Select()
        self.stats["sheets_designed"].append("00_HOME")

    # ── NAVIGATION ─────────────────────────────────────────────────────
    def design_navigation(self) -> None:
        ws = self._focus("NAVIGATION")
        self._delete_all_shapes(ws)
        ws.Cells.Clear()
        self._canvas(ws, hide_headings=True)
        for col in range(1, 15):
            ws.Columns(col).ColumnWidth = 8.5

        hero = self._rect(ws, "NavHero", 10, 36, 760, 70, NAVY)
        try:
            hero.Fill.TwoColorGradient(MSO_GRADIENT_HORIZ, 1)
            hero.Fill.ForeColor.RGB = _rgb(NAVY)
            hero.Fill.BackColor.RGB = _rgb(NAVY_DEEP)
        except Exception:
            pass
        t = self._rect(ws, "NavTitle", 22, 48, 400, 24, NAVY)
        t.Line.Visible = MSO_FALSE
        self._set_text(t, "Navigation", size=20, bold=True, color=WHITE, align="left")
        s = self._rect(ws, "NavSub", 22, 78, 480, 16, NAVY)
        s.Line.Visible = MSO_FALSE
        self._set_text(
            s,
            "Jump directly to any PPWR management module",
            size=9,
            bold=False,
            color="B8C7D6",
            align="left",
        )
        if self.logo_path.exists():
            try:
                pic = ws.Shapes.AddPicture(
                    str(self.logo_path.resolve()),
                    False,
                    True,
                    620,
                    48,
                    140,
                    48,
                )
                pic.Name = "InciAkuLogo_Nav"
                self.stats["shapes_created"] += 1
            except Exception:
                pass

        self._ensure_home_button("NAVIGATION")
        self._top_nav_strip(ws, "NAVIGATION")

        sections = [
            (
                "DASHBOARD & SEARCH",
                [
                    ("HOME", "Executive cockpit", "00_HOME", ""),
                    ("SEARCH", "Global inquiry workspace", "SEARCH", ""),
                ],
            ),
            (
                "PACKAGING DATA",
                [
                    ("PACKAGING CONFIGURATIONS", "247 controlled configurations", "PACKAGING_CONFIGURATIONS", "→ Open module"),
                    ("PRODUCT MASTER", "2,046 products", "PRODUCT_MASTER", "→ Open module"),
                    ("COMPONENT MASTER", "112 components", "COMPONENT_MASTER", "→ Open module"),
                ],
            ),
            (
                "DOCUMENT CONTROL",
                [
                    ("DOCUMENT CENTER", "988 linked controlled documents", "DOCUMENT_CENTER", "→ Open module"),
                    ("TECHNICAL FILES", "247 technical files", "TECHNICAL_FILES", "→ Open module"),
                    ("DECLARATIONS", "247 EU DoCs", "DECLARATIONS_OF_CONFORMITY", "→ Open module"),
                    ("LABELS", "247 identification labels", "LABELS", "→ Open module"),
                    ("SHIPMENT STATEMENTS", "247 statements", "SHIPMENT_STATEMENTS", "→ Open module"),
                ],
            ),
            (
                "SYSTEM & OPERATIONS",
                [
                    ("SHIPMENTS", "Transactional register", "SHIPMENTS", "→ Open module"),
                    ("DOC ENGINE MAP", "Read-only engine map", "DOC_ENGINE_MAP", "→ Open module"),
                ],
            ),
        ]

        y = 120
        for sec_title, cards in sections:
            sh = self._rect(ws, f"Sec_{sec_title[:8]}", 12, y, 760, 16, IVORY)
            sh.Line.Visible = MSO_FALSE
            self._set_text(sh, sec_title, size=9, bold=True, color=NAVY, align="left")
            y += 22
            cw, ch, cg = 244, 64, 12
            for i, (title, desc, target, cta) in enumerate(cards):
                if not self._has(target):
                    continue
                col = i % 3
                if i > 0 and col == 0:
                    y += ch + cg
                left = 12 + col * (cw + cg)
                card = self._round_card(ws, f"NavCard_{target}", left, y, cw, ch, fill=CARD)
                self._rect(ws, f"NavAccent_{target}", left, y + 10, 3, ch - 20, GOLD)
                tt = self._rect(ws, f"NavT_{target}", left + 12, y + 10, cw - 24, 18, CARD)
                tt.Line.Visible = MSO_FALSE
                self._set_text(tt, title, size=9, bold=True, color=NAVY, align="left")
                dd = self._rect(ws, f"NavD_{target}", left + 12, y + 30, cw - 24, 14, CARD)
                dd.Line.Visible = MSO_FALSE
                self._set_text(dd, desc, size=8, bold=False, color=MUTED, align="left")
                if cta:
                    cc = self._rect(ws, f"NavC_{target}", left + 12, y + 46, cw - 24, 12, CARD)
                    cc.Line.Visible = MSO_FALSE
                    self._set_text(cc, cta, size=7.5, bold=True, color=STEEL, align="left")
                self._link_shape(ws, card, target)
            y += ch + 20

        self.stats["sheets_designed"].append("NAVIGATION")

    # ── SEARCH ─────────────────────────────────────────────────────────
    def design_search(self) -> None:
        ws = self._focus("SEARCH")
        self._delete_all_shapes(ws)
        # Keep one input cell; clear the rest of chrome
        ws.Cells.Clear()
        self._canvas(ws, hide_headings=True)
        for col in range(1, 15):
            ws.Columns(col).ColumnWidth = 8.5

        hero = self._rect(ws, "SearchHero", 10, 36, 760, 58, NAVY)
        t = self._rect(ws, "SearchTitle", 22, 44, 500, 22, NAVY)
        t.Line.Visible = MSO_FALSE
        self._set_text(t, "Global Search", size=18, bold=True, color=WHITE, align="left")
        s = self._rect(ws, "SearchSub", 22, 70, 560, 14, NAVY)
        s.Line.Visible = MSO_FALSE
        self._set_text(
            s,
            "Search packaging sets, configurations, source BOMs and document IDs",
            size=8.5,
            bold=False,
            color="B8C7D6",
            align="left",
        )

        self._ensure_home_button("SEARCH")
        self._top_nav_strip(ws, "SEARCH")

        # Search card
        card = self._round_card(ws, "SearchCard", 12, 110, 760, 100, fill=CARD)
        lab = self._rect(ws, "SearchBy", 28, 122, 200, 14, CARD)
        lab.Line.Visible = MSO_FALSE
        self._set_text(lab, "SEARCH BY", size=8, bold=True, color=MUTED, align="left")
        hint = self._rect(ws, "SearchHint", 28, 138, 400, 14, CARD)
        hint.Line.Visible = MSO_FALSE
        self._set_text(
            hint,
            "Packaging Set / Configuration ID",
            size=9,
            bold=True,
            color=NAVY,
            align="left",
        )

        # Native input cell inside card area (Excel-safe)
        # Place at C10 visually near card — use cell for XLOOKUP key
        ws.Range("C10").Value = ""
        ws.Range("C10").Interior.Color = _rgb("FFF8E8")
        ws.Range("C10").Font.Name = FONT
        ws.Range("C10").Font.Size = 14
        ws.Range("C10").Font.Bold = True
        ws.Range("C10").Font.Color = _rgb(NAVY)
        ws.Range("C10").Borders.Color = _rgb(GOLD)
        ws.Range("C10").ColumnWidth = 28
        ws.Range("C10").RowHeight = 26

        ex = self._rect(ws, "SearchEx", 28, 180, 500, 14, CARD)
        ex.Line.Visible = MSO_FALSE
        self._set_text(
            ex,
            "Example:  ST-051-STD-01   ·   CNT-20-STD-01   ·   IND-24V-01",
            size=8,
            bold=False,
            color=MUTED,
            align="left",
        )

        # Results card
        res = self._round_card(ws, "ResultCard", 12, 230, 760, 200, fill=CARD)
        rh = self._rect(ws, "ResultHdr", 28, 242, 400, 16, CARD)
        rh.Line.Visible = MSO_FALSE
        self._set_text(rh, "RESULTS", size=9, bold=True, color=NAVY, align="left")
        self._rect(ws, "ResultGold", 28, 260, 36, 2.5, GOLD)

        # Labels in shapes + formulas in cells D12:D17 (aligned under results)
        fields = [
            (12, "Configuration ID", "PACKAGING_CONFIGURATIONS!B:B"),
            (13, "Family", "PACKAGING_CONFIGURATIONS!D:D"),
            (14, "Source Configuration ID", "PACKAGING_CONFIGURATIONS!C:C"),
            (15, "Technical File ID", "PACKAGING_CONFIGURATIONS!K:K"),
            (16, "DoC / Label / Statement", "PACKAGING_CONFIGURATIONS!L:L"),
            (17, "Packaging Tare kg", "PACKAGING_CONFIGURATIONS!H:H"),
        ]
        # Position result cells
        ws.Range("B12").Value = "Field"
        ws.Range("D12").Value = "Value"
        for i, (row, label, colref) in enumerate(fields):
            r = 12 + i
            ws.Range(f"B{r}").Value = label
            ws.Range(f"B{r}").Font.Name = FONT
            ws.Range(f"B{r}").Font.Size = 9
            ws.Range(f"B{r}").Font.Color = _rgb(MUTED)
            ws.Range(f"B{r}").Interior.Color = _rgb(CARD)
            ws.Range(f"D{r}").Formula = (
                f'=IF($C$10="","",IFERROR(XLOOKUP($C$10,PACKAGING_CONFIGURATIONS!A:A,{colref}),'
                f'"Not found — use AutoFilter on Document Center"))'
            )
            ws.Range(f"D{r}").Font.Name = FONT
            ws.Range(f"D{r}").Font.Size = 10
            ws.Range(f"D{r}").Font.Color = _rgb(INK)
            ws.Range(f"D{r}").Interior.Color = _rgb(STONE)

        # Status line
        ws.Range("B18").Value = "Document Pack Status"
        ws.Range("D18").Value = "988 / 988 LINKED  ·  Rev.00"
        ws.Range("B18").Font.Color = _rgb(MUTED)
        ws.Range("D18").Font.Color = _rgb(SUCCESS_FG)
        ws.Range("D18").Font.Bold = True
        ws.Range("D18").Interior.Color = _rgb(SUCCESS_BG)

        # Quick link cards
        y = 450
        qh = self._rect(ws, "SQHdr", 12, y, 400, 14, IVORY)
        qh.Line.Visible = MSO_FALSE
        self._set_text(qh, "QUICK LINKS", size=9, bold=True, color=NAVY, align="left")
        links = [
            ("Document Center", "DOCUMENT_CENTER"),
            ("Packaging Configurations", "PACKAGING_CONFIGURATIONS"),
            ("Technical Files", "TECHNICAL_FILES"),
            ("Product Master", "PRODUCT_MASTER"),
        ]
        for i, (label, target) in enumerate(links):
            left = 12 + i * 190
            c = self._round_card(ws, f"SQ_{target}", left, y + 20, 180, 40, fill=CARD)
            self._rect(ws, f"SQG_{target}", left, y + 28, 3, 24, GOLD)
            tt = self._rect(ws, f"SQT_{target}", left + 10, y + 28, 160, 24, CARD)
            tt.Line.Visible = MSO_FALSE
            self._set_text(tt, label, size=8.5, bold=True, color=NAVY, align="left")
            self._link_shape(ws, c, target)

        self.stats["sheets_designed"].append("SEARCH")

    # ── Register sheets ────────────────────────────────────────────────
    def design_register(self, name: str) -> None:
        meta = REGISTER_META[name]
        title, subtitle, chips = meta
        ws = self._focus(name)
        # Do NOT clear data / hyperlinks — only remove prior PhaseN shapes and add chrome
        for i in range(int(ws.Shapes.Count), 0, -1):
            try:
                sn = str(ws.Shapes(i).Name)
                if sn.startswith(("HomeBtn", "TopNav", "NavPill", "RevBadge", "RegHero", "Chip_", "RegAccent", "InciAku")):
                    ws.Shapes(i).Delete()
            except Exception:
                pass

        win = self.excel.ActiveWindow
        win.DisplayGridlines = False
        try:
            win.DisplayHeadings = True
        except Exception:
            pass

        # Soften used range borders slightly — header detect via row 5 typical after Phase L/M
        # Add floating header bar
        hero = self._rect(ws, f"RegHero_{name}", 90, 4, 620, 48, NAVY)
        try:
            hero.Fill.TwoColorGradient(MSO_GRADIENT_HORIZ, 1)
            hero.Fill.ForeColor.RGB = _rgb(NAVY)
            hero.Fill.BackColor.RGB = _rgb(STEEL)
        except Exception:
            pass
        tt = self._rect(ws, f"RegTitle_{name}", 100, 8, 400, 20, NAVY)
        tt.Line.Visible = MSO_FALSE
        self._set_text(tt, title, size=12, bold=True, color=WHITE, align="left")
        st = self._rect(ws, f"RegSub_{name}", 100, 30, 480, 14, NAVY)
        st.Line.Visible = MSO_FALSE
        self._set_text(st, subtitle, size=8, bold=False, color="D5DEE8", align="left")

        # chips
        x = 100
        for i, (lab, val) in enumerate(chips):
            chip = self._round_card(ws, f"Chip_{name}_{i}", x, 56, 100, 28, fill=CARD, line=STEEL)
            try:
                chip.Shadow.Visible = MSO_FALSE
            except Exception:
                pass
            self._set_text(chip, f"{lab}\n{val}", size=7.5, bold=True, color=NAVY, align="center")
            x += 108

        self._ensure_home_button(name)
        # compact top nav without overlapping hero too much
        for left, label, target in (
            (12, "Nav", "NAVIGATION"),
            (48, "Search", "SEARCH"),
        ):
            if target == name:
                continue
            # small text links in cells near A1 already have home shape
            pass

        # Style header row if found (row with Packaging Set / Product Code etc.)
        header_row = self._find_header_row_com(ws)
        if header_row:
            last_col = ws.UsedRange.Columns.Count
            last_row = ws.UsedRange.Rows.Count
            # UsedRange may start mid-sheet; use broader approach
            hdr = ws.Rows(header_row)
            hdr.Font.Name = FONT
            hdr.Font.Bold = True
            hdr.Font.Color = _rgb(WHITE)
            hdr.Interior.Color = _rgb(NAVY)
            hdr.RowHeight = 26
            # band data rows lightly (cap for performance)
            max_band = min(header_row + 400, ws.UsedRange.Row + ws.UsedRange.Rows.Count)
            for r in range(header_row + 1, max_band + 1):
                if r % 2 == 0:
                    ws.Rows(r).Interior.Color = _rgb("F3F6F9")
                else:
                    ws.Rows(r).Interior.Color = _rgb(WHITE)
            try:
                ws.Application.ActiveWindow.FreezePanes = False
                ws.Rows(header_row + 1).Select()
                self.excel.ActiveWindow.FreezePanes = True
            except Exception:
                pass
            # AutoFilter preserve/reapply
            try:
                if ws.AutoFilterMode:
                    pass
                else:
                    ws.Rows(header_row).AutoFilter()
            except Exception:
                pass

        ws.Range("A1").Select()
        self.stats["sheets_designed"].append(name)

    def _find_header_row_com(self, ws) -> int | None:
        for r in range(1, 12):
            v = ws.Cells(r, 1).Value
            if v is None:
                continue
            s = str(v)
            if s.startswith("=") or s.startswith("◆"):
                continue
            keys = (
                "Packaging Set",
                "Product Code",
                "ERP Component",
                "Label ID",
                "Technical File",
                "Declaration",
                "Statement",
                "Configuration",
            )
            if any(k in s for k in keys):
                return r
            v2 = ws.Cells(r, 2).Value
            if v2 and any(
                k in str(v2)
                for k in ("Configuration", "Product", "Component", "Document", "Family")
            ):
                return r
        return None

    def _reorder_ui_first(self) -> None:
        # Move UI sheets to front
        for idx, name in enumerate(UI_SHEETS, start=1):
            if not self._has(name):
                continue
            try:
                self.wb.Worksheets(name).Move(Before=self.wb.Worksheets(idx))
            except Exception:
                pass
        self._ungroup()
