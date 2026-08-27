"""Phase O2 — Class A full application canvas (HOME / NAVIGATION / SEARCH)."""

from __future__ import annotations

from pathlib import Path
from typing import Any

import win32com.client

MSO_RECT = 1
MSO_ROUNDED = 5
MSO_TRUE = -1
MSO_FALSE = 0
XL_FREE_FLOATING = 3  # xlFreeFloating — do not move/size with cells
MSO_GRADIENT_HORIZ = 1

NAVY = "0E2A47"
NAVY_DEEP = "0A1F35"
STEEL = "315E87"
GOLD = "C8A24A"
GOLD_SOFT = "D2B15B"
IVORY = "F7F5F0"
CARD = "FFFFFF"
STONE = "F3F1EB"
TRACK = "E4E8EE"
OK_BG = "DDEBDD"
OK_FG = "2F5D3A"
INK = "1C2430"
MUTED = "5C6B7A"
WHITE = "FFFFFF"
FONT = "Tahoma"


def _rgb(h: str) -> int:
    h = h.lstrip("#")
    return int(h[4:6] + h[2:4] + h[0:2], 16)


class ClassACanvas:
    def __init__(self, excel, wb, logo_path: Path) -> None:
        self.excel = excel
        self.wb = wb
        self.logo_path = logo_path
        self.shapes_created = 0

    def _focus(self, name: str):
        ws = self.wb.Worksheets(name)
        ws.Select()
        ws.Activate()
        ws.Range("A1").Select()
        self.excel.CutCopyMode = False
        return ws

    def _clear_shapes(self, ws) -> int:
        n = int(ws.Shapes.Count)
        removed = n
        for i in range(n, 0, -1):
            try:
                ws.Shapes(i).Delete()
            except Exception:
                pass
        return removed

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
            sh.OffsetX = 1.5
            sh.OffsetY = 2.0
            sh.Transparency = 0.72
            try:
                sh.Blur = 5
            except Exception:
                pass
        except Exception:
            pass

    def _round(self, ws, name, left, top, w, h, fill=CARD, line=None):
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
        self._shadow(shp)
        self.shapes_created += 1
        return shp

    def _rect(self, ws, name, left, top, w, h, fill, line=None):
        shp = ws.Shapes.AddShape(MSO_RECT, left, top, w, h)
        shp.Name = name
        shp.Fill.Solid()
        shp.Fill.ForeColor.RGB = _rgb(fill)
        if line:
            shp.Line.Visible = MSO_TRUE
            shp.Line.ForeColor.RGB = _rgb(line)
        else:
            shp.Line.Visible = MSO_FALSE
        self._free(shp)
        self.shapes_created += 1
        return shp

    def _text(self, shp, text, size=10, bold=False, color=INK, align="left"):
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
            tf.MarginLeft = 8
            tf.MarginRight = 8
            tf.MarginTop = 3
            tf.MarginBottom = 3
            tf.WordWrap = MSO_TRUE
        except Exception:
            pass

    def _link(self, ws, shp, sheet: str) -> None:
        try:
            ws.Hyperlinks.Add(Anchor=shp, Address="", SubAddress=f"'{sheet}'!A1")
        except Exception:
            pass

    def _canvas(self, ws, zoom=92):
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
        # column widths BEFORE shapes
        for c in range(1, 16):
            ws.Columns(c).ColumnWidth = 9
        rng = ws.Range("A1:O45")
        rng.Interior.Color = _rgb(IVORY)
        rng.Font.Name = FONT
        rng.Borders.LineStyle = 0

    def design_home(self) -> int:
        ws = self._focus("00_HOME")
        removed = self._clear_shapes(ws)
        ws.Cells.Clear()
        self._canvas(ws, 92)

        # Top nav shapes (single hyperlink owner each)
        home = self._round(ws, "HomeBtn", 12, 8, 72, 22, fill=NAVY)
        self._text(home, "← HOME", 8, True, WHITE, "center")
        self._link(ws, home, "00_HOME")
        for left, label, target in (
            (92, "Navigation", "NAVIGATION"),
            (188, "Search", "SEARCH"),
            (268, "Documents", "DOCUMENT_CENTER"),
        ):
            p = self._round(ws, f"HomeNav_{target}", left, 10, 88, 18, fill=CARD, line=STEEL)
            try:
                p.Shadow.Visible = MSO_FALSE
            except Exception:
                pass
            self._text(p, label, 7.5, True, STEEL, "center")
            self._link(ws, p, target)
        self._round(ws, "BadgeExcel", 480, 10, 110, 18, fill=OK_BG)
        self._text(ws.Shapes("BadgeExcel"), "EXCEL VALIDATED", 7.5, True, OK_FG, "center")
        try:
            ws.Shapes("BadgeExcel").Shadow.Visible = MSO_FALSE
        except Exception:
            pass
        self._round(ws, "BadgeRev", 600, 10, 50, 18, fill=GOLD)
        self._text(ws.Shapes("BadgeRev"), "REV.00", 8, True, NAVY, "center")
        try:
            ws.Shapes("BadgeRev").Shadow.Visible = MSO_FALSE
        except Exception:
            pass

        # Hero
        hero = self._rect(ws, "Hero", 10, 36, 780, 70, NAVY)
        try:
            hero.Fill.TwoColorGradient(MSO_GRADIENT_HORIZ, 1)
            hero.Fill.ForeColor.RGB = _rgb(NAVY)
            hero.Fill.BackColor.RGB = _rgb(NAVY_DEEP)
        except Exception:
            pass
        b = self._rect(ws, "HeroBrand", 22, 42, 360, 12, NAVY)
        b.Line.Visible = MSO_FALSE
        self._text(b, "İNCİ AKÜ  •  PPWR PIMS", 8, True, GOLD_SOFT, "left")
        t = self._rect(ws, "HeroTitle", 22, 54, 520, 26, NAVY)
        t.Line.Visible = MSO_FALSE
        self._text(
            t,
            "İnci Akü PPWR Packaging Information Management System",
            13,
            True,
            WHITE,
            "left",
        )
        s = self._rect(ws, "HeroSub", 22, 84, 500, 14, NAVY)
        s.Line.Visible = MSO_FALSE
        self._text(s, "Controlled Packaging Data & Compliance Workspace", 8, False, "B8C7D6", "left")
        if self.logo_path.exists():
            try:
                pic = ws.Shapes.AddPicture(
                    str(self.logo_path.resolve()), False, True, 640, 46, 140, 48
                )
                pic.Name = "InciAkuLogo"
                self._free(pic)
                self.shapes_created += 1
            except Exception:
                pass

        # Primary KPIs (4 large)
        sec = self._rect(ws, "SecPrimary", 12, 116, 300, 14, IVORY)
        sec.Line.Visible = MSO_FALSE
        self._text(sec, "PRIMARY INDICATORS", 9, True, NAVY, "left")
        primaries = [
            ("247", "PACKAGING CONFIGURATIONS"),
            ("2,046", "PRODUCTS"),
            ("988", "CONTROLLED DOCUMENTS"),
            ("0", "BLOCKING ERRORS"),
        ]
        pw, ph, gap = 186, 86, 12
        for i, (val, lab) in enumerate(primaries):
            left = 12 + i * (pw + gap)
            top = 134
            card = self._round(ws, f"PKPI_{i}", left, top, pw, ph, CARD)
            self._rect(ws, f"PKPI_g_{i}", left + 12, top + 10, pw - 24, 2.5, GOLD)
            lv = self._rect(ws, f"PKPI_v_{i}", left + 10, top + 22, pw - 20, 30, CARD)
            lv.Line.Visible = MSO_FALSE
            self._text(lv, val, 22, True, NAVY, "left")
            ll = self._rect(ws, f"PKPI_l_{i}", left + 10, top + 56, pw - 20, 22, CARD)
            ll.Line.Visible = MSO_FALSE
            self._text(ll, lab, 7.5, True, MUTED, "left")

        # Secondary strip
        sec2 = self._rect(ws, "SecSecondary", 12, 232, 300, 14, IVORY)
        sec2.Line.Visible = MSO_FALSE
        self._text(sec2, "SECONDARY STRIP", 9, True, NAVY, "left")
        secondaries = [
            ("112", "Components"),
            ("1,690", "BOM Lines"),
            ("247", "Technical Files"),
            ("247", "EU DoCs"),
            ("247", "Labels"),
            ("247", "Statements"),
        ]
        sw, sh = 124, 52
        for i, (val, lab) in enumerate(secondaries):
            left = 12 + i * (sw + 8)
            top = 250
            card = self._round(ws, f"SKPI_{i}", left, top, sw, sh, CARD)
            self._rect(ws, f"SKPI_g_{i}", left, top + 8, 3, sh - 16, GOLD)
            vv = self._rect(ws, f"SKPI_v_{i}", left + 10, top + 8, sw - 16, 20, CARD)
            vv.Line.Visible = MSO_FALSE
            self._text(vv, val, 14, True, NAVY, "left")
            ll = self._rect(ws, f"SKPI_l_{i}", left + 10, top + 30, sw - 16, 14, CARD)
            ll.Line.Visible = MSO_FALSE
            self._text(ll, lab.upper(), 7, True, MUTED, "left")

        # Three panels
        panel_top = 318
        # System health
        sp = self._round(ws, "SysHealth", 12, panel_top, 250, 150, CARD)
        sh = self._rect(ws, "SysHdr", 24, panel_top + 10, 220, 14, CARD)
        sh.Line.Visible = MSO_FALSE
        self._text(sh, "SYSTEM HEALTH", 9, True, NAVY, "left")
        self._rect(ws, "SysGold", 24, panel_top + 28, 28, 2.5, GOLD)
        lines = [
            ("Master Data", "READY"),
            ("Golden Register", "247/247"),
            ("Document Pack", "988/988"),
            ("Registry", "LINKED"),
            ("Excel Validation", "PASS"),
            ("Blocking Errors", "0"),
        ]
        for i, (k, v) in enumerate(lines):
            y = panel_top + 40 + i * 17
            row = self._rect(ws, f"SysR_{i}", 24, y, 140, 14, CARD)
            row.Line.Visible = MSO_FALSE
            self._text(row, f"● {k}", 7.5, False, MUTED, "left")
            pill = self._round(ws, f"SysP_{i}", 170, y, 70, 14, OK_BG)
            try:
                pill.Shadow.Visible = MSO_FALSE
            except Exception:
                pass
            self._text(pill, v, 7, True, OK_FG, "center")

        # Portfolio mix
        mp = self._round(ws, "Portfolio", 276, panel_top, 250, 150, CARD)
        mh = self._rect(ws, "PortHdr", 288, panel_top + 10, 220, 14, CARD)
        mh.Line.Visible = MSO_FALSE
        self._text(mh, "PACKAGING PORTFOLIO", 9, True, NAVY, "left")
        track_l, track_t, track_w, track_h = 288, panel_top + 50, 214, 14
        self._round(ws, "MixTrack", track_l, track_t, track_w, track_h, TRACK)
        w1 = track_w * (240 / 247)
        w2 = max(track_w * (3 / 247), 3)
        w3 = max(track_w - w1 - w2, 4)
        self._rect(ws, "MixS", track_l, track_t, w1, track_h, NAVY)
        self._rect(ws, "MixI", track_l + w1, track_t, w2, track_h, STEEL)
        self._rect(ws, "MixC", track_l + w1 + w2, track_t, w3, track_h, GOLD)
        leg = self._rect(ws, "MixLeg", 288, panel_top + 78, 214, 40, CARD)
        leg.Line.Visible = MSO_FALSE
        self._text(
            leg,
            "Starter 240\nIndustrial 3\nContainer 4",
            8,
            False,
            MUTED,
            "left",
        )

        # Completeness
        dp = self._round(ws, "Completeness", 540, panel_top, 250, 150, CARD)
        dh = self._rect(ws, "CompHdr", 552, panel_top + 10, 220, 14, CARD)
        dh.Line.Visible = MSO_FALSE
        self._text(dh, "DOCUMENT COMPLETENESS", 9, True, NAVY, "left")
        for i, lab in enumerate(
            ["Technical Files", "EU DoCs", "Labels", "Statements"]
        ):
            y = panel_top + 36 + i * 26
            lr = self._rect(ws, f"CompL_{i}", 552, y, 100, 12, CARD)
            lr.Line.Visible = MSO_FALSE
            self._text(lr, lab, 7.5, False, INK, "left")
            self._round(ws, f"CompT_{i}", 552, y + 12, 160, 8, TRACK)
            bar = self._round(ws, f"CompB_{i}", 552, y + 12, 160, 8, GOLD)
            try:
                bar.Shadow.Visible = MSO_FALSE
                ws.Shapes(f"CompT_{i}").Shadow.Visible = MSO_FALSE
            except Exception:
                pass
            rr = self._rect(ws, f"CompR_{i}", 720, y, 50, 12, CARD)
            rr.Line.Visible = MSO_FALSE
            self._text(rr, "100%", 7.5, True, NAVY, "right")

        # Quick actions — first screen
        qa_top = 486
        qh = self._rect(ws, "QAHdr", 12, qa_top, 200, 14, IVORY)
        qh.Line.Visible = MSO_FALSE
        self._text(qh, "QUICK ACTIONS", 9, True, NAVY, "left")
        actions = [
            ("PACKAGING", "Packaging Configurations", "PACKAGING_CONFIGURATIONS"),
            ("PRODUCTS", "Product Master", "PRODUCT_MASTER"),
            ("COMPONENTS", "Component Master", "COMPONENT_MASTER"),
            ("DOCUMENTS", "Document Center", "DOCUMENT_CENTER"),
            ("TECHNICAL", "Technical Files", "TECHNICAL_FILES"),
            ("DECLARATIONS", "EU DoCs", "DECLARATIONS_OF_CONFORMITY"),
            ("LABELS", "Labels", "LABELS"),
            ("STATEMENTS", "Shipment Statements", "SHIPMENT_STATEMENTS"),
            ("SEARCH", "Global Search", "SEARCH"),
            ("SYSTEM", "Navigation", "NAVIGATION"),
        ]
        aw, ah, ag = 150, 48, 8
        for i, (eyebrow, title, target) in enumerate(actions):
            r, c = divmod(i, 5)
            left = 12 + c * (aw + ag)
            top = qa_top + 18 + r * (ah + ag)
            card = self._round(ws, f"Act_{target}", left, top, aw, ah, CARD)
            self._rect(ws, f"ActG_{target}", left, top + 8, 3, ah - 16, GOLD)
            eb = self._rect(ws, f"ActE_{target}", left + 10, top + 6, aw - 16, 12, CARD)
            eb.Line.Visible = MSO_FALSE
            self._text(eb, eyebrow, 7, True, GOLD, "left")
            tt = self._rect(ws, f"ActT_{target}", left + 10, top + 22, aw - 16, 18, CARD)
            tt.Line.Visible = MSO_FALSE
            self._text(tt, title, 8.5, True, NAVY, "left")
            # ONE hyperlink owner = card front
            self._link(ws, card, target)

        foot = self._rect(ws, "Footer", 12, qa_top + 18 + 2 * (ah + ag) + 6, 780, 22, IVORY)
        foot.Line.Visible = MSO_FALSE
        self._text(
            foot,
            "İnci Akü Sanayi ve Ticaret A.Ş.  ·  PPWR Packaging Information Management System  ·  Rev.00 Controlled Baseline",
            7.5,
            False,
            MUTED,
            "left",
        )
        ws.Range("A1").Select()
        return removed

    def design_navigation(self) -> int:
        ws = self._focus("NAVIGATION")
        removed = self._clear_shapes(ws)
        ws.Cells.Clear()
        self._canvas(ws, 95)
        home = self._round(ws, "HomeBtn", 12, 8, 72, 22, NAVY)
        self._text(home, "← HOME", 8, True, WHITE, "center")
        self._link(ws, home, "00_HOME")

        hero = self._rect(ws, "NavHero", 10, 36, 780, 60, NAVY)
        t = self._rect(ws, "NavTitle", 22, 44, 400, 24, NAVY)
        t.Line.Visible = MSO_FALSE
        self._text(t, "Navigation", 18, True, WHITE, "left")
        s = self._rect(ws, "NavSub", 22, 72, 500, 14, NAVY)
        s.Line.Visible = MSO_FALSE
        self._text(s, "Jump directly to any PPWR management module", 9, False, "B8C7D6", "left")
        if self.logo_path.exists():
            try:
                pic = ws.Shapes.AddPicture(
                    str(self.logo_path.resolve()), False, True, 640, 42, 140, 48
                )
                pic.Name = "InciAkuLogo_Nav"
                self._free(pic)
                self.shapes_created += 1
            except Exception:
                pass

        sections = [
            (
                "DASHBOARD & SEARCH",
                [
                    ("HOME", "Executive cockpit", "00_HOME"),
                    ("SEARCH", "Global inquiry workspace", "SEARCH"),
                ],
            ),
            (
                "PACKAGING DATA",
                [
                    ("PACKAGING CONFIGURATIONS", "247 controlled configurations", "PACKAGING_CONFIGURATIONS"),
                    ("PRODUCT MASTER", "2,046 products", "PRODUCT_MASTER"),
                    ("COMPONENT MASTER", "112 components", "COMPONENT_MASTER"),
                ],
            ),
            (
                "DOCUMENT CONTROL",
                [
                    ("DOCUMENT CENTER", "988 linked documents", "DOCUMENT_CENTER"),
                    ("TECHNICAL FILES", "247 technical files", "TECHNICAL_FILES"),
                    ("DECLARATIONS", "247 EU DoCs", "DECLARATIONS_OF_CONFORMITY"),
                    ("LABELS", "247 labels", "LABELS"),
                    ("SHIPMENT STATEMENTS", "247 statements", "SHIPMENT_STATEMENTS"),
                ],
            ),
            (
                "SYSTEM & OPERATIONS",
                [
                    ("SHIPMENTS", "Transactional register", "SHIPMENTS"),
                    ("DOC ENGINE MAP", "Read-only mapping", "DOC_ENGINE_MAP"),
                ],
            ),
        ]
        y = 110
        for sec, cards in sections:
            sh = self._rect(ws, f"Sec_{sec[:6]}", 12, y, 780, 14, IVORY)
            sh.Line.Visible = MSO_FALSE
            self._text(sh, sec, 9, True, NAVY, "left")
            y += 20
            cw, ch, cg = 250, 58, 12
            for i, (title, desc, target) in enumerate(cards):
                col = i % 3
                if i and col == 0:
                    y += ch + cg
                left = 12 + col * (cw + cg)
                card = self._round(ws, f"NavCard_{target}", left, y, cw, ch, CARD)
                self._rect(ws, f"NavG_{target}", left, y + 10, 3, ch - 20, GOLD)
                tt = self._rect(ws, f"NavT_{target}", left + 12, y + 10, cw - 24, 18, CARD)
                tt.Line.Visible = MSO_FALSE
                self._text(tt, title, 9, True, NAVY, "left")
                dd = self._rect(ws, f"NavD_{target}", left + 12, y + 30, cw - 24, 16, CARD)
                dd.Line.Visible = MSO_FALSE
                self._text(dd, desc + "  → Open", 8, False, MUTED, "left")
                self._link(ws, card, target)
            y += ch + 18
        return removed

    def design_search(self) -> int:
        ws = self._focus("SEARCH")
        removed = self._clear_shapes(ws)
        ws.Cells.Clear()
        self._canvas(ws, 95)
        home = self._round(ws, "HomeBtn", 12, 8, 72, 22, NAVY)
        self._text(home, "← HOME", 8, True, WHITE, "center")
        self._link(ws, home, "00_HOME")

        hero = self._rect(ws, "SearchHero", 10, 36, 780, 52, NAVY)
        t = self._rect(ws, "SearchTitle", 22, 42, 500, 22, NAVY)
        t.Line.Visible = MSO_FALSE
        self._text(t, "Global Search", 16, True, WHITE, "left")
        s = self._rect(ws, "SearchSub", 22, 68, 560, 12, NAVY)
        s.Line.Visible = MSO_FALSE
        self._text(
            s,
            "Search packaging sets, configurations, source BOMs and document IDs",
            8.5,
            False,
            "B8C7D6",
            "left",
        )

        # Search card frame (shape) — input remains a cell
        self._round(ws, "SearchCard", 12, 104, 780, 90, CARD)
        lab = self._rect(ws, "SearchLab", 28, 116, 240, 14, CARD)
        lab.Line.Visible = MSO_FALSE
        self._text(lab, "SEARCH BY  ·  Packaging Set / Configuration ID", 8, True, MUTED, "left")
        # Input cell C8
        ws.Range("C8").Value = ""
        ws.Range("C8").Interior.Color = _rgb("FFF8E8")
        ws.Range("C8").Font.Name = FONT
        ws.Range("C8").Font.Size = 14
        ws.Range("C8").Font.Bold = True
        ws.Range("C8").Font.Color = _rgb(NAVY)
        ws.Range("C8").Borders.Color = _rgb(GOLD)
        ws.Range("C8").RowHeight = 26
        ex = self._rect(ws, "SearchEx", 28, 168, 500, 14, CARD)
        ex.Line.Visible = MSO_FALSE
        self._text(ex, "Example: ST-051-STD-01  ·  CNT-20-STD-01  ·  IND-24V-01", 8, False, MUTED, "left")

        self._round(ws, "ResultCard", 12, 210, 780, 180, CARD)
        rh = self._rect(ws, "ResultHdr", 28, 222, 200, 14, CARD)
        rh.Line.Visible = MSO_FALSE
        self._text(rh, "RESULTS", 9, True, NAVY, "left")
        fields = [
            (10, "Configuration ID", "PACKAGING_CONFIGURATIONS!B:B"),
            (11, "Family", "PACKAGING_CONFIGURATIONS!D:D"),
            (12, "Source Configuration ID", "PACKAGING_CONFIGURATIONS!C:C"),
            (13, "Technical File ID", "PACKAGING_CONFIGURATIONS!K:K"),
            (14, "Packaging Tare kg", "PACKAGING_CONFIGURATIONS!H:H"),
        ]
        for r, label, colref in fields:
            ws.Range(f"B{r}").Value = label
            ws.Range(f"B{r}").Font.Name = FONT
            ws.Range(f"B{r}").Font.Size = 9
            ws.Range(f"B{r}").Font.Color = _rgb(MUTED)
            ws.Range(f"D{r}").Formula = (
                f'=IF($C$8="","",IFERROR(XLOOKUP($C$8,PACKAGING_CONFIGURATIONS!A:A,{colref}),'
                f'"Not found — use AutoFilter on Document Center"))'
            )
            ws.Range(f"D{r}").Font.Name = FONT
            ws.Range(f"D{r}").Font.Size = 10
            ws.Range(f"D{r}").Font.Color = _rgb(INK)
            ws.Range(f"D{r}").Interior.Color = _rgb(STONE)
        ws.Range("B15").Value = "Document Pack Status"
        ws.Range("D15").Value = "988 / 988 LINKED  ·  Rev.00"
        ws.Range("B15").Font.Name = FONT
        ws.Range("D15").Font.Name = FONT
        ws.Range("D15").Font.Bold = True
        ws.Range("D15").Font.Color = _rgb(OK_FG)
        ws.Range("D15").Interior.Color = _rgb(OK_BG)

        y = 410
        for i, (label, target) in enumerate(
            [
                ("Document Center", "DOCUMENT_CENTER"),
                ("Packaging Configurations", "PACKAGING_CONFIGURATIONS"),
                ("Technical Files", "TECHNICAL_FILES"),
                ("Product Master", "PRODUCT_MASTER"),
            ]
        ):
            left = 12 + i * 195
            c = self._round(ws, f"SQ_{target}", left, y, 185, 40, CARD)
            self._rect(ws, f"SQG_{target}", left, y + 8, 3, 24, GOLD)
            tt = self._rect(ws, f"SQT_{target}", left + 10, y + 10, 165, 22, CARD)
            tt.Line.Visible = MSO_FALSE
            self._text(tt, label, 8.5, True, NAVY, "left")
            self._link(ws, c, target)
        return removed
