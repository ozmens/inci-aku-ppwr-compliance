"""Phase O4 — Class A V5 premium executive canvas (HOME / NAV / SEARCH)."""

from __future__ import annotations

from pathlib import Path
from typing import Any

MSO_RECT = 1
MSO_ROUNDED = 5
MSO_TRUE = -1
MSO_FALSE = 0
XL_FREE_FLOATING = 3
MSO_GRADIENT_HORIZ = 1

# V5 approved palette
NAVY = "0B2341"
NAVY2 = "123A63"
MIDNIGHT = "081A2F"
GOLD = "C9A24A"
DARK_GOLD = "A9842E"
IVORY = "F5F1E8"
CARD = "FCFBF8"
PALE = "E9EEF3"
LINE = "D6DCE2"
OK_BG = "E5F0E7"
OK_FG = "325D3E"
INK = "1E2C3A"
MUTED = "5C6B7A"
WHITE = "FFFFFF"
FONT = "Tahoma"


def _rgb(h: str) -> int:
    h = h.lstrip("#")
    return int(h[4:6] + h[2:4] + h[0:2], 16)


class ClassAV5Canvas:
    def __init__(self, excel, wb, logo_path: Path) -> None:
        self.excel = excel
        self.wb = wb
        self.logo_path = logo_path
        self.shapes_created = 0
        self.hyperlinks_added = 0
        self.locked_shapes = 0

    def _focus(self, name: str):
        ws = self.wb.Worksheets(name)
        ws.Select()
        ws.Activate()
        ws.Range("A1").Select()
        self.excel.CutCopyMode = False
        return ws

    def _clear_shapes(self, ws) -> int:
        n = int(ws.Shapes.Count)
        for i in range(n, 0, -1):
            try:
                ws.Shapes(i).Delete()
            except Exception:
                pass
        return n

    def _free(self, shp) -> None:
        try:
            shp.Placement = XL_FREE_FLOATING
        except Exception:
            pass

    def _shadow(self, shp, soft: bool = True) -> None:
        try:
            sh = shp.Shadow
            sh.Visible = MSO_TRUE
            sh.Style = 1
            sh.OffsetX = 1.0 if soft else 1.5
            sh.OffsetY = 1.5 if soft else 2.0
            sh.Transparency = 0.82 if soft else 0.75
            try:
                sh.Blur = 5 if soft else 7
            except Exception:
                pass
        except Exception:
            pass

    def _no_shadow(self, shp) -> None:
        try:
            shp.Shadow.Visible = MSO_FALSE
        except Exception:
            pass

    def _round(self, ws, name, left, top, w, h, fill=CARD, line=None, shadow=True):
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
        self._no_shadow(shp)
        self.shapes_created += 1
        return shp

    def _navy_grad(self, shp) -> None:
        try:
            shp.Fill.TwoColorGradient(MSO_GRADIENT_HORIZ, 1)
            shp.Fill.ForeColor.RGB = _rgb(MIDNIGHT)
            shp.Fill.BackColor.RGB = _rgb(NAVY)
        except Exception:
            shp.Fill.Solid()
            shp.Fill.ForeColor.RGB = _rgb(NAVY)

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
            tf.MarginLeft = 10
            tf.MarginRight = 8
            tf.MarginTop = 4
            tf.MarginBottom = 4
            tf.WordWrap = MSO_TRUE
        except Exception:
            pass

    def _link(self, ws, shp, sheet: str) -> None:
        try:
            ws.Hyperlinks.Add(Anchor=shp, Address="", SubAddress=f"'{sheet}'!A1")
            self.hyperlinks_added += 1
        except Exception:
            pass

    def _lock_all(self, ws) -> int:
        n = 0
        for i in range(1, int(ws.Shapes.Count) + 1):
            try:
                ws.Shapes(i).Locked = True
                n += 1
            except Exception:
                pass
        self.locked_shapes += n
        return n

    def protect_ui_sheet(self, name: str, unlock_cells: list[str] | None = None) -> dict[str, Any]:
        ws = self._focus(name)
        locked = self._lock_all(ws)
        try:
            ws.Cells.Locked = False
        except Exception:
            pass
        if unlock_cells:
            for addr in unlock_cells:
                try:
                    ws.Range(addr).Locked = False
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
            mode = "DrawingObjects+Contents"
        except Exception:
            try:
                ws.Protect(
                    Password="",
                    DrawingObjects=True,
                    Contents=False,
                    Scenarios=False,
                    UserInterfaceOnly=True,
                )
                mode = "DrawingObjects_only"
            except Exception as exc:
                return {"sheet": name, "ok": False, "error": str(exc), "locked": locked}
        return {"sheet": name, "ok": True, "mode": mode, "locked_shapes": locked}

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
        for c in range(1, 16):
            ws.Columns(c).ColumnWidth = 9.2
        rng = ws.Range("A1:O42")
        rng.Interior.Color = _rgb(IVORY)
        rng.Font.Name = FONT
        rng.Borders.LineStyle = 0

    def _top_pills(self, ws) -> None:
        home = self._round(ws, "HomeBtn", 14, 10, 72, 22, NAVY, shadow=False)
        self._text(home, "← HOME", 8, True, WHITE, "center")
        self._link(ws, home, "00_HOME")
        for left, label, target in (
            (96, "Navigation", "NAVIGATION"),
            (188, "Search", "SEARCH"),
            (266, "Documents", "DOCUMENT_CENTER"),
        ):
            p = self._round(ws, f"TopNav_{target}", left, 11, 84, 20, CARD, line=LINE, shadow=False)
            self._text(p, label, 7.5, True, NAVY2, "center")
            self._link(ws, p, target)
        for left, label, fill, fg, name in (
            (520, "EXCEL VALIDATED", OK_BG, OK_FG, "PillExcel"),
            (648, "REV.00", GOLD, MIDNIGHT, "PillRev"),
            (710, "CONTROLLED", PALE, NAVY, "PillCtrl"),
        ):
            w = 118 if name == "PillExcel" else 54 if name != "PillCtrl" else 78
            if name == "PillCtrl":
                left = 710
            pill = self._round(ws, name, left, 11, w, 20, fill, shadow=False)
            self._text(pill, label, 7, True, fg, "center")

    def _hero(self, ws, title: str, subtitle: str, logo_name: str) -> None:
        hero = self._rect(ws, "Hero", 12, 40, 820, 78, NAVY)
        self._navy_grad(hero)
        gold = self._rect(ws, "HeroGold", 12, 118, 820, 3, GOLD)
        try:
            gold.Fill.TwoColorGradient(MSO_GRADIENT_HORIZ, 1)
            gold.Fill.ForeColor.RGB = _rgb(DARK_GOLD)
            gold.Fill.BackColor.RGB = _rgb(GOLD)
        except Exception:
            pass
        under = self._rect(ws, "HeroUnder", 12, 121, 820, 8, NAVY2)
        brand = self._rect(ws, "HeroBrand", 28, 48, 420, 12, MIDNIGHT)
        brand.Line.Visible = MSO_FALSE
        self._text(brand, "İNCİ AKÜ  ·  PPWR PIMS", 8, True, GOLD, "left")
        t = self._rect(ws, "HeroTitle", 28, 62, 560, 28, MIDNIGHT)
        t.Line.Visible = MSO_FALSE
        self._text(t, title, 20, True, WHITE, "left")
        s = self._rect(ws, "HeroSub", 28, 94, 560, 14, MIDNIGHT)
        s.Line.Visible = MSO_FALSE
        self._text(s, subtitle, 9, False, "A8B9C8", "left")
        if self.logo_path.exists():
            try:
                pic = ws.Shapes.AddPicture(
                    str(self.logo_path.resolve()), False, True, 680, 50, 136, 52
                )
                pic.Name = logo_name
                self._free(pic)
                self.shapes_created += 1
            except Exception:
                pass

    def design_home(self) -> int:
        ws = self._focus("00_HOME")
        removed = self._clear_shapes(ws)
        ws.Cells.Clear()
        self._canvas(ws, 90)
        self._top_pills(ws)
        self._hero(
            ws,
            "İnci Akü PPWR Packaging Information Management System",
            "Controlled Packaging Data & Compliance Workspace  ·  Executive Cockpit",
            "InciAkuLogo",
        )

        # Section label
        sec = self._rect(ws, "SecKPI", 14, 140, 260, 14, IVORY)
        sec.Line.Visible = MSO_FALSE
        self._text(sec, "PRIMARY INDICATORS", 9, True, NAVY, "left")

        primaries = [
            ("247", "PACKAGING CONFIGURATIONS"),
            ("2,046", "PRODUCTS"),
            ("988", "CONTROLLED DOCUMENTS"),
            ("0", "BLOCKING ERRORS"),
        ]
        pw, ph, gap = 198, 96, 10
        for i, (val, lab) in enumerate(primaries):
            left = 14 + i * (pw + gap)
            top = 158
            card = self._round(ws, f"PKPI_{i}", left, top, pw, ph, CARD, line=PALE)
            accent = self._rect(ws, f"PKPI_g_{i}", left + 16, top + 14, pw - 32, 3, GOLD)
            try:
                accent.Fill.TwoColorGradient(MSO_GRADIENT_HORIZ, 1)
                accent.Fill.ForeColor.RGB = _rgb(DARK_GOLD)
                accent.Fill.BackColor.RGB = _rgb(GOLD)
            except Exception:
                pass
            lv = self._rect(ws, f"PKPI_v_{i}", left + 14, top + 26, pw - 28, 36, CARD)
            lv.Line.Visible = MSO_FALSE
            self._text(lv, val, 26, True, NAVY, "left")
            ll = self._rect(ws, f"PKPI_l_{i}", left + 14, top + 66, pw - 28, 22, CARD)
            ll.Line.Visible = MSO_FALSE
            self._text(ll, lab, 8, True, MUTED, "left")

        # Secondary strip
        sec2 = self._rect(ws, "SecSec", 14, 266, 260, 14, IVORY)
        sec2.Line.Visible = MSO_FALSE
        self._text(sec2, "SECONDARY METRICS", 9, True, NAVY, "left")
        secondaries = [
            ("112", "Components"),
            ("1,690", "BOM Lines"),
            ("247", "Technical Files"),
            ("247", "EU DoCs"),
            ("247", "Labels"),
            ("247", "Statements"),
        ]
        sw, sh = 132, 50
        for i, (val, lab) in enumerate(secondaries):
            left = 14 + i * (sw + 6)
            top = 284
            card = self._round(ws, f"SKPI_{i}", left, top, sw, sh, CARD, line=PALE)
            self._rect(ws, f"SKPI_g_{i}", left, top + 10, 3.5, sh - 20, GOLD)
            vv = self._rect(ws, f"SKPI_v_{i}", left + 12, top + 8, sw - 18, 18, CARD)
            vv.Line.Visible = MSO_FALSE
            self._text(vv, val, 14, True, NAVY, "left")
            ll = self._rect(ws, f"SKPI_l_{i}", left + 12, top + 28, sw - 18, 14, CARD)
            ll.Line.Visible = MSO_FALSE
            self._text(ll, lab.upper(), 7, True, MUTED, "left")

        # Insight panels
        panel_top = 348
        # System Health
        self._round(ws, "SysHealth", 14, panel_top, 268, 148, CARD, line=PALE)
        sh = self._rect(ws, "SysHdr", 28, panel_top + 12, 230, 14, CARD)
        sh.Line.Visible = MSO_FALSE
        self._text(sh, "SYSTEM HEALTH", 9, True, NAVY, "left")
        self._rect(ws, "SysGold", 28, panel_top + 28, 28, 2.5, GOLD)
        for i, (k, v) in enumerate(
            [
                ("Master Data", "READY"),
                ("Golden Register", "247/247"),
                ("Document Pack", "988/988"),
                ("Registry", "LINKED"),
                ("Excel Validation", "PASS"),
                ("Blocking Errors", "0"),
            ]
        ):
            y = panel_top + 40 + i * 16
            row = self._rect(ws, f"SysR_{i}", 28, y, 140, 14, CARD)
            row.Line.Visible = MSO_FALSE
            self._text(row, f"●  {k}", 7.5, False, MUTED, "left")
            pill = self._round(ws, f"SysP_{i}", 178, y, 78, 14, OK_BG, shadow=False)
            self._text(pill, v, 7, True, OK_FG, "center")

        # Portfolio
        self._round(ws, "Portfolio", 292, panel_top, 268, 148, CARD, line=PALE)
        mh = self._rect(ws, "PortHdr", 306, panel_top + 12, 230, 14, CARD)
        mh.Line.Visible = MSO_FALSE
        self._text(mh, "PACKAGING PORTFOLIO", 9, True, NAVY, "left")
        track_l, track_t, track_w, track_h = 306, panel_top + 48, 232, 14
        self._round(ws, "MixTrack", track_l, track_t, track_w, track_h, PALE, shadow=False)
        w1 = track_w * (240 / 247)
        w2 = max(track_w * (3 / 247), 4)
        w3 = max(track_w - w1 - w2, 5)
        self._rect(ws, "MixS", track_l, track_t, w1, track_h, NAVY)
        self._rect(ws, "MixI", track_l + w1, track_t, w2, track_h, NAVY2)
        self._rect(ws, "MixC", track_l + w1 + w2, track_t, w3, track_h, GOLD)
        leg = self._rect(ws, "MixLeg", 306, panel_top + 78, 232, 50, CARD)
        leg.Line.Visible = MSO_FALSE
        self._text(leg, "Starter 240\nIndustrial 3\nContainer 4", 8.5, False, MUTED, "left")

        # Completeness
        self._round(ws, "Completeness", 570, panel_top, 262, 148, CARD, line=PALE)
        dh = self._rect(ws, "CompHdr", 584, panel_top + 12, 230, 14, CARD)
        dh.Line.Visible = MSO_FALSE
        self._text(dh, "DOCUMENT COMPLETENESS", 9, True, NAVY, "left")
        for i, lab in enumerate(["Technical Files", "EU DoCs", "Labels", "Statements"]):
            y = panel_top + 36 + i * 26
            lr = self._rect(ws, f"CompL_{i}", 584, y, 110, 12, CARD)
            lr.Line.Visible = MSO_FALSE
            self._text(lr, lab, 7.5, False, INK, "left")
            self._round(ws, f"CompT_{i}", 584, y + 13, 170, 8, PALE, shadow=False)
            self._round(ws, f"CompB_{i}", 584, y + 13, 170, 8, GOLD, shadow=False)
            rr = self._rect(ws, f"CompR_{i}", 762, y, 50, 12, CARD)
            rr.Line.Visible = MSO_FALSE
            self._text(rr, "100%", 7.5, True, NAVY, "right")

        # Quick actions
        qa_top = 512
        qh = self._rect(ws, "QAHdr", 14, qa_top, 220, 14, IVORY)
        qh.Line.Visible = MSO_FALSE
        self._text(qh, "QUICK ACTIONS", 9, True, NAVY, "left")
        actions = [
            ("DOCUMENT CENTER", "988 linked controlled documents", "DOCUMENT_CENTER"),
            ("PACKAGING CONFIGS", "247 controlled configurations", "PACKAGING_CONFIGURATIONS"),
            ("TECHNICAL FILES", "247 controlled Rev.00 files", "TECHNICAL_FILES"),
            ("EU DECLARATIONS", "247 EU declarations", "DECLARATIONS_OF_CONFORMITY"),
            ("LABELS", "247 packaging labels", "LABELS"),
            ("STATEMENTS", "247 shipment statements", "SHIPMENT_STATEMENTS"),
            ("PRODUCT MASTER", "2,046 products", "PRODUCT_MASTER"),
            ("COMPONENT MASTER", "112 packaging components", "COMPONENT_MASTER"),
            ("GLOBAL SEARCH", "Packaging set inquiry", "SEARCH"),
            ("NAVIGATION", "Module launchpad", "NAVIGATION"),
        ]
        aw, ah, ag = 158, 54, 8
        for i, (title, desc, target) in enumerate(actions):
            r, c = divmod(i, 5)
            left = 14 + c * (aw + ag)
            top = qa_top + 18 + r * (ah + ag)
            card = self._round(ws, f"Act_{target}", left, top, aw, ah, CARD, line=PALE)
            self._rect(ws, f"ActG_{target}", left, top + 10, 3.5, ah - 20, GOLD)
            tt = self._rect(ws, f"ActT_{target}", left + 12, top + 8, aw - 20, 16, CARD)
            tt.Line.Visible = MSO_FALSE
            self._text(tt, title, 8, True, NAVY, "left")
            dd = self._rect(ws, f"ActD_{target}", left + 12, top + 26, aw - 20, 22, CARD)
            dd.Line.Visible = MSO_FALSE
            self._text(dd, f"{desc}\nOpen →", 7, False, MUTED, "left")
            self._link(ws, card, target)

        foot = self._rect(ws, "Footer", 14, qa_top + 18 + 2 * (ah + ag) + 6, 820, 18, IVORY)
        foot.Line.Visible = MSO_FALSE
        self._text(
            foot,
            "İnci Akü Sanayi ve Ticaret A.Ş.  ·  PPWR Packaging Information Management System  ·  Rev.00 Controlled Baseline",
            7.5,
            False,
            MUTED,
            "left",
        )
        self._lock_all(ws)
        ws.Range("A1").Select()
        return removed

    def design_navigation(self) -> int:
        ws = self._focus("NAVIGATION")
        removed = self._clear_shapes(ws)
        ws.Cells.Clear()
        self._canvas(ws, 92)
        self._top_pills(ws)
        self._hero(
            ws,
            "Navigation",
            "Jump directly to any PPWR management module",
            "InciAkuLogo_Nav",
        )

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
                    ("COMPONENT MASTER", "112 packaging components", "COMPONENT_MASTER"),
                ],
            ),
            (
                "DOCUMENT CONTROL",
                [
                    ("DOCUMENT CENTER", "988 linked documents", "DOCUMENT_CENTER"),
                    ("TECHNICAL FILES", "247 controlled Rev.00 files", "TECHNICAL_FILES"),
                    ("DECLARATIONS", "247 EU declarations", "DECLARATIONS_OF_CONFORMITY"),
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
        y = 146
        for sec, cards in sections:
            sh = self._rect(ws, f"Sec_{sec[:10]}", 14, y, 820, 14, IVORY)
            sh.Line.Visible = MSO_FALSE
            self._text(sh, sec, 9, True, NAVY, "left")
            y += 20
            cw, ch, cg = 266, 68, 12
            for i, (title, desc, target) in enumerate(cards):
                col = i % 3
                if i and col == 0:
                    y += ch + cg
                left = 14 + col * (cw + cg)
                card = self._round(ws, f"NavCard_{target}", left, y, cw, ch, CARD, line=PALE)
                self._rect(ws, f"NavG_{target}", left, y + 12, 4, ch - 24, GOLD)
                tt = self._rect(ws, f"NavT_{target}", left + 16, y + 12, cw - 30, 20, CARD)
                tt.Line.Visible = MSO_FALSE
                self._text(tt, title, 9.5, True, NAVY, "left")
                dd = self._rect(ws, f"NavD_{target}", left + 16, y + 36, cw - 30, 24, CARD)
                dd.Line.Visible = MSO_FALSE
                self._text(dd, f"{desc}\nOpen →", 8, False, MUTED, "left")
                self._link(ws, card, target)
            y += ch + 18
        self._lock_all(ws)
        return removed

    def design_search(self) -> int:
        """Critical: visible editable search input — no opaque shapes over C8:G9."""
        ws = self._focus("SEARCH")
        removed = self._clear_shapes(ws)
        ws.Cells.Clear()
        self._canvas(ws, 92)
        self._top_pills(ws)
        self._hero(
            ws,
            "Global Search",
            "Search packaging sets, configurations, source BOMs and document IDs",
            "InciAkuLogo_Search",
        )

        # Cell-based search card — fully clickable input (no covering shapes)
        for r in range(7, 12):
            for c in range(1, 15):
                ws.Cells(r, c).Interior.Color = _rgb(CARD)
        for r in range(7, 12):
            ws.Cells(r, 1).Interior.Color = _rgb(GOLD)
        ws.Columns(1).ColumnWidth = 1.2

        ws.Range("B7").Value = "SEARCH BY  ·  Packaging Set / Configuration ID"
        ws.Range("B7").Font.Name = FONT
        ws.Range("B7").Font.Size = 9
        ws.Range("B7").Font.Bold = True
        ws.Range("B7").Font.Color = _rgb(MUTED)
        ws.Rows(7).RowHeight = 20

        try:
            ws.Range("C8:G9").UnMerge()
        except Exception:
            pass
        ws.Range("C8:G9").Merge()
        inp = ws.Range("C8")
        inp.Value = ""
        inp.Interior.Color = _rgb(WHITE)
        inp.Font.Name = FONT
        inp.Font.Size = 16
        inp.Font.Bold = True
        inp.Font.Color = _rgb(NAVY)
        inp.HorizontalAlignment = -4131
        inp.VerticalAlignment = -4108
        for edge in (7, 8, 9, 10):
            try:
                b = ws.Range("C8:G9").Borders(edge)
                b.LineStyle = 1
                b.Weight = 3
                b.Color = _rgb(GOLD)
            except Exception:
                pass
        ws.Rows(8).RowHeight = 28
        ws.Rows(9).RowHeight = 28
        ws.Range("C8:G9").Locked = False

        ws.Range("B8").Value = "⌕"
        ws.Range("B8").Font.Name = FONT
        ws.Range("B8").Font.Size = 18
        ws.Range("B8").Font.Color = _rgb(GOLD)
        ws.Range("B8").HorizontalAlignment = -4108
        ws.Range("B8").VerticalAlignment = -4108

        ws.Range("B10").Value = (
            "Results update automatically as you type   ·   "
            "Examples:  ST-051-STD-01   ·   CNT-20-STD-01   ·   IND-24V-01"
        )
        ws.Range("B10").Font.Name = FONT
        ws.Range("B10").Font.Size = 8
        ws.Range("B10").Font.Color = _rgb(MUTED)
        ws.Rows(10).RowHeight = 18
        ws.Rows(11).RowHeight = 10

        for r in range(12, 20):
            for c in range(1, 15):
                ws.Cells(r, c).Interior.Color = _rgb(CARD)
        for r in range(12, 20):
            ws.Cells(r, 1).Interior.Color = _rgb(GOLD)

        ws.Range("B12").Value = "RESULTS"
        ws.Range("B12").Font.Name = FONT
        ws.Range("B12").Font.Size = 9
        ws.Range("B12").Font.Bold = True
        ws.Range("B12").Font.Color = _rgb(NAVY)

        ws.Range("B13").Formula = (
            '=IF($C$8="","Enter a Packaging Set or Configuration ID to view controlled data.","")'
        )
        ws.Range("B13").Font.Name = FONT
        ws.Range("B13").Font.Size = 9
        ws.Range("B13").Font.Italic = True
        ws.Range("B13").Font.Color = _rgb(MUTED)

        fields = [
            (14, "Configuration ID", "PACKAGING_CONFIGURATIONS!B:B"),
            (15, "Family", "PACKAGING_CONFIGURATIONS!D:D"),
            (16, "Source Configuration ID", "PACKAGING_CONFIGURATIONS!C:C"),
            (17, "Technical File ID", "PACKAGING_CONFIGURATIONS!K:K"),
            (18, "Packaging Tare kg", "PACKAGING_CONFIGURATIONS!H:H"),
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
            ws.Range(f"D{r}").Interior.Color = _rgb(PALE)
            ws.Range(f"B{r}").Locked = True
            ws.Range(f"D{r}").Locked = True

        ws.Range("B19").Value = "Document Pack Status"
        ws.Range("D19").Value = "988 / 988 LINKED  ·  Rev.00"
        ws.Range("B19").Font.Name = FONT
        ws.Range("D19").Font.Name = FONT
        ws.Range("D19").Font.Bold = True
        ws.Range("D19").Font.Color = _rgb(OK_FG)
        ws.Range("D19").Interior.Color = _rgb(OK_BG)

        # Shortcut tiles well below the cell results area
        y = 420
        for i, (label, target) in enumerate(
            [
                ("Document Center", "DOCUMENT_CENTER"),
                ("Packaging Configurations", "PACKAGING_CONFIGURATIONS"),
                ("Technical Files", "TECHNICAL_FILES"),
                ("Product Master", "PRODUCT_MASTER"),
            ]
        ):
            left = 14 + i * 206
            c = self._round(ws, f"SQ_{target}", left, y, 196, 44, CARD, line=PALE)
            self._rect(ws, f"SQG_{target}", left, y + 10, 3.5, 24, GOLD)
            tt = self._rect(ws, f"SQT_{target}", left + 14, y + 12, 170, 22, CARD)
            tt.Line.Visible = MSO_FALSE
            self._text(tt, f"{label}  →", 9, True, NAVY, "left")
            self._link(ws, c, target)

        self._lock_all(ws)
        ws.Range("C8:G9").Locked = False
        self._clear_shapes_over_range(ws, "C8")
        ws.Range("C8").Select()
        return removed

    def _clear_shapes_over_range(self, ws, addr: str) -> int:
        cell = ws.Range(addr)
        cl, ct = float(cell.Left), float(cell.Top)
        cw, ch = float(cell.Width), float(cell.Height)
        cx, cy = cl + cw / 2, ct + ch / 2
        removed = 0
        for i in range(int(ws.Shapes.Count), 0, -1):
            shp = ws.Shapes(i)
            name = str(shp.Name)
            if name.startswith(("Home", "TopNav", "Pill", "Hero", "SQ_", "Inci")):
                continue
            try:
                sl, st = float(shp.Left), float(shp.Top)
                sw, sh = float(shp.Width), float(shp.Height)
            except Exception:
                continue
            if not (sl <= cx <= sl + sw and st <= cy <= st + sh):
                continue
            try:
                if float(shp.Fill.Transparency) >= 0.95:
                    continue
            except Exception:
                pass
            try:
                shp.Delete()
                removed += 1
            except Exception:
                pass
        return removed
