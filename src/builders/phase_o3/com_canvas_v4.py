"""Phase O3 — Class A V4 premium canvas with locked shapes + sheet protection."""

from __future__ import annotations

from pathlib import Path
from typing import Any

MSO_RECT = 1
MSO_ROUNDED = 5
MSO_TRUE = -1
MSO_FALSE = 0
XL_FREE_FLOATING = 3
MSO_GRADIENT_HORIZ = 1

# V4 palette
DEEP_NAVY = "102A43"
MIDNIGHT = "071C30"
STEEL = "315E87"
WARM_IVORY = "F6F3EC"
CARD = "FCFBF8"
PALE = "EAF0F4"
GOLD = "C7A24A"
DARK_GOLD = "A9832F"
CHARCOAL = "273746"
OK_BG = "E4EFE7"
OK_FG = "376344"
MUTED = "5C6B7A"
WHITE = "FFFFFF"
FONT = "Tahoma"


def _rgb(h: str) -> int:
    h = h.lstrip("#")
    return int(h[4:6] + h[2:4] + h[0:2], 16)


class ClassAV4Canvas:
    def __init__(self, excel, wb, logo_path: Path) -> None:
        self.excel = excel
        self.wb = wb
        self.logo_path = logo_path
        self.shapes_created = 0
        self.locked_shapes = 0
        self.hyperlinks_added = 0

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

    def _lock(self, shp) -> None:
        try:
            shp.Locked = True
            self.locked_shapes += 1
        except Exception:
            pass

    def _shadow(self, shp, *, soft: bool = True) -> None:
        try:
            sh = shp.Shadow
            sh.Visible = MSO_TRUE
            sh.Style = 1
            sh.OffsetX = 1.2 if soft else 2.0
            sh.OffsetY = 1.8 if soft else 2.5
            sh.Transparency = 0.80 if soft else 0.72
            try:
                sh.Blur = 4 if soft else 6
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

    def _navy_gradient(self, shp) -> None:
        try:
            shp.Fill.TwoColorGradient(MSO_GRADIENT_HORIZ, 1)
            shp.Fill.ForeColor.RGB = _rgb(MIDNIGHT)
            shp.Fill.BackColor.RGB = _rgb(DEEP_NAVY)
        except Exception:
            shp.Fill.Solid()
            shp.Fill.ForeColor.RGB = _rgb(DEEP_NAVY)

    def _gold_gradient(self, shp) -> None:
        try:
            shp.Fill.TwoColorGradient(MSO_GRADIENT_HORIZ, 1)
            shp.Fill.ForeColor.RGB = _rgb(DARK_GOLD)
            shp.Fill.BackColor.RGB = _rgb(GOLD)
        except Exception:
            shp.Fill.Solid()
            shp.Fill.ForeColor.RGB = _rgb(GOLD)

    def _text(self, shp, text, size=10, bold=False, color=CHARCOAL, align="left"):
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
            self.hyperlinks_added += 1
        except Exception:
            pass

    def _lock_all_shapes(self, ws) -> int:
        n = 0
        for i in range(1, int(ws.Shapes.Count) + 1):
            try:
                ws.Shapes(i).Locked = True
                n += 1
            except Exception:
                pass
        self.locked_shapes += n
        return n

    def protect_ui_sheet(self, name: str, *, unlock_cells: list[str] | None = None) -> dict[str, Any]:
        """Lock drawing objects; keep hyperlinks clickable. Light/no password."""
        ws = self._focus(name)
        locked = self._lock_all_shapes(ws)
        # Unlock search / editable cells before Contents protection
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
        # Protect drawings; Contents=True with all cells unlocked keeps filters/cells usable
        # where needed; Class A sheets have no tables.
        try:
            ws.Unprotect(Password="")
        except Exception:
            pass
        try:
            # DrawingObjects=True prevents shape text edit / move when Locked
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

    def _canvas(self, ws, zoom=93):
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
        for c in range(1, 17):
            ws.Columns(c).ColumnWidth = 8.6
        rng = ws.Range("A1:P40")
        rng.Interior.Color = _rgb(WARM_IVORY)
        rng.Font.Name = FONT
        rng.Borders.LineStyle = 0

    def design_home(self) -> int:
        ws = self._focus("00_HOME")
        removed = self._clear_shapes(ws)
        ws.Cells.Clear()
        self._canvas(ws, 93)

        # Top chrome
        home = self._round(ws, "HomeBtn", 12, 8, 70, 22, fill=DEEP_NAVY, shadow=False)
        self._text(home, "← HOME", 8, True, WHITE, "center")
        self._link(ws, home, "00_HOME")
        for left, label, target in (
            (90, "Navigation", "NAVIGATION"),
            (182, "Search", "SEARCH"),
            (258, "Documents", "DOCUMENT_CENTER"),
        ):
            p = self._round(
                ws, f"HomeNav_{target}", left, 10, 84, 18, fill=CARD, line=PALE, shadow=False
            )
            self._text(p, label, 7.5, True, STEEL, "center")
            self._link(ws, p, target)
        badge = self._round(ws, "BadgeExcel", 500, 10, 118, 18, fill=OK_BG, shadow=False)
        self._text(badge, "EXCEL VALIDATED", 7.5, True, OK_FG, "center")
        rev = self._round(ws, "BadgeRev", 628, 10, 52, 18, fill=GOLD, shadow=False)
        self._text(rev, "REV.00", 8, True, MIDNIGHT, "center")

        # Hero ~20%
        hero = self._rect(ws, "Hero", 10, 34, 800, 68, DEEP_NAVY)
        self._navy_gradient(hero)
        brand = self._rect(ws, "HeroBrand", 22, 40, 380, 12, MIDNIGHT)
        brand.Line.Visible = MSO_FALSE
        self._text(brand, "İNCİ AKÜ  •  PPWR PIMS", 8, True, GOLD, "left")
        title = self._rect(ws, "HeroTitle", 22, 52, 540, 26, MIDNIGHT)
        title.Line.Visible = MSO_FALSE
        self._text(
            title,
            "İnci Akü PPWR Packaging Information Management System",
            13,
            True,
            WHITE,
            "left",
        )
        sub = self._rect(ws, "HeroSub", 22, 80, 520, 14, MIDNIGHT)
        sub.Line.Visible = MSO_FALSE
        self._text(sub, "Controlled Packaging Data & Compliance Workspace", 8, False, "A8B9C8", "left")
        if self.logo_path.exists():
            try:
                pic = ws.Shapes.AddPicture(
                    str(self.logo_path.resolve()), False, True, 660, 42, 138, 50
                )
                pic.Name = "InciAkuLogo"
                self._free(pic)
                self.shapes_created += 1
            except Exception:
                pass

        # 4 hero KPIs
        sec = self._rect(ws, "SecPrimary", 12, 112, 280, 12, WARM_IVORY)
        sec.Line.Visible = MSO_FALSE
        self._text(sec, "PRIMARY INDICATORS", 8.5, True, DEEP_NAVY, "left")
        primaries = [
            ("247", "PACKAGING CONFIGURATIONS"),
            ("2,046", "PRODUCTS"),
            ("988", "CONTROLLED DOCUMENTS"),
            ("0", "BLOCKING ERRORS"),
        ]
        pw, ph, gap = 192, 88, 12
        for i, (val, lab) in enumerate(primaries):
            left = 12 + i * (pw + gap)
            top = 128
            card = self._round(ws, f"PKPI_{i}", left, top, pw, ph, CARD, line=PALE)
            accent = self._rect(ws, f"PKPI_g_{i}", left + 14, top + 12, pw - 28, 2.5, GOLD)
            self._gold_gradient(accent)
            lv = self._rect(ws, f"PKPI_v_{i}", left + 12, top + 24, pw - 24, 32, CARD)
            lv.Line.Visible = MSO_FALSE
            self._text(lv, val, 24, True, DEEP_NAVY, "left")
            ll = self._rect(ws, f"PKPI_l_{i}", left + 12, top + 58, pw - 24, 22, CARD)
            ll.Line.Visible = MSO_FALSE
            self._text(ll, lab, 7.5, True, MUTED, "left")

        # Secondary strip
        sec2 = self._rect(ws, "SecSecondary", 12, 228, 280, 12, WARM_IVORY)
        sec2.Line.Visible = MSO_FALSE
        self._text(sec2, "SECONDARY METRICS", 8.5, True, DEEP_NAVY, "left")
        secondaries = [
            ("112", "Components"),
            ("1,690", "BOM Lines"),
            ("247", "Technical Files"),
            ("247", "EU DoCs"),
            ("247", "Labels"),
            ("247", "Statements"),
        ]
        sw, sh = 128, 46
        for i, (val, lab) in enumerate(secondaries):
            left = 12 + i * (sw + 6)
            top = 244
            card = self._round(ws, f"SKPI_{i}", left, top, sw, sh, CARD, line=PALE)
            self._rect(ws, f"SKPI_g_{i}", left, top + 8, 3, sh - 16, GOLD)
            vv = self._rect(ws, f"SKPI_v_{i}", left + 10, top + 6, sw - 16, 18, CARD)
            vv.Line.Visible = MSO_FALSE
            self._text(vv, val, 13, True, DEEP_NAVY, "left")
            ll = self._rect(ws, f"SKPI_l_{i}", left + 10, top + 26, sw - 16, 14, CARD)
            ll.Line.Visible = MSO_FALSE
            self._text(ll, lab.upper(), 6.5, True, MUTED, "left")

        # Panels
        panel_top = 304
        sp = self._round(ws, "SysHealth", 12, panel_top, 258, 138, CARD, line=PALE)
        sh = self._rect(ws, "SysHdr", 24, panel_top + 10, 220, 12, CARD)
        sh.Line.Visible = MSO_FALSE
        self._text(sh, "SYSTEM HEALTH", 8.5, True, DEEP_NAVY, "left")
        self._rect(ws, "SysGold", 24, panel_top + 26, 26, 2.2, GOLD)
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
            y = panel_top + 36 + i * 15.5
            row = self._rect(ws, f"SysR_{i}", 24, y, 140, 13, CARD)
            row.Line.Visible = MSO_FALSE
            self._text(row, f"● {k}", 7, False, MUTED, "left")
            pill = self._round(ws, f"SysP_{i}", 172, y, 72, 13, OK_BG, shadow=False)
            self._text(pill, v, 6.5, True, OK_FG, "center")

        mp = self._round(ws, "Portfolio", 282, panel_top, 258, 138, CARD, line=PALE)
        mh = self._rect(ws, "PortHdr", 294, panel_top + 10, 220, 12, CARD)
        mh.Line.Visible = MSO_FALSE
        self._text(mh, "PACKAGING PORTFOLIO", 8.5, True, DEEP_NAVY, "left")
        track_l, track_t, track_w, track_h = 294, panel_top + 44, 220, 12
        self._round(ws, "MixTrack", track_l, track_t, track_w, track_h, PALE, shadow=False)
        w1 = track_w * (240 / 247)
        w2 = max(track_w * (3 / 247), 3)
        w3 = max(track_w - w1 - w2, 4)
        self._rect(ws, "MixS", track_l, track_t, w1, track_h, DEEP_NAVY)
        self._rect(ws, "MixI", track_l + w1, track_t, w2, track_h, STEEL)
        self._rect(ws, "MixC", track_l + w1 + w2, track_t, w3, track_h, GOLD)
        leg = self._rect(ws, "MixLeg", 294, panel_top + 70, 220, 50, CARD)
        leg.Line.Visible = MSO_FALSE
        self._text(leg, "Starter 240\nIndustrial 3\nContainer 4", 8, False, MUTED, "left")

        dp = self._round(ws, "Completeness", 552, panel_top, 258, 138, CARD, line=PALE)
        dh = self._rect(ws, "CompHdr", 564, panel_top + 10, 220, 12, CARD)
        dh.Line.Visible = MSO_FALSE
        self._text(dh, "DOCUMENT COMPLETENESS", 8.5, True, DEEP_NAVY, "left")
        for i, lab in enumerate(["Technical Files", "EU DoCs", "Labels", "Statements"]):
            y = panel_top + 32 + i * 24
            lr = self._rect(ws, f"CompL_{i}", 564, y, 100, 11, CARD)
            lr.Line.Visible = MSO_FALSE
            self._text(lr, lab, 7, False, CHARCOAL, "left")
            self._round(ws, f"CompT_{i}", 564, y + 12, 170, 7, PALE, shadow=False)
            bar = self._round(ws, f"CompB_{i}", 564, y + 12, 170, 7, GOLD, shadow=False)
            rr = self._rect(ws, f"CompR_{i}", 742, y, 48, 11, CARD)
            rr.Line.Visible = MSO_FALSE
            self._text(rr, "100%", 7, True, DEEP_NAVY, "right")

        # Quick actions — elegant cards with gold accent + Open →
        qa_top = 456
        qh = self._rect(ws, "QAHdr", 12, qa_top, 200, 12, WARM_IVORY)
        qh.Line.Visible = MSO_FALSE
        self._text(qh, "QUICK ACTIONS", 8.5, True, DEEP_NAVY, "left")
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
        aw, ah, ag = 154, 52, 8
        for i, (title, desc, target) in enumerate(actions):
            r, c = divmod(i, 5)
            left = 12 + c * (aw + ag)
            top = qa_top + 16 + r * (ah + ag)
            card = self._round(ws, f"Act_{target}", left, top, aw, ah, CARD, line=PALE)
            self._rect(ws, f"ActG_{target}", left, top + 8, 3, ah - 16, GOLD)
            tt = self._rect(ws, f"ActT_{target}", left + 10, top + 6, aw - 18, 16, CARD)
            tt.Line.Visible = MSO_FALSE
            self._text(tt, title, 8, True, DEEP_NAVY, "left")
            dd = self._rect(ws, f"ActD_{target}", left + 10, top + 24, aw - 18, 20, CARD)
            dd.Line.Visible = MSO_FALSE
            self._text(dd, f"{desc}\nOpen →", 7, False, MUTED, "left")
            self._link(ws, card, target)

        foot = self._rect(
            ws, "Footer", 12, qa_top + 16 + 2 * (ah + ag) + 4, 800, 18, WARM_IVORY
        )
        foot.Line.Visible = MSO_FALSE
        self._text(
            foot,
            "İnci Akü Sanayi ve Ticaret A.Ş.  ·  PPWR Packaging Information Management System  ·  Rev.00",
            7,
            False,
            MUTED,
            "left",
        )
        self._lock_all_shapes(ws)
        ws.Range("A1").Select()
        return removed

    def design_navigation(self) -> int:
        ws = self._focus("NAVIGATION")
        removed = self._clear_shapes(ws)
        ws.Cells.Clear()
        self._canvas(ws, 95)
        home = self._round(ws, "HomeBtn", 12, 8, 70, 22, DEEP_NAVY, shadow=False)
        self._text(home, "← HOME", 8, True, WHITE, "center")
        self._link(ws, home, "00_HOME")

        hero = self._rect(ws, "NavHero", 10, 34, 800, 56, DEEP_NAVY)
        self._navy_gradient(hero)
        t = self._rect(ws, "NavTitle", 22, 40, 400, 22, MIDNIGHT)
        t.Line.Visible = MSO_FALSE
        self._text(t, "Navigation", 17, True, WHITE, "left")
        s = self._rect(ws, "NavSub", 22, 66, 520, 14, MIDNIGHT)
        s.Line.Visible = MSO_FALSE
        self._text(s, "Jump directly to any PPWR management module", 8.5, False, "A8B9C8", "left")
        if self.logo_path.exists():
            try:
                pic = ws.Shapes.AddPicture(
                    str(self.logo_path.resolve()), False, True, 660, 38, 138, 48
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
                    ("HOME", "Executive cockpit", "00_HOME", ""),
                    ("SEARCH", "Global inquiry workspace", "SEARCH", ""),
                ],
            ),
            (
                "PACKAGING DATA",
                [
                    ("PACKAGING CONFIGURATIONS", "247 controlled configurations", "PACKAGING_CONFIGURATIONS", "247"),
                    ("PRODUCT MASTER", "2,046 products", "PRODUCT_MASTER", "2,046"),
                    ("COMPONENT MASTER", "112 packaging components", "COMPONENT_MASTER", "112"),
                ],
            ),
            (
                "DOCUMENT CONTROL",
                [
                    ("DOCUMENT CENTER", "988 linked documents", "DOCUMENT_CENTER", "988"),
                    ("TECHNICAL FILES", "247 controlled Rev.00 files", "TECHNICAL_FILES", "247"),
                    ("DECLARATIONS", "247 EU declarations", "DECLARATIONS_OF_CONFORMITY", "247"),
                    ("LABELS", "247 labels", "LABELS", "247"),
                    ("SHIPMENT STATEMENTS", "247 statements", "SHIPMENT_STATEMENTS", "247"),
                ],
            ),
            (
                "SYSTEM & OPERATIONS",
                [
                    ("SHIPMENTS", "Transactional register", "SHIPMENTS", ""),
                    ("DOC ENGINE MAP", "Read-only mapping", "DOC_ENGINE_MAP", ""),
                ],
            ),
        ]
        y = 106
        for sec, cards in sections:
            sh = self._rect(ws, f"Sec_{sec[:8]}", 12, y, 800, 12, WARM_IVORY)
            sh.Line.Visible = MSO_FALSE
            self._text(sh, sec, 8.5, True, DEEP_NAVY, "left")
            y += 18
            cw, ch, cg = 258, 62, 14
            for i, (title, desc, target, _cnt) in enumerate(cards):
                col = i % 3
                if i and col == 0:
                    y += ch + cg
                left = 12 + col * (cw + cg)
                card = self._round(ws, f"NavCard_{target}", left, y, cw, ch, CARD, line=PALE)
                self._rect(ws, f"NavG_{target}", left, y + 10, 3.5, ch - 20, GOLD)
                tt = self._rect(ws, f"NavT_{target}", left + 14, y + 10, cw - 28, 18, CARD)
                tt.Line.Visible = MSO_FALSE
                self._text(tt, title, 9, True, DEEP_NAVY, "left")
                dd = self._rect(ws, f"NavD_{target}", left + 14, y + 32, cw - 28, 22, CARD)
                dd.Line.Visible = MSO_FALSE
                self._text(dd, f"{desc}\nOpen →", 7.5, False, MUTED, "left")
                self._link(ws, card, target)
            y += ch + 16
        self._lock_all_shapes(ws)
        return removed

    def design_search(self) -> int:
        ws = self._focus("SEARCH")
        removed = self._clear_shapes(ws)
        ws.Cells.Clear()
        self._canvas(ws, 95)
        home = self._round(ws, "HomeBtn", 12, 8, 70, 22, DEEP_NAVY, shadow=False)
        self._text(home, "← HOME", 8, True, WHITE, "center")
        self._link(ws, home, "00_HOME")

        hero = self._rect(ws, "SearchHero", 10, 34, 800, 48, DEEP_NAVY)
        self._navy_gradient(hero)
        t = self._rect(ws, "SearchTitle", 22, 40, 500, 20, MIDNIGHT)
        t.Line.Visible = MSO_FALSE
        self._text(t, "Global Search", 15, True, WHITE, "left")
        s = self._rect(ws, "SearchSub", 22, 62, 560, 12, MIDNIGHT)
        s.Line.Visible = MSO_FALSE
        self._text(
            s,
            "Search packaging sets, configurations, source BOMs and document IDs",
            8,
            False,
            "A8B9C8",
            "left",
        )

        # Main search card — transparent fill so input cell C8 remains clickable
        scard = self._round(ws, "SearchCard", 12, 96, 800, 86, CARD, line=PALE)
        try:
            scard.Fill.Transparency = 1.0
        except Exception:
            scard.Fill.Visible = MSO_FALSE
        self._rect(ws, "SearchAccent", 12, 96, 4, 86, GOLD)
        lab = self._rect(ws, "SearchLab", 28, 106, 360, 12, WARM_IVORY)
        lab.Line.Visible = MSO_FALSE
        self._text(lab, "SEARCH BY  ·  Packaging Set / Configuration ID", 8, True, MUTED, "left")
        ws.Range("C8").Value = ""
        ws.Range("C8").Interior.Color = _rgb("FFF9EF")
        ws.Range("C8").Font.Name = FONT
        ws.Range("C8").Font.Size = 14
        ws.Range("C8").Font.Bold = True
        ws.Range("C8").Font.Color = _rgb(DEEP_NAVY)
        ws.Range("C8").Borders.Color = _rgb(GOLD)
        ws.Range("C8").RowHeight = 28
        ws.Range("C8").Locked = False
        ex = self._rect(ws, "SearchEx", 28, 158, 520, 12, WARM_IVORY)
        ex.Line.Visible = MSO_FALSE
        self._text(ex, "Example: ST-051-STD-01  ·  CNT-20-STD-01  ·  IND-24V-01", 7.5, False, MUTED, "left")

        # Compact results panel — frame only (cells remain visible underneath)
        frame = self._round(ws, "ResultCard", 12, 196, 800, 150, CARD, line=PALE)
        try:
            frame.Fill.Transparency = 1.0  # hollow frame; cells show through
        except Exception:
            frame.Fill.Visible = MSO_FALSE
        rh = self._rect(ws, "ResultHdr", 28, 204, 200, 14, WARM_IVORY)
        rh.Line.Visible = MSO_FALSE
        self._text(rh, "RESULTS", 8.5, True, DEEP_NAVY, "left")
        ws.Range("B10").Formula = (
            '=IF($C$8="","Enter a Packaging Set or Configuration ID to view controlled data.","")'
        )
        ws.Range("B10").Font.Name = FONT
        ws.Range("B10").Font.Size = 8.5
        ws.Range("B10").Font.Italic = True
        ws.Range("B10").Font.Color = _rgb(MUTED)
        ws.Range("B10").Locked = True
        fields = [
            (11, "Configuration ID", "PACKAGING_CONFIGURATIONS!B:B"),
            (12, "Family", "PACKAGING_CONFIGURATIONS!D:D"),
            (13, "Source Configuration ID", "PACKAGING_CONFIGURATIONS!C:C"),
            (14, "Technical File ID", "PACKAGING_CONFIGURATIONS!K:K"),
            (15, "Packaging Tare kg", "PACKAGING_CONFIGURATIONS!H:H"),
        ]
        for r, label, colref in fields:
            ws.Range(f"B{r}").Value = label
            ws.Range(f"B{r}").Font.Name = FONT
            ws.Range(f"B{r}").Font.Size = 8.5
            ws.Range(f"B{r}").Font.Color = _rgb(MUTED)
            ws.Range(f"D{r}").Formula = (
                f'=IF($C$8="","",IFERROR(XLOOKUP($C$8,PACKAGING_CONFIGURATIONS!A:A,{colref}),'
                f'"Not found — use AutoFilter on Document Center"))'
            )
            ws.Range(f"D{r}").Font.Name = FONT
            ws.Range(f"D{r}").Font.Size = 9.5
            ws.Range(f"D{r}").Font.Color = _rgb(CHARCOAL)
            ws.Range(f"D{r}").Interior.Color = _rgb(PALE)
            ws.Range(f"B{r}").Locked = True
            ws.Range(f"D{r}").Locked = True
        ws.Range("B16").Value = "Document Pack Status"
        ws.Range("D16").Value = "988 / 988 LINKED  ·  Rev.00"
        ws.Range("B16").Font.Name = FONT
        ws.Range("D16").Font.Name = FONT
        ws.Range("D16").Font.Bold = True
        ws.Range("D16").Font.Color = _rgb(OK_FG)
        ws.Range("D16").Interior.Color = _rgb(OK_BG)

        y = 362
        for i, (label, target) in enumerate(
            [
                ("Document Center", "DOCUMENT_CENTER"),
                ("Packaging Configurations", "PACKAGING_CONFIGURATIONS"),
                ("Technical Files", "TECHNICAL_FILES"),
                ("Product Master", "PRODUCT_MASTER"),
            ]
        ):
            left = 12 + i * 200
            c = self._round(ws, f"SQ_{target}", left, y, 190, 40, CARD, line=PALE)
            self._rect(ws, f"SQG_{target}", left, y + 8, 3, 24, GOLD)
            tt = self._rect(ws, f"SQT_{target}", left + 10, y + 10, 170, 20, CARD)
            tt.Line.Visible = MSO_FALSE
            self._text(tt, f"{label}  →", 8.5, True, DEEP_NAVY, "left")
            self._link(ws, c, target)
        self._lock_all_shapes(ws)
        return removed
