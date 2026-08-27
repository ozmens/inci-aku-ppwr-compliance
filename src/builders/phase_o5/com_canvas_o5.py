"""Phase O5 — complete front-end rebuild of HOME / NAVIGATION / SEARCH.

Must look unmistakably different from V5 (layout, cards, spacing, hierarchy).
"""

from __future__ import annotations

from pathlib import Path
from typing import Any

MSO_RECT = 1
MSO_ROUNDED = 5
MSO_TRUE = -1
MSO_FALSE = 0
XL_FREE_FLOATING = 3
MSO_GRADIENT_HORIZ = 1

NAVY = "0B2341"
NAVY2 = "123A63"
MIDNIGHT = "07182B"
GOLD = "C9A24A"
DARK_GOLD = "A9842E"
IVORY = "F5F1E8"
CARD = "FFFEFB"
PALE = "E8EEF4"
LINE = "D5DCE3"
OK_BG = "E5F0E7"
OK_FG = "325D3E"
INK = "1E2C3A"
MUTED = "6A7785"
WHITE = "FFFFFF"
FONT = "Tahoma"


def _rgb(h: str) -> int:
    h = h.lstrip("#")
    return int(h[4:6] + h[2:4] + h[0:2], 16)


class ClassAO5Canvas:
    """Blank-canvas Class A rebuild — not an incremental V5 restyle."""

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
        self.excel.CutCopyMode = False
        return ws

    def _clear_all(self, ws) -> None:
        for i in range(int(ws.Shapes.Count), 0, -1):
            try:
                ws.Shapes(i).Delete()
            except Exception:
                pass
        ws.Cells.Clear()
        try:
            ws.Cells.UnMerge()
        except Exception:
            pass

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
            sh.OffsetX = 0.8
            sh.OffsetY = 1.6
            sh.Transparency = 0.78
            try:
                sh.Blur = 6
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

    def _navy_grad(self, shp) -> None:
        try:
            shp.Fill.TwoColorGradient(MSO_GRADIENT_HORIZ, 1)
            shp.Fill.ForeColor.RGB = _rgb(MIDNIGHT)
            shp.Fill.BackColor.RGB = _rgb(NAVY)
        except Exception:
            shp.Fill.Solid()
            shp.Fill.ForeColor.RGB = _rgb(NAVY)

    def _canvas(self, ws, zoom=90, cols=16, rows=40):
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
        for c in range(1, cols + 1):
            ws.Columns(c).ColumnWidth = 8.8
        end = f"{chr(64 + min(cols, 26))}{rows}" if cols <= 26 else f"P{rows}"
        # P = col 16
        rng = ws.Range(f"A1:P{rows}")
        rng.Interior.Color = _rgb(IVORY)
        rng.Font.Name = FONT
        rng.Borders.LineStyle = 0

    def protect_ui_sheet(self, name: str, unlock: list[str] | None = None) -> dict[str, Any]:
        ws = self._focus(name)
        locked = self._lock_all(ws)
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
            mode = "DrawingObjects+Contents"
        except Exception as exc:
            return {"sheet": name, "ok": False, "error": str(exc), "locked": locked}
        return {"sheet": name, "ok": True, "mode": mode, "locked_shapes": locked}

    def _clear_shapes_over(self, ws, addr: str) -> int:
        cell = ws.Range(addr)
        cx = float(cell.Left) + float(cell.Width) / 2
        cy = float(cell.Top) + float(cell.Height) / 2
        removed = 0
        for i in range(int(ws.Shapes.Count), 0, -1):
            shp = ws.Shapes(i)
            name = str(shp.Name)
            if name.startswith(("NavBar", "NavPill", "Hero", "Inci", "SQ_", "SearchBtn")):
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

    # ═══════════════════════════════════════════════════════════
    # HOME — new hierarchy
    # ═══════════════════════════════════════════════════════════
    def design_home(self) -> None:
        ws = self._focus("00_HOME")
        self._clear_all(ws)
        self._canvas(ws, 90, 16, 42)

        # --- ROW 1: full-width application nav bar ---
        bar = self._rect(ws, "NavBar", 10, 8, 860, 28, NAVY)
        self._navy_grad(bar)
        # left pills
        for left, label, target in (
            (18, "HOME", "00_HOME"),
            (88, "NAVIGATION", "NAVIGATION"),
            (188, "SEARCH", "SEARCH"),
            (268, "DOCUMENTS", "DOCUMENT_CENTER"),
        ):
            p = self._round(ws, f"NavPill_{target}", left, 12, 72 if label != "NAVIGATION" else 92, 20, WHITE, shadow=False)
            if label == "HOME":
                p.Fill.ForeColor.RGB = _rgb(GOLD)
                self._text(p, label, 8, True, MIDNIGHT, "center")
            else:
                self._text(p, label, 7.5, True, NAVY, "center")
            self._link(ws, p, target)
        # right status
        for left, label, fill, fg in (
            (560, "EXCEL VALIDATED", OK_BG, OK_FG),
            (690, "REV.00", GOLD, MIDNIGHT),
            (760, "CONTROLLED", PALE, NAVY),
        ):
            w = 120 if "EXCEL" in label else 60 if label == "REV.00" else 95
            p = self._round(ws, f"Status_{label[:6]}", left, 12, w, 20, fill, shadow=False)
            self._text(p, label, 7, True, fg, "center")

        # --- HERO (rows 2–6 visual band) ---
        hero = self._rect(ws, "Hero", 10, 44, 860, 100, NAVY)
        self._navy_grad(hero)
        brand = self._rect(ws, "HeroBrand", 28, 54, 280, 14, MIDNIGHT)
        brand.Line.Visible = MSO_FALSE
        self._text(brand, "İNCİ AKÜ PPWR", 9, True, GOLD, "left")
        t1 = self._rect(ws, "HeroT1", 28, 72, 520, 28, MIDNIGHT)
        t1.Line.Visible = MSO_FALSE
        self._text(t1, "Packaging Information", 22, True, WHITE, "left")
        t2 = self._rect(ws, "HeroT2", 28, 100, 520, 22, MIDNIGHT)
        t2.Line.Visible = MSO_FALSE
        self._text(t2, "Management System", 20, True, WHITE, "left")
        sub = self._rect(ws, "HeroSub", 28, 126, 520, 12, MIDNIGHT)
        sub.Line.Visible = MSO_FALSE
        self._text(
            sub,
            "Controlled Packaging Data  •  Compliance  •  Document Registry",
            8.5,
            False,
            "A9BBC9",
            "left",
        )
        if self.logo_path.exists():
            try:
                pic = ws.Shapes.AddPicture(
                    str(self.logo_path.resolve()), False, True, 720, 58, 130, 70
                )
                pic.Name = "InciAkuLogo"
                self._free(pic)
                self.shapes_created += 1
            except Exception:
                pass
        gold_line = self._rect(ws, "HeroGold", 10, 144, 860, 3.5, GOLD)
        try:
            gold_line.Fill.TwoColorGradient(MSO_GRADIENT_HORIZ, 1)
            gold_line.Fill.ForeColor.RGB = _rgb(DARK_GOLD)
            gold_line.Fill.BackColor.RGB = _rgb(GOLD)
        except Exception:
            pass

        # --- FOUR LARGE PRIMARY KPI CARDS ---
        primaries = [
            ("247", "PACKAGING\nCONFIGURATIONS"),
            ("2,046", "PRODUCTS"),
            ("988", "CONTROLLED\nDOCUMENTS"),
            ("0", "BLOCKING\nERRORS"),
        ]
        pw, ph, gap = 208, 118, 10
        for i, (val, lab) in enumerate(primaries):
            left = 10 + i * (pw + gap)
            top = 162
            card = self._round(ws, f"KPI_{i}", left, top, pw, ph, CARD, line=PALE)
            # top gold hairline accent
            self._rect(ws, f"KPI_g_{i}", left + 18, top + 16, 36, 3.5, GOLD)
            vv = self._rect(ws, f"KPI_v_{i}", left + 16, top + 28, pw - 32, 42, CARD)
            vv.Line.Visible = MSO_FALSE
            self._text(vv, val, 28, True, NAVY, "left")
            ll = self._rect(ws, f"KPI_l_{i}", left + 16, top + 74, pw - 32, 34, CARD)
            ll.Line.Visible = MSO_FALSE
            self._text(ll, lab, 9, True, MUTED, "left")

        # --- SECONDARY METRICS — one slim bar of chips (NOT six cards) ---
        strip = self._round(ws, "SecStrip", 10, 296, 860, 42, CARD, line=PALE)
        self._no_shadow(strip)
        self._shadow(strip)
        chips = [
            ("112", "Components"),
            ("1,690", "BOM Lines"),
            ("247", "Technical Files"),
            ("247", "EU DoCs"),
            ("247", "Labels"),
            ("247", "Statements"),
        ]
        for i, (val, lab) in enumerate(chips):
            left = 24 + i * 140
            vv = self._rect(ws, f"ChipV_{i}", left, 302, 50, 16, CARD)
            vv.Line.Visible = MSO_FALSE
            self._text(vv, val, 11, True, NAVY, "left")
            ll = self._rect(ws, f"ChipL_{i}", left + 48, 302, 85, 16, CARD)
            ll.Line.Visible = MSO_FALSE
            self._text(ll, lab, 8, False, MUTED, "left")
            if i < 5:
                div = self._rect(ws, f"ChipD_{i}", left + 128, 306, 1, 22, LINE)
                div.Line.Visible = MSO_FALSE

        # --- THREE OPERATING PANELS ---
        panel_top = 354
        ph = 156
        # LEFT System Health
        self._round(ws, "PanelHealth", 10, panel_top, 280, ph, CARD, line=PALE)
        h = self._rect(ws, "PH_title", 24, panel_top + 12, 240, 14, CARD)
        h.Line.Visible = MSO_FALSE
        self._text(h, "SYSTEM HEALTH", 9, True, NAVY, "left")
        self._rect(ws, "PH_gold", 24, panel_top + 28, 28, 2.5, GOLD)
        for i, (k, v) in enumerate(
            [
                ("Master Data", "READY"),
                ("Golden Register", "247/247"),
                ("Document Registry", "LINKED"),
                ("Excel Validation", "PASS"),
                ("Blocking Errors", "0"),
            ]
        ):
            y = panel_top + 42 + i * 20
            rk = self._rect(ws, f"PH_k_{i}", 24, y, 150, 16, CARD)
            rk.Line.Visible = MSO_FALSE
            self._text(rk, k, 8, False, MUTED, "left")
            rv = self._round(ws, f"PH_v_{i}", 180, y, 84, 16, OK_BG, shadow=False)
            self._text(rv, v, 7.5, True, OK_FG, "center")

        # CENTER Portfolio — row bars (not stacked single bar)
        self._round(ws, "PanelPort", 300, panel_top, 280, ph, CARD, line=PALE)
        h = self._rect(ws, "PP_title", 314, panel_top + 12, 240, 14, CARD)
        h.Line.Visible = MSO_FALSE
        self._text(h, "PACKAGING PORTFOLIO", 9, True, NAVY, "left")
        self._rect(ws, "PP_gold", 314, panel_top + 28, 28, 2.5, GOLD)
        for i, (lab, n, color, frac) in enumerate(
            [
                ("Starter", 240, NAVY, 240 / 247),
                ("Industrial", 3, NAVY2, 0.35),
                ("Container", 4, GOLD, 0.40),
            ]
        ):
            y = panel_top + 48 + i * 32
            rk = self._rect(ws, f"PP_k_{i}", 314, y, 90, 14, CARD)
            rk.Line.Visible = MSO_FALSE
            self._text(rk, lab, 8.5, False, INK, "left")
            rn = self._rect(ws, f"PP_n_{i}", 500, y, 50, 14, CARD)
            rn.Line.Visible = MSO_FALSE
            self._text(rn, str(n), 9, True, NAVY, "right")
            self._round(ws, f"PP_tr_{i}", 314, y + 16, 230, 8, PALE, shadow=False)
            self._round(ws, f"PP_br_{i}", 314, y + 16, max(230 * frac, 8), 8, color, shadow=False)

        # RIGHT Completeness
        self._round(ws, "PanelComp", 590, panel_top, 280, ph, CARD, line=PALE)
        h = self._rect(ws, "PC_title", 604, panel_top + 12, 240, 14, CARD)
        h.Line.Visible = MSO_FALSE
        self._text(h, "DOCUMENT COMPLETENESS", 9, True, NAVY, "left")
        self._rect(ws, "PC_gold", 604, panel_top + 28, 28, 2.5, GOLD)
        for i, lab in enumerate(["Technical Files", "EU DoCs", "Labels", "Statements"]):
            y = panel_top + 44 + i * 26
            rk = self._rect(ws, f"PC_k_{i}", 604, y, 120, 12, CARD)
            rk.Line.Visible = MSO_FALSE
            self._text(rk, lab, 8, False, INK, "left")
            rp = self._rect(ws, f"PC_p_{i}", 780, y, 60, 12, CARD)
            rp.Line.Visible = MSO_FALSE
            self._text(rp, "100%", 8, True, NAVY, "right")
            self._round(ws, f"PC_tr_{i}", 604, y + 13, 230, 7, PALE, shadow=False)
            self._round(ws, f"PC_br_{i}", 604, y + 13, 230, 7, GOLD, shadow=False)

        # --- QUICK ACTION LAUNCHPAD 5×2 ---
        qa_top = 528
        qh = self._rect(ws, "QA_title", 10, qa_top, 300, 14, IVORY)
        qh.Line.Visible = MSO_FALSE
        self._text(qh, "QUICK ACTION LAUNCHPAD", 9, True, NAVY, "left")
        actions = [
            ("Document Center", "988 linked controlled documents", "DOCUMENT_CENTER"),
            ("Packaging Configurations", "247 controlled configurations", "PACKAGING_CONFIGURATIONS"),
            ("Technical Files", "247 controlled Rev.00 files", "TECHNICAL_FILES"),
            ("EU Declarations", "247 EU declarations", "DECLARATIONS_OF_CONFORMITY"),
            ("Labels", "247 packaging labels", "LABELS"),
            ("Shipment Statements", "247 shipment statements", "SHIPMENT_STATEMENTS"),
            ("Product Master", "2,046 products", "PRODUCT_MASTER"),
            ("Component Master", "112 packaging components", "COMPONENT_MASTER"),
            ("Global Search", "Packaging set inquiry", "SEARCH"),
            ("Navigation", "Module launchpad", "NAVIGATION"),
        ]
        aw, ah, agx, agy = 168, 52, 8, 8
        for i, (title, desc, target) in enumerate(actions):
            r, c = divmod(i, 5)
            left = 10 + c * (aw + agx)
            top = qa_top + 20 + r * (ah + agy)
            card = self._round(ws, f"Act_{target}", left, top, aw, ah, CARD, line=PALE)
            self._rect(ws, f"ActG_{target}", left + 8, top + 10, 3, ah - 20, GOLD)
            tt = self._rect(ws, f"ActT_{target}", left + 18, top + 8, aw - 40, 16, CARD)
            tt.Line.Visible = MSO_FALSE
            self._text(tt, title, 8.5, True, NAVY, "left")
            dd = self._rect(ws, f"ActD_{target}", left + 18, top + 26, aw - 50, 18, CARD)
            dd.Line.Visible = MSO_FALSE
            self._text(dd, desc, 7, False, MUTED, "left")
            ar = self._rect(ws, f"ActA_{target}", left + aw - 28, top + 16, 18, 18, CARD)
            ar.Line.Visible = MSO_FALSE
            self._text(ar, "→", 12, True, GOLD, "center")
            self._link(ws, card, target)

        foot = self._rect(ws, "Footer", 10, qa_top + 20 + 2 * (ah + agy) + 4, 860, 16, IVORY)
        foot.Line.Visible = MSO_FALSE
        self._text(
            foot,
            "İnci Akü Sanayi ve Ticaret A.Ş.  ·  PPWR PIMS  ·  Rev.00 Controlled Baseline",
            7.5,
            False,
            MUTED,
            "left",
        )
        self._lock_all(ws)
        ws.Range("A1").Select()

    # ═══════════════════════════════════════════════════════════
    # NAVIGATION — launchpad rebuild
    # ═══════════════════════════════════════════════════════════
    def design_navigation(self) -> None:
        ws = self._focus("NAVIGATION")
        self._clear_all(ws)
        self._canvas(ws, 90, 16, 36)

        bar = self._rect(ws, "NavBar", 10, 8, 860, 28, NAVY)
        self._navy_grad(bar)
        for left, label, target in (
            (18, "HOME", "00_HOME"),
            (88, "NAVIGATION", "NAVIGATION"),
            (188, "SEARCH", "SEARCH"),
            (268, "DOCUMENTS", "DOCUMENT_CENTER"),
        ):
            fill = GOLD if label == "NAVIGATION" else WHITE
            fg = MIDNIGHT if label == "NAVIGATION" else NAVY
            p = self._round(
                ws,
                f"NavPill_{target}",
                left,
                12,
                72 if label != "NAVIGATION" else 92,
                20,
                fill,
                shadow=False,
            )
            self._text(p, label, 7.5, True, fg, "center")
            self._link(ws, p, target)

        hero = self._rect(ws, "Hero", 10, 44, 860, 72, NAVY)
        self._navy_grad(hero)
        t = self._rect(ws, "HeroT", 28, 54, 500, 28, MIDNIGHT)
        t.Line.Visible = MSO_FALSE
        self._text(t, "NAVIGATION", 22, True, WHITE, "left")
        s = self._rect(ws, "HeroS", 28, 88, 560, 14, MIDNIGHT)
        s.Line.Visible = MSO_FALSE
        self._text(s, "Access all controlled PPWR management modules", 9, False, "A9BBC9", "left")
        self._rect(ws, "HeroGold", 10, 116, 860, 3.5, GOLD)
        if self.logo_path.exists():
            try:
                pic = ws.Shapes.AddPicture(
                    str(self.logo_path.resolve()), False, True, 730, 52, 120, 52
                )
                pic.Name = "InciAkuLogo_Nav"
                self._free(pic)
                self.shapes_created += 1
            except Exception:
                pass

        # Dense 3-column launchpad filling canvas
        sections = [
            (
                "WORKSPACE",
                [
                    ("HOME", "Executive cockpit", "00_HOME", "Cockpit"),
                    ("Global Search", "Packaging set inquiry", "SEARCH", "Lookup"),
                ],
            ),
            (
                "MASTER DATA",
                [
                    ("Packaging Configurations", "247 controlled configurations", "PACKAGING_CONFIGURATIONS", "247"),
                    ("Product Master", "2,046 products", "PRODUCT_MASTER", "2,046"),
                    ("Component Master", "112 packaging components", "COMPONENT_MASTER", "112"),
                ],
            ),
            (
                "DOCUMENT CONTROL",
                [
                    ("Document Center", "988 linked documents", "DOCUMENT_CENTER", "988"),
                    ("Technical Files", "247 controlled Rev.00 files", "TECHNICAL_FILES", "247"),
                    ("EU Declarations", "247 EU declarations", "DECLARATIONS_OF_CONFORMITY", "247"),
                    ("Labels", "247 labels", "LABELS", "247"),
                    ("Shipment Statements", "247 statements", "SHIPMENT_STATEMENTS", "247"),
                ],
            ),
            (
                "OPERATIONS & SYSTEM",
                [
                    ("Shipments", "Transactional register", "SHIPMENTS", "Empty"),
                    ("Document Engine Map", "Read-only mapping", "DOC_ENGINE_MAP", "Map"),
                ],
            ),
        ]
        # Place as flowing 3-column grid
        y = 136
        col_w, card_h, gap_x, gap_y = 278, 62, 12, 10
        for sec_title, cards in sections:
            sh = self._rect(ws, f"Sec_{sec_title[:8]}", 10, y, 860, 16, IVORY)
            sh.Line.Visible = MSO_FALSE
            self._text(sh, sec_title, 9, True, NAVY, "left")
            y += 22
            for i, (title, desc, target, meta) in enumerate(cards):
                col = i % 3
                if i and col == 0:
                    y += card_h + gap_y
                left = 10 + col * (col_w + gap_x)
                card = self._round(ws, f"NavCard_{target}", left, y, col_w, card_h, CARD, line=PALE)
                self._rect(ws, f"NavG_{target}", left, y + 10, 4, card_h - 20, GOLD)
                tt = self._rect(ws, f"NavT_{target}", left + 16, y + 8, col_w - 70, 18, CARD)
                tt.Line.Visible = MSO_FALSE
                self._text(tt, title, 9.5, True, NAVY, "left")
                dd = self._rect(ws, f"NavD_{target}", left + 16, y + 28, col_w - 80, 26, CARD)
                dd.Line.Visible = MSO_FALSE
                self._text(dd, f"{desc}\n{meta}  ·  Open →", 7.5, False, MUTED, "left")
                self._link(ws, card, target)
            y += card_h + 18

        self._lock_all(ws)
        ws.Range("A1").Select()

    # ═══════════════════════════════════════════════════════════
    # SEARCH — complete rebuild, obvious input
    # ═══════════════════════════════════════════════════════════
    def design_search(self) -> None:
        ws = self._focus("SEARCH")
        self._clear_all(ws)
        self._canvas(ws, 90, 16, 36)

        bar = self._rect(ws, "NavBar", 10, 8, 860, 28, NAVY)
        self._navy_grad(bar)
        for left, label, target in (
            (18, "HOME", "00_HOME"),
            (88, "NAVIGATION", "NAVIGATION"),
            (188, "SEARCH", "SEARCH"),
            (268, "DOCUMENTS", "DOCUMENT_CENTER"),
        ):
            fill = GOLD if label == "SEARCH" else WHITE
            fg = MIDNIGHT if label == "SEARCH" else NAVY
            p = self._round(
                ws,
                f"NavPill_{target}",
                left,
                12,
                72 if label != "NAVIGATION" else 92,
                20,
                fill,
                shadow=False,
            )
            self._text(p, label, 7.5, True, fg, "center")
            self._link(ws, p, target)

        hero = self._rect(ws, "Hero", 10, 44, 860, 68, NAVY)
        self._navy_grad(hero)
        t = self._rect(ws, "HeroT", 28, 52, 600, 26, MIDNIGHT)
        t.Line.Visible = MSO_FALSE
        self._text(t, "GLOBAL SEARCH", 20, True, WHITE, "left")
        s = self._rect(ws, "HeroS", 28, 82, 620, 14, MIDNIGHT)
        s.Line.Visible = MSO_FALSE
        self._text(
            s,
            "Find a packaging set, configuration, source BOM or document.",
            9,
            False,
            "A9BBC9",
            "left",
        )
        self._rect(ws, "HeroGold", 10, 112, 860, 3.5, GOLD)
        if self.logo_path.exists():
            try:
                pic = ws.Shapes.AddPicture(
                    str(self.logo_path.resolve()), False, True, 740, 52, 110, 48
                )
                pic.Name = "InciAkuLogo_Search"
                self._free(pic)
                self.shapes_created += 1
            except Exception:
                pass

        # --- SEARCH CARD: cell surface only for input zone ---
        # Paint card background with cells (rows 7–12), no opaque shape over C8:H9
        for r in range(7, 13):
            for c in range(2, 15):
                ws.Cells(r, c).Interior.Color = _rgb(CARD)
        for r in range(7, 13):
            ws.Cells(r, 2).Interior.Color = _rgb(GOLD)
            ws.Columns(2).ColumnWidth = 1.4

        ws.Range("C7").Value = "SEARCH TERM"
        ws.Range("C7").Font.Name = FONT
        ws.Range("C7").Font.Size = 9
        ws.Range("C7").Font.Bold = True
        ws.Range("C7").Font.Color = _rgb(MUTED)
        ws.Rows(7).RowHeight = 18

        # REAL INPUT C8:H9 — must remain uncovered
        try:
            ws.Range("C8:H9").UnMerge()
        except Exception:
            pass
        ws.Range("C8:H9").Merge()
        inp = ws.Range("C8")
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
                b = ws.Range("C8:H9").Borders(edge)
                b.LineStyle = 1
                b.Weight = 3  # ~2pt visual weight
                b.Color = _rgb(GOLD)
            except Exception:
                pass
        ws.Rows(8).RowHeight = 26
        ws.Rows(9).RowHeight = 26
        ws.Range("C8:H9").Locked = False

        # Visual SEARCH button (shape) to the right of input — does NOT cover C8:H9
        # Place by points after rows exist
        btn = self._round(ws, "SearchBtn", 620, 148, 110, 40, NAVY, shadow=True)
        self._text(btn, "SEARCH", 11, True, WHITE, "center")
        # no hyperlink — visual affordance only

        ws.Range("C10").Value = "Results update automatically after entering a valid ID."
        ws.Range("C10").Font.Name = FONT
        ws.Range("C10").Font.Size = 8
        ws.Range("C10").Font.Italic = True
        ws.Range("C10").Font.Color = _rgb(MUTED)

        ws.Range("C11").Value = "Examples:   ST-051-STD-01     CNT-20-STD-01     IND-24V-01"
        ws.Range("C11").Font.Name = FONT
        ws.Range("C11").Font.Size = 8
        ws.Range("C11").Font.Color = _rgb(MUTED)
        ws.Rows(12).RowHeight = 8

        # --- RESULT CARD (compact, cell-based) ---
        for r in range(13, 25):
            for c in range(2, 15):
                ws.Cells(r, c).Interior.Color = _rgb(CARD)
        for r in range(13, 25):
            ws.Cells(r, 2).Interior.Color = _rgb(GOLD)

        ws.Range("C13").Value = "RESULTS"
        ws.Range("C13").Font.Name = FONT
        ws.Range("C13").Font.Size = 9
        ws.Range("C13").Font.Bold = True
        ws.Range("C13").Font.Color = _rgb(NAVY)

        ws.Range("C14").Formula = (
            '=IF($C$8="","Enter a Packaging Set or Configuration ID above'
            ' to view controlled packaging data.","")'
        )
        ws.Range("C14").Font.Name = FONT
        ws.Range("C14").Font.Size = 9
        ws.Range("C14").Font.Italic = True
        ws.Range("C14").Font.Color = _rgb(MUTED)

        # XLOOKUP fields — Packaging Set Code = A, Config=B, Family=D, Source=C, Tare=H,
        # TF=K, DoC=L, Label=M, Statement=N
        fields = [
            (15, "Packaging Set Code", 'IF($C$8="","",$C$8)'),
            (16, "Configuration ID", "PACKAGING_CONFIGURATIONS!B:B"),
            (17, "Family", "PACKAGING_CONFIGURATIONS!D:D"),
            (18, "Source Configuration ID", "PACKAGING_CONFIGURATIONS!C:C"),
            (19, "Packaging Tare", "PACKAGING_CONFIGURATIONS!H:H"),
            (20, "Technical File ID", "PACKAGING_CONFIGURATIONS!K:K"),
            (21, "EU DoC ID", "PACKAGING_CONFIGURATIONS!L:L"),
            (22, "Label ID", "PACKAGING_CONFIGURATIONS!M:M"),
            (23, "Statement ID", "PACKAGING_CONFIGURATIONS!N:N"),
        ]
        for r, label, colref in fields:
            ws.Range(f"C{r}").Value = label
            ws.Range(f"C{r}").Font.Name = FONT
            ws.Range(f"C{r}").Font.Size = 8.5
            ws.Range(f"C{r}").Font.Color = _rgb(MUTED)
            if colref.startswith("IF("):
                ws.Range(f"E{r}").Formula = f"={colref}"
            else:
                ws.Range(f"E{r}").Formula = (
                    f'=IF($C$8="","",IFERROR(XLOOKUP($C$8,PACKAGING_CONFIGURATIONS!A:A,{colref}),'
                    f'"Not found"))'
                )
            ws.Range(f"E{r}").Font.Name = FONT
            ws.Range(f"E{r}").Font.Size = 9.5
            ws.Range(f"E{r}").Font.Color = _rgb(INK)
            ws.Range(f"E{r}").Interior.Color = _rgb(PALE)
            ws.Range(f"C{r}").Locked = True
            ws.Range(f"E{r}").Locked = True

        ws.Range("C24").Value = "Document Pack Status"
        ws.Range("E24").Value = "988 / 988 LINKED  ·  Rev.00"
        ws.Range("C24").Font.Name = FONT
        ws.Range("E24").Font.Name = FONT
        ws.Range("E24").Font.Bold = True
        ws.Range("E24").Font.Color = _rgb(OK_FG)
        ws.Range("E24").Interior.Color = _rgb(OK_BG)

        # Quick links — 4 tiles
        y = 455
        for i, (label, target) in enumerate(
            [
                ("Document Center", "DOCUMENT_CENTER"),
                ("Packaging Configurations", "PACKAGING_CONFIGURATIONS"),
                ("Technical Files", "TECHNICAL_FILES"),
                ("Product Master", "PRODUCT_MASTER"),
            ]
        ):
            left = 10 + i * 216
            c = self._round(ws, f"SQ_{target}", left, y, 206, 44, CARD, line=PALE)
            self._rect(ws, f"SQG_{target}", left, y + 10, 4, 24, GOLD)
            tt = self._rect(ws, f"SQT_{target}", left + 14, y + 12, 170, 22, CARD)
            tt.Line.Visible = MSO_FALSE
            self._text(tt, f"{label}  →", 9, True, NAVY, "left")
            self._link(ws, c, target)

        self._lock_all(ws)
        ws.Range("C8:H9").Locked = False
        self._clear_shapes_over(ws, "C8")
        # Reposition SearchBtn if it landed over input — force known safe coords
        try:
            btn = ws.Shapes("SearchBtn")
            btn.Left = 720
            btn.Top = 148
        except Exception:
            pass
        # Final cover check again after move
        self._clear_shapes_over(ws, "C8")
        # Recreate button if deleted
        try:
            _ = ws.Shapes("SearchBtn")
        except Exception:
            btn = self._round(ws, "SearchBtn", 720, 148, 110, 40, NAVY, shadow=True)
            self._text(btn, "SEARCH", 11, True, WHITE, "center")
            btn.Locked = True
        ws.Range("C8").Select()
