"""Phase O7 — low-risk in-place visual polish of O6 front-end sheets.

Preserves shapes, hyperlinks, formulas, geometry. Style-only adjustments.
"""

from __future__ import annotations

from typing import Any

MSO_TRUE = -1
MSO_FALSE = 0
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

FRONT = ("00_HOME", "NAVIGATION", "SEARCH")


def _rgb(h: str) -> int:
    h = h.lstrip("#")
    return int(h[4:6] + h[2:4] + h[0:2], 16)


def _soft_shadow(shp) -> None:
    try:
        sh = shp.Shadow
        sh.Visible = MSO_TRUE
        sh.Style = 1
        sh.OffsetX = 0.5
        sh.OffsetY = 1.0
        sh.Transparency = 0.88
        try:
            sh.Blur = 3.5
        except Exception:
            pass
    except Exception:
        pass


def _no_shadow(shp) -> None:
    try:
        shp.Shadow.Visible = MSO_FALSE
    except Exception:
        pass


def _tahoma(shp, *, size: float | None = None, bold: bool | None = None, color: str | None = None) -> None:
    try:
        f = shp.TextFrame.Characters().Font
        f.Name = FONT
        if size is not None:
            f.Size = size
        if bold is not None:
            f.Bold = bold
        if color is not None:
            f.Color = _rgb(color)
    except Exception:
        pass


def _ensure_ivory_canvas(ws, rows: int = 40) -> None:
    for c in range(1, 17):
        try:
            ws.Columns(c).ColumnWidth = 8.6
        except Exception:
            pass
    rng = ws.Range(f"A1:P{rows}")
    # Only set fill/font — do NOT clear values (SEARCH formulas live here)
    rng.Interior.Color = _rgb(IVORY)
    try:
        # Don't overwrite SEARCH result/input cell fonts wholesale via range —
        # apply name only where safe: skip if formula present in used band
        pass
    except Exception:
        pass
    # Apply Tahoma to empty/label cells carefully: set default font on unused area
    try:
        ws.Range("A1:P2").Font.Name = FONT
    except Exception:
        pass


def polish_workbook_frontend(excel, wb) -> dict[str, Any]:
    """In-place polish of three Class A sheets. Returns stats."""
    stats: dict[str, Any] = {"sheets": {}, "accents_added": 0, "shapes_touched": 0}

    for name in FRONT:
        ws = wb.Worksheets(name)
        ws.Select()
        try:
            ws.Unprotect(Password="")
        except Exception:
            pass

        win = excel.ActiveWindow
        win.DisplayGridlines = False
        try:
            win.DisplayHeadings = False
        except Exception:
            pass
        try:
            win.Zoom = 92
        except Exception:
            pass
        win.ScrollRow = 1
        win.ScrollColumn = 1

        _ensure_ivory_canvas(ws, 40)

        # SEARCH: preserve white/gold input area after ivory wash
        if name == "SEARCH":
            try:
                for r in range(9, 15):
                    for c in range(2, 17):
                        ws.Cells(r, c).Interior.Color = _rgb(CARD)
                    ws.Cells(r, 1).Interior.Color = _rgb(GOLD)
                for r in range(16, 28):
                    for c in range(2, 17):
                        ws.Cells(r, c).Interior.Color = _rgb(CARD)
                    ws.Cells(r, 1).Interior.Color = _rgb(GOLD)
                ws.Range("B11:H12").Interior.Color = _rgb(WHITE)
                for edge in (7, 8, 9, 10):
                    b = ws.Range("B11:H12").Borders(edge)
                    b.LineStyle = 1
                    b.Weight = 3
                    b.Color = _rgb(GOLD)
                ws.Range("B11").Font.Name = FONT
                ws.Range("B11").Font.Size = 12
                ws.Range("B11").Font.Bold = True
                ws.Range("B11").Font.Color = _rgb(NAVY)
            except Exception:
                pass

        touched = 0
        accents = 0
        for i in range(1, int(ws.Shapes.Count) + 1):
            shp = ws.Shapes(i)
            n = str(shp.Name)
            touched += 1

            # Typography
            try:
                if shp.TextFrame.Characters().Text:
                    _tahoma(shp)
            except Exception:
                pass

            # Hero: richer gradient + text hierarchy
            if n == "Hero":
                try:
                    shp.Fill.TwoColorGradient(MSO_GRADIENT_HORIZ, 1)
                    shp.Fill.ForeColor.RGB = _rgb(MIDNIGHT)
                    shp.Fill.BackColor.RGB = _rgb(NAVY)
                except Exception:
                    pass
                _tahoma(shp, size=13, bold=True, color=WHITE)
                try:
                    tf = shp.TextFrame
                    tf.MarginLeft = 16
                    tf.MarginTop = 10
                    tf.MarginBottom = 8
                except Exception:
                    pass
                _no_shadow(shp)

            elif n == "HeroGold":
                try:
                    shp.Fill.TwoColorGradient(MSO_GRADIENT_HORIZ, 1)
                    shp.Fill.ForeColor.RGB = _rgb(DARK_GOLD)
                    shp.Fill.BackColor.RGB = _rgb(GOLD)
                except Exception:
                    shp.Fill.Solid()
                    shp.Fill.ForeColor.RGB = _rgb(GOLD)
                _no_shadow(shp)

            elif n == "NavBar":
                try:
                    shp.Fill.TwoColorGradient(MSO_GRADIENT_HORIZ, 1)
                    shp.Fill.ForeColor.RGB = _rgb(MIDNIGHT)
                    shp.Fill.BackColor.RGB = _rgb(NAVY)
                except Exception:
                    pass
                _no_shadow(shp)

            elif n.startswith("NavPill_"):
                _no_shadow(shp)
                _tahoma(shp, size=8, bold=True)
                try:
                    shp.TextFrame.HorizontalAlignment = 2  # center
                    shp.TextFrame.VerticalAlignment = 2
                    shp.TextFrame.MarginLeft = 2
                    shp.TextFrame.MarginRight = 2
                except Exception:
                    pass

            elif n.startswith("Status_"):
                _no_shadow(shp)
                _tahoma(shp, size=6.5, bold=True)
                try:
                    shp.TextFrame.HorizontalAlignment = 2
                    shp.TextFrame.VerticalAlignment = 2
                except Exception:
                    pass

            elif n.startswith("KPI_"):
                shp.Fill.Solid()
                shp.Fill.ForeColor.RGB = _rgb(CARD)
                shp.Line.Visible = MSO_TRUE
                shp.Line.ForeColor.RGB = _rgb(PALE)
                shp.Line.Weight = 0.75
                _soft_shadow(shp)
                _tahoma(shp, size=13, bold=True, color=NAVY)
                try:
                    shp.TextFrame.MarginLeft = 14
                    shp.TextFrame.MarginTop = 12
                    shp.TextFrame.MarginBottom = 10
                except Exception:
                    pass
                # Thin gold top accent inside card (only if not already present)
                accent_name = f"{n}_accent"
                exists = False
                for j in range(1, int(ws.Shapes.Count) + 1):
                    if str(ws.Shapes(j).Name) == accent_name:
                        exists = True
                        break
                if not exists:
                    try:
                        acc = ws.Shapes.AddShape(
                            1,  # msoShapeRectangle
                            float(shp.Left) + 12,
                            float(shp.Top) + 10,
                            28,
                            3,
                        )
                        acc.Name = accent_name
                        acc.Fill.Solid()
                        acc.Fill.ForeColor.RGB = _rgb(GOLD)
                        acc.Line.Visible = MSO_FALSE
                        try:
                            acc.Placement = 3
                        except Exception:
                            pass
                        _no_shadow(acc)
                        acc.Locked = True
                        accents += 1
                    except Exception:
                        pass

            elif n in ("SecStrip",) or n.startswith("Panel"):
                shp.Fill.Solid()
                shp.Fill.ForeColor.RGB = _rgb(CARD)
                shp.Line.Visible = MSO_TRUE
                shp.Line.ForeColor.RGB = _rgb(PALE)
                shp.Line.Weight = 0.75
                _soft_shadow(shp)
                _tahoma(shp, size=9, bold=(n.startswith("Panel")), color=INK)
                try:
                    shp.TextFrame.MarginLeft = 12
                    shp.TextFrame.MarginTop = 10
                except Exception:
                    pass

            elif n.startswith("Act_") or n.startswith("NavCard_") or n.startswith("SQ_"):
                shp.Fill.Solid()
                shp.Fill.ForeColor.RGB = _rgb(CARD)
                shp.Line.Visible = MSO_TRUE
                shp.Line.ForeColor.RGB = _rgb(PALE)
                shp.Line.Weight = 0.75
                _soft_shadow(shp)
                _tahoma(shp, size=9, bold=True, color=NAVY)
                try:
                    shp.TextFrame.MarginLeft = 12
                    shp.TextFrame.MarginTop = 8
                    shp.TextFrame.MarginBottom = 8
                except Exception:
                    pass

            elif n == "SearchBtn":
                shp.Fill.Solid()
                shp.Fill.ForeColor.RGB = _rgb(NAVY)
                shp.Line.Visible = MSO_FALSE
                _soft_shadow(shp)
                _tahoma(shp, size=11, bold=True, color=WHITE)
                try:
                    shp.TextFrame.HorizontalAlignment = 2
                    shp.TextFrame.VerticalAlignment = 2
                except Exception:
                    pass

            elif n.startswith("Inci"):
                _no_shadow(shp)

            else:
                # generic card-like rounded shapes
                try:
                    if "Round" in str(getattr(shp, "AutoShapeType", "")) or True:
                        _soft_shadow(shp)
                except Exception:
                    pass

        # Lock all shapes again
        for i in range(1, int(ws.Shapes.Count) + 1):
            try:
                ws.Shapes(i).Locked = True
            except Exception:
                pass

        # Light protect
        try:
            ws.Cells.Locked = False
        except Exception:
            pass
        if name == "SEARCH":
            try:
                ws.Range("B11:H12").Locked = False
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
        except Exception:
            pass

        stats["sheets"][name] = {
            "shapes": int(ws.Shapes.Count),
            "touched": touched,
            "accents_added": accents,
        }
        stats["accents_added"] += accents
        stats["shapes_touched"] += touched

    wb.Worksheets("00_HOME").Select()
    return stats
