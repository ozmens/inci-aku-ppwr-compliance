"""
Final production hardening: protect all system/hidden sheets.
"""

from __future__ import annotations

from openpyxl.workbook.workbook import Workbook

from .inci_ops_ux import VISIBLE_SHEET_TITLES


def protect_all_system_sheets(wb: Workbook, tech_password: str = "PIMS_TECH") -> int:
    """Protect every non-visible sheet (engines, lookups, masters, DOC_ENGINE_VARS)."""
    visible = set(VISIBLE_SHEET_TITLES.values()) | {"Dashboard"}
    count = 0
    for name in wb.sheetnames:
        ws = wb[name]
        if name in visible and ws.sheet_state == "visible":
            continue
        # Force hidden system sheets to stay hidden
        if name.startswith(("ENG_", "SYS_", "LKP_", "DOC_")) or name == "SYS_RELATIONSHIPS":
            if ws.sheet_state == "visible":
                ws.sheet_state = "hidden"
        if ws.sheet_state != "visible":
            ws.protection.sheet = True
            ws.protection.password = tech_password
            count += 1
    return count
