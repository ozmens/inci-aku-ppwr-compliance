"""
Variant Basis (TR/EN) — explains WHY a BOM is a separate variant.

Must not live inside Word builders.
"""

from __future__ import annotations

from collections import defaultdict

from models.packaging_configuration import PackagingConfigurationLine
from utils.constants import LINE_ROLE_TO_BASIS, VARIANT_BASIS_ROLE_PRIORITY
from utils.text import normalize_ws


class VariantBasisService:
    """Derive customer/auditor-facing Variant Basis from exact BOM differences."""

    def _bucket(self, line: PackagingConfigurationLine) -> str:
        role = (line.line_role or "OTHER").upper()
        return LINE_ROLE_TO_BASIS.get(role, "OTHER")

    def _describe_line(self, line: PackagingConfigurationLine, lang: str) -> str:
        name = line.name_tr if lang == "TR" else line.name_en
        name = name or line.component_name
        erp = line.component_erp_code or line.component_id
        qty = line.quantity
        if lang == "TR":
            return f"{normalize_ws(name)} (ERP {erp}) | {qty:g} adet"
        return f"{normalize_ws(name)} (ERP {erp}) | qty {qty:g}"

    def derive(
        self,
        *,
        battery_or_pack_type: str,
        lines: list[PackagingConfigurationLine],
        max_parts: int = 3,
    ) -> tuple[str, str]:
        """
        Priority: type → outer carton → upper cover → separator → edge → film → strap → other.
        Returns (VARIANT_BASIS_TR, VARIANT_BASIS_EN).
        """
        by_bucket: dict[str, list[PackagingConfigurationLine]] = defaultdict(list)
        for line in lines:
            by_bucket[self._bucket(line)].append(line)

        parts_tr: list[str] = [normalize_ws(battery_or_pack_type)]
        parts_en: list[str] = [normalize_ws(battery_or_pack_type)]

        for bucket in VARIANT_BASIS_ROLE_PRIORITY:
            if bucket == "BATTERY_TYPE":
                continue
            candidates = by_bucket.get(bucket) or []
            if not candidates:
                continue
            # Prefer highest quantity / first sort
            candidates = sorted(candidates, key=lambda x: (-x.quantity, x.sort_order, x.component_erp_code))
            line = candidates[0]
            parts_tr.append(self._describe_line(line, "TR"))
            parts_en.append(self._describe_line(line, "EN"))
            if len(parts_tr) >= max_parts:
                break

        # If still thin, append other distinctive ERP differences
        if len(parts_tr) < max_parts:
            for line in sorted(lines, key=lambda x: (x.sort_order, x.component_erp_code)):
                desc_tr = self._describe_line(line, "TR")
                if desc_tr in parts_tr:
                    continue
                parts_tr.append(desc_tr)
                parts_en.append(self._describe_line(line, "EN"))
                if len(parts_tr) >= max_parts:
                    break

        return (" | ".join(parts_tr), " | ".join(parts_en))

    def is_unique_within_family(
        self,
        basis_tr: str,
        existing_bases_tr: list[str],
    ) -> bool:
        """Two active variants in same parent family may not share Variant Basis."""
        norm = normalize_ws(basis_tr).casefold()
        return all(normalize_ws(b).casefold() != norm for b in existing_bases_tr)
