"""
BOM identity / equivalence — Golden Variant exact-BOM principle.

ONE physically distinct packaging BOM = ONE final Packaging Configuration.
"""

from __future__ import annotations

import hashlib
import json
from dataclasses import dataclass

from models.packaging_configuration import PackagingConfigurationLine


@dataclass(frozen=True, slots=True)
class NormalizedBomItem:
    component_key: str
    quantity: float
    uom: str
    line_role: str


class BomService:
    """Exact-BOM signature and physical equivalence checks."""

    def normalize_lines(
        self, lines: list[PackagingConfigurationLine]
    ) -> list[NormalizedBomItem]:
        items: list[NormalizedBomItem] = []
        for line in lines:
            key = (line.component_erp_code or line.component_id or "").strip().upper()
            if not key:
                continue
            items.append(
                NormalizedBomItem(
                    component_key=key,
                    quantity=float(line.quantity),
                    uom=(line.uom or "PCS").strip().upper(),
                    line_role=(line.line_role or "OTHER").strip().upper(),
                )
            )
        items.sort(key=lambda x: (x.component_key, x.line_role, x.uom, x.quantity))
        return items

    def bom_signature(self, lines: list[PackagingConfigurationLine]) -> str:
        """
        Internal control fingerprint — NOT a customer-facing identifier.
        """
        payload = [
            {
                "c": i.component_key,
                "q": i.quantity,
                "u": i.uom,
                "r": i.line_role,
            }
            for i in self.normalize_lines(lines)
        ]
        raw = json.dumps(payload, separators=(",", ":"), ensure_ascii=True)
        return hashlib.sha256(raw.encode("utf-8")).hexdigest()

    def are_physically_equivalent(
        self,
        a: list[PackagingConfigurationLine],
        b: list[PackagingConfigurationLine],
    ) -> bool:
        """True only when normalized BOM content is identical."""
        return self.normalize_lines(a) == self.normalize_lines(b)

    def require_separate_configurations(
        self,
        a: list[PackagingConfigurationLine],
        b: list[PackagingConfigurationLine],
    ) -> bool:
        """
        Different ERP code / qty / UOM / role ⇒ separate final configurations.
        Same pallet count alone must NEVER force aggregation.
        """
        return not self.are_physically_equivalent(a, b)
