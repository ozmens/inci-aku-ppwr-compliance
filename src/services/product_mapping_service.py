"""
Product ↔ Packaging Configuration mapping for Battery DPP traceability.

Battery Product / SKU
→ Source Configuration
→ Final Packaging Configuration
→ Exact Packaging BOM
→ Components → Materials → Mass → PPWR documents
"""

from __future__ import annotations

from dataclasses import dataclass, field

from models.product import Product


@dataclass(slots=True)
class ProductPackagingMap:
    product_id: str
    product_code: str
    source_configuration_id: str | None
    final_configuration_id: str
    packaging_set_code: str


@dataclass(slots=True)
class ProductMappingService:
    """In-memory mapping registry (Excel repository later)."""

    _maps: list[ProductPackagingMap] = field(default_factory=list)

    def link(
        self,
        product: Product,
        *,
        final_configuration_id: str,
        packaging_set_code: str,
        source_configuration_id: str | None = None,
    ) -> ProductPackagingMap:
        entry = ProductPackagingMap(
            product_id=product.product_id,
            product_code=product.product_code,
            source_configuration_id=source_configuration_id,
            final_configuration_id=final_configuration_id,
            packaging_set_code=packaging_set_code,
        )
        self._maps.append(entry)
        if final_configuration_id not in product.linked_final_configuration_ids:
            product.linked_final_configuration_ids.append(final_configuration_id)
        if source_configuration_id and source_configuration_id not in product.linked_source_configuration_ids:
            product.linked_source_configuration_ids.append(source_configuration_id)
        return entry

    def by_product(self, product_code: str) -> list[ProductPackagingMap]:
        code = product_code.strip().upper()
        return [m for m in self._maps if m.product_code.upper() == code]

    def by_final_configuration(self, final_configuration_id: str) -> list[ProductPackagingMap]:
        fid = final_configuration_id.strip().upper()
        return [m for m in self._maps if m.final_configuration_id.upper() == fid]

    def orphan_products(self, products: list[Product]) -> list[Product]:
        linked = {m.product_id for m in self._maps}
        return [p for p in products if p.product_id not in linked]
