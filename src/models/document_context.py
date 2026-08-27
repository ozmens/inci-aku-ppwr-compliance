"""Shared DocumentContext for all future Golden Word builders."""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any

from .packaging_configuration import PackagingConfiguration, PackagingConfigurationLine
from .product import Product
from .shipment import Shipment
from .technical_file import Article5Assessment


@dataclass(slots=True)
class MaterialSummaryRow:
    material_family: str
    weight_g: float
    share_pct: float | None = None


@dataclass(slots=True)
class LegalEntityInfo:
    legal_name: str
    address: str
    email: str
    website: str


@dataclass(slots=True)
class DocumentIds:
    packaging_set_code: str
    final_configuration_id: str
    technical_file_id: str
    doc_id: str
    label_id: str
    statement_id: str
    revision_code: str = "R00"
    revision_display: str = "Rev.00"


@dataclass(slots=True)
class DocumentContext:
    """
    Controlled context consumed by TF / DoC / Label / Statement builders.

    Technical File builders must use `for_technical_file()` which strips
    customer / OEM / market fields.
    """

    legal_entity: LegalEntityInfo
    configuration: PackagingConfiguration
    document_ids: DocumentIds
    bom_lines: list[PackagingConfigurationLine]
    material_summary: list[MaterialSummaryRow]
    total_tare_g: float
    products: list[Product] = field(default_factory=list)
    article5: Article5Assessment | None = None
    annex_drawings_status: str = "PENDING – DRAWINGS / PHOTOGRAPHS"
    shipment: Shipment | None = None
    # Commercial — must NOT flow into Technical File output
    customer_name: str | None = None
    oem_name: str | None = None
    customer_market: str | None = None
    extras: dict[str, Any] = field(default_factory=dict)

    def for_technical_file(self) -> DocumentContext:
        """Return a privacy-safe copy without commercial identifiers."""
        return DocumentContext(
            legal_entity=self.legal_entity,
            configuration=self.configuration,
            document_ids=self.document_ids,
            bom_lines=list(self.bom_lines),
            material_summary=list(self.material_summary),
            total_tare_g=self.total_tare_g,
            products=list(self.products),
            article5=self.article5,
            annex_drawings_status=self.annex_drawings_status,
            shipment=None,
            customer_name=None,
            oem_name=None,
            customer_market=None,
            extras={
                k: v
                for k, v in self.extras.items()
                if k
                not in {
                    "customer_name",
                    "oem_name",
                    "customer_market",
                    "destination_country",
                    "incoterm",
                }
            },
        )
