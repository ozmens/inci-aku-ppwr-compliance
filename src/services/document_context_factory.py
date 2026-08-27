"""Build DocumentContext from services — shared by all document builders."""

from __future__ import annotations

import config
from models.document_context import (
    DocumentContext,
    DocumentIds,
    LegalEntityInfo,
    MaterialSummaryRow,
)
from models.packaging_configuration import PackagingConfiguration
from models.product import Product
from models.shipment import Shipment
from models.technical_file import Article5Assessment
from services.id_service import IdService
from services.weight_service import WeightService
from utils.constants import ANNEX_DRAWINGS_PENDING, ARTICLE5_BASIS_LABEL


class DocumentContextFactory:
    """Single controlled context for TF / DoC / Label / Statement builders."""

    def __init__(
        self,
        id_service: IdService | None = None,
        weight_service: WeightService | None = None,
    ) -> None:
        self.ids = id_service or IdService(
            default_revision_code=config.DOCUMENT_REVISION_CODE
        )
        self.weights = weight_service or WeightService()

    def build(
        self,
        configuration: PackagingConfiguration,
        *,
        products: list[Product] | None = None,
        shipment: Shipment | None = None,
        customer_name: str | None = None,
        oem_name: str | None = None,
        customer_market: str | None = None,
        material_summary: list[MaterialSummaryRow] | None = None,
        article5: Article5Assessment | None = None,
    ) -> DocumentContext:
        wr = self.weights.calculate_tare(configuration.lines)
        doc_ids = DocumentIds(
            packaging_set_code=configuration.packaging_set_code,
            final_configuration_id=configuration.final_configuration_id,
            technical_file_id=self.ids.technical_file_id(
                configuration.packaging_set_code, configuration.revision_code
            ),
            doc_id=self.ids.doc_id(
                configuration.packaging_set_code, configuration.revision_code
            ),
            label_id=self.ids.label_id(
                configuration.packaging_set_code, configuration.revision_code
            ),
            statement_id=self.ids.statement_id(
                configuration.packaging_set_code, configuration.revision_code
            ),
            revision_code=configuration.revision_code,
            revision_display=config.DOCUMENT_REVISION,
        )
        if article5 is None:
            article5 = Article5Assessment(
                basis_label=ARTICLE5_BASIS_LABEL,
                evidence_references=[],
                numerical_claim=None,
            )
        return DocumentContext(
            legal_entity=LegalEntityInfo(
                legal_name=config.COMPANY_LEGAL_NAME,
                address=config.COMPANY_ADDRESS,
                email=config.COMPANY_EMAIL,
                website=config.COMPANY_WEBSITE,
            ),
            configuration=configuration,
            document_ids=doc_ids,
            bom_lines=list(configuration.lines),
            material_summary=list(material_summary or []),
            total_tare_g=wr.total_tare_g,
            products=list(products or []),
            article5=article5,
            annex_drawings_status=ANNEX_DRAWINGS_PENDING,
            shipment=shipment,
            customer_name=customer_name,
            oem_name=oem_name,
            customer_market=customer_market,
        )
