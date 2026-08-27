"""Domain and schema models for PIMS (Phase D architecture)."""

from .registry import SchemaRegistry, TableDefinition
from .schema_version import SCHEMA_VERSION
from .component import Component
from .packaging_configuration import PackagingConfiguration, PackagingConfigurationLine, SourceLineage
from .product import Product
from .document_context import DocumentContext, DocumentIds, LegalEntityInfo
from .technical_file import TechnicalFile, Article5Assessment
from .declaration import Declaration
from .shipment import Shipment
from .statement import Statement
from .document import Document, DocumentLink
from .commercial_scenario import CommercialScenario
from .transport_configuration import TransportConfiguration
from .material import Material

__all__ = [
    "SchemaRegistry",
    "TableDefinition",
    "SCHEMA_VERSION",
    "Component",
    "PackagingConfiguration",
    "PackagingConfigurationLine",
    "SourceLineage",
    "Product",
    "DocumentContext",
    "DocumentIds",
    "LegalEntityInfo",
    "TechnicalFile",
    "Article5Assessment",
    "Declaration",
    "Shipment",
    "Statement",
    "Document",
    "DocumentLink",
    "CommercialScenario",
    "TransportConfiguration",
    "Material",
]
