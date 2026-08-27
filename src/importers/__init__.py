"""Import template generation and Phase F production migration."""

from .template_builder import generate_import_templates

try:
    from .production import ProductionMigrationService, VariantDescriptionCodec
except ImportError:  # pragma: no cover
    ProductionMigrationService = None  # type: ignore
    VariantDescriptionCodec = None  # type: ignore

__all__ = [
    "generate_import_templates",
    "ProductionMigrationService",
    "VariantDescriptionCodec",
]
