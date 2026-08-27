"""Phase G — Golden Word template integration / pilot generation."""

from .service import PhaseGPilotService, PilotResult
from .runtime_template_builder import build_runtime_templates
from .tokens import GOLDEN_FILES, RUNTIME_FILES

__all__ = [
    "PhaseGPilotService",
    "PilotResult",
    "build_runtime_templates",
    "GOLDEN_FILES",
    "RUNTIME_FILES",
]
