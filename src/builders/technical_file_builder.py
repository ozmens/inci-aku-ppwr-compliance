"""
Document builders — Phase D architecture stubs.

They consume DocumentContext only. No Word generation until ENABLE_WORD_GENERATION.
Weight/ID logic lives in services — not here.
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

import config
from models.document_context import DocumentContext
from utils.docx import assert_word_allowed


@dataclass(frozen=True, slots=True)
class BuilderPlan:
    document_kind: str
    output_stem: str
    uses_technical_file_privacy: bool
    status: str = "architecture_ready_no_render"


class TechnicalFileBuilder:
    """Future Golden Technical File — never receives customer/OEM/market fields."""

    def plan(self, ctx: DocumentContext) -> BuilderPlan:
        safe = ctx.for_technical_file()
        return BuilderPlan(
            document_kind="TECHNICAL_FILE",
            output_stem=safe.document_ids.technical_file_id,
            uses_technical_file_privacy=True,
        )

    def render(self, ctx: DocumentContext, output_dir: Path) -> Path:
        assert_word_allowed(config.ENABLE_WORD_GENERATION)
        raise AssertionError("unreachable")  # pragma: no cover


class DeclarationBuilder:
    def plan(self, ctx: DocumentContext) -> BuilderPlan:
        return BuilderPlan(
            document_kind="DECLARATION_OF_CONFORMITY",
            output_stem=ctx.document_ids.doc_id,
            uses_technical_file_privacy=False,
        )

    def render(self, ctx: DocumentContext, output_dir: Path) -> Path:
        assert_word_allowed(config.ENABLE_WORD_GENERATION)
        raise AssertionError("unreachable")  # pragma: no cover


class LabelBuilder:
    def plan(self, ctx: DocumentContext) -> BuilderPlan:
        return BuilderPlan(
            document_kind="PACKAGING_IDENTIFICATION_LABEL",
            output_stem=ctx.document_ids.label_id,
            uses_technical_file_privacy=False,
        )

    def render(self, ctx: DocumentContext, output_dir: Path) -> Path:
        assert_word_allowed(config.ENABLE_WORD_GENERATION)
        raise AssertionError("unreachable")  # pragma: no cover


class ShipmentStatementBuilder:
    def plan(self, ctx: DocumentContext) -> BuilderPlan:
        return BuilderPlan(
            document_kind="SHIPMENT_PACKAGING_INFORMATION_STATEMENT",
            output_stem=ctx.document_ids.statement_id,
            uses_technical_file_privacy=False,
        )

    def render(self, ctx: DocumentContext, output_dir: Path) -> Path:
        assert_word_allowed(config.ENABLE_WORD_GENERATION)
        raise AssertionError("unreachable")  # pragma: no cover


class DocumentPackageBuilder:
    """Orchestrates the four Golden outputs from one DocumentContext."""

    def __init__(self) -> None:
        self.tf = TechnicalFileBuilder()
        self.doc = DeclarationBuilder()
        self.label = LabelBuilder()
        self.statement = ShipmentStatementBuilder()

    def plan_all(self, ctx: DocumentContext) -> list[BuilderPlan]:
        return [
            self.tf.plan(ctx),
            self.doc.plan(ctx),
            self.label.plan(ctx),
            self.statement.plan(ctx),
        ]
