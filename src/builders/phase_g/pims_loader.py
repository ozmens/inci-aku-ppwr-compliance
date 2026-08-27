"""Load packaging configuration + products from production PIMS workbook."""

from __future__ import annotations

import re
from pathlib import Path

from models.packaging_configuration import PackagingConfiguration, PackagingConfigurationLine, SourceLineage
from models.product import Product
from repositories.excel_repository import ExcelRepository
from importers.production.variant_description_codec import VariantDescriptionCodec
from importers.production.normalizer import _line_weight_inputs
from services.id_service import IdService


class ProductionDocumentLoader:
    """Table-oriented loader for Phase G document generation."""

    def __init__(self, workbook_path: Path) -> None:
        self.workbook_path = workbook_path
        self.repo = ExcelRepository(workbook_path, enabled=True)
        self.codec = VariantDescriptionCodec()
        self.ids = IdService()

    def open(self) -> None:
        self.repo.open()

    def close(self) -> None:
        self.repo.close()

    def list_set_codes_by_family(self, family: str) -> list[str]:
        rows = self.repo.iter_data_rows("PACKAGING_CONFIGURATION")
        out = []
        for r in rows:
            notes = str(r.get("NOTES") or "")
            if f"FAMILY={family.upper()}" in notes:
                out.append(str(r["CONFIG_GROUP_CODE"]))
        return out

    def load_configuration(self, packaging_set_code: str) -> tuple[PackagingConfiguration, list[Product]]:
        pc_rows = self.repo.iter_data_rows("PACKAGING_CONFIGURATION")
        pc = next(
            (r for r in pc_rows if str(r.get("CONFIG_GROUP_CODE")) == packaging_set_code),
            None,
        )
        if not pc:
            raise KeyError(f"Configuration not found: {packaging_set_code}")
        pc_id = str(pc["PACKAGING_CONFIGURATION_ID"])
        notes = str(pc.get("NOTES") or "")
        final_id = _note_field(notes, "FINAL_CONFIGURATION_ID") or self.ids.final_configuration_id(
            packaging_set_code
        )
        source_cfg = _note_field(notes, "SOURCE_CONFIGURATION_ID") or ""
        mass_kg = _note_field(notes, "PACKAGING_MASS_KG")
        family = _note_field(notes, "FAMILY") or ""
        try:
            vb_tr, vb_en = self.codec.deserialize(str(pc.get("DESCRIPTION") or ""))
        except Exception:
            vb_tr, vb_en = "", ""

        parent = self.ids.normalize_family_code(packaging_set_code)
        m = re.search(r"-(\d{2})$", packaging_set_code)
        variant_no = int(m.group(1)) if m else 1

        # BOM lines
        line_rows = [
            r
            for r in self.repo.iter_data_rows("PACKAGING_CONFIGURATION_LINE")
            if str(r.get("PACKAGING_CONFIGURATION_ID")) == pc_id
        ]
        line_rows.sort(key=lambda r: int(r.get("SORT_ORDER") or 0))
        components = {
            str(r["COMPONENT_ID"]): r for r in self.repo.iter_data_rows("COMPONENT")
        }
        bom: list[PackagingConfigurationLine] = []
        for r in line_rows:
            comp = components.get(str(r.get("COMPONENT_ID")), {})
            erp = str(comp.get("COMPONENT_CODE") or comp.get("EXTERNAL_REF") or "")
            name = str(comp.get("COMPONENT_NAME") or "")
            qty = float(r.get("QUANTITY") or 0)
            notes_l = str(r.get("NOTES") or "")
            uom = _note_field(notes_l, "UOM") or "ADT"
            weight_basis = _note_field(notes_l, "WEIGHT_BASIS") or ""
            line_wt = _note_field(notes_l, "LINE_WEIGHT_KG")
            unit_wt = None
            # reconstruct weight_g for WeightService using same rules as migration
            class _BL:
                pass

            bl = _BL()
            bl.quantity = qty
            bl.uom = uom
            bl.weight_basis = weight_basis
            bl.unit_weight_kg = None
            bl.line_weight_kg = float(line_wt) if line_wt else 0.0
            # Prefer component WEIGHT_G when not mass-input
            comp_wg = comp.get("WEIGHT_G")
            if weight_basis == "RECIPE_QUANTITY_IS_KG" or (
                uom.upper() == "KG" and (comp_wg is None or float(comp_wg or 0) == 1000)
            ):
                weight_g, qty_use = 1000.0, qty
            elif comp_wg not in (None, ""):
                weight_g, qty_use = float(comp_wg), qty
            else:
                weight_g, qty_use = _line_weight_inputs(bl)
            bom.append(
                PackagingConfigurationLine(
                    component_id=str(r.get("COMPONENT_ID")),
                    component_erp_code=erp,
                    component_name=name,
                    quantity=qty_use,
                    uom=uom,
                    sort_order=int(r.get("SORT_ORDER") or 0),
                    weight_g=weight_g,
                    name_tr=name,
                    name_en=name,
                )
            )

        # Products via commercial scenarios
        products: list[Product] = []
        scenarios = self.repo.iter_data_rows("COMMERCIAL_SCENARIO")
        transports = {
            str(t["TRANSPORT_CONFIGURATION_ID"]): t
            for t in self.repo.iter_data_rows("TRANSPORT_CONFIGURATION")
        }
        product_rows = {str(p["PRODUCT_ID"]): p for p in self.repo.iter_data_rows("PRODUCT")}
        seen = set()
        for sc in scenarios:
            tc = transports.get(str(sc.get("TRANSPORT_CONFIGURATION_ID")))
            if not tc:
                continue
            if str(tc.get("PACKAGING_CONFIGURATION_ID")) != pc_id:
                continue
            pid = str(sc.get("PRODUCT_ID"))
            if pid in seen:
                continue
            seen.add(pid)
            pr = product_rows.get(pid)
            if not pr:
                continue
            products.append(
                Product(
                    product_id=pid,
                    product_code=str(pr.get("PRODUCT_CODE")),
                    product_name=str(pr.get("PRODUCT_NAME") or ""),
                )
            )

        # Split configuration name into TR/EN if bilingual stored as single
        name = str(pc.get("PACKAGING_CONFIGURATION_NAME") or packaging_set_code)
        name_tr, name_en = name, name
        if " | " in name and "Variant" in name:
            # production names are often English-only from Golden register
            pass

        cfg = PackagingConfiguration(
            packaging_set_code=packaging_set_code,
            final_configuration_id=final_id,
            parent_family_code=parent,
            variant_number=variant_no,
            name=name,
            revision_no=int(pc.get("REVISION_NO") or 0),
            revision_code="R00",
            description=str(pc.get("DESCRIPTION") or ""),
            variant_basis_tr=vb_tr,
            variant_basis_en=vb_en,
            lines=bom,
            lineage=SourceLineage(source_configuration_id=source_cfg),
            status="ACTIVE",
            extras={
                "family": family,
                "packaging_mass_kg": float(mass_kg) if mass_kg else None,
                "name_tr": name_tr,
                "name_en": name_en,
                "nominal_product_qty": _guess_qty(name, packaging_set_code),
                "revision_date": "08.08.2026",
                "phone": "",
            },
        )
        return cfg, products


def _note_field(notes: str, key: str) -> str | None:
    m = re.search(rf"{re.escape(key)}=([^;]+)", notes)
    return m.group(1).strip() if m else None


def _guess_qty(name: str, set_code: str) -> str:
    m = re.search(r"(\d+)\s*-?\s*unit", name, re.I)
    if m:
        return m.group(1)
    m = re.search(r"ST-(\d+)-", set_code)
    if m:
        return str(int(m.group(1)))
    return ""
