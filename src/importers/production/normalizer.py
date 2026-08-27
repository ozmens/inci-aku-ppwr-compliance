"""Normalize staged Golden data into Schema 1.0.0 row dicts."""

from __future__ import annotations

from collections import defaultdict
from datetime import datetime, timezone
from typing import Any

from models.packaging_configuration import PackagingConfigurationLine
from services.id_service import IdService
from services.weight_service import WeightService

from .staging import StagingBundle
from .variant_description_codec import VariantDescriptionCodec


NOW = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
EFFECTIVE_FROM = "2026-08-08"
STATUS_ACTIVE = "2"
MASS_INPUT_BASES = frozenset(
    {
        "RECIPE_QUANTITY_IS_KG",
        "DIRECT_SOURCE_MASS_UOM",
        "DIRECT_SOURCE_APPLIED_MASS",
        "DIRECT_SOURCE_LINE_MASS",
        "DIRECT_SOURCE_OPERATIONAL_PALLET_MASS",
    }
)


class NormalizedStore:
    def __init__(self) -> None:
        self.tables: dict[str, list[dict[str, Any]]] = defaultdict(list)
        self.set_code_to_pc_id: dict[str, str] = {}
        self.erp_to_component_id: dict[str, str] = {}
        self.product_code_to_id: dict[str, str] = {}
        self.uom_code_to_id: dict[str, str] = {}
        self.blocking_errors: list[str] = []
        self.warnings: list[str] = []
        self.stats: dict[str, Any] = {}


def normalize_bundle(bundle: StagingBundle) -> NormalizedStore:
    store = NormalizedStore()
    codec = VariantDescriptionCodec()
    ids = IdService()
    weights = WeightService()

    # --- Lookups extensions (M, SET) ---
    # Base lookups already in template; promoter may append if missing.
    store.uom_code_to_id = {"G": "1", "KG": "2", "MM": "3", "PCS": "4", "PAL": "5", "TEU": "6", "M": "7", "SET": "8", "ADT": "4"}

    # --- Org seeds ---
    store.tables["LEGAL_ENTITY"].append(
        {
            "LEGAL_ENTITY_ID": "1",
            "LEGAL_ENTITY_CODE": "INCI_AKU",
            "LEGAL_ENTITY_NAME": "İnci Akü Sanayi ve Ticaret A.Ş.",
            "COUNTRY_ID": "1",
            "STATUS_ID": STATUS_ACTIVE,
            "NOTES": "Platform manufacturer seed — Phase F",
        }
    )
    store.tables["PERSON"].append(
        {
            "PERSON_ID": "1",
            "PERSON_CODE": "P-COMPLIANCE",
            "FULL_NAME": "PPWR Compliance Officer",
            "EMAIL": "compliance@inciaku.com",
            "JOB_TITLE": "Compliance Officer",
            "STATUS_ID": STATUS_ACTIVE,
        }
    )
    store.tables["PLANT"].append(
        {
            "PLANT_ID": "1",
            "PLANT_CODE": "MANISA",
            "PLANT_NAME": "Manisa Plant",
            "LEGAL_ENTITY_ID": "1",
            "COUNTRY_ID": "1",
            "STATUS_ID": STATUS_ACTIVE,
            "EXTERNAL_REF": "PLANT-01",
        }
    )

    # --- Components ---
    for i, (erp, comp) in enumerate(sorted(bundle.components.items()), start=1):
        cid = str(i)
        store.erp_to_component_id[erp] = cid
        weight_g = _canonical_component_weight_g(erp, bundle)
        notes = f"SOURCE_ERP={erp}"
        if comp.aliases:
            notes += f" | ALIASES={'; '.join(comp.aliases[:5])}"
        if comp.lineage:
            notes += f" | SRC={comp.lineage.source_file}"
        store.tables["COMPONENT"].append(
            {
                "COMPONENT_ID": cid,
                "COMPONENT_CODE": erp,
                "COMPONENT_NAME": comp.description or erp,
                "COMPONENT_TYPE_ID": "14",  # OTHER default
                "PACKAGING_LEVEL_ID": "3",
                "PACKAGING_FUNCTION_ID": "3",
                "OWNERSHIP_TYPE_ID": "1",
                "SUPPLIER_ID": None,
                "WEIGHT_G": weight_g,
                "LENGTH_MM": None,
                "WIDTH_MM": None,
                "HEIGHT_MM": None,
                "RECYCLED_CONTENT_PCT": None,
                "RECYCLABILITY_CLASS_ID": None,
                "REUSE_CYCLE_TARGET": None,
                "SPEC_REF": None,
                "STATUS_ID": STATUS_ACTIVE,
                "EFFECTIVE_FROM": EFFECTIVE_FROM,
                "EFFECTIVE_TO": None,
                "EXTERNAL_REF": erp,
                "NOTES": notes,
                "CREATED_AT": NOW,
                "UPDATED_AT": NOW,
            }
        )
        # Material family if known — map loosely
        if comp.material_family:
            mid = _material_id_for_family(comp.material_family)
            if mid:
                store.tables["COMPONENT_MATERIAL"].append(
                    {
                        "COMPONENT_MATERIAL_ID": str(len(store.tables["COMPONENT_MATERIAL"]) + 1),
                        "COMPONENT_ID": cid,
                        "MATERIAL_ID": mid,
                        "SHARE_PCT": 100.0,
                        "SORT_ORDER": 1,
                        "NOTES": f"Family-level only from Level-2 ({comp.material_family}); not invented grade",
                    }
                )

    # --- Products (unique by code; prefer in-scope descriptions) ---
    products: dict[str, dict] = {}
    for pm in bundle.product_maps:
        existing = products.get(pm.product_code)
        prefer = "IN OEM" in (pm.status or "").upper() or "EU PACKAGE" in (pm.status or "").upper()
        if existing is None or (prefer and "IN OEM" not in (existing.get("_status") or "").upper()):
            cat = "1" if (pm.battery_type or "").upper() not in {"IND", "INDUSTRIAL"} else "2"
            # crude: industrial product codes often start differently — use family from mapping later
            products[pm.product_code] = {
                "PRODUCT_CODE": pm.product_code,
                "PRODUCT_NAME": pm.product_description,
                "PRODUCT_CATEGORY_ID": cat,
                "NET_WEIGHT_G": 0,
                "LENGTH_MM": None,
                "WIDTH_MM": None,
                "HEIGHT_MM": None,
                "STATUS_ID": STATUS_ACTIVE,
                "EFFECTIVE_FROM": EFFECTIVE_FROM,
                "EFFECTIVE_TO": None,
                "EXTERNAL_REF": pm.product_code,
                "NOTES": f"STATUS={pm.status}; SOURCE_CFG={pm.source_configuration_id}; MARKET={pm.customer_market}",
                "CREATED_AT": NOW,
                "UPDATED_AT": NOW,
                "_status": pm.status,
            }
    for i, (code, row) in enumerate(sorted(products.items()), start=1):
        pid = str(i)
        store.product_code_to_id[code] = pid
        out = {k: v for k, v in row.items() if not k.startswith("_")}
        out["PRODUCT_ID"] = pid
        store.tables["PRODUCT"].append(out)

    # --- Packaging configurations + BOM + documents ---
    bom_by_set: dict[str, list] = defaultdict(list)
    for line in bundle.bom_lines:
        bom_by_set[line.packaging_set_code].append(line)

    line_id = 1
    doc_lib_id = 1
    doc_link_id = 1
    tf_id = 1
    doc_id_pk = 1

    for i, cfg in enumerate(bundle.configurations, start=1):
        pc_id = str(i)
        store.set_code_to_pc_id[cfg.packaging_set_code] = pc_id

        # Validate Golden IDs vs IdService
        expected_final = ids.final_configuration_id(cfg.packaging_set_code)
        if cfg.final_configuration_id != expected_final:
            store.blocking_errors.append(
                f"Final ID mismatch {cfg.packaging_set_code}: {cfg.final_configuration_id} != {expected_final}"
            )
            continue
        for field, actual, exp in [
            ("TF", cfg.technical_file_id, ids.technical_file_id(cfg.packaging_set_code)),
            ("DOC", cfg.doc_id, ids.doc_id(cfg.packaging_set_code)),
            ("LBL", cfg.label_id, ids.label_id(cfg.packaging_set_code)),
            ("STM", cfg.statement_id, ids.statement_id(cfg.packaging_set_code)),
        ]:
            if actual != exp:
                store.blocking_errors.append(
                    f"Doc ID mismatch {cfg.packaging_set_code} {field}: {actual} != {exp}"
                )

        try:
            description = codec.serialize(cfg.variant_basis_tr, cfg.variant_basis_en)
        except Exception as exc:  # noqa: BLE001
            store.blocking_errors.append(f"Variant codec failed {cfg.packaging_set_code}: {exc}")
            continue

        notes = (
            f"FINAL_CONFIGURATION_ID={cfg.final_configuration_id};"
            f"SOURCE_CONFIGURATION_ID={cfg.source_configuration_id};"
            f"FAMILY={cfg.family};"
            f"PACKAGING_MASS_KG={cfg.packaging_mass_kg};"
            f"EVIDENCE_BASIS={cfg.evidence_basis}"
        )
        if cfg.lineage:
            notes += f";SOURCE_REF={cfg.lineage.source_reference}"

        store.tables["PACKAGING_CONFIGURATION"].append(
            {
                "PACKAGING_CONFIGURATION_ID": pc_id,
                "CONFIG_GROUP_CODE": cfg.packaging_set_code,
                "REVISION_NO": 0,
                "PACKAGING_CONFIGURATION_NAME": cfg.configuration_name,
                "DESCRIPTION": description,
                "SUPERSEDES_ID": None,
                "STATUS_ID": STATUS_ACTIVE,
                "EFFECTIVE_FROM": EFFECTIVE_FROM,
                "EFFECTIVE_TO": None,
                "NOTES": notes,
                "CREATED_AT": NOW,
                "UPDATED_AT": NOW,
            }
        )

        # Exact BOM lines + weight check
        lines = bom_by_set.get(cfg.packaging_set_code, [])
        if not lines:
            store.blocking_errors.append(f"Missing BOM for {cfg.packaging_set_code}")
            continue

        pclines: list[PackagingConfigurationLine] = []
        for sort_i, bl in enumerate(lines, start=1):
            cid = store.erp_to_component_id.get(bl.component_code)
            if not cid:
                store.blocking_errors.append(
                    f"Unresolved component {bl.component_code} on {cfg.packaging_set_code}"
                )
                continue
            weight_g, qty = _line_weight_inputs(bl)
            store.tables["PACKAGING_CONFIGURATION_LINE"].append(
                {
                    "PACKAGING_CONFIGURATION_LINE_ID": str(line_id),
                    "PACKAGING_CONFIGURATION_ID": pc_id,
                    "COMPONENT_ID": cid,
                    "QUANTITY": qty,
                    "LINE_ROLE_ID": "18",  # OTHER
                    "SORT_ORDER": sort_i,
                    "IS_OPTIONAL": False,
                    "NOTES": (
                        f"UOM={bl.uom};WEIGHT_BASIS={bl.weight_basis};"
                        f"LINE_WEIGHT_KG={bl.line_weight_kg};"
                        f"SOURCE_CFG={bl.source_configuration_id};"
                        f"SRC_REF={bl.lineage.source_reference};"
                        f"SRC_FILE={bl.lineage.source_file}"
                    ),
                }
            )
            line_id += 1
            pclines.append(
                PackagingConfigurationLine(
                    component_id=cid,
                    component_erp_code=bl.component_code,
                    component_name=bl.erp_description,
                    quantity=qty,
                    uom=bl.uom,
                    weight_g=weight_g,
                    sort_order=sort_i,
                )
            )

        wr = weights.calculate_tare(pclines)
        claimed_g = float(cfg.packaging_mass_kg) * 1000.0
        if not weights.assert_consistent(
            calculated_tare_g=wr.total_tare_g,
            claimed_tare_g=claimed_g,
            tolerance_g=0.05,  # 0.05 g
        ):
            # also allow kg-space tolerance 1e-4 kg = 0.1 g
            if abs(wr.total_tare_g - claimed_g) > 0.1:
                store.blocking_errors.append(
                    f"Tare mismatch {cfg.packaging_set_code}: calc_g={wr.total_tare_g} claimed_g={claimed_g}"
                )

        # Transport configs for Industrial / Container
        if cfg.family in {"INDUSTRIAL", "CONTAINER"}:
            tc_id = str(len(store.tables["TRANSPORT_CONFIGURATION"]) + 1)
            unit_type = "3" if cfg.family == "CONTAINER" else "2"  # CONTAINER / PALLET
            store.tables["TRANSPORT_CONFIGURATION"].append(
                {
                    "TRANSPORT_CONFIGURATION_ID": tc_id,
                    "CONFIG_GROUP_CODE": cfg.packaging_set_code,
                    "REVISION_NO": 0,
                    "TRANSPORT_CONFIGURATION_NAME": cfg.configuration_name,
                    "PACKAGING_CONFIGURATION_ID": pc_id,
                    "TRANSPORT_UNIT_TYPE_ID": unit_type,
                    "UNITS_PER_LAYER": None,
                    "LAYERS_PER_UNIT": None,
                    "CONTAINER_PAYLOAD_UNITS": None,
                    "MAX_GROSS_WEIGHT_KG": None,
                    "SUPERSEDES_ID": None,
                    "STATUS_ID": STATUS_ACTIVE,
                    "EFFECTIVE_FROM": EFFECTIVE_FROM,
                    "EFFECTIVE_TO": None,
                    "NOTES": f"SOURCE_CONFIGURATION_ID={cfg.source_configuration_id};FAMILY={cfg.family}",
                    "CREATED_AT": NOW,
                    "UPDATED_AT": NOW,
                }
            )
            for sort_i, bl in enumerate(lines, start=1):
                cid = store.erp_to_component_id.get(bl.component_code)
                if not cid:
                    continue
                store.tables["TRANSPORT_CONFIGURATION_LINE"].append(
                    {
                        "TRANSPORT_CONFIGURATION_LINE_ID": str(
                            len(store.tables["TRANSPORT_CONFIGURATION_LINE"]) + 1
                        ),
                        "TRANSPORT_CONFIGURATION_ID": tc_id,
                        "COMPONENT_ID": cid,
                        "QUANTITY_PER_TRANSPORT_UNIT": bl.quantity,
                        "LINE_ROLE_ID": "18",
                        "SORT_ORDER": sort_i,
                        "NOTES": f"UOM={bl.uom};LINE_WEIGHT_KG={bl.line_weight_kg}",
                    }
                )

        # Document metadata — TF entity + DOCUMENT_LIBRARY for all four outputs
        store.tables["TECHNICAL_FILE"].append(
            {
                "TECHNICAL_FILE_ID": str(tf_id),
                "TECHNICAL_FILE_CODE": cfg.technical_file_id,
                "TITLE": f"Technical File — {cfg.packaging_set_code}",
                "COMPONENT_ID": None,
                "PACKAGING_CONFIGURATION_ID": pc_id,
                "TRANSPORT_CONFIGURATION_ID": None,
                "REVISION_NO": 0,
                "ASSESSMENT_DATE": None,
                "RECYCLABILITY_SUMMARY": None,
                "SUBSTANCE_OF_CONCERN_NOTES": (
                    "REV00 CURRENT EVIDENCE BASIS — ARTICLE 5 ASSESSMENT BASIS"
                ),
                "DESIGN_FOR_RECYCLING_NOTES": "PENDING – DRAWINGS / PHOTOGRAPHS",
                "OWNER_PERSON_ID": "1",
                "STATUS_ID": STATUS_ACTIVE,
                "EFFECTIVE_FROM": EFFECTIVE_FROM,
                "EFFECTIVE_TO": None,
                "NOTES": "Metadata only — Word not generated in Phase F",
                "CREATED_AT": NOW,
                "UPDATED_AT": NOW,
            }
        )
        tf_pk = str(tf_id)
        tf_id += 1

        store.tables["DECLARATION_OF_CONFORMITY"].append(
            {
                "DECLARATION_OF_CONFORMITY_ID": str(doc_id_pk),
                "DOC_NUMBER": cfg.doc_id,
                "TITLE": f"EU DoC — {cfg.packaging_set_code}",
                "LEGAL_ENTITY_ID": "1",
                "PRODUCT_ID": None,
                "PACKAGING_CONFIGURATION_ID": pc_id,
                "TRANSPORT_CONFIGURATION_ID": None,
                "TECHNICAL_FILE_ID": tf_pk,
                "RESPONSIBLE_PERSON_ID": "1",
                "REGULATION_REFERENCE": "EU PPWR",
                "CONFORMITY_STATEMENT": (
                    "REV00 CURRENT EVIDENCE BASIS — ARTICLE 5 ASSESSMENT BASIS"
                ),
                "ISSUE_DATE": EFFECTIVE_FROM,
                "VALID_UNTIL": None,
                "REVISION_NO": 0,
                "STATUS_ID": STATUS_ACTIVE,
                "APPROVED_AT": None,
                "NOTES": "Metadata only — Word/PDF not generated in Phase F",
                "CREATED_AT": NOW,
                "UPDATED_AT": NOW,
            }
        )
        doc_id_pk += 1

        for code, dtype in [
            (cfg.technical_file_id, "1"),  # TECH_FILE
            (cfg.doc_id, "4"),  # DOC_PDF
            (cfg.label_id, "11"),  # LABEL
            (cfg.statement_id, "12"),  # STATEMENT
        ]:
            did = str(doc_lib_id)
            store.tables["DOCUMENT_LIBRARY"].append(
                {
                    "DOCUMENT_ID": did,
                    "DOCUMENT_CODE": code,
                    "DOCUMENT_TITLE": code,
                    "DOCUMENT_TYPE_ID": dtype,
                    "FILE_URI": f"pending://metadata-only/{code}",
                    "FILE_HASH": None,
                    "ISSUE_DATE": EFFECTIVE_FROM,
                    "STATUS_ID": STATUS_ACTIVE,
                    "NOTES": "PENDING – DRAWINGS / PHOTOGRAPHS where applicable; Word not generated",
                    "CREATED_AT": NOW,
                    "UPDATED_AT": NOW,
                }
            )
            store.tables["DOCUMENT_LINK"].append(
                {
                    "DOCUMENT_LINK_ID": str(doc_link_id),
                    "DOCUMENT_ID": did,
                    "COMPONENT_ID": None,
                    "PRODUCT_ID": None,
                    "PACKAGING_CONFIGURATION_ID": pc_id,
                    "TRANSPORT_CONFIGURATION_ID": None,
                    "TECHNICAL_FILE_ID": tf_pk if dtype == "1" else None,
                    "DECLARATION_OF_CONFORMITY_ID": None,
                    "STATEMENT_ID": None,
                    "SORT_ORDER": doc_link_id,
                    "NOTES": None,
                }
            )
            doc_lib_id += 1
            doc_link_id += 1

    # One engineering transport per packaging config (for commercial FK); reuse IND/CNT
    pc_to_tc: dict[str, str] = {
        tc["PACKAGING_CONFIGURATION_ID"]: tc["TRANSPORT_CONFIGURATION_ID"]
        for tc in store.tables["TRANSPORT_CONFIGURATION"]
    }
    for cfg in bundle.configurations:
        pc_id = store.set_code_to_pc_id.get(cfg.packaging_set_code)
        if not pc_id or pc_id in pc_to_tc:
            continue
        tc_id = str(len(store.tables["TRANSPORT_CONFIGURATION"]) + 1)
        store.tables["TRANSPORT_CONFIGURATION"].append(
            {
                "TRANSPORT_CONFIGURATION_ID": tc_id,
                "CONFIG_GROUP_CODE": cfg.packaging_set_code,
                "REVISION_NO": 0,
                "TRANSPORT_CONFIGURATION_NAME": f"Default transport — {cfg.packaging_set_code}",
                "PACKAGING_CONFIGURATION_ID": pc_id,
                "TRANSPORT_UNIT_TYPE_ID": "2",
                "UNITS_PER_LAYER": None,
                "LAYERS_PER_UNIT": None,
                "CONTAINER_PAYLOAD_UNITS": None,
                "MAX_GROSS_WEIGHT_KG": None,
                "SUPERSEDES_ID": None,
                "STATUS_ID": STATUS_ACTIVE,
                "EFFECTIVE_FROM": EFFECTIVE_FROM,
                "EFFECTIVE_TO": None,
                "NOTES": "Engineering default transport for commercial mapping FK",
                "CREATED_AT": NOW,
                "UPDATED_AT": NOW,
            }
        )
        pc_to_tc[pc_id] = tc_id

    # Commercial scenarios for in-scope product mappings only
    scen_id = 1
    for pm in bundle.product_maps:
        if not pm.final_set_code or not pm.final_configuration_id:
            continue
        if "DOMESTIC" in (pm.status or "").upper():
            continue
        pc_id = store.set_code_to_pc_id.get(pm.final_set_code)
        pid = store.product_code_to_id.get(pm.product_code)
        if not pc_id or not pid:
            continue
        tc_id = pc_to_tc.get(pc_id)
        if not tc_id:
            continue
        store.tables["COMMERCIAL_SCENARIO"].append(
            {
                "COMMERCIAL_SCENARIO_ID": str(scen_id),
                "COMMERCIAL_SCENARIO_CODE": f"SC-{pm.product_code}-{pm.final_set_code}",
                "COMMERCIAL_SCENARIO_NAME": f"{pm.product_code} → {pm.final_set_code}",
                "SCENARIO_TYPE_ID": "2",
                "PRODUCT_ID": pid,
                "TRANSPORT_CONFIGURATION_ID": tc_id,
                "CUSTOMER_ID": None,
                "DESTINATION_COUNTRY_ID": "2",
                "INCOTERM_ID": None,
                "TRANSPORT_MODE_ID": None,
                "STATUS_ID": STATUS_ACTIVE,
                "VALID_FROM": EFFECTIVE_FROM,
                "VALID_TO": None,
                "NOTES": (
                    f"IN_SCOPE_FINAL; SOURCE_CFG={pm.source_configuration_id}; "
                    f"MARKET={pm.customer_market}; STATUS={pm.status}"
                ),
                "CREATED_AT": NOW,
                "UPDATED_AT": NOW,
            }
        )
        scen_id += 1

    # Statement control metadata (one per final config) — no shipment fabrications
    for i, cfg in enumerate(bundle.configurations, start=1):
        store.tables["STATEMENT"].append(
            {
                "STATEMENT_ID": str(i),
                "STATEMENT_CODE": cfg.statement_id,
                "STATEMENT_TYPE_ID": "1",
                "LEGAL_ENTITY_ID": "1",
                "COUNTRY_ID": "2",
                "PERIOD_YEAR": 2026,
                "PERIOD_MONTH": None,
                "PERIOD_FROM": EFFECTIVE_FROM,
                "PERIOD_TO": EFFECTIVE_FROM,
                "STATUS_ID": STATUS_ACTIVE,
                "GENERATED_AT": None,
                "APPROVED_BY_PERSON_ID": None,
                "APPROVED_AT": None,
                "NOTES": (
                    f"Control metadata for {cfg.packaging_set_code}; "
                    f"Word not generated; PACKAGING_CONFIGURATION via document link"
                ),
            }
        )

    # Variant basis uniqueness QA
    by_family: dict[str, list[str]] = defaultdict(list)
    for cfg in bundle.configurations:
        parent = ids.normalize_family_code(cfg.packaging_set_code)
        by_family[parent].append(cfg.variant_basis_en.casefold().strip())
    for parent, bases in by_family.items():
        seen = set()
        for b in bases:
            if not b:
                continue
            if b in seen:
                store.blocking_errors.append(f"Duplicate Variant Basis in family {parent}")
            seen.add(b)

    store.stats = {
        "configurations": len(store.tables["PACKAGING_CONFIGURATION"]),
        "bom_lines": len(store.tables["PACKAGING_CONFIGURATION_LINE"]),
        "components": len(store.tables["COMPONENT"]),
        "products": len(store.tables["PRODUCT"]),
        "component_materials": len(store.tables["COMPONENT_MATERIAL"]),
        "technical_files": len(store.tables["TECHNICAL_FILE"]),
        "docs": len(store.tables["DECLARATION_OF_CONFORMITY"]),
        "document_library": len(store.tables["DOCUMENT_LIBRARY"]),
        "commercial_scenarios": len(store.tables["COMMERCIAL_SCENARIO"]),
        "transport_configurations": len(store.tables["TRANSPORT_CONFIGURATION"]),
        "blocking_errors": len(store.blocking_errors),
        "starter": sum(1 for c in bundle.configurations if c.family == "STARTER"),
        "industrial": sum(1 for c in bundle.configurations if c.family == "INDUSTRIAL"),
        "container": sum(1 for c in bundle.configurations if c.family == "CONTAINER"),
    }
    return store


def _line_weight_inputs(bl) -> tuple[float, float]:
    """Return (weight_g, quantity) for WeightService: always qty * weight_g = line grams."""
    basis = (bl.weight_basis or "").upper()
    uom = (bl.uom or "").upper()
    if basis == "RECIPE_QUANTITY_IS_KG" or (uom == "KG" and bl.unit_weight_kg in (None, 0)):
        # quantity is mass in kg
        return 1000.0, float(bl.quantity)
    if uom == "KG" and bl.unit_weight_kg is not None:
        # qty * unit_weight_kg = line kg
        return float(bl.unit_weight_kg) * 1000.0, float(bl.quantity)
    if bl.unit_weight_kg is not None:
        return float(bl.unit_weight_kg) * 1000.0, float(bl.quantity)
    # fallback: derive from line weight
    if bl.quantity:
        return (float(bl.line_weight_kg) / float(bl.quantity)) * 1000.0, float(bl.quantity)
    return float(bl.line_weight_kg) * 1000.0, 1.0


def _canonical_component_weight_g(erp: str, bundle: StagingBundle) -> float | None:
    """Prefer PROCESSED unit weights; mass-input-only components → 1000 g/kg."""
    lines = [bl for bl in bundle.bom_lines if bl.component_code == erp]
    unit_weights = [
        bl.unit_weight_kg
        for bl in lines
        if bl.unit_weight_kg is not None
        and bl.weight_basis not in {"RECIPE_QUANTITY_IS_KG"}
    ]
    if unit_weights:
        # median
        unit_weights.sort()
        mid = unit_weights[len(unit_weights) // 2]
        return float(mid) * 1000.0
    # all mass-input
    if lines and all(
        bl.weight_basis == "RECIPE_QUANTITY_IS_KG" or (bl.uom or "").upper() == "KG"
        for bl in lines
    ):
        return 1000.0
    comp = bundle.components.get(erp)
    if comp and comp.unit_weight_kg is not None:
        return float(comp.unit_weight_kg) * 1000.0
    return None


def _material_id_for_family(family: str) -> str | None:
    f = family.upper()
    mapping = {
        "PLASTIC": "1",  # LDPE as generic plastic token — family only
        "PAPER": "5",
        "PAPER_CARDBOARD": "5",
        "WOOD": "7",
        "METAL": "8",
        "COMPOSITE": "10",
    }
    for k, v in mapping.items():
        if k in f:
            return v
    return None
