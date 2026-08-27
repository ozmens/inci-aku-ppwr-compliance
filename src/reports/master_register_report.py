"""Master register report stub for future Packaging Set / Final Config register."""

from __future__ import annotations

from dataclasses import dataclass

from models.packaging_configuration import PackagingConfiguration


@dataclass(slots=True)
class MasterRegisterReport:
    configurations: list[PackagingConfiguration]

    def rows(self) -> list[dict[str, str]]:
        return [
            {
                "packaging_set_code": c.packaging_set_code,
                "final_configuration_id": c.final_configuration_id,
                "parent_family_code": c.parent_family_code,
                "variant_number": str(c.variant_number),
                "variant_basis_tr": c.variant_basis_tr,
                "bom_signature": c.bom_signature or "",
                "status": c.status,
            }
            for c in self.configurations
        ]
