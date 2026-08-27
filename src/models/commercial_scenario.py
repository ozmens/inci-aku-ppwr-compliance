"""Commercial scenario domain model (Incoterms / DoC variant driver)."""

from __future__ import annotations

from dataclasses import dataclass


@dataclass(slots=True)
class CommercialScenario:
    commercial_scenario_id: str
    code: str
    name: str
    product_id: str
    transport_configuration_id: str
    destination_country: str | None = None
    incoterm: str | None = None
    customer_id: str | None = None
    customer_name: str | None = None
    status: str = "ACTIVE"
