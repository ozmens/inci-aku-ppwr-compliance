"""Shipment domain model."""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date


@dataclass(slots=True)
class Shipment:
    shipment_id: str
    shipment_number: str
    commercial_scenario_id: str
    packaging_configuration_id: str
    transport_configuration_id: str
    qty_product_units: float
    plant_id: str | None = None
    ship_date: date | None = None
    lot_number: str | None = None  # maps to EXTERNAL_REF in frozen schema
    destination_country: str | None = None
    status: str = "DRAFT"
