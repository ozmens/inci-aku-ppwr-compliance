"""
Lightweight domain model stubs.

These dataclasses describe in-memory objects for future builders/validators.
They do not perform Excel I/O.
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date
from decimal import Decimal


@dataclass(slots=True)
class Component:
    component_id: int
    component_code: str
    component_name: str
    weight_g: Decimal
    ownership_type_id: int
    packaging_level_id: int
    packaging_function_id: int
    component_type_id: int
    status_id: int


@dataclass(slots=True)
class Product:
    product_id: int
    product_code: str
    product_name: str
    product_category_id: int
    net_weight_g: Decimal
    status_id: int


@dataclass(slots=True)
class CommercialScenario:
    commercial_scenario_id: int
    commercial_scenario_code: str
    product_id: int
    transport_configuration_id: int
    destination_country_id: int
    scenario_type_id: int
    customer_id: int | None
    valid_from: date
    status_id: int


@dataclass(slots=True)
class Shipment:
    shipment_id: int
    shipment_number: str
    commercial_scenario_id: int
    plant_id: int
    ship_date: date
    qty_product_units: Decimal
    packaging_configuration_id: int
    transport_configuration_id: int
    status_id: int
