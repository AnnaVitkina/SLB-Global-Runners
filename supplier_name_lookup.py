"""Map origin country + supplier to FRED carrier name for Supplier name column."""

from __future__ import annotations

import re
from dataclasses import dataclass

import pandas as pd

ANY_ORIGIN = frozenset({"*"})


def _cell_text(value: object) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip()


def _normalize_origin(value: object) -> str:
    return _cell_text(value).upper()


def _canonical_supplier(value: object) -> str:
    text = _cell_text(value).lower()
    if not text:
        return ""

    aliases = (
        ("aramex", "aramex"),
        ("bws", "bws"),
        ("bluewater", "bws"),
        ("dhl", "dhl"),
        ("dsv", "dsv"),
        ("expeditors", "expeditors"),
        ("geodis", "geodis"),
        ("kuehne", "kuehne nagel"),
        ("rulewave", "rulewave"),
        ("crane worldwide", "crane"),
        ("pentagon freight", "pentagon"),
        ("db schenker", "db schenker"),
        ("bollore", "bollore"),
    )
    for needle, canonical in aliases:
        if needle in text:
            return canonical
    return text


@dataclass(frozen=True)
class SupplierCarrierRule:
    supplier_id: str
    origins: frozenset[str]
    carrier_name: str
    carrier_name_ocean: str | None = None
    carrier_name_air: str | None = None

    def resolve_carrier_name(self, transport_mode: str | None) -> str:
        mode = (transport_mode or "").strip().lower()
        if mode == "ocean" and self.carrier_name_ocean:
            return self.carrier_name_ocean
        if mode == "air" and self.carrier_name_air:
            return self.carrier_name_air
        return self.carrier_name


SUPPLIER_CARRIER_RULES: tuple[SupplierCarrierRule, ...] = (
    SupplierCarrierRule("aramex", ANY_ORIGIN, "ARX AE; ARX SG; ARX US; ARX CN; ARX HK; ARX IE"),
    SupplierCarrierRule("bws", frozenset({"CN"}), "Bluewater CN"),
    SupplierCarrierRule("bws", frozenset({"VN", "US", "SG", "MY"}), "Bluewater US"),
    SupplierCarrierRule("bws", frozenset({"FR", "NL"}), "Bluewater NL"),
    SupplierCarrierRule("dhl", frozenset({"CN"}), "DHL CN"),
    SupplierCarrierRule("dhl", frozenset({"IN", "SG"}), "DHL SG"),
    SupplierCarrierRule("dhl", frozenset({"JP"}), "DHL JP"),
    SupplierCarrierRule("dhl", frozenset({"US"}), "DHL UX"),
    SupplierCarrierRule("dhl", frozenset({"IT"}), "DHL NL"),
    SupplierCarrierRule("dhl", frozenset({"GB"}), "DHL UK"),
    SupplierCarrierRule("dsv", frozenset({"AE"}), "DSV AE"),
    SupplierCarrierRule("dsv", frozenset({"CN"}), "DSV CN"),
    SupplierCarrierRule("dsv", frozenset({"RO"}), "DSV UK"),
    SupplierCarrierRule("dsv", frozenset({"US"}), "DSV US"),
    SupplierCarrierRule("dsv", ANY_ORIGIN, "DSV US"),
    SupplierCarrierRule("expeditors", frozenset({"IN"}), "EI AE Int"),
    SupplierCarrierRule(
        "expeditors",
        frozenset({"DE", "FR", "IN", "SI", "SK", "TH", "TW", "US", "VN", "AT", "LU"}),
        "EI IAH",
    ),
    SupplierCarrierRule("expeditors", frozenset({"RO", "IT"}), "EI NL"),
    SupplierCarrierRule("expeditors", frozenset({"MY", "SG"}), "EI SG"),
    SupplierCarrierRule("expeditors", frozenset({"CN"}), "EI SHA"),
    SupplierCarrierRule("expeditors", frozenset({"GB"}), "EI UK"),
    SupplierCarrierRule("expeditors", frozenset({"JP"}), "EI JP"),
    SupplierCarrierRule("geodis", frozenset({"AE", "IN"}), "Geodis AE"),
    SupplierCarrierRule("geodis", frozenset({"CN"}), "Geodis CN"),
    SupplierCarrierRule("geodis", frozenset({"SG"}), "Geodis SG"),
    SupplierCarrierRule("geodis", frozenset({"US"}), "Geodis US"),
    SupplierCarrierRule("geodis", frozenset({"FR", "IT"}), "Geodis NL"),
    SupplierCarrierRule("kuehne nagel", frozenset({"CN"}), "K&N SHA"),
    SupplierCarrierRule("kuehne nagel", frozenset({"IN"}), "K&N AE"),
    SupplierCarrierRule("kuehne nagel", frozenset({"US", "FR"}), "K&N"),
    SupplierCarrierRule("kuehne nagel", frozenset({"SG"}), "K&N SD"),
    SupplierCarrierRule("rulewave", frozenset({"BE", "NL"}), "RLWV"),
    SupplierCarrierRule("rulewave", frozenset({"CN", "SG"}), "RLWV SG"),
    SupplierCarrierRule(
        "rulewave",
        frozenset({"US"}),
        "RLWUS",
        carrier_name_ocean="RLWUS/RLWV",
        carrier_name_air="RLWUS",
    ),
    SupplierCarrierRule("rulewave", frozenset({"AE"}), "RLWV AE"),
    SupplierCarrierRule("crane", frozenset({"DE", "NL", "GB"}), "CRANE NL"),
    SupplierCarrierRule("crane", frozenset({"IN"}), "CRANE SG"),
    SupplierCarrierRule("pentagon", frozenset({"US"}), "PNTG HOU"),
    SupplierCarrierRule("pentagon", frozenset({"AE"}), "PNTG AE"),
    SupplierCarrierRule("db schenker", frozenset({"US"}), "SCH US"),
    SupplierCarrierRule("bollore", frozenset({"VN", "US"}), "Bollore US"),
    SupplierCarrierRule("bollore", frozenset({"MX"}), "Bollore MX"),
)


def _rule_specificity(origins: frozenset[str]) -> int:
    if origins == ANY_ORIGIN:
        return 10_000
    return len(origins)


def lookup_fred_supplier_name(
    origin_country: object,
    supplier: object,
    *,
    transport_mode: str | None = None,
) -> str:
    origin = _normalize_origin(origin_country)
    supplier_id = _canonical_supplier(supplier)
    if not origin or not supplier_id:
        return ""

    matching_rules = [
        rule
        for rule in SUPPLIER_CARRIER_RULES
        if rule.supplier_id == supplier_id
        and (rule.origins == ANY_ORIGIN or origin in rule.origins)
    ]
    if not matching_rules:
        return ""

    best_rule = min(matching_rules, key=lambda rule: _rule_specificity(rule.origins))
    return best_rule.resolve_carrier_name(transport_mode)


_DSV_CARRIER_PATTERN = re.compile(r"^DSV(\s|$|/|-)", re.IGNORECASE)


def is_dsv_supplier_text(value: object) -> bool:
    return _canonical_supplier(value) == "dsv"


def is_dsv_carrier_name(value: object) -> bool:
    text = _cell_text(value)
    if not text:
        return False
    upper = text.upper()
    if upper in {"DSV", "DSV US/DSV CN"}:
        return True
    if _DSV_CARRIER_PATTERN.match(text):
        return True
    return is_dsv_supplier_text(value)


def apply_dsv_carrier_display(
    carrier_name: object,
    origin_country: object,
    destination_country: object | None = None,
) -> str:
    text = _cell_text(carrier_name)
    if not text and not is_dsv_supplier_text(carrier_name):
        return text
    if not is_dsv_carrier_name(text) and not is_dsv_supplier_text(carrier_name):
        return text

    origin = _normalize_origin(origin_country)
    destination = _normalize_origin(destination_country)
    if origin == "CN" and destination == "MX":
        return "DSV US/DSV CN"
    return "DSV US"


def coerce_dsv_display_name(
    value: object,
    origin_country: object,
    destination_country: object | None = None,
) -> str:
    text = _cell_text(value)
    if not is_dsv_carrier_name(text) and not is_dsv_supplier_text(value):
        return text
    display = apply_dsv_carrier_display(text or "DSV", origin_country, destination_country)
    if _cell_text(display).upper() == "DSV":
        return "DSV US"
    return display


def lookup_fred_supplier_display_name(
    origin_country: object,
    supplier: object,
    *,
    transport_mode: str | None = None,
    destination_country: object | None = None,
) -> str:
    resolved = lookup_fred_supplier_name(
        origin_country,
        supplier,
        transport_mode=transport_mode,
    )
    if not resolved and is_dsv_supplier_text(supplier):
        resolved = "DSV"

    display = coerce_dsv_display_name(
        resolved or supplier,
        origin_country,
        destination_country,
    )
    return display


def map_fred_supplier_names(
    origin_countries: pd.Series,
    suppliers: pd.Series,
    *,
    transport_mode: str | None = None,
    destination_countries: pd.Series | None = None,
) -> pd.Series:
    if destination_countries is None:
        destinations: list[object] = [None] * len(origin_countries)
    else:
        destinations = list(destination_countries)

    return pd.Series(
        [
            lookup_fred_supplier_display_name(
                origin,
                supplier,
                transport_mode=transport_mode,
                destination_country=destination,
            )
            for origin, supplier, destination in zip(
                origin_countries,
                suppliers,
                destinations,
                strict=False,
            )
        ],
        index=origin_countries.index,
        dtype="object",
    )
