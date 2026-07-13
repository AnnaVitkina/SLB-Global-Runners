"""Map origin country + supplier to FRED carrier name for Supplier name column."""

from __future__ import annotations

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


def map_fred_supplier_names(
    origin_countries: pd.Series,
    suppliers: pd.Series,
    *,
    transport_mode: str | None = None,
) -> pd.Series:
    return pd.Series(
        [
            lookup_fred_supplier_name(origin, supplier, transport_mode=transport_mode)
            for origin, supplier in zip(origin_countries, suppliers, strict=False)
        ],
        index=origin_countries.index,
        dtype="object",
    )
