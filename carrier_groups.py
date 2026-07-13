"""Carrier group definitions for region-based sea cost blocks."""

from __future__ import annotations

CARRIER_GROUP_HOU = "HOU"
CARRIER_GROUP_MEA = "MEA"
CARRIER_GROUP_SC = "SC"
CARRIER_GROUP_RTM = "RTM"

HOU_ORIGIN_COUNTRIES = frozenset({"US", "CA", "MX"})
MEA_ORIGIN_COUNTRIES = frozenset({"AE"})
SC_ORIGIN_COUNTRIES = frozenset({"CN", "SG", "JP"})
RTM_ORIGIN_COUNTRIES = frozenset(
    {
        "AT",
        "BE",
        "BG",
        "CH",
        "CY",
        "CZ",
        "DE",
        "DK",
        "EE",
        "ES",
        "FI",
        "FR",
        "GB",
        "GR",
        "HR",
        "HU",
        "IE",
        "IS",
        "IT",
        "LI",
        "LT",
        "LU",
        "LV",
        "MT",
        "NL",
        "NO",
        "PL",
        "PT",
        "RO",
        "SE",
        "SI",
        "SK",
    }
)

_ORIGIN_TO_GROUP: dict[str, str] = {}
for origin in HOU_ORIGIN_COUNTRIES:
    _ORIGIN_TO_GROUP[origin] = CARRIER_GROUP_HOU
for origin in MEA_ORIGIN_COUNTRIES:
    _ORIGIN_TO_GROUP[origin] = CARRIER_GROUP_MEA
for origin in SC_ORIGIN_COUNTRIES:
    _ORIGIN_TO_GROUP[origin] = CARRIER_GROUP_SC
for origin in RTM_ORIGIN_COUNTRIES:
    _ORIGIN_TO_GROUP[origin] = CARRIER_GROUP_RTM


def carrier_group_for_origin(origin_country: object) -> str | None:
    if origin_country is None:
        return None
    origin = str(origin_country).strip().upper()
    if not origin:
        return None
    return _ORIGIN_TO_GROUP.get(origin)


def is_arx_carrier_supplier_name(supplier_name: object) -> bool:
    """True when supplier resolves to ARX (ARX AE merges into ARX via conditions)."""
    if supplier_name is None:
        return False
    text = str(supplier_name).strip()
    if not text:
        return False

    from build_conditions import normalize_condition_value

    return normalize_condition_value(text) == "ARX"


def is_dzs_aei_carrier_supplier_name(supplier_name: object) -> bool:
    """True when supplier resolves to EI AE Int (merged into EI AE Int/EI IAH)."""
    if supplier_name is None:
        return False
    text = str(supplier_name).strip()
    if not text:
        return False

    from build_conditions import normalize_condition_value

    normalized = normalize_condition_value(text)
    return normalized == "EI AE Int/EI IAH" or text == "EI AE Int"
