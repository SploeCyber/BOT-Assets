#!/usr/bin/env python3
"""
Card data overrides and synonym normalization.

Contains hard-coded overrides for specific cards and
canonical forms for synonym field values.
"""

import logging
from typing import Any, Dict, List

logger = logging.getLogger(__name__)

# =============================================================================
# SYNONYM GROUPS - Normalize field values to canonical forms
# =============================================================================

EX_SYNONYM_GROUPS: List[Dict[str, List[str]]] = [
    {"canonical": "โดนใจ", "aliases": ["โดนใจ", "#โดนใจ"]},
    {"canonical": "Only #1", "aliases": ["Only #1", "#Only1", "Only1"]},
]

SYMBOL_SYNONYM_GROUPS: List[Dict[str, List[str]]] = [
    {"canonical": "กะปอม", "aliases": ["กะปอม", "กระปอม"]},
    {"canonical": "กลยุทธ์", "aliases": ["กลยุทธ์", "กลยุททหาร", "กลยุทธ"]},
    {"canonical": "ฤๅษี", "aliases": ["ฤๅษี", "ฤษี", "ฤาษี", "คาถาฤษี", "คาถาฤาษี", "คาถาฤๅษี"]},
    {"canonical": "เครื่องจักร", "aliases": ["เครื่องจักร", "หุ่นยนต์", "เครื่องจักร์"]},
    {"canonical": "คน", "aliases": ["คน", "ตน"]},
    {"canonical": "ชาวต่างชาติ", "aliases": ["ชาวต่างชาติ", "ต่างชาติ"]},
    {"canonical": "รัททาทุย", "aliases": ["รัททาทุย", "รักททาทุย"]},
    {"canonical": "สัตว์มหัศจรรย์", "aliases": ["สัตว์มหัศจรรย์", "สัตว์วิเศษ"]},
    {"canonical": "เอเลี่ยน", "aliases": ["เอเลี่ยน", "เอเลื่ยน"]},
    {"canonical": "นรก", "aliases": ["นรก", "นัก"]},
    {"canonical": "สัตว์", "aliases": ["สัตว์", "สัตย์"]},
]

# =============================================================================
# CARD-SPECIFIC OVERRIDES
# =============================================================================

CARD_OVERRIDES: Dict[str, Dict[str, Any]] = {
    # SD07-017: Add SubType "React"
    "SD07-017": {
        "set": {
            "SubType": "React"
        }
    },
    # SD07-016: Add SubType "React"
    "SD07-016": {
        "set": {
            "SubType": "React"
        }
    },
    # BT02-030: Remove "Ex" field (was "Only #1")
    "BT02-030": {
        "remove": ["Ex"]
    },
    # BT03-033: cost 4
    "BT03-033": {
        "set": {
            "Cost": 4
        }
    },
    # BT09-022: gem 2
    "BT09-022": {
        "set": {
            "Gem": 2
        }
    },
    # BT09-055: rare "R"
    "BT09-055": {
        "set": {
            "Rare": "R"
        }
    }
}


# =============================================================================
# HELPER FUNCTIONS
# =============================================================================

def _normalize_value(value: Any, synonym_groups: List[Dict[str, Any]]) -> Any:
    """
    Normalize a value using synonym groups.
    
    Args:
        value: The value to normalize
        synonym_groups: List of synonym group dictionaries
        
    Returns:
        The canonical form if found, otherwise original value
    """
    if not value:
        return value

    value_str = str(value).strip()

    for group in synonym_groups:
        if value_str in group["aliases"]:
            return group["canonical"]

    return value_str


def normalize_ex(value: Any) -> Any:
    """
    Normalize Ex field value to canonical form.
    
    Args:
        value: The Ex field value to normalize
        
    Returns:
        Normalized value
    """
    return _normalize_value(value, EX_SYNONYM_GROUPS)


def normalize_symbol(value: Any) -> Any:
    """
    Normalize Symbol field value to canonical form.
    
    Args:
        value: The Symbol field value to normalize
        
    Returns:
        Normalized value
    """
    return _normalize_value(value, SYMBOL_SYNONYM_GROUPS)


def apply_card_overrides(card_data: Dict[str, Any], print_code: str) -> Dict[str, Any]:
    """
    Apply hard-coded overrides and normalization to a card.

    Args:
        card_data: The card dictionary to modify
        print_code: The Print code of the card (e.g., "SD07-017")
        
    Returns:
        The modified card_data dictionary
    """
    # 1. Normalize Ex field
    if "Ex" in card_data and card_data["Ex"]:
        card_data["Ex"] = normalize_ex(card_data["Ex"])

    # 2. Normalize Symbol field (at top level only)
    if "Symbol" in card_data and card_data["Symbol"]:
        card_data["Symbol"] = normalize_symbol(card_data["Symbol"])

    # 3. Apply card-specific overrides
    # Get base print code (e.g., "SD07-017" from "SD07-017-2")
    parts = print_code.split("-")
    base_print = "-".join(parts[:2]) if len(parts) >= 2 else print_code

    if base_print in CARD_OVERRIDES:
        override = CARD_OVERRIDES[base_print]

        # Apply "set" overrides
        if "set" in override:
            for key, value in override["set"].items():
                card_data[key] = value

        # Apply "remove" overrides
        if "remove" in override:
            for key in override["remove"]:
                card_data.pop(key, None)

    return card_data
