#!/usr/bin/env python3
"""
Merge all card-data.json files into a single consolidated JSON file.

Reads asset-map.json to find all card-data.json files, merges them into
one array, and adds metadata about the source set.

Usage:
    python merge_datasets.py                          # Merge all, output to assets/all-cards.json
    python merge_datasets.py --output assets/full.json # Custom output path
    python merge_datasets.py --exclude-source-field   # Don't add _source field
    python merge_datasets.py --include BT01 BT02       # Only merge specific sets
    python merge_datasets.py --exclude CC01 CC02       # Exclude specific sets
"""

import os
import json
import logging
import argparse
from pathlib import Path
from typing import Dict, List, Optional

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)
logger = logging.getLogger(__name__)

CARD_DATA_FILENAME = "card-data.json"
ALL_CARDS_FILENAME = "all-cards.json"


def load_asset_map(project_root):
    """Load and parse the asset-map.json file."""
    asset_map_path = os.path.join(project_root, "asset-map.json")
    if not os.path.exists(asset_map_path):
        raise FileNotFoundError(f"asset-map.json not found at {asset_map_path}")

    with open(asset_map_path, "r", encoding="utf-8") as f:
        return json.load(f)


def load_dataset(dataset_path: str) -> List[Dict]:
    """Load a single dataset.json file. Returns empty list if not found."""
    if not os.path.exists(dataset_path):
        logger.warning(f"⚠️  Not found: {dataset_path}")
        return []

    with open(dataset_path, "r", encoding="utf-8") as f:
        data = json.load(f)
        if not isinstance(data, list):
            logger.warning(f"⚠️  Expected array, got {type(data).__name__} in {dataset_path}")
            return []
        return data


def merge_datasets(
    project_root: str,
    output_path: Optional[str] = None,
    exclude_source_field: bool = False,
    include_sets: Optional[List[str]] = None,
    exclude_sets: Optional[List[str]] = None
) -> tuple:
    """
    Merge all dataset.json files into one.

    Args:
        project_root: Root directory of the project
        output_path: Output file path (default: assets/all-cards.json)
        exclude_source_field: If True, don't add _source metadata field
        include_sets: List of set names to include (None = all)
        exclude_sets: List of set names to exclude (None = exclude nothing)
        
    Returns:
        Tuple of (total_cards, total_sets_processed)
    """
    if output_path is None:
        output_path = os.path.join(project_root, "assets", ALL_CARDS_FILENAME)

    asset_map = load_asset_map(project_root)

    merged_cards = []
    stats = {}
    total_sets_processed = 0

    logger.info(f"📋 Found {len(asset_map)} sets in asset-map.json")
    if include_sets:
        logger.info(f"🎯 Including only: {', '.join(include_sets)}")
    if exclude_sets:
        logger.info(f"🚫 Excluding: {', '.join(exclude_sets)}")
    logger.info("")

    for asset_entry in asset_map:
        set_name = asset_entry["name"]
        # Update path to use new card-data.json filename
        old_path = asset_entry["path"]
        if old_path.endswith("/dataset.json"):
            new_path = old_path.replace("/dataset.json", f"/{CARD_DATA_FILENAME}")
        else:
            new_path = old_path
        dataset_path = os.path.join(project_root, new_path)

        # Apply include/exclude filters
        if include_sets and not any(filter_term in set_name for filter_term in include_sets):
            logger.info(f"  ⏭️  Skipped: {set_name}")
            continue

        if exclude_sets and any(filter_term in set_name for filter_term in exclude_sets):
            logger.info(f"  ⏭️  Excluded: {set_name}")
            continue

        # Skip the merged dataset itself if it's in the asset map
        if set_name == "All Cards":
            continue

        logger.info(f"📦 Loading: {set_name}")
        cards = load_dataset(dataset_path)

        if not cards:
            logger.warning(f"  ⚠️  No cards found")
            continue

        # Extract the folder path from the dataset path (e.g., "assets/BT01 - Welcome ตลิ่งชัน")
        # asset_entry["path"] is like "assets/BT01 - Welcome ตลิ่งชัน/card-data.json"
        source_folder = str(Path(new_path).parent).replace("\\", "/")

        # Add source metadata and fix ImagePath to include full relative path
        if not exclude_source_field:
            for card in cards:
                if isinstance(card, dict):
                    card["_source"] = set_name
                    # Update ImagePath to include the source folder path
                    if "ImagePath" in card:
                        original_path = card["ImagePath"]
                        # Only prepend if it's not already a path
                        if "/" not in original_path and "\\" not in original_path:
                            card["ImagePath"] = f"{source_folder}/{original_path}"

        merged_cards.extend(cards)
        stats[set_name] = len(cards)
        total_sets_processed += 1
        logger.info(f"  ✅ {len(cards)} cards")

    # Write merged output
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(merged_cards, f, ensure_ascii=False, indent=2)

    # Log summary
    logger.info("")
    logger.info("=" * 60)
    logger.info("📊 MERGE SUMMARY")
    logger.info("=" * 60)
    logger.info(f"  Total sets processed: {total_sets_processed}")
    logger.info(f"  Total cards merged:   {len(merged_cards)}")
    logger.info("")
    logger.info("  Cards per set:")
    for set_name, count in sorted(stats.items()):
        logger.info(f"    {set_name:40s} {count:5d}")
    logger.info("=" * 60)
    logger.info(f"✅ Merged dataset saved to: {output_path}")
    logger.info(f"📁 File size: {os.path.getsize(output_path) / 1024:.1f} KB")

    return len(merged_cards), total_sets_processed


def main() -> int:
    """Main entry point for dataset merging."""
    parser = argparse.ArgumentParser(
        description="Merge all card-data.json files into a single consolidated JSON file.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python merge_datasets.py                          # Merge all sets
  python merge_datasets.py --output full.json       # Custom output path
  python merge_datasets.py --exclude-source-field   # Don't add _source field
  python merge_datasets.py --include BT01 BT02       # Only merge BT01 and BT02
  python merge_datasets.py --exclude CC01            # Exclude CC01
        """
    )
    parser.add_argument(
        "--output", "-o",
        type=str,
        default=None,
        help="Output file path (default: assets/all-cards.json)"
    )
    parser.add_argument(
        "--exclude-source-field",
        action="store_true",
        help="Don't add _source metadata field to cards"
    )
    parser.add_argument(
        "--include",
        nargs="+",
        metavar="SET",
        help="Only merge these sets (partial match supported)"
    )
    parser.add_argument(
        "--exclude",
        nargs="+",
        metavar="SET",
        help="Exclude these sets (partial match supported)"
    )

    args = parser.parse_args()

    # Determine project root (parent of scripts folder)
    project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

    try:
        total_cards, total_sets = merge_datasets(
            project_root=project_root,
            output_path=args.output,
            exclude_source_field=args.exclude_source_field,
            include_sets=args.include,
            exclude_sets=args.exclude
        )
    except FileNotFoundError as e:
        logger.error(f"❌ Error: {e}")
        return 1
    except Exception as e:
        logger.error(f"❌ Unexpected error: {e}")
        import traceback
        traceback.print_exc()
        return 1

    return 0


if __name__ == "__main__":
    exit(main())
