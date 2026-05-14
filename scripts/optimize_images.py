"""
Image optimization script for BOT Assets.

Produces three image variants from each source image:
  1. _original.png  — Untouched archival copy (lossless, full resolution)
  2. .png           — Standard variant (2K, max 2048px, quantized)
  3. _optimize.png  — Optimized variant (max 1024px, quantized, small file size)

All variants are derived from the _original file.
"""

import os
import sys
import argparse
import logging
import time
from pathlib import Path
from typing import Optional, Tuple, List
from PIL import Image
import concurrent.futures

try:
    import imagequant
    HAS_IMAGEQUANT = True
except ImportError:
    HAS_IMAGEQUANT = False

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)
logger = logging.getLogger(__name__)

# Configuration
STANDARD_SIZE = 2048   # Maximum dimension for standard variant (2K)
OPTIMIZE_SIZE = 1024   # Maximum dimension for optimized variant
PNG_OPTIMIZE = True    # Enable PNG optimization
WEBP_QUALITY = 85      # WebP quality (0-100)
MAX_WORKERS = min(os.cpu_count() or 2, 8)  # Match CPU cores (CI runners have 2-4)


def _quantize_image(img: Image.Image) -> Image.Image:
    """Quantize an image to 256 colors for PNG size reduction."""
    if HAS_IMAGEQUANT:
        return imagequant.quantize_pil_image(img, dithering_level=1.0, max_colors=256)
    else:
        if img.mode == "RGBA":
            alpha = img.getchannel('A')
            img_p = img.convert('RGB').quantize(colors=256, method=Image.Quantize.MAXCOVERAGE)
            img = img_p.convert('RGBA')
            img.putalpha(alpha)
        elif img.mode != "P":
            img = img.quantize(colors=256, method=Image.Quantize.MAXCOVERAGE)
        return img


def _create_variant(source_img: Image.Image, max_size: int, output_path: str,
                    output_format: str = "png") -> Optional[int]:
    """
    Create a single image variant by resizing and quantizing.
    
    Args:
        source_img: Source PIL Image (will not be modified)
        max_size: Maximum dimension (width or height)
        output_path: Path to save the variant
        output_format: Output format ('png' or 'webp')
        
    Returns:
        File size in bytes of the saved variant, or None if failed
    """
    try:
        img = source_img.copy()

        # Resize if necessary
        if img.width > max_size or img.height > max_size:
            img.thumbnail((max_size, max_size), Image.Resampling.LANCZOS)

        # Optimize and save
        if output_format == "webp":
            if img.mode != "RGBA":
                img = img.convert("RGBA")
            img.save(output_path, 'WEBP', quality=WEBP_QUALITY, method=6)
        else:
            img = _quantize_image(img)
            img.save(output_path, 'PNG', optimize=PNG_OPTIMIZE)

        return os.path.getsize(output_path)
    except Exception as e:
        logger.error(f"Error creating variant {output_path}: {e}")
        return None


def optimize_image(file_path: str, output_format: str = "png",
                   force: bool = False) -> Optional[List[str]]:
    """
    Optimize an image by producing standard (2K) and optimized (1024px) variants.
    
    Both variants are derived from the _original file. If the _original doesn't
    exist yet, the current file is preserved as _original first.
    
    Args:
        file_path: Path to the input image (the base name, e.g. BT01-001-SCR.png)
        output_format: Output format ('png' or 'webp')
        force: If True, regenerate variants even if they already exist
        
    Returns:
        List of paths to created variants, or None if failed/skipped
    """
    try:
        # Check if the file is an image
        if not file_path.lower().endswith(('.png', '.jpg', '.jpeg', '.webp')):
            return None

        # Skip files that are already marked as original or optimize
        basename_lower = os.path.basename(file_path).lower()
        if "_original." in basename_lower or "_optimize." in basename_lower:
            return None

        base, ext = os.path.splitext(file_path)
        original_file_path = f"{base}_original{ext}"
        standard_file_path = file_path  # Same as original name
        optimize_file_path = f"{base}_optimize{ext}"

        # Determine if we need to create the _original
        if not os.path.exists(original_file_path):
            if not os.path.exists(file_path):
                return None
            # Rename the current file to the _original path
            logger.info(f"Preserving original: {os.path.basename(file_path)} -> {os.path.basename(original_file_path)}")
            os.rename(file_path, original_file_path)
        else:
            # _original already exists — check if we should re-derive
            standard_exists = os.path.exists(standard_file_path)
            optimize_exists = os.path.exists(optimize_file_path)
            if standard_exists and optimize_exists and not force:
                logger.debug(f"Skipping (all variants exist): {os.path.basename(file_path)}")
                return None

        # Open the image from the preserved original path
        img = Image.open(original_file_path)
        img.load()

        original_size = os.path.getsize(original_file_path)
        logger.info(
            f"Optimizing: {os.path.basename(file_path)} "
            f"({img.width}x{img.height}, {original_size / 1024 / 1024:.2f} MB)"
        )

        created = []

        # === Cascade resize approach ===
        # 1. Resize original → 2048px (standard)
        # 2. Resize the 2048px copy → 1024px (optimize)
        # This avoids processing the full-res original twice.

        # Step 1: Resize to standard size
        std_img = img.copy()
        if std_img.width > STANDARD_SIZE or std_img.height > STANDARD_SIZE:
            std_img.thumbnail((STANDARD_SIZE, STANDARD_SIZE), Image.Resampling.LANCZOS)

        # Step 2: Cascade — resize the already-smaller standard image to optimize size
        opt_img = std_img.copy()
        if opt_img.width > OPTIMIZE_SIZE or opt_img.height > OPTIMIZE_SIZE:
            opt_img.thumbnail((OPTIMIZE_SIZE, OPTIMIZE_SIZE), Image.Resampling.LANCZOS)

        # Free full-res original from memory immediately
        img.close()

        # Step 3: Quantize and save both variants
        try:
            if output_format == "webp":
                if std_img.mode != "RGBA":
                    std_img = std_img.convert("RGBA")
                std_img.save(standard_file_path, 'WEBP', quality=WEBP_QUALITY, method=6)
            else:
                std_img = _quantize_image(std_img)
                std_img.save(standard_file_path, 'PNG', optimize=PNG_OPTIMIZE)

            std_size = os.path.getsize(standard_file_path)
            reduction = ((original_size - std_size) / original_size) * 100
            logger.info(
                f"  ✅ Standard: {os.path.basename(standard_file_path)} "
                f"({std_size / 1024:.0f} KB, {reduction:.1f}% smaller)"
            )
            created.append(standard_file_path)
        except Exception as e:
            logger.error(f"Error creating standard variant: {e}")

        try:
            if output_format == "webp":
                if opt_img.mode != "RGBA":
                    opt_img = opt_img.convert("RGBA")
                opt_img.save(optimize_file_path, 'WEBP', quality=WEBP_QUALITY, method=6)
            else:
                opt_img = _quantize_image(opt_img)
                opt_img.save(optimize_file_path, 'PNG', optimize=PNG_OPTIMIZE)

            opt_size = os.path.getsize(optimize_file_path)
            reduction = ((original_size - opt_size) / original_size) * 100
            logger.info(
                f"  ✅ Optimize: {os.path.basename(optimize_file_path)} "
                f"({opt_size / 1024:.0f} KB, {reduction:.1f}% smaller)"
            )
            created.append(optimize_file_path)
        except Exception as e:
            logger.error(f"Error creating optimize variant: {e}")

        return created if created else None

    except Exception as e:
        logger.error(f"Error optimizing {file_path}: {e}")
        return None


def process_directory(target_dir: str, output_format: str = "png",
                      force: bool = False) -> Tuple[int, int, int]:
    """
    Process all images in a directory.
    
    Args:
        target_dir: Directory to scan for images
        output_format: Output format ('png' or 'webp')
        force: If True, regenerate variants even if they already exist
        
    Returns:
        Tuple of (success_count, skipped_count, error_count)
    """
    if not os.path.exists(target_dir):
        logger.error(f"Directory not found: {target_dir}")
        return 0, 0, 0

    logger.info(f"Scanning directory: {target_dir}")

    # Find all candidate images (base names only, not _original or _optimize)
    files_to_process = []
    for root, _, files in os.walk(target_dir):
        for file in files:
            if file.lower().endswith(('.png', '.jpg', '.jpeg', '.webp')):
                fl = file.lower()
                if "_original." not in fl and "_optimize." not in fl:
                    files_to_process.append(os.path.join(root, file))

    # Also find _original files that don't have a base name counterpart yet
    # (e.g., if the standard variant was deleted but _original remains)
    for root, _, files in os.walk(target_dir):
        for file in files:
            if "_original." in file.lower():
                # Derive the base name
                base_name = file.replace("_original", "")
                base_path = os.path.join(root, base_name)
                if base_path not in files_to_process:
                    files_to_process.append(base_path)

    logger.info(f"Found {len(files_to_process)} images to optimize")
    
    if not files_to_process:
        logger.info("No images to process")
        return 0, 0, 0

    start_time = time.time()
    success_count = 0
    skipped_count = 0
    error_count = 0

    # Use ThreadPoolExecutor for parallel processing
    # Pillow releases GIL during image I/O operations
    with concurrent.futures.ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
        futures = {
            executor.submit(optimize_image, file_path, output_format, force): file_path
            for file_path in files_to_process
        }
        
        for future in concurrent.futures.as_completed(futures):
            file_path = futures[future]
            try:
                result = future.result()
                if result:
                    success_count += 1
                else:
                    skipped_count += 1
            except Exception as e:
                logger.error(f"Failed to process {file_path}: {e}")
                error_count += 1

    end_time = time.time()
    
    logger.info(
        f"\n✅ Optimization completed in {end_time - start_time:.2f}s: "
        f"{success_count} succeeded, {skipped_count} skipped, {error_count} errors"
    )
    
    return success_count, skipped_count, error_count


def main():
    """Main entry point for image optimization."""
    parser = argparse.ArgumentParser(description="Optimize BOT card images into multiple variants.")
    parser.add_argument("folder", nargs="?", default=None,
                        help="Specific folder name inside assets/ to optimize (optional)")
    parser.add_argument("format", nargs="?", default="png", choices=["png", "webp"],
                        help="Output format (default: png)")
    parser.add_argument("--force", action="store_true",
                        help="Regenerate all variants even if they already exist")
    args = parser.parse_args()

    script_dir = os.path.dirname(os.path.abspath(__file__))
    base_dir = os.path.dirname(script_dir)
    assets_dir = os.path.join(base_dir, 'assets')

    if args.folder:
        target_dir = os.path.join(assets_dir, args.folder)
        if not os.path.exists(target_dir):
            logger.error(f"Directory not found: {target_dir}")
            return
    else:
        target_dir = assets_dir
        if not os.path.exists(target_dir):
            logger.error(f"Directory not found: {target_dir}")
            return

    success, skipped, errors = process_directory(target_dir, args.format, args.force)
    
    if errors > 0:
        logger.warning(f"Completed with {errors} error(s)")
        sys.exit(1)


if __name__ == "__main__":
    main()

