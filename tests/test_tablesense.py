"""
TableSense Dataset Evaluation for SpreadsheetLLM

This script evaluates SpreadsheetLLM's table detection performance
on the TableSense dataset by comparing detected regions with ground truth annotations.

Metrics:
- Precision: Correct detections / All detections
- Recall: Correct detections / All ground truth tables
- F1 Score: Harmonic mean of precision and recall
- IoU: Intersection over Union for region overlap
"""

import json
import logging
import os
import tarfile
import tempfile
from datetime import datetime
from pathlib import Path
from typing import List, Tuple

from huggingface_hub import hf_hub_download
from rich.console import Console
from rich.logging import RichHandler
from rich.table import Table

from spreadsheet_llm import SpreadsheetLLMWrapper
from spreadsheet_llm.cell_range_utils import (
    col_to_index,
    index_to_col,
    parse_excel_range,
    box_to_range,
)

# Initialize Rich console
console = Console()

# Configuration: Enable/disable detailed logs from spreadsheet_llm modules
ENABLE_SPREADSHEET_LLM_LOGS = True  # Set to True to see detailed compression logs

# Configure logging with timestamp-based log file
log_dir = Path("logs")
log_dir.mkdir(exist_ok=True)
timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
log_file = log_dir / f"tablesense_evaluation_{timestamp}.log"

# Create file formatter (rich handler has its own formatting for console)
file_formatter = logging.Formatter(
    "%(asctime)s - [%(levelname)s] - %(name)s - %(message)s"
)

# Configure root logger to capture logs from all modules
root_logger = logging.getLogger()
root_logger.setLevel(logging.DEBUG)  # Capture all levels, handlers will filter

# Rich console handler - shows INFO and above with rich formatting
rich_handler = RichHandler(
    console=console,
    level=logging.INFO,
    show_time=True,
    rich_tracebacks=True,  # Enable beautiful tracebacks
)
root_logger.addHandler(rich_handler)

# File handler - saves DEBUG and above for all modules
file_handler = logging.FileHandler(log_file, encoding="utf-8")
file_handler.setLevel(logging.DEBUG)
file_handler.setFormatter(file_formatter)
root_logger.addHandler(file_handler)

# Get logger for this module
logger = logging.getLogger(__name__)

# Configure spreadsheet_llm module loggers based on the switch
if not ENABLE_SPREADSHEET_LLM_LOGS:
    # Suppress INFO and DEBUG logs from spreadsheet_llm modules
    # Only show WARNING and above (errors will still be visible)
    logging.getLogger("spreadsheet_llm").setLevel(logging.WARNING)
    logger.info(
        "SpreadsheetLLM detailed logs: DISABLED (set ENABLE_SPREADSHEET_LLM_LOGS=True to enable)"
    )
else:
    logger.info("SpreadsheetLLM detailed logs: ENABLED")

logger.info(f"Logging to file: {log_file}")


def calculate_iou(
    box1: Tuple[int, int, int, int], box2: Tuple[int, int, int, int]
) -> float:
    """Calculate Intersection over Union (IoU) between two boxes.

    Args:
        box1: (row_start, col_start, row_end, col_end) - 0-indexed
        box2: (row_start, col_start, row_end, col_end) - 0-indexed

    Returns:
        IoU score between 0 and 1
    """
    r1_start, c1_start, r1_end, c1_end = box1
    r2_start, c2_start, r2_end, c2_end = box2

    # Calculate intersection
    inter_r_start = max(r1_start, r2_start)
    inter_c_start = max(c1_start, c2_start)
    inter_r_end = min(r1_end, r2_end)
    inter_c_end = min(c1_end, c2_end)

    # Check if there's actual intersection
    if inter_r_start > inter_r_end or inter_c_start > inter_c_end:
        return 0.0

    # Calculate areas
    inter_area = (inter_r_end - inter_r_start + 1) * (inter_c_end - inter_c_start + 1)
    box1_area = (r1_end - r1_start + 1) * (c1_end - c1_start + 1)
    box2_area = (r2_end - r2_start + 1) * (c2_end - c2_start + 1)

    # Calculate union
    union_area = box1_area + box2_area - inter_area

    return inter_area / union_area if union_area > 0 else 0.0


def display_metrics_table(result: dict, file_name: str, is_cached: bool = False):
    """Display evaluation metrics in a rich table format.

    Args:
        result: Evaluation result dictionary
        file_name: Name of the file being evaluated
        is_cached: Whether this result came from cache
    """
    table = Table(title=f"{'[CACHED] ' if is_cached else ''}{file_name}")

    # Add columns
    table.add_column("IoU Threshold", style="cyan", justify="center")
    table.add_column("TP", style="green", justify="right")
    table.add_column("FP", style="red", justify="right")
    table.add_column("FN", style="yellow", justify="right")
    table.add_column("Precision", style="magenta", justify="right")
    table.add_column("Recall", style="blue", justify="right")
    table.add_column("F1", style="bold green", justify="right")

    # Add summary row first
    table.add_row(
        "Summary",
        str(result["num_detected"]),
        "-",
        str(result["num_ground_truth"]),
        "-",
        "-",
        "-",
        style="dim",
    )
    table.add_section()

    # Add rows for each IoU threshold
    for threshold in result["iou_thresholds"]:
        threshold_str = str(threshold)
        metrics = result["details"]["metrics_by_iou"][threshold_str]

        table.add_row(
            f"{threshold:.2f}",
            str(metrics["true_positives"]),
            str(metrics["false_positives"]),
            str(metrics["false_negatives"]),
            f"{metrics['precision']:.4f}",
            f"{metrics['recall']:.4f}",
            f"{metrics['f1']:.4f}",
        )

    console.print(table)


def evaluate_file(
    xlsx_path: str,
    sheet_name: str,
    ground_truth_regions: List[str],
    wrapper: SpreadsheetLLMWrapper,
    model,
    iou_thresholds: List[float] = [1, 0.6],
):
    """Evaluate SpreadsheetLLM on a single file with multiple IoU thresholds.

    Args:
        xlsx_path: Path to Excel file
        sheet_name: Name of the sheet to evaluate
        ground_truth_regions: List of ground truth table regions
        wrapper: SpreadsheetLLM wrapper instance
        model: LangChain ChatModel for LLM-based recognition
        iou_thresholds: List of IoU thresholds to evaluate (default: [0.5, 0.75, 1.0])

    Returns:
        Dictionary with evaluation metrics for each IoU threshold
    """
    if iou_thresholds is None:
        iou_thresholds = [0.5, 0.75, 1.0]
    try:
        # Read and compress spreadsheet
        logger.debug("Step 1: Reading spreadsheet...")
        wb = wrapper.read_spreadsheet(xlsx_path)
        if wb is None:
            logger.warning(f"Failed to read: {xlsx_path}")
            return None

        logger.debug("Step 2: Compressing spreadsheet...")
        result = wrapper.compress_spreadsheet(
            wb, format_aware=True
        )  # Use format-aware for better recognition
        if result is None:
            logger.warning(f"Failed to compress: {xlsx_path}")
            return None

        # Use LLM-based recognition to detect table regions
        logger.debug("Step 3: Running LLM-based recognition...")
        recognition_result = wrapper.recognize_original(
            compress_dict=result.compress_dict,
            sheet_compressor=result.sheet_compressor,
            model=model,
            user_prompt="Identify all table regions in this spreadsheet",
        )
        logger.debug("Step 4: Recognition completed, parsing results...")

        # Capture LLM reasoning and items
        llm_reasoning = (
            recognition_result.reasoning
            if hasattr(recognition_result, "reasoning")
            else None
        )

        # Convert LLM-detected ranges to boxes (0-indexed)
        detected_boxes = []
        recognition_items = []
        for item in recognition_result.items:
            # Store item details for caching
            recognition_items.append(
                {
                    "title": item.title,
                    "range": item.range,
                }
            )

            # Parse each range (e.g., "A1:D10" or "A1,B2:B5")
            ranges = item.range.split(",")
            for range_str in ranges:
                range_str = range_str.strip()
                if ":" in range_str:
                    try:
                        start_col, start_row, end_col, end_row = parse_excel_range(
                            range_str
                        )
                        r_start = start_row - 1  # Convert to 0-indexed
                        c_start = col_to_index(start_col)
                        r_end = end_row - 1
                        c_end = col_to_index(end_col)
                        detected_boxes.append((r_start, c_start, r_end, c_end))
                    except Exception as parse_err:
                        logger.warning(
                            f"Failed to parse range '{range_str}': {parse_err}"
                        )
                        continue

        # Convert ground truth regions to boxes (0-indexed)
        logger.debug("Step 5: Converting ground truth regions to boxes...")
        gt_boxes = []
        for region in ground_truth_regions:
            start_col, start_row, end_col, end_row = parse_excel_range(region)
            r_start = start_row - 1  # Convert to 0-indexed
            c_start = col_to_index(start_col)
            r_end = end_row - 1
            c_end = col_to_index(end_col)
            gt_boxes.append((r_start, c_start, r_end, c_end))

        logger.debug(
            f"  Detected boxes: {len(detected_boxes)}, GT boxes: {len(gt_boxes)}"
        )

        # Pre-compute all IoU values (optimize for multiple thresholds)
        logger.debug("Step 6: Computing IoU matrix...")
        iou_matrix = {}
        for det_idx, det_box in enumerate(detected_boxes):
            for gt_idx, gt_box in enumerate(gt_boxes):
                iou = calculate_iou(det_box, gt_box)
                iou_matrix[(det_idx, gt_idx)] = iou

        # Evaluate for each IoU threshold
        metrics_by_threshold = {}

        for threshold in iou_thresholds:
            # Match detected boxes with ground truth using current threshold
            matched_gt = set()
            matched_det = set()

            for det_idx, det_box in enumerate(detected_boxes):
                best_iou = 0
                best_gt_idx = -1

                for gt_idx, gt_box in enumerate(gt_boxes):
                    if gt_idx in matched_gt:
                        continue

                    iou = iou_matrix[(det_idx, gt_idx)]
                    if iou > best_iou:
                        best_iou = iou
                        best_gt_idx = gt_idx

                # If IoU exceeds threshold, consider it a match
                if best_iou >= threshold:
                    matched_gt.add(best_gt_idx)
                    matched_det.add(det_idx)

            # Calculate metrics for this threshold
            true_positives = len(matched_gt)
            false_positives = len(detected_boxes) - len(matched_det)
            false_negatives = len(gt_boxes) - len(matched_gt)

            precision = (
                true_positives / (true_positives + false_positives)
                if (true_positives + false_positives) > 0
                else 0
            )
            recall = (
                true_positives / (true_positives + false_negatives)
                if (true_positives + false_negatives) > 0
                else 0
            )
            f1 = (
                2 * precision * recall / (precision + recall)
                if (precision + recall) > 0
                else 0
            )

            metrics_by_threshold[threshold] = {
                "true_positives": true_positives,
                "false_positives": false_positives,
                "false_negatives": false_negatives,
                "precision": precision,
                "recall": recall,
                "f1": f1,
                "matched_gt_indices": sorted(matched_gt),
                "matched_det_indices": sorted(matched_det),
            }

        # Convert boxes to readable format for caching
        detected_ranges = [box_to_range(box) for box in detected_boxes]
        gt_ranges = ground_truth_regions  # Already in Excel format

        # Build match details for each threshold
        matches_by_threshold = {}
        for threshold, metrics in metrics_by_threshold.items():
            matched_gt = set(metrics["matched_gt_indices"])
            matched_det = set(metrics["matched_det_indices"])

            matches = []
            for gt_idx in matched_gt:
                # Find which detected box matched this ground truth
                for det_idx in matched_det:
                    iou = iou_matrix[(det_idx, gt_idx)]
                    if iou >= threshold:
                        matches.append(
                            {
                                "ground_truth": gt_ranges[gt_idx],
                                "detected": detected_ranges[det_idx],
                                "iou": round(iou, 4),
                            }
                        )
                        break

            # Unmatched detections (false positives)
            false_positive_ranges = [
                detected_ranges[i]
                for i in range(len(detected_boxes))
                if i not in matched_det
            ]

            # Unmatched ground truths (false negatives)
            false_negative_ranges = [
                gt_ranges[i] for i in range(len(gt_boxes)) if i not in matched_gt
            ]

            matches_by_threshold[threshold] = {
                "matches": matches,
                "false_positives": false_positive_ranges,
                "false_negatives": false_negative_ranges,
            }

        # Sort metrics_by_iou by threshold (convert to sorted dict)
        sorted_metrics_by_iou = dict(
            sorted(
                {
                    str(threshold): {
                        "true_positives": metrics["true_positives"],
                        "false_positives": metrics["false_positives"],
                        "false_negatives": metrics["false_negatives"],
                        "precision": metrics["precision"],
                        "recall": metrics["recall"],
                        "f1": metrics["f1"],
                        "matches": matches_by_threshold[threshold]["matches"],
                        "false_positives_ranges": matches_by_threshold[threshold][
                            "false_positives"
                        ],
                        "false_negatives_ranges": matches_by_threshold[threshold][
                            "false_negatives"
                        ],
                    }
                    for threshold, metrics in metrics_by_threshold.items()
                }.items()
            )
        )

        return {
            "file": xlsx_path,
            "sheet": sheet_name,
            "num_detected": len(detected_boxes),
            "num_ground_truth": len(gt_boxes),
            "iou_thresholds": iou_thresholds,
            "llm_reasoning": llm_reasoning,
            "details": {
                "ground_truth_ranges": gt_ranges,
                "recognition_items": recognition_items,
                # Metrics grouped by IoU threshold (sorted by threshold)
                "metrics_by_iou": sorted_metrics_by_iou,
            },
        }

    except Exception as e:
        logger.error(f"Error processing {xlsx_path}: {e}")
        import traceback

        # Log full traceback to help identify the error location
        logger.error("Full traceback:")
        logger.error(traceback.format_exc())
        return None


def load_cache(cache_file: Path) -> dict[str, dict]:
    """Load cached evaluation results.

    Args:
        cache_file: Path to cache file

    Returns:
        Dictionary mapping file paths to evaluation results
    """
    if cache_file.exists():
        try:
            with open(cache_file, "r") as f:
                cache_data = json.load(f)
                logger.info(
                    f"   ✓ Loaded {len(cache_data)} cached results from {cache_file}"
                )
                return cache_data
        except Exception as e:
            logger.warning(f"   Failed to load cache: {e}")
            return {}
    else:
        logger.info(f"   No existing cache found at {cache_file}")
        return {}


def save_cache(cache_file: Path, cache_data: dict[str, dict]):
    """Save evaluation results to cache.

    Args:
        cache_file: Path to cache file
        cache_data: Dictionary of cached results
    """
    try:
        with open(cache_file, "w", encoding="utf-8") as f:
            json.dump(cache_data, f, indent=2, ensure_ascii=False)
        logger.debug(f"   ✓ Saved {len(cache_data)} results to cache")
    except Exception as e:
        logger.error(f"   Failed to save cache: {e}")


def main():
    """Main evaluation pipeline."""

    logger.info("=" * 70)
    logger.info("TABLESENSE EVALUATION FOR SPREADSHEETLLM (with LLM Recognition)")
    logger.info("=" * 70)

    # Setup cache file
    cache_file = Path("cache") / "tablesense_evaluation_cache.json"
    cache_file.parent.mkdir(exist_ok=True)

    # Load existing cache
    logger.info("\n1. Loading cache...")
    cache = load_cache(cache_file)

    # Download dataset
    logger.info("\n2. Downloading TableSense dataset...")
    annotations_file = hf_hub_download(
        repo_id="kl3269/tablesense", filename="annotations.jsonl", repo_type="dataset"
    )

    data_archive = hf_hub_download(
        repo_id="kl3269/tablesense", filename="data.tar.gz", repo_type="dataset"
    )

    logger.info("   ✓ Downloaded annotations and data")

    # Extract data
    logger.info("\n3. Extracting dataset...")
    temp_dir = tempfile.mkdtemp()
    with tarfile.open(data_archive, "r:gz") as tar:
        tar.extractall(temp_dir, filter="data")

    data_dir = os.path.join(temp_dir, "data")
    logger.info(f"   ✓ Extracted to: {data_dir}")

    # Load annotations
    logger.info("\n4. Loading annotations...")
    with open(annotations_file, "r") as f:
        annotations = [json.loads(line) for line in f]

    # Filter to test set only (or use a sample for faster testing)
    test_annotations = [ann for ann in annotations if ann.get("split") == "testing_set"]
    logger.info(f"   ✓ Loaded {len(test_annotations)} test annotations")

    # Initialize LLM model
    logger.info("\n5. Initializing LLM model...")
    try:
        from langchain_openai import ChatOpenAI

        model_name = "google/gemini-2.5-pro"
        model = ChatOpenAI(model=model_name)
        logger.info(f"   ✓ Initialized model: {model_name}")
    except Exception as e:
        logger.error(f"   ✗ Failed to initialize LLM: {e}")
        logger.error("   Please set OPENAI_API_KEY environment variable")
        return

    # Initialize SpreadsheetLLM
    logger.info("\n6. Initializing SpreadsheetLLM...")
    wrapper = SpreadsheetLLMWrapper()
    logger.info("   ✓ Initialized")

    # Evaluate on test set
    logger.info(f"\n7. Evaluating on test set ({len(test_annotations)} files)...")
    logger.info("   Using LLM-based table recognition (recognize_original)...")
    logger.info("   Cache will be used to skip already processed files")

    results = []
    processed_count = 0
    cached_count = 0

    for idx, annotation in enumerate(test_annotations, 1):
        xlsx_path = os.path.join(data_dir, annotation["clean_file"])

        if not os.path.exists(xlsx_path):
            logger.warning(
                f"   [{idx}/{len(test_annotations)}] File not found: {xlsx_path}"
            )
            continue

        # Create cache key using relative path
        cache_key = annotation["clean_file"]

        # Check if result is already cached
        if cache_key in cache:
            cached_result = cache[cache_key]
            current_thresholds = [0.5, 0.75, 1.0]

            # Check if cached result has all required IoU thresholds
            cached_thresholds = set(
                cached_result.get("details", {}).get("metrics_by_iou", {}).keys()
            )
            required_thresholds = set(str(t) for t in current_thresholds)
            missing_thresholds = required_thresholds - cached_thresholds

            if missing_thresholds:
                logger.info(
                    f"   [{idx}/{len(test_annotations)}] Cache found but missing IoU thresholds {missing_thresholds}, recomputing metrics: {annotation['clean_file']}"
                )

                # Recompute metrics for missing thresholds using cached detection results
                if (
                    "details" in cached_result
                    and "ground_truth_ranges" in cached_result["details"]
                ):
                    try:
                        # Get cached detection and ground truth data
                        recognition_items = cached_result["details"].get(
                            "recognition_items", []
                        )
                        gt_ranges = cached_result["details"]["ground_truth_ranges"]

                        # Reconstruct detected and ground truth boxes (reuse code from evaluate_file)
                        detected_boxes = []
                        for item in recognition_items:
                            ranges = item["range"].split(",")
                            for range_str in ranges:
                                range_str = range_str.strip()
                                if ":" in range_str:
                                    start_col, start_row, end_col, end_row = (
                                        parse_excel_range(range_str)
                                    )
                                    r_start = start_row - 1
                                    c_start = col_to_index(start_col)
                                    r_end = end_row - 1
                                    c_end = col_to_index(end_col)
                                    detected_boxes.append(
                                        (r_start, c_start, r_end, c_end)
                                    )

                        gt_boxes = []
                        for region in gt_ranges:
                            start_col, start_row, end_col, end_row = parse_excel_range(
                                region
                            )
                            r_start = start_row - 1
                            c_start = col_to_index(start_col)
                            r_end = end_row - 1
                            c_end = col_to_index(end_col)
                            gt_boxes.append((r_start, c_start, r_end, c_end))

                        # Recompute IoU matrix
                        iou_matrix = {}
                        for det_idx, det_box in enumerate(detected_boxes):
                            for gt_idx, gt_box in enumerate(gt_boxes):
                                iou = calculate_iou(det_box, gt_box)
                                iou_matrix[(det_idx, gt_idx)] = iou

                        # Compute metrics for missing thresholds
                        for threshold_str in missing_thresholds:
                            threshold = float(threshold_str)

                            # Match detected boxes with ground truth
                            matched_gt = set()
                            matched_det = set()

                            for det_idx, det_box in enumerate(detected_boxes):
                                best_iou = 0
                                best_gt_idx = -1

                                for gt_idx, gt_box in enumerate(gt_boxes):
                                    if gt_idx in matched_gt:
                                        continue

                                    iou = iou_matrix[(det_idx, gt_idx)]
                                    if iou > best_iou:
                                        best_iou = iou
                                        best_gt_idx = gt_idx

                                if best_iou >= threshold:
                                    matched_gt.add(best_gt_idx)
                                    matched_det.add(det_idx)

                            # Calculate metrics
                            true_positives = len(matched_gt)
                            false_positives = len(detected_boxes) - len(matched_det)
                            false_negatives = len(gt_boxes) - len(matched_gt)

                            precision = (
                                true_positives / (true_positives + false_positives)
                                if (true_positives + false_positives) > 0
                                else 0
                            )
                            recall = (
                                true_positives / (true_positives + false_negatives)
                                if (true_positives + false_negatives) > 0
                                else 0
                            )
                            f1 = (
                                2 * precision * recall / (precision + recall)
                                if (precision + recall) > 0
                                else 0
                            )

                            # Build match details
                            detected_ranges = [
                                box_to_range(box) for box in detected_boxes
                            ]

                            matches = []
                            for gt_idx in matched_gt:
                                for det_idx in matched_det:
                                    iou = iou_matrix[(det_idx, gt_idx)]
                                    if iou >= threshold:
                                        matches.append(
                                            {
                                                "ground_truth": gt_ranges[gt_idx],
                                                "detected": detected_ranges[det_idx],
                                                "iou": round(iou, 4),
                                            }
                                        )
                                        break

                            false_positive_ranges = [
                                detected_ranges[i]
                                for i in range(len(detected_boxes))
                                if i not in matched_det
                            ]
                            false_negative_ranges = [
                                gt_ranges[i]
                                for i in range(len(gt_boxes))
                                if i not in matched_gt
                            ]

                            # Add to cached result
                            cached_result["details"]["metrics_by_iou"][
                                threshold_str
                            ] = {
                                "true_positives": true_positives,
                                "false_positives": false_positives,
                                "false_negatives": false_negatives,
                                "precision": precision,
                                "recall": recall,
                                "f1": f1,
                                "matches": matches,
                                "false_positives_ranges": false_positive_ranges,
                                "false_negatives_ranges": false_negative_ranges,
                            }

                        # Update iou_thresholds in cached result
                        cached_result["iou_thresholds"] = current_thresholds

                        # Sort metrics_by_iou by key (lexicographic order)
                        cached_result["details"]["metrics_by_iou"] = dict(
                            sorted(cached_result["details"]["metrics_by_iou"].items())
                        )

                        # Update cache
                        cache[cache_key] = cached_result
                        save_cache(cache_file, cache)

                        logger.info(
                            f"   ✓ Recomputed metrics for thresholds: {missing_thresholds}"
                        )

                    except Exception as e:
                        logger.error(f"   Failed to recompute metrics: {e}")
                        # Fall through to use cached result as-is
            else:
                logger.info(
                    f"   [{idx}/{len(test_annotations)}] Using cached result: {annotation['clean_file']}"
                )

            result = cached_result
            results.append(result)
            cached_count += 1

            # Display metrics table
            display_metrics_table(result, annotation["clean_file"], is_cached=True)
            continue

        # Process new file with status display
        logger.info(
            f"   [{idx}/{len(test_annotations)}] Processing: {annotation['clean_file']}"
        )

        with console.status(f"[bold green]Processing: {xlsx_path}"):
            result = evaluate_file(
                xlsx_path=xlsx_path,
                sheet_name=annotation["sheet_name"],
                ground_truth_regions=annotation["table_regions"],
                wrapper=wrapper,
                model=model,
                iou_thresholds=[0.5, 0.75, 1.0],
            )

        if result:
            results.append(result)
            processed_count += 1

            # Save to cache immediately after successful processing
            cache[cache_key] = result
            save_cache(cache_file, cache)

            # Display metrics table
            display_metrics_table(result, annotation["clean_file"], is_cached=False)

    logger.info(
        f"\n   ✓ Processed {processed_count} new files, reused {cached_count} cached results"
    )

    # Calculate aggregate metrics
    if results:
        logger.info("\n" + "=" * 70)
        logger.info("EVALUATION RESULTS")
        logger.info("=" * 70)

        # Get all IoU thresholds from the first result
        iou_thresholds = results[0].get("iou_thresholds", [0.5])

        logger.info(f"\nFiles evaluated: {len(results)}")
        logger.info(
            f"Total ground truth tables: {sum(r['num_ground_truth'] for r in results)}"
        )
        logger.info(f"Total detected tables: {sum(r['num_detected'] for r in results)}")

        # Calculate metrics for each IoU threshold
        logger.info("\n" + "-" * 70)
        logger.info("METRICS BY IoU THRESHOLD")
        logger.info("-" * 70)

        overall_metrics_by_iou = {}

        for threshold in iou_thresholds:
            threshold_str = str(threshold)

            # Aggregate TP/FP/FN across all files for this threshold
            total_tp = sum(
                r["details"]["metrics_by_iou"][threshold_str]["true_positives"]
                for r in results
                if "details" in r
                and "metrics_by_iou" in r["details"]
                and threshold_str in r["details"]["metrics_by_iou"]
            )
            total_fp = sum(
                r["details"]["metrics_by_iou"][threshold_str]["false_positives"]
                for r in results
                if "details" in r
                and "metrics_by_iou" in r["details"]
                and threshold_str in r["details"]["metrics_by_iou"]
            )
            total_fn = sum(
                r["details"]["metrics_by_iou"][threshold_str]["false_negatives"]
                for r in results
                if "details" in r
                and "metrics_by_iou" in r["details"]
                and threshold_str in r["details"]["metrics_by_iou"]
            )

            overall_precision = (
                total_tp / (total_tp + total_fp) if (total_tp + total_fp) > 0 else 0
            )
            overall_recall = (
                total_tp / (total_tp + total_fn) if (total_tp + total_fn) > 0 else 0
            )
            overall_f1 = (
                2
                * overall_precision
                * overall_recall
                / (overall_precision + overall_recall)
                if (overall_precision + overall_recall) > 0
                else 0
            )

            # Average F1 per file
            f1_scores = [
                r["details"]["metrics_by_iou"][threshold_str]["f1"]
                for r in results
                if "details" in r
                and "metrics_by_iou" in r["details"]
                and threshold_str in r["details"]["metrics_by_iou"]
            ]
            avg_f1 = sum(f1_scores) / len(f1_scores) if f1_scores else 0

            overall_metrics_by_iou[threshold_str] = {
                "true_positives": total_tp,
                "false_positives": total_fp,
                "false_negatives": total_fn,
                "precision": overall_precision,
                "recall": overall_recall,
                "f1": overall_f1,
                "avg_f1_per_file": avg_f1,
            }

        # Display aggregate metrics in a table
        aggregate_table = Table(title=f"Overall Metrics ({len(results)} files)")
        aggregate_table.add_column("IoU Threshold", style="cyan", justify="center")
        aggregate_table.add_column("TP", style="green", justify="right")
        aggregate_table.add_column("FP", style="red", justify="right")
        aggregate_table.add_column("FN", style="yellow", justify="right")
        aggregate_table.add_column("Precision", style="magenta", justify="right")
        aggregate_table.add_column("Recall", style="blue", justify="right")
        aggregate_table.add_column("F1 (Overall)", style="bold green", justify="right")
        aggregate_table.add_column("F1 (Avg)", style="bold cyan", justify="right")

        for threshold in iou_thresholds:
            threshold_str = str(threshold)
            metrics = overall_metrics_by_iou[threshold_str]
            aggregate_table.add_row(
                f"{threshold:.2f}",
                str(metrics["true_positives"]),
                str(metrics["false_positives"]),
                str(metrics["false_negatives"]),
                f"{metrics['precision']:.4f}",
                f"{metrics['recall']:.4f}",
                f"{metrics['f1']:.4f}",
                f"{metrics['avg_f1_per_file']:.4f}",
            )

        console.print("\n")
        console.print(aggregate_table)

        # Save detailed results
        output_file = Path("tablesense_evaluation_results.json")
        with open(output_file, "w") as f:
            json.dump(
                {
                    "overall_metrics_by_iou": overall_metrics_by_iou,
                    "per_file_results": results,
                },
                f,
                indent=2,
            )

        logger.info(f"\n✓ Detailed results saved to: {output_file}")
    else:
        logger.error("\n❌ No successful evaluations")

    logger.info("\n" + "=" * 70)


if __name__ == "__main__":
    main()
