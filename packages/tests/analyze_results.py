"""
Analyze TableSense Evaluation Results

This script provides detailed analysis of SpreadsheetLLM's performance
on the TableSense dataset, including:
- Overall metrics by IoU threshold
- Error case analysis
- Table size distribution
- IoU distribution analysis
- Success/failure patterns
- AI-generated insights using Gemini
"""

import json
import statistics
from pathlib import Path
from typing import Dict

from rich.console import Console
from rich.panel import Panel
from rich.progress import track
from rich.table import Table

console = Console()

# Configuration
ENABLE_AI_ANALYSIS = True  # Set to True to generate AI analysis report
AI_MODEL_NAME = "google/gemini-2.5-pro"  # Gemini model to use


def load_results(cache_file: Path) -> Dict:
    """Load evaluation results from cache file."""
    with open(cache_file, "r") as f:
        return json.load(f)


def parse_range(range_str: str) -> tuple[int, int]:
    """Parse Excel range to get number of rows and columns."""
    if ":" not in range_str:
        return 1, 1

    start, end = range_str.split(":")

    # Parse start
    start_col = "".join(c for c in start if c.isalpha())
    start_row = int("".join(c for c in start if c.isdigit()))

    # Parse end
    end_col = "".join(c for c in end if c.isalpha())
    end_row = int("".join(c for c in end if c.isdigit()))

    # Calculate dimensions
    def col_to_num(col):
        num = 0
        for c in col:
            num = num * 26 + (ord(c) - ord("A") + 1)
        return num

    rows = end_row - start_row + 1
    cols = col_to_num(end_col) - col_to_num(start_col) + 1

    return rows, cols


def analyze_overall_metrics(results: Dict) -> Dict:
    """Calculate overall metrics across all files for each IoU threshold."""
    console.print("\n[bold cyan]Computing Overall Metrics...[/bold cyan]")

    # Get IoU thresholds from first result
    first_result = next(iter(results.values()))
    iou_thresholds = first_result.get("iou_thresholds", [0.5, 0.75, 1.0])

    overall_metrics = {}

    for threshold in track(iou_thresholds, description="Processing thresholds"):
        threshold_str = str(threshold)

        total_tp = 0
        total_fp = 0
        total_fn = 0
        total_gt = 0
        total_detected = 0

        for file_result in results.values():
            if (
                "details" not in file_result
                or "metrics_by_iou" not in file_result["details"]
            ):
                continue

            if threshold_str in file_result["details"]["metrics_by_iou"]:
                metrics = file_result["details"]["metrics_by_iou"][threshold_str]
                total_tp += metrics["true_positives"]
                total_fp += metrics["false_positives"]
                total_fn += metrics["false_negatives"]

            total_gt += file_result.get("num_ground_truth", 0)
            total_detected += file_result.get("num_detected", 0)

        precision = total_tp / (total_tp + total_fp) if (total_tp + total_fp) > 0 else 0
        recall = total_tp / (total_tp + total_fn) if (total_tp + total_fn) > 0 else 0
        f1 = (
            2 * precision * recall / (precision + recall)
            if (precision + recall) > 0
            else 0
        )

        overall_metrics[threshold_str] = {
            "threshold": threshold,
            "true_positives": total_tp,
            "false_positives": total_fp,
            "false_negatives": total_fn,
            "precision": precision,
            "recall": recall,
            "f1": f1,
            "total_ground_truth": total_gt,
            "total_detected": total_detected,
        }

    return overall_metrics


def analyze_iou_distribution(results: Dict) -> Dict:
    """Analyze distribution of IoU scores for matched tables."""
    console.print("\n[bold cyan]Analyzing IoU Distribution...[/bold cyan]")

    iou_scores = []
    perfect_matches = 0
    high_iou = 0  # IoU >= 0.9
    medium_iou = 0  # 0.7 <= IoU < 0.9
    low_iou = 0  # IoU < 0.7

    for file_result in track(results.values(), description="Processing files"):
        if "details" not in file_result:
            continue

        # Use highest IoU threshold to get all matches
        metrics_by_iou = file_result["details"].get("metrics_by_iou", {})
        if not metrics_by_iou:
            continue

        # Get matches from the first threshold (they should all have the same matches)
        first_threshold = list(metrics_by_iou.keys())[0]
        matches = metrics_by_iou[first_threshold].get("matches", [])

        for match in matches:
            iou = match["iou"]
            iou_scores.append(iou)

            if iou == 1.0:
                perfect_matches += 1
            elif iou >= 0.9:
                high_iou += 1
            elif iou >= 0.7:
                medium_iou += 1
            else:
                low_iou += 1

    return {
        "iou_scores": iou_scores,
        "perfect_matches": perfect_matches,
        "high_iou_matches": high_iou,
        "medium_iou_matches": medium_iou,
        "low_iou_matches": low_iou,
        "mean_iou": statistics.mean(iou_scores) if iou_scores else 0,
        "median_iou": statistics.median(iou_scores) if iou_scores else 0,
        "min_iou": min(iou_scores) if iou_scores else 0,
        "max_iou": max(iou_scores) if iou_scores else 0,
    }


def analyze_error_cases(results: Dict) -> Dict:
    """Analyze false positives and false negatives."""
    console.print("\n[bold cyan]Analyzing Error Cases...[/bold cyan]")

    # Use IoU threshold 0.75 for error analysis
    threshold_str = "0.75"

    false_positives = []
    false_negatives = []
    perfect_detections = []
    partial_matches = []

    for filename, file_result in track(
        results.items(), description="Processing error cases"
    ):
        if "details" not in file_result:
            continue

        metrics_by_iou = file_result["details"].get("metrics_by_iou", {})
        if threshold_str not in metrics_by_iou:
            continue

        metrics = metrics_by_iou[threshold_str]

        # Categorize files
        if metrics["false_positives"] > 0:
            false_positives.append(
                {
                    "file": filename,
                    "fp_count": metrics["false_positives"],
                    "fp_ranges": metrics.get("false_positives_ranges", []),
                    "detected": file_result["num_detected"],
                    "ground_truth": file_result["num_ground_truth"],
                }
            )

        if metrics["false_negatives"] > 0:
            false_negatives.append(
                {
                    "file": filename,
                    "fn_count": metrics["false_negatives"],
                    "fn_ranges": metrics.get("false_negatives_ranges", []),
                    "detected": file_result["num_detected"],
                    "ground_truth": file_result["num_ground_truth"],
                }
            )

        if metrics["precision"] == 1.0 and metrics["recall"] == 1.0:
            perfect_detections.append(filename)
        elif metrics["true_positives"] > 0:
            partial_matches.append(
                {
                    "file": filename,
                    "precision": metrics["precision"],
                    "recall": metrics["recall"],
                    "f1": metrics["f1"],
                }
            )

    return {
        "false_positives": false_positives,
        "false_negatives": false_negatives,
        "perfect_detections": perfect_detections,
        "partial_matches": partial_matches,
        "fp_count": len(false_positives),
        "fn_count": len(false_negatives),
        "perfect_count": len(perfect_detections),
        "partial_count": len(partial_matches),
    }


def analyze_table_sizes(results: Dict) -> Dict:
    """Analyze distribution of table sizes (ground truth)."""
    console.print("\n[bold cyan]Analyzing Table Sizes...[/bold cyan]")

    table_sizes = []
    size_categories = {
        "tiny": 0,  # < 100 cells
        "small": 0,  # 100-500 cells
        "medium": 0,  # 500-2000 cells
        "large": 0,  # 2000-10000 cells
        "huge": 0,  # > 10000 cells
    }

    size_to_success = {
        "tiny": {"success": 0, "total": 0},
        "small": {"success": 0, "total": 0},
        "medium": {"success": 0, "total": 0},
        "large": {"success": 0, "total": 0},
        "huge": {"success": 0, "total": 0},
    }

    for file_result in track(results.values(), description="Processing table sizes"):
        if "details" not in file_result:
            continue

        gt_ranges = file_result["details"].get("ground_truth_ranges", [])

        for range_str in gt_ranges:
            rows, cols = parse_range(range_str)
            cells = rows * cols
            table_sizes.append(cells)

            # Categorize size
            if cells < 100:
                category = "tiny"
            elif cells < 500:
                category = "small"
            elif cells < 2000:
                category = "medium"
            elif cells < 10000:
                category = "large"
            else:
                category = "huge"

            size_categories[category] += 1
            size_to_success[category]["total"] += 1

            # Check if successfully detected (using 0.75 IoU threshold)
            metrics_by_iou = file_result["details"].get("metrics_by_iou", {})
            if "0.75" in metrics_by_iou:
                metrics = metrics_by_iou["0.75"]
                if metrics["true_positives"] > 0:
                    size_to_success[category]["success"] += 1

    # Calculate success rates by size
    success_rates = {}
    for category, data in size_to_success.items():
        if data["total"] > 0:
            success_rates[category] = data["success"] / data["total"]
        else:
            success_rates[category] = 0

    return {
        "table_sizes": table_sizes,
        "size_categories": size_categories,
        "success_rates": success_rates,
        "mean_size": statistics.mean(table_sizes) if table_sizes else 0,
        "median_size": statistics.median(table_sizes) if table_sizes else 0,
        "min_size": min(table_sizes) if table_sizes else 0,
        "max_size": max(table_sizes) if table_sizes else 0,
    }


def display_results(overall_metrics, iou_dist, error_cases, table_sizes):
    """Display all analysis results using Rich tables and panels."""

    # Overall Metrics Table
    console.print("\n")
    console.print(
        Panel.fit(
            "[bold cyan]TABLESENSE EVALUATION ANALYSIS[/bold cyan]", border_style="cyan"
        )
    )

    # 1. Overall Performance by IoU Threshold
    console.print("\n[bold]1. Overall Performance by IoU Threshold[/bold]")
    metrics_table = Table(title="Performance Metrics", show_header=True)
    metrics_table.add_column("IoU Threshold", style="cyan", justify="center")
    metrics_table.add_column("Precision", style="green", justify="right")
    metrics_table.add_column("Recall", style="blue", justify="right")
    metrics_table.add_column("F1 Score", style="magenta", justify="right")
    metrics_table.add_column("TP", style="green", justify="right")
    metrics_table.add_column("FP", style="red", justify="right")
    metrics_table.add_column("FN", style="yellow", justify="right")

    for metrics in sorted(overall_metrics.values(), key=lambda x: x["threshold"]):
        metrics_table.add_row(
            f"{metrics['threshold']:.2f}",
            f"{metrics['precision']:.4f}",
            f"{metrics['recall']:.4f}",
            f"{metrics['f1']:.4f}",
            str(metrics["true_positives"]),
            str(metrics["false_positives"]),
            str(metrics["false_negatives"]),
        )

    console.print(metrics_table)

    # 2. IoU Distribution Analysis
    console.print("\n[bold]2. IoU Distribution Analysis[/bold]")
    iou_table = Table(title="IoU Score Distribution", show_header=True)
    iou_table.add_column("Metric", style="cyan")
    iou_table.add_column("Value", style="white", justify="right")

    iou_table.add_row(
        "Perfect Matches (IoU = 1.0)",
        (
            f"{iou_dist['perfect_matches']} ({iou_dist['perfect_matches']/len(iou_dist['iou_scores'])*100:.1f}%)"
            if iou_dist["iou_scores"]
            else "0"
        ),
    )
    iou_table.add_row("High IoU (0.9 ≤ IoU < 1.0)", f"{iou_dist['high_iou_matches']}")
    iou_table.add_row(
        "Medium IoU (0.7 ≤ IoU < 0.9)", f"{iou_dist['medium_iou_matches']}"
    )
    iou_table.add_row("Low IoU (IoU < 0.7)", f"{iou_dist['low_iou_matches']}")
    iou_table.add_row("", "")
    iou_table.add_row("Mean IoU", f"{iou_dist['mean_iou']:.4f}")
    iou_table.add_row("Median IoU", f"{iou_dist['median_iou']:.4f}")
    iou_table.add_row("Min IoU", f"{iou_dist['min_iou']:.4f}")
    iou_table.add_row("Max IoU", f"{iou_dist['max_iou']:.4f}")

    console.print(iou_table)

    # 3. Error Analysis
    console.print("\n[bold]3. Error Analysis (at IoU = 0.75)[/bold]")
    error_table = Table(title="Error Cases", show_header=True)
    error_table.add_column("Category", style="cyan")
    error_table.add_column("Count", style="white", justify="right")
    error_table.add_column("Percentage", style="white", justify="right")

    total_files = (
        error_cases["perfect_count"]
        + error_cases["partial_count"]
        + error_cases["fp_count"]
        + error_cases["fn_count"]
    )

    error_table.add_row(
        "Perfect Detections",
        str(error_cases["perfect_count"]),
        f"{error_cases['perfect_count']/total_files*100:.1f}%",
    )
    error_table.add_row(
        "Partial Matches",
        str(error_cases["partial_count"]),
        f"{error_cases['partial_count']/total_files*100:.1f}%",
    )
    error_table.add_row(
        "Files with False Positives",
        str(error_cases["fp_count"]),
        f"{error_cases['fp_count']/total_files*100:.1f}%",
    )
    error_table.add_row(
        "Files with False Negatives",
        str(error_cases["fn_count"]),
        f"{error_cases['fn_count']/total_files*100:.1f}%",
    )

    console.print(error_table)

    # Show top error cases
    if error_cases["false_positives"]:
        console.print("\n[yellow]Top 5 Files with False Positives:[/yellow]")
        for i, fp_case in enumerate(
            sorted(
                error_cases["false_positives"],
                key=lambda x: x["fp_count"],
                reverse=True,
            )[:5],
            1,
        ):
            console.print(f"  {i}. {fp_case['file']}")
            console.print(
                f"     FP: {fp_case['fp_count']}, Detected: {fp_case['detected']}, GT: {fp_case['ground_truth']}"
            )

    if error_cases["false_negatives"]:
        console.print("\n[yellow]Top 5 Files with False Negatives:[/yellow]")
        for i, fn_case in enumerate(
            sorted(
                error_cases["false_negatives"],
                key=lambda x: x["fn_count"],
                reverse=True,
            )[:5],
            1,
        ):
            console.print(f"  {i}. {fn_case['file']}")
            console.print(
                f"     FN: {fn_case['fn_count']}, Detected: {fn_case['detected']}, GT: {fn_case['ground_truth']}"
            )

    # 4. Table Size Analysis
    console.print("\n[bold]4. Table Size Analysis[/bold]")
    size_table = Table(title="Table Size Distribution", show_header=True)
    size_table.add_column("Size Category", style="cyan")
    size_table.add_column("Cell Range", style="white")
    size_table.add_column("Count", style="white", justify="right")
    size_table.add_column("Success Rate", style="green", justify="right")

    size_ranges = {
        "tiny": "< 100",
        "small": "100-500",
        "medium": "500-2,000",
        "large": "2,000-10,000",
        "huge": "> 10,000",
    }

    for category in ["tiny", "small", "medium", "large", "huge"]:
        count = table_sizes["size_categories"][category]
        success_rate = table_sizes["success_rates"][category]
        size_table.add_row(
            category.capitalize(),
            size_ranges[category],
            str(count),
            f"{success_rate*100:.1f}%" if count > 0 else "N/A",
        )

    console.print(size_table)

    # Table size statistics
    console.print(f"\n  Mean table size: {table_sizes['mean_size']:.0f} cells")
    console.print(f"  Median table size: {table_sizes['median_size']:.0f} cells")
    console.print(f"  Min table size: {table_sizes['min_size']} cells")
    console.print(f"  Max table size: {table_sizes['max_size']} cells")

    # 5. Key Insights
    console.print("\n[bold]5. Key Insights & Recommendations[/bold]")
    insights = []

    # Performance insights
    f1_075 = overall_metrics["0.75"]["f1"]
    if f1_075 >= 0.9:
        insights.append(
            "✓ [green]Excellent performance[/green] with F1 score > 0.9 at IoU=0.75"
        )
    elif f1_075 >= 0.8:
        insights.append(
            "✓ [yellow]Good performance[/yellow] with F1 score > 0.8 at IoU=0.75"
        )
    else:
        insights.append(
            "⚠ [red]Room for improvement[/red] - F1 score < 0.8 at IoU=0.75"
        )

    # IoU insights
    perfect_rate = (
        iou_dist["perfect_matches"] / len(iou_dist["iou_scores"]) * 100
        if iou_dist["iou_scores"]
        else 0
    )
    if perfect_rate >= 80:
        insights.append(
            f"✓ [green]High accuracy[/green] - {perfect_rate:.1f}% perfect matches (IoU = 1.0)"
        )
    else:
        insights.append(
            f"⚠ [yellow]Accuracy opportunity[/yellow] - Only {perfect_rate:.1f}% perfect matches"
        )

    # Size insights
    best_size = max(table_sizes["success_rates"].items(), key=lambda x: x[1])
    worst_size = min(table_sizes["success_rates"].items(), key=lambda x: x[1])
    insights.append(
        f"• Best performance on {best_size[0]} tables ({best_size[1]*100:.1f}% success)"
    )
    insights.append(
        f"• Challenges with {worst_size[0]} tables ({worst_size[1]*100:.1f}% success)"
    )

    # Error insights
    if error_cases["fp_count"] > error_cases["fn_count"]:
        insights.append(
            "⚠ More false positives than false negatives - model may be over-detecting"
        )
    elif error_cases["fn_count"] > error_cases["fp_count"]:
        insights.append(
            "⚠ More false negatives than false positives - model may be under-detecting"
        )

    for insight in insights:
        console.print(f"  {insight}")

    console.print("\n" + "=" * 70)


def generate_ai_analysis(overall_metrics, iou_dist, error_cases, table_sizes):
    """Generate detailed analysis report using Gemini AI."""
    console.print("\n[bold cyan]Generating AI Analysis Report...[/bold cyan]")

    try:
        from langchain_openai import ChatOpenAI

        # Initialize Gemini model
        console.print(f"  Initializing {AI_MODEL_NAME}...")
        model = ChatOpenAI(model=AI_MODEL_NAME)

        # Prepare analysis data summary
        analysis_summary = {
            "overall_metrics": overall_metrics,
            "iou_distribution": {
                "perfect_matches": iou_dist["perfect_matches"],
                "high_iou_matches": iou_dist["high_iou_matches"],
                "medium_iou_matches": iou_dist["medium_iou_matches"],
                "low_iou_matches": iou_dist["low_iou_matches"],
                "mean_iou": iou_dist["mean_iou"],
                "median_iou": iou_dist["median_iou"],
                "total_matches": len(iou_dist["iou_scores"]),
            },
            "error_analysis": {
                "perfect_count": error_cases["perfect_count"],
                "partial_count": error_cases["partial_count"],
                "fp_count": error_cases["fp_count"],
                "fn_count": error_cases["fn_count"],
                "top_false_positives": (
                    error_cases["false_positives"][:5]
                    if error_cases["false_positives"]
                    else []
                ),
                "top_false_negatives": (
                    error_cases["false_negatives"][:5]
                    if error_cases["false_negatives"]
                    else []
                ),
            },
            "table_size_analysis": {
                "size_categories": table_sizes["size_categories"],
                "success_rates": table_sizes["success_rates"],
                "mean_size": table_sizes["mean_size"],
                "median_size": table_sizes["median_size"],
                "min_size": table_sizes["min_size"],
                "max_size": table_sizes["max_size"],
            },
        }

        # Create prompt for AI analysis
        prompt = f"""You are an expert in evaluating table detection systems. Analyze the following evaluation results from SpreadsheetLLM on the TableSense benchmark dataset and provide a comprehensive research-quality analysis report.

## Evaluation Data:

{json.dumps(analysis_summary, indent=2)}

## Instructions:

Please provide a detailed analysis report in Markdown format with the following sections:

1. **Executive Summary**
   - Brief overview of overall performance
   - Key findings (3-5 bullet points)
   - Overall assessment (Excellent/Good/Fair/Needs Improvement)

2. **Performance Analysis**
   - Analyze performance across different IoU thresholds
   - Compare precision, recall, and F1 scores
   - Discuss what these metrics reveal about the model's behavior

3. **IoU Distribution Insights**
   - Interpret the IoU distribution pattern
   - Discuss the rate of perfect matches vs partial matches
   - What does this tell us about detection accuracy?

4. **Error Pattern Analysis**
   - Analyze false positive patterns (over-detection)
   - Analyze false negative patterns (under-detection)
   - Which error type is more prevalent and why might this be?
   - Examine specific error cases if notable patterns exist

5. **Table Size Performance**
   - How does performance vary with table size?
   - Which table sizes are handled best/worst?
   - Potential reasons for performance differences across sizes

6. **Strengths and Weaknesses**
   - Clear strengths demonstrated by the results
   - Areas of weakness or limitation
   - Comparison to typical baselines (if relevant)

7. **Recommendations for Improvement**
   - Specific, actionable recommendations based on the error analysis
   - Prioritized list of potential improvements
   - Suggested experiments or modifications

8. **Research Implications**
   - What do these results mean for practical table detection applications?
   - Suitability for different use cases
   - Comparison to state-of-the-art (general discussion)

9. **Conclusion**
   - Overall assessment
   - Main takeaways
   - Future directions

Please write in a professional, research-oriented tone suitable for a technical report or paper. Use specific numbers from the data to support your analysis. Be critical but constructive in your assessment."""

        console.print("  Sending analysis request to AI model...")
        with console.status("[bold green]Waiting for AI response..."):
            response = model.invoke(prompt)

        ai_report = response.content
        console.print("  [green]✓[/green] AI analysis completed")

        return ai_report

    except ImportError:
        console.print(
            "  [yellow]⚠[/yellow] langchain_openai not available, skipping AI analysis"
        )
        return None
    except Exception as e:
        console.print(f"  [red]✗[/red] Failed to generate AI analysis: {e}")
        return None


def main():
    """Main analysis pipeline."""
    cache_file = Path("cache/tablesense_evaluation_cache.json")

    if not cache_file.exists():
        console.print("[red]Error: Cache file not found![/red]")
        console.print(f"Expected location: {cache_file}")
        return

    # Load results
    console.print(f"[cyan]Loading results from:[/cyan] {cache_file}")
    results = load_results(cache_file)
    console.print(f"[green]✓[/green] Loaded {len(results)} file results")

    # Run analyses
    overall_metrics = analyze_overall_metrics(results)
    iou_dist = analyze_iou_distribution(results)
    error_cases = analyze_error_cases(results)
    table_sizes = analyze_table_sizes(results)

    # Display results
    display_results(overall_metrics, iou_dist, error_cases, table_sizes)

    # Save detailed analysis (JSON)
    json_output_file = Path("analysis_report.json")
    with open(json_output_file, "w") as f:
        json.dump(
            {
                "overall_metrics": overall_metrics,
                "iou_distribution": {
                    k: v
                    for k, v in iou_dist.items()
                    if k != "iou_scores"  # Don't save all individual scores
                },
                "error_summary": {
                    "perfect_count": error_cases["perfect_count"],
                    "partial_count": error_cases["partial_count"],
                    "fp_count": error_cases["fp_count"],
                    "fn_count": error_cases["fn_count"],
                },
                "table_size_summary": {
                    k: v
                    for k, v in table_sizes.items()
                    if k != "table_sizes"  # Don't save all individual sizes
                },
            },
            f,
            indent=2,
        )

    console.print(f"\n[green]✓[/green] Detailed analysis saved to: {json_output_file}")

    # Generate AI analysis report
    if ENABLE_AI_ANALYSIS:
        ai_report = generate_ai_analysis(
            overall_metrics, iou_dist, error_cases, table_sizes
        )

        if ai_report and isinstance(ai_report, str):
            # Save AI-generated report
            md_output_file = Path("analysis_report_ai.md")
            with open(md_output_file, "w", encoding="utf-8") as f:
                f.write("# SpreadsheetLLM TableSense Evaluation Analysis\n\n")
                f.write(f"*AI-Generated Report using {AI_MODEL_NAME}*\n\n")
                f.write("---\n\n")
                f.write(ai_report)

            console.print(
                f"[green]✓[/green] AI-generated analysis saved to: {md_output_file}"
            )

            # Display a preview
            console.print("\n[bold cyan]AI Analysis Preview:[/bold cyan]")
            preview_lines = ai_report.split("\n")[:15]  # First 15 lines
            for line in preview_lines:
                console.print(f"  {line}")
            if len(ai_report.split("\n")) > 15:
                console.print(
                    f"  [dim]... (see {md_output_file} for full report)[/dim]"
                )
    else:
        console.print(
            "\n[dim]AI analysis disabled (set ENABLE_AI_ANALYSIS=True to enable)[/dim]"
        )


if __name__ == "__main__":
    main()
