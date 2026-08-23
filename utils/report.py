"""
Report generation module.
Creates summary reports and exports data to CSV and JSON formats.
"""

import csv
import json
import os
from datetime import datetime
from typing import Any, Dict


class ReportGenerator:
    """Generates and exports reports from update results."""

    @staticmethod
    def generate_report(
        diff_result: Dict[str, Any], session_data: Dict[str, Any]
    ) -> Dict[str, Any]:
        """Generate a comprehensive report from the update process.

        Args:
            diff_result: Results from the Differ.compare() method.
            session_data: Session configuration and mappings.

        Returns:
            Report dictionary with all relevant information.
        """
        column_mapping = session_data["column_mapping"]

        target_rows_old = diff_result.get("target_rows", 0)
        new_skus = diff_result.get("new_skus", 0)

        return {
            "timestamp": datetime.now().isoformat(),
            "source_file": session_data.get("source_path", ""),
            "target_file": session_data["target_path"],
            "output_file": session_data.get("output_path", ""),
            "source_sku_column": session_data["source_sku_col"],
            "target_sku_column": session_data["target_sku_col"],
            "column_mapping_updates": column_mapping.get("updates", {}),
            "column_mapping_new": column_mapping.get("new_columns", {}),
            "prefix_used": column_mapping.get("prefix", ""),
            "source_rows": diff_result.get("source_rows", 0),
            "target_rows_old": target_rows_old,
            "target_rows_new": target_rows_old + new_skus,
            "columns_matched": diff_result.get("columns_matched", 0),
            "columns_total": (
                len(column_mapping.get("updates", {}))
                + len(column_mapping.get("new_columns", {}))
            ),
            "new_columns_added": diff_result.get("new_columns_count", 0),
            "cells_updated": len(diff_result.get("updates", [])),
            "new_skus": new_skus,
            "sku_warnings": len(diff_result.get("warnings", [])),
            "warnings": diff_result.get("warnings", []),
            "updates": diff_result.get("updates", []),
        }

    @staticmethod
    def export_report_csv(report: Dict[str, Any], output_path: str):
        """Export report to CSV files.

        Creates:
        - {base_name}_summary.csv: Summary statistics
        - {base_name}_updates.csv: Detailed cell updates

        Args:
            report: Report dictionary.
            output_path: Base output path for generating file names.
        """
        dir_path = os.path.dirname(output_path)
        base_name = os.path.splitext(os.path.basename(output_path))[0]

        # Export summary
        summary_path = os.path.join(dir_path, f"{base_name}_summary.csv")

        with open(summary_path, "w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerow(["Metric", "Value"])
            writer.writerow(["Timestamp", report["timestamp"]])
            writer.writerow(["Source Rows", report["source_rows"]])
            writer.writerow(["Target Rows (Old)", report["target_rows_old"]])
            writer.writerow(["Target Rows (New)", report["target_rows_new"]])
            writer.writerow(
                [
                    "Columns Matched",
                    f"{report['columns_matched']}/{report['columns_total']}",
                ]
            )
            writer.writerow(["New Columns Added", report["new_columns_added"]])
            writer.writerow(["Cells Updated", report["cells_updated"]])
            writer.writerow(["New SKUs Added", report["new_skus"]])
            writer.writerow(["SKU Warnings", report["sku_warnings"]])
            writer.writerow(["Prefix Used", report["prefix_used"]])

        # Export detailed updates
        if report["updates"]:
            updates_path = os.path.join(dir_path, f"{base_name}_updates.csv")

            with open(updates_path, "w", newline="", encoding="utf-8") as f:
                writer = csv.DictWriter(
                    f,
                    fieldnames=[
                        "sku",
                        "column",
                        "source_column",
                        "old_value",
                        "new_value",
                    ],
                )
                writer.writeheader()
                writer.writerows(report["updates"])

        # Export warnings if any
        if report["warnings"]:
            warnings_path = os.path.join(
                dir_path, f"{base_name}_warnings.json"
            )

            with open(warnings_path, "w", encoding="utf-8") as f:
                json.dump(report["warnings"], f, indent=2, default=str)

    @staticmethod
    def export_report_json(report: Dict[str, Any], output_path: str):
        """Export full report as JSON file.

        Limits the updates array to first 100 entries to keep
        file size manageable.

        Args:
            report: Report dictionary.
            output_path: Base output path for generating file name.
        """
        json_path = output_path.replace(".xlsx", "_report.json")

        # Create a copy without large update arrays
        report_copy = report.copy()
        if "updates" in report_copy:
            report_copy["updates_count"] = len(report_copy["updates"])
            report_copy["updates_sample"] = report_copy["updates"][:100]
            del report_copy["updates"]

        with open(json_path, "w", encoding="utf-8") as f:
            json.dump(report_copy, f, indent=2, default=str)
