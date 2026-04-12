import argparse
import os
from datetime import datetime
from typing import Any

from dotenv import load_dotenv
from openpyxl import Workbook
from openpyxl.styles import Alignment

from agent.orchestrator import ReportOrchestrator
import yaml


def _load_config(config_path: str) -> dict:
    with open(config_path, "r", encoding="utf-8") as f:
        return yaml.safe_load(f) or {}


def _ensure_sample_excel_exists(excel_path: str) -> None:
    if os.path.exists(excel_path):
        return

    os.makedirs(os.path.dirname(os.path.abspath(excel_path)), exist_ok=True)

    wb = Workbook()

    # Helper for percent cells: store as ratio 0..1.
    def set_percent(ws, cell, value_ratio):
        cell.value = float(value_ratio)
        cell.number_format = "0.00%"
        cell.alignment = Alignment(horizontal="center")

    # Sheet 1: Cover Page
    ws_cover = wb.active
    ws_cover.title = "Cover Page"

    ws_cover.append(["MIM", "Delivery Percent (>70%)", "Target: Delivery Percent (>70%)"])
    ws_cover.append(["", "First-Time Right Delivery (>85%)", "Target: First-Time Right Delivery (>85%)"])

    ws_cover.append(["Change", "Change Success Ratio (>99%)", "Target: Change Success Ratio (>99%)"])
    ws_cover.append(["", "Change Causing MI vs Other MIs (<10%)", "Target: Change Causing MI vs Other MIs (<10%)"])
    ws_cover.append(["", "Failed Change Ratio (<1%)", "Target: Failed Change Ratio (<1%)"])
    ws_cover.append(["", "Unauthorised Change rate (<1%)", "Target: Unauthorised Change rate (<1%)"])

    ws_cover.append(["Problem", "Avg Resolution Days (<5)", "Target: Avg Resolution Days (<5)"])
    ws_cover.append(["", "Recurring Issues (<3)", "Target: Recurring Issues (<3)"])

    ws_cover.append(["CMDB", "CI Accuracy (>95%)", "Target: CI Accuracy (>95%)"])
    ws_cover.append(["", "Orphan CIs", "Target: Orphan CIs (Lower is better)"])
    ws_cover.append(["", "Stale Records (>30 days)", "Target: Stale Records (>30 days) (Lower is better)"])
    ws_cover.append(["", "Audit Compliance (>98%)", "Target: Audit Compliance (>98%)"])

    ws_cover.append(["NOC", "MTTR Hours (<2)", "Target: MTTR Hours (<2)"])
    ws_cover.append(["", "SLA Compliance (>95%)", "Target: SLA Compliance (>95%)"])

    # Sheet 2: MIM
    ws_mim = wb.create_sheet("MIM")
    ws_mim.append(
        [
            "Sprint No",
            "No. of Stories taken in Sprint planning (tickets)",
            "Adhoc Stories taken",
            "Total stories handled in sprint",
            "Stories Completed",
            "Delivery Percent (>70%)",
            "First-Time Right Delivery (>85%)",
        ]
    )
    ws_mim.append(["Sprint 16/01", 20, 3, 23, 11, None, None])
    # row 2 col F/G are percent cells (0-indexed in append)
    set_percent(ws_mim, ws_mim["F2"], 0.4783)
    set_percent(ws_mim, ws_mim["G2"], 0.78)

    # Sheet 3: Change
    ws_change = wb.create_sheet("Change")
    ws_change.append(
        [
            "Week No",
            "Change Success Ratio (>99%)",
            "Change Causing MI vs Other MIs (<10%)",
            "Failed Change Ratio (<1%)",
            "Unauthorised Change rate (<1%)",
        ]
    )
    ws_change.append(["4", None, None, None, None])
    set_percent(ws_change, ws_change["B2"], 1.00)
    set_percent(ws_change, ws_change["C2"], 0.00)
    set_percent(ws_change, ws_change["D2"], 0.00)
    set_percent(ws_change, ws_change["E2"], 0.0061)

    # Sheet 4: Problem
    ws_problem = wb.create_sheet("Problem")
    ws_problem.append(["Week No", "Open Problems", "Problems Resolved", "Avg Resolution Days (<5)", "Recurring Issues (<3)"])
    ws_problem.append(["Week 4", 12, 8, 4.2, 2])

    # Sheet 5: CMDB
    ws_cmdb = wb.create_sheet("CMDB")
    ws_cmdb.append(["Week No", "CI Accuracy (>95%)", "Orphan CIs", "Stale Records (>30 days)", "Audit Compliance (>98%)"])
    ws_cmdb.append(["Week 4", None, 5, 12, None])
    set_percent(ws_cmdb, ws_cmdb["B2"], 0.975)
    set_percent(ws_cmdb, ws_cmdb["E2"], 0.991)

    # Sheet 6: NOC
    ws_noc = wb.create_sheet("NOC")
    ws_noc.append(["Week No", "P1 Incidents", "P2 Incidents", "MTTR Hours (<2)", "SLA Compliance (>95%)"])
    ws_noc.append(["Week 4", 2, 7, 1.8, None])
    set_percent(ws_noc, ws_noc["E2"], 0.964)

    # Save.
    wb.save(excel_path)


def main():
    parser = argparse.ArgumentParser(description="ITSM Report Automation Agent — POC (Local Machine)")
    parser.add_argument("--run-once", action="store_true", help="Run the pipeline once and exit")
    args = parser.parse_args()

    # Load env (optional).
    load_dotenv()

    repo_dir = os.path.dirname(os.path.abspath(__file__))
    config_path = os.path.join(repo_dir, "config.yaml")
    config = _load_config(config_path)

    excel_watch_path = config.get("excel", {}).get("watch_path") or "./sample_data/sample_itsm.xlsx"
    excel_watch_path_abs = os.path.abspath(os.path.join(repo_dir, excel_watch_path)) if not os.path.isabs(excel_watch_path) else excel_watch_path

    _ensure_sample_excel_exists(excel_watch_path_abs)

    orchestrator = ReportOrchestrator(config)

    if args.run_once:
        result = orchestrator.run_pipeline(excel_watch_path_abs)
        print("\nPipeline Result:")
        for k in ["status", "slides_generated", "slides_skipped", "output_path", "duration_seconds"]:
            print(f"{k}: {result.get(k)}")
        return

    print(f"Watching Excel file for changes: {excel_watch_path_abs}")

    # Lazy import to keep startup fast.
    from tools.t1_file_watcher import ExcelFileWatcher

    def on_modified(_excel_path: str):
        result = orchestrator.run_pipeline(_excel_path)
        # Keep console output simple for POC.
        print(f"Pipeline finished: {result.get('status')} -> {result.get('output_path')}")

    watcher = ExcelFileWatcher(excel_watch_path_abs, on_modified)
    watcher.start()


if __name__ == "__main__":
    main()

