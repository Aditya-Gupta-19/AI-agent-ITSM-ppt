import os
import time
import json
import hashlib
from datetime import datetime
from typing import Any, Dict, List, Optional

from tools.t1_file_watcher import ExcelFileWatcher  # noqa: F401
from tools.t2_excel_reader import ExcelReader
from tools.t3_kpi_parser import KPIParser
from tools.t4_ai_engine import AIEngine
from tools.t5_chart_generator import ChartGenerator
from tools.t6_pptx_builder import PPTXBuilder
from tools.t8_logger import get_logger

from workflows.w1_ingest import run as w1_run
from workflows.w2_parse import run as w2_run
from workflows.w3_ai_process import run as w3_run
from workflows.w4_charts import run as w4_run
from workflows.w5_slide_builder import run as w5_run


class ReportOrchestrator:
    def __init__(
        self,
        config: dict,
        logger: Any = None,
        excel_reader: Optional[ExcelReader] = None,
        kpi_parser: Optional[KPIParser] = None,
        ai_engine: Optional[AIEngine] = None,
        chart_generator: Optional[ChartGenerator] = None,
        pptx_builder: Optional[PPTXBuilder] = None,
    ):
        self.config = config
        log_cfg = config.get("logging", {}) if isinstance(config, dict) else {}
        log_file = log_cfg.get("log_file")

        self.logger = logger or get_logger("orchestrator", log_file=log_file)
        self.excel_reader = excel_reader or ExcelReader(config)
        self.kpi_parser = kpi_parser or KPIParser()
        self.ai_engine = ai_engine or AIEngine(config)
        self.chart_generator = chart_generator or ChartGenerator()
        self.pptx_builder = pptx_builder or PPTXBuilder()

    def _state_file_path(self, output_folder: str, filename_prefix: str) -> str:
        return os.path.join(output_folder, f"{filename_prefix}_sheet_state.json")

    def _load_state(self, state_path: str) -> Dict[str, str]:
        if not os.path.exists(state_path):
            return {}
        try:
            with open(state_path, "r", encoding="utf-8") as f:
                payload = json.load(f) or {}
            return payload.get("sheet_hashes", {}) if isinstance(payload, dict) else {}
        except Exception:
            return {}

    def _save_state(self, state_path: str, sheet_hashes: Dict[str, str]) -> None:
        os.makedirs(os.path.dirname(state_path), exist_ok=True)
        payload = {"sheet_hashes": sheet_hashes}
        with open(state_path, "w", encoding="utf-8") as f:
            json.dump(payload, f, indent=2)

    def _hash_sheet(self, sheet_info: Dict[str, Any]) -> str:
        df = sheet_info.get("dataframe")
        if df is None:
            return "empty"
        serialized = df.fillna("").to_json(orient="split", date_format="iso")
        return hashlib.sha256(serialized.encode("utf-8")).hexdigest()

    def run_pipeline(self, excel_path: str) -> Dict[str, Any]:
        """
        Full pipeline execution. Returns a result dict with:
        {
          "status": "success" | "partial" | "failed",
          "slides_generated": int,
          "slides_skipped": int,
          "output_path": str,
          "errors": list[str],
          "duration_seconds": float
        }
        """

        start = time.time()
        errors: List[str] = []
        slides_generated = 0
        slides_skipped = 0

        output_cfg = self.config.get("output", {}) if isinstance(self.config, dict) else {}
        output_folder = output_cfg.get("folder", "./output")
        filename_prefix = output_cfg.get("filename_prefix", "ITSM_Report")
        report_date = datetime.now().strftime("%Y%m%d")
        output_path = os.path.join(output_folder, f"{filename_prefix}_{report_date}.pptx")
        state_path = self._state_file_path(output_folder, filename_prefix)

        excel_cfg = self.config.get("excel", {}) if isinstance(self.config, dict) else {}
        sheets_to_skip = set(excel_cfg.get("sheets_to_skip") or [])
        kpi_sheet_name = excel_cfg.get("kpi_sheet_name") or "Cover Page"

        try:
            # Step 1: Load and parse Excel (W1 + W2)
            self.logger.info("Starting pipeline.")
            sheets_data = w1_run(excel_reader=self.excel_reader, excel_path=excel_path, logger=self.logger)
            kpi_dict, sheets_data = w2_run(
                sheets_data=sheets_data,
                kpi_sheet_name=kpi_sheet_name,
                kpi_parser=self.kpi_parser,
                logger=self.logger,
            )

            prev_hashes = self._load_state(state_path)
            current_hashes: Dict[str, str] = {}
            for sheet_name, sheet_info in sheets_data.items():
                current_hashes[sheet_name] = self._hash_sheet(sheet_info)

            kpi_changed = prev_hashes.get(kpi_sheet_name) != current_hashes.get(kpi_sheet_name)

            all_processable = [
                s for s in sheets_data.keys() if s not in sheets_to_skip
            ]
            changed_sheets = set()
            if not os.path.exists(output_path) or not prev_hashes or kpi_changed:
                changed_sheets = set(all_processable)
            else:
                for s in all_processable:
                    if prev_hashes.get(s) != current_hashes.get(s):
                        changed_sheets.add(s)

            if not changed_sheets:
                duration = time.time() - start
                self.logger.info("No sheet changes detected. PPT not updated.")
                return {
                    "status": "success",
                    "slides_generated": 0,
                    "slides_skipped": 0,
                    "output_path": output_path,
                    "errors": [],
                    "duration_seconds": duration,
                }

            # Step 3: For each changed sheet, run AI + charts, then queue slides.
            slide_queue: List[Dict[str, Any]] = []
            for sheet_name, sheet_info in sheets_data.items():
                if sheet_name in sheets_to_skip:
                    self.logger.info(f"Skipping sheet '{sheet_name}' (in skip list).")
                    slides_skipped += 1
                    continue
                if sheet_name not in changed_sheets:
                    continue

                # Sheet must have data.
                try:
                    is_empty = bool(sheet_info.get("is_empty"))
                    row_count = int(sheet_info.get("row_count", 0) or 0)
                    df = sheet_info.get("dataframe")
                    if is_empty or row_count == 0 or df is None or df.empty:
                        self.logger.info(f"Skipping sheet '{sheet_name}' (empty).")
                        slides_skipped += 1
                        continue

                    # KPI evaluation rows for slide table.
                    thresholds = self.kpi_parser.parse_header_thresholds(sheet_info.get("headers") or [])
                    first_row = df.iloc[0]
                    row_data = {str(col): first_row.get(col) for col in df.columns}
                    kpi_rows = self.kpi_parser.evaluate_kpis(row_data=row_data, thresholds=thresholds)

                    # Step 3b: Run W3 (AI)
                    ai_output = w3_run(
                        sheet_name=sheet_name,
                        sheet_info=sheet_info,
                        ai_engine=self.ai_engine,
                        kpi_dict=kpi_dict,
                        logger=self.logger,
                    )

                    # Step 3c: Run W4 (Charts)
                    chart_bytes = w4_run(
                        sheet_name=sheet_name,
                        sheet_info=sheet_info,
                        chart_generator=self.chart_generator,
                        logger=self.logger,
                    )

                    slide_queue.append(
                        {
                            "sheet_name": sheet_name,
                            "ai_output": ai_output,
                            "kpi_rows": kpi_rows,
                            "chart_bytes": chart_bytes,
                        }
                    )
                    slides_generated += 1
                except Exception as sheet_exc:
                    msg = f"Sheet '{sheet_name}' failed: {sheet_exc}"
                    self.logger.warning(msg)
                    errors.append(msg)
                    slides_skipped += 1

            # Step 4: Run W5 (build PPTX)
            if not slide_queue:
                raise RuntimeError("No slides could be generated (all sheets skipped/failed).")

            self.pptx_builder.start_or_load_report(output_path)
            w5_run(
                pptx_builder=self.pptx_builder,
                report_date=report_date,
                groups_queue=slide_queue,
                total_group_count=len(all_processable),
                output_path=output_path,
            )
            self._save_state(state_path, current_hashes)

            # Step 5: Log summary
            duration = time.time() - start
            summary_status = "success" if not errors else "partial"
            self.logger.info(
                f"✅ Pipeline complete\nSheets processed : {len(slide_queue)}\nSheets skipped   : {slides_skipped}\n"
                f"Output saved to  : {output_path}\nDuration         : {duration:.1f} seconds"
            )

            return {
                "status": summary_status,
                "slides_generated": slides_generated,
                "slides_skipped": slides_skipped,
                "output_path": output_path,
                "errors": errors,
                "duration_seconds": duration,
            }
        except Exception as e:
            duration = time.time() - start
            err_msg = f"Pipeline failed: {e}"
            self.logger.error(err_msg)
            errors.append(err_msg)
            return {
                "status": "failed",
                "slides_generated": slides_generated,
                "slides_skipped": slides_skipped,
                "output_path": output_path,
                "errors": errors,
                "duration_seconds": duration,
            }

