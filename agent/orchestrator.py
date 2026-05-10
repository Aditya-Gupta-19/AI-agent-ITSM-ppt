import os
import re
import time
import json
import hashlib
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime
from typing import Any, Dict, List, Optional

from tools.t1_file_watcher import ExcelFileWatcher  # noqa: F401
from tools.t2_excel_reader import ExcelReader
from tools.t3_kpi_parser import KPIParser
from tools.t4_ai_engine import RuleEngine
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
        rule_engine: Optional[RuleEngine] = None,
        chart_generator: Optional[ChartGenerator] = None,
        pptx_builder: Optional[PPTXBuilder] = None,
    ):
        self.config = config
        log_cfg = config.get("logging", {}) if isinstance(config, dict) else {}
        log_file = log_cfg.get("log_file")

        self.logger = logger or get_logger("orchestrator", log_file=log_file)
        self.excel_reader = excel_reader or ExcelReader(config)
        self.kpi_parser = kpi_parser or KPIParser()
        self.rule_engine = rule_engine or RuleEngine(config)
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

    def _hash_sheet(self, sheet_info: Dict[str, Any], team_config: Optional[dict] = None) -> str:
        df = sheet_info.get("dataframe")
        if df is None:
            return "empty"
        serialized = df.fillna("").to_json(orient="split", date_format="iso")
        if team_config:
            serialized += json.dumps(team_config, sort_keys=True, default=str)
        return hashlib.sha256(serialized.encode("utf-8")).hexdigest()

    @staticmethod
    def _resolve_team_config(sheet_name: str, report_config: dict) -> dict:
        """
        Look up report_config for sheet_name with fuzzy fallback.
        Handles the Excel 31-char sheet name truncation: if the exact key is
        missing, try the longest config key that the sheet_name starts with
        (prefix match, case-insensitive, min 10 chars).
        """
        if not sheet_name or not report_config:
            return {}
        # Exact match first.
        exact = report_config.get(sheet_name)
        if exact is not None:
            return exact
        sn_lower = sheet_name.lower().strip()
        best_key, best_len = None, 0
        for key in report_config:
            kl = key.lower().strip()
            min_len = min(len(sn_lower), len(kl))
            if min_len < 10:
                continue
            # sheet_name starts with key prefix OR key starts with sheet_name
            if sn_lower.startswith(kl[:min_len]) or kl.startswith(sn_lower[:min_len]):
                if min_len > best_len:
                    best_key, best_len = key, min_len
        return report_config.get(best_key, {}) if best_key else {}

    @staticmethod
    def _deduplicate_sheets(sheet_names: list) -> list:
        """Skip 'MIM1' if 'MIM' exists; 'MIM (2)' if 'MIM' exists."""
        base_names = set()
        for name in sheet_names:
            base = re.sub(r'[\s\(\)]*\d+[\s\(\)]*$', '', name).strip()
            if base == name:
                base_names.add(name)

        result = []
        for name in sheet_names:
            base = re.sub(r'[\s\(\)]*\d+[\s\(\)]*$', '', name).strip()
            if name in base_names or base not in base_names:
                result.append(name)
        return result

    def _process_one_sheet(
        self,
        sheet_name: str,
        sheet_info: Dict[str, Any],
        report_config: dict,
        kpi_dict: dict,
        kpi_definitions: dict,
        default_summary_mode: str,
        excel_summary_max_chars: int,
    ) -> Optional[Dict[str, Any]]:
        """Process one sheet: KPI eval + AI + charts. Returns slide item or None if empty."""
        # Ghost-slide guard: skip blank sheet names
        if not sheet_name or str(sheet_name).strip().lower() in ("", "nan"):
            return None

        is_empty = bool(sheet_info.get("is_empty"))
        row_count = int(sheet_info.get("row_count", 0) or 0)
        df = sheet_info.get("dataframe")
        if is_empty or row_count == 0 or df is None or df.empty:
            self.logger.info(f"Skipping sheet '{sheet_name}' (empty).")
            return None
        # Drop all-NaN rows; re-check after cleaning
        df = df.dropna(how="all")
        if df.empty:
            self.logger.info(f"Skipping sheet '{sheet_name}' (all rows empty after cleaning).")
            return None

        team_config = self._resolve_team_config(sheet_name, report_config)
        layout = team_config.get("layout", "standard")
        rag_thresholds = {
            "green": float(team_config.get("green_threshold", 95.0)),
            "amber": float(team_config.get("amber_threshold", 90.0)),
        }
        summary_mode = team_config.get("summary_mode") or default_summary_mode
        if layout == "multigroup":
            summary_mode = "use_excel"
        excel_summary = self.excel_reader.get_user_summary(df)
        excel_summary = (excel_summary or "")[:excel_summary_max_chars]

        parsed_sections = ExcelReader.parse_weekly_comments(excel_summary)

        thresholds = self.kpi_parser.parse_header_thresholds(sheet_info.get("headers") or [])
        last_row = df.iloc[-1]
        row_data = {str(col): last_row[col] for col in df.columns}
        column_types = sheet_info.get("column_types") or {}
        kpi_rows = self.kpi_parser.evaluate_kpis(
            row_data=row_data, thresholds=thresholds, column_types=column_types
        )

        skip_context_kw = {"comment", "achievement", "week no", "sprint no",
                           "week", "sprint", "month", "quarter", "date"}
        context_rows = []
        for col in df.columns:
            col_str = str(col)
            if col_str in thresholds or column_types.get(col_str) != "numeric":
                continue
            if any(kw in col_str.lower() for kw in skip_context_kw):
                continue
            series = df[col].dropna()
            if series.empty:
                continue
            val = series.iloc[-1]
            try:
                fval = float(val)
                disp = f"{int(fval)}" if fval == int(fval) else f"{fval:.1f}"
            except Exception:
                disp = str(val)
            context_rows.append({"kpi_name": col_str, "actual_value": disp,
                                  "threshold": "—", "status": "INFO"})

        multigroup_data: dict = {}
        if layout == "multigroup":
            multigroup_data = self.excel_reader.read_multigroup_data(df)

        ai_output = w3_run(
            sheet_name=sheet_name, sheet_info=sheet_info,
            ai_engine=self.rule_engine, kpi_dict=kpi_dict, logger=self.logger,
            excel_summary=excel_summary, summary_mode=summary_mode,
            kpi_rows=kpi_rows,
        )
        if layout == "multigroup":
            chart_bytes, charts_list = None, []
        else:
            chart_bytes, charts_list = w4_run(
                sheet_name=sheet_name, sheet_info=sheet_info,
                chart_generator=self.chart_generator, logger=self.logger,
                report_config=report_config,
            )

        overall_rag = ai_output.get("overall_rag", "AMBER")
        all_kpi_rows = kpi_rows + context_rows + [{
            "kpi_name": "Overall Performance", "actual_value": overall_rag,
            "threshold": "—", "status": overall_rag,
        }]
        return {
            "sheet_name": sheet_name,
            "ai_output": ai_output,
            "kpi_rows": all_kpi_rows,
            "chart_bytes": chart_bytes,
            "charts_list": charts_list,
            "summary_mode": summary_mode,
            "df": df,
            "column_types": column_types,
            "kpi_definitions": kpi_definitions.get(sheet_name, []),
            "layout": layout,
            "multigroup_data": multigroup_data,
            "rag_thresholds": rag_thresholds,
            "parsed_sections": parsed_sections,
        }

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
        summary_cfg = self.config.get("summary", {}) if isinstance(self.config, dict) else {}
        default_summary_mode = summary_cfg.get("default_mode", "ai_write")
        excel_summary_max_chars = int(summary_cfg.get("excel_summary_max_chars", 500))

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

            # Load Report_Config from Excel (user-controlled chart/summary settings).
            try:
                report_config = self.excel_reader.read_report_config(excel_path)
                self.logger.info(f"Loaded report config for {len(report_config)} teams.")
            except Exception as e:
                self.logger.warning(f"Could not read Report_Config: {e}")
                report_config = {}

            # Load KPI definitions from Cover Page (for definition table on slides).
            try:
                kpi_definitions = self.excel_reader.read_kpi_definitions(excel_path)
            except Exception:
                kpi_definitions = {}

            prev_hashes = self._load_state(state_path)
            current_hashes: Dict[str, str] = {}
            for sheet_name, sheet_info in sheets_data.items():
                # Include team's report_config in the hash so chart-type changes
                # only trigger that specific team's slide, not all teams.
                team_cfg = self._resolve_team_config(sheet_name, report_config) if sheet_name not in sheets_to_skip else None
                current_hashes[sheet_name] = self._hash_sheet(sheet_info, team_cfg)

            kpi_changed = prev_hashes.get(kpi_sheet_name) != current_hashes.get(kpi_sheet_name)

            # Deduplicate numbered sheet variants (e.g. MIM1, APAC SD1).
            all_processable = self._deduplicate_sheets([
                s for s in sheets_data.keys() if s not in sheets_to_skip
            ])

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

            # Step 3: Collect sheets to process, then run all in parallel.
            processable_items = []
            for sheet_name, sheet_info in sheets_data.items():
                if not sheet_name or str(sheet_name).strip().lower() in ("", "nan"):
                    continue
                if sheet_name in sheets_to_skip:
                    self.logger.info(f"Skipping sheet '{sheet_name}' (in skip list).")
                    slides_skipped += 1
                    continue
                if sheet_name not in changed_sheets:
                    continue
                team_config = self._resolve_team_config(sheet_name, report_config)
                if team_config.get("skip", False):
                    self.logger.info(f"Skipping '{sheet_name}' (Report_Config skip=yes).")
                    slides_skipped += 1
                    continue
                processable_items.append((sheet_name, sheet_info))

            slide_queue: List[Dict[str, Any]] = []
            n_workers = min(len(processable_items), 4)
            with ThreadPoolExecutor(max_workers=n_workers) as pool:
                futures = [
                    pool.submit(
                        self._process_one_sheet,
                        sn, si, report_config, kpi_dict, kpi_definitions,
                        default_summary_mode, excel_summary_max_chars,
                    )
                    for sn, si in processable_items
                ]
                for (sheet_name, _), fut in zip(processable_items, futures):
                    try:
                        result = fut.result()
                        if result is None:
                            slides_skipped += 1
                        else:
                            slide_queue.append(result)
                            slides_generated += 1
                    except Exception as sheet_exc:
                        msg = f"Sheet '{sheet_name}' failed: {sheet_exc}"
                        self.logger.warning(msg)
                        errors.append(msg)
                        slides_skipped += 1

            # Step 4: Run W5 (build PPTX)
            if not slide_queue:
                raise RuntimeError("No slides could be generated (all sheets skipped/failed).")

            # Save state BEFORE PPTX write so next run skips unchanged sheets
            # even if the PPTX save fails (e.g. file locked in PowerPoint).
            self._save_state(state_path, current_hashes)

            self.pptx_builder.start_or_load_report(output_path)
            saved_path = w5_run(
                pptx_builder=self.pptx_builder,
                report_date=report_date,
                groups_queue=slide_queue,
                total_group_count=len(all_processable),
                output_path=output_path,
            )
            output_path = saved_path  # may be _draft.pptx if original was locked

            duration = time.time() - start
            summary_status = "success" if not errors else "partial"
            self.logger.info(
                f"Pipeline complete\nSheets processed : {len(slide_queue)}\nSheets skipped   : {slides_skipped}\n"
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
