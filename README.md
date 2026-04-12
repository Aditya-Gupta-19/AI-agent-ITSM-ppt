
## Prerequisites
- Python 3.11+
- Ollama installed: https://ollama.ai/download
- Run: `ollama pull phi3` (downloads ~2GB, one-time)
- VS Code with Python extension

## Setup
```bash
# 1. Clone / open project in VS Code
# 2. Create virtual environment
python -m venv venv
venv\Scripts\activate   # Windows

# 3. Install dependencies
pip install -r requirements.txt

# 4. Start Ollama (must be running before agent starts)
ollama serve

# 5. Edit config.yaml — set excel.watch_path to your Excel file

# 6. Run once (for testing)
python main.py --run-once

# 7. Run in watch mode (continuous)
python main.py
```

## Expected Output
- PPTX saved to ./output/ITSM_Report_YYYYMMDD.pptx
- Log file at ./output/pipeline.log
- Console shows progress for each sheet

## Switching to Azure OpenAI
Change config.yaml:
```yaml
ai:
  provider: "azure_openai"
  model: "gpt-4o"
  azure_endpoint: "https://YOUR-RESOURCE.openai.azure.com/"
  api_version: "2024-02-01"
```
Add to .env: AZURE_OPENAI_KEY=your_key_here

## CRITICAL RULES for Claude Code

1. **Build all files in the exact structure above** — do not flatten into one file
2. **Every tool is a class** — not a collection of functions
3. **Orchestrator uses dependency injection** — tools passed in constructor
4. **Never hardcode file paths** — always read from config.yaml
5. **All errors are caught at the sheet level** — one bad sheet never crashes the pipeline
6. **AI output is always validated** — if JSON missing any key, use fallback
7. **Charts are always optional** — if matplotlib fails, slide is built without chart
8. **Sample Excel file must be generated programmatically** — not a manual file
9. **Run tests after building**: `pytest tests/ -v` — all tests must pass
10. **Print a clear pipeline summary at the end**:
✅ Pipeline complete
Sheets processed : 5
Sheets skipped   : 0
Output saved to  : ./output/ITSM_Report_20260407.pptx
Duration         : 23.4 seconds

## Build Order for Claude Code

Build in this exact sequence:
1. requirements.txt
2. config.yaml + .env.example
3. sample_data/sample_itsm.xlsx (programmatically)
4. tools/t8_logger.py
5. tools/t2_excel_reader.py + tests
6. tools/t3_kpi_parser.py + tests
7. tools/t4_ai_engine.py + tests
8. tools/t5_chart_generator.py
9. tools/t6_pptx_builder.py
10. tools/t1_file_watcher.py
11. workflows/w1_ingest.py through w5_slide_builder.py
12. agent/orchestrator.py
13. main.py
14. README.md
15. Run: python main.py --run-once
16. Fix any errors, re-run until output PPTX is generated.


