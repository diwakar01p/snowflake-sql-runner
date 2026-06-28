Connection_snowflake

Purpose
- Run SQL text files in a folder against Snowflake, collect results into a timestamped Excel workbook and optional CSV, and write per-run logs.

Quick start
1. (Optional) Create a virtual environment and install deps:

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

2. By default credentials are embedded in `Connection_snowflake.py`. To run as-is:

```powershell
python Connection_snowflake.py --folder "C:\Users\diwakar\OneDrive\Desktop\python_files"
```

3. To enable CSV output as well:

```powershell
python Connection_snowflake.py --folder "C:\Users\diwakar\OneDrive\Desktop\python_files" --output-format both
```

Notes
- Credentials are currently embedded in the script (per your request). To remove them later, set `SNOWFLAKE_ACCOUNT`, `SNOWFLAKE_USER`, `SNOWFLAKE_PASSWORD` environment variables or add them to `config.yaml` and install `PyYAML`.
- Logs are written to `logs/run_<timestamp>.log` under the output folder.
- Output Excel and CSV files are timestamped to avoid overwriting previous runs.

Files
- `Connection_snowflake.py` — main script (credentials currently embedded)
- `config.yaml` — optional config template
- `requirements.txt` — Python dependencies

Next steps
- Replace embedded credentials with environment/config-based secrets when ready.
