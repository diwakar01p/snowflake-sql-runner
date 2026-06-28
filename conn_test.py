# Script to run SQL files against Snowflake using config.yaml for connection details
# Co-authored with CoCo
from datetime import datetime
from pathlib import Path
import argparse
import csv
import getpass
import logging
import time
try:
    import yaml
except Exception:
    yaml = None
import os

import snowflake.connector
from openpyxl import Workbook

# Default config file location (same directory as this script)
CONFIG_FILE = Path(__file__).parent / 'config.yaml'


def load_config(path: Path):
    """Load YAML config file."""
    if not path or not path.exists():
        return {}
    if yaml is None:
        print('ERROR: PyYAML is required. Install it with: pip install pyyaml')
        return {}
    with open(path, 'r', encoding='utf-8') as f:
        return yaml.safe_load(f) or {}


def setup_logging(log_path: Path, verbose: bool = False):
    """Configure file+console logging."""
    log_path.parent.mkdir(parents=True, exist_ok=True)
    level = logging.DEBUG if verbose else logging.INFO
    logging.basicConfig(
        level=level,
        format='%(asctime)s %(levelname)s: %(message)s',
        handlers=[
            logging.FileHandler(log_path, encoding='utf-8'),
            logging.StreamHandler()
        ]
    )


def list_sql_files(folder: Path, pattern: str):
    """Return sorted list of files in `folder` matching `pattern`."""
    files = sorted(folder.glob(pattern))
    files = [f for f in files if f.is_file() and not f.name.startswith('_')]
    return files


def run_sql_files(conn_params, files, output_excel: Path, output_csv, log):
    """Connect to Snowflake, run each SQL statement, and write results."""
    conn = snowflake.connector.connect(**conn_params)
    cursor = conn.cursor()

    wb = Workbook()
    ws = wb.active
    ws.title = 'summary'
    ws.append(['OBJ', 'TEST_CASE', 'CNT', 'source_file'])

    csv_file = None
    csv_writer = None
    if output_csv:
        output_csv.parent.mkdir(parents=True, exist_ok=True)
        csv_file = open(output_csv, 'w', newline='', encoding='utf-8')
        csv_writer = csv.writer(csv_file)
        csv_writer.writerow(['OBJ', 'TEST_CASE', 'CNT', 'source_file'])

    total_files = 0
    total_statements = 0
    total_rows = 0

    try:
        for sql_file in files:
            total_files += 1
            sql_text = sql_file.read_text(encoding='utf-8')
            if not sql_text.strip():
                logging.info('Skipping empty file %s', sql_file.name)
                continue

            statements = [s.strip() for s in sql_text.split(';') if s.strip()]
            for idx, stmt in enumerate(statements, start=1):
                total_statements += 1
                logging.debug('Running %s statement %d', sql_file.name, idx)
                try:
                    cursor.execute(stmt)
                except Exception as e:
                    logging.error('Error executing %s statement %d: %s', sql_file.name, idx, e)
                    log.write(f'ERROR {sql_file.name} stmt {idx}: {e}\n')
                    continue

                try:
                    rows = cursor.fetchall()
                except Exception:
                    rows = []

                if not rows:
                    ws.append(['NO RESULT', '', '', sql_file.name])
                    if csv_writer:
                        csv_writer.writerow(['NO RESULT', '', '', sql_file.name])
                    continue

                for row in rows:
                    values = [str(v) for v in row]
                    values += [''] * (3 - len(values))
                    ws.append([values[0], values[1], values[2], sql_file.name])
                    if csv_writer:
                        csv_writer.writerow([values[0], values[1], values[2], sql_file.name])
                    total_rows += 1
    finally:
        output_excel.parent.mkdir(parents=True, exist_ok=True)
        wb.save(output_excel)
        if csv_file:
            csv_file.close()
        cursor.close()
        conn.close()

    return {'files': total_files, 'statements': total_statements, 'rows': total_rows}


def main():
    p = argparse.ArgumentParser(description='Run SQL files against Snowflake and save a summary')
    p.add_argument('--config', help='Path to config.yaml', type=Path, default=CONFIG_FILE)
    p.add_argument('--folder', help='SQL folder (overrides config)', type=Path)
    p.add_argument('--output-format', choices=['xlsx', 'csv', 'both'], default='xlsx')
    p.add_argument('--repeat', type=int, help='Repeat every N minutes')
    p.add_argument('--verbose', action='store_true')
    args = p.parse_args()

    # Load config file
    cfg = load_config(args.config)
    if not cfg:
        print(f'Config file not found: {args.config}')
        print('Create a config.yaml with your connection details.')
        return

    print(f'Loaded config from: {args.config}')

    # Resolve folder and pattern from config
    folder = args.folder or Path(cfg.get('folder', '.'))
    pattern = cfg.get('pattern', '*.txt')

    if not folder.exists():
        print('Folder does not exist:', folder)
        return

    files = list_sql_files(folder, pattern)
    if not files:
        print('No files found matching pattern:', pattern)
        return

    print(f'Found {len(files)} file(s):')
    for f in files:
        print(f'  - {f.name}')

    # Only prompt for password at runtime (everything else from config)
    password = getpass.getpass('\nEnter Snowflake password: ')

    conn_params = dict(
        account=cfg.get('account', 'EXITYOV-QE33421'),
        user=cfg.get('user', 'DIWAKAR01P'),
        password=password,
        role=cfg.get('role', 'ACCOUNTADMIN'),
        warehouse=cfg.get('warehouse', 'COMPUTE_WH'),
        database=cfg.get('database', 'SNOWFLAKE_SAMPLE_DATA'),
        schema=cfg.get('schema', 'TPCDS_SF100TCL'),
    )

    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    output_base = folder
    output_excel = output_base / f'summary_{timestamp}.xlsx'
    output_csv = None
    if args.output_format in ('csv', 'both'):
        output_csv = output_base / f'summary_{timestamp}.csv'

    log_path = output_base / 'logs' / f'run_{timestamp}.log'
    setup_logging(log_path, args.verbose)
    logging.info('Connecting as %s to %s', conn_params['user'], conn_params['account'])
    logging.info('Starting run; files=%d', len(files))

    while True:
        with open(log_path, 'a', encoding='utf-8') as log:
            log.write(f'Run started: {datetime.now()}\n')
            log.write(f'Files: {len(files)}\n')
            stats = run_sql_files(conn_params, files, output_excel, output_csv, log)
            log.write(f"Completed: files={stats['files']} statements={stats['statements']} rows={stats['rows']}\n")
        logging.info('Completed run: %s', stats)
        if not args.repeat:
            break
        logging.info('Sleeping for %d minutes', args.repeat)
        time.sleep(args.repeat * 60)


if __name__ == '__main__':
    main()
