from datetime import datetime  # get current timestamps for filenames and logs
from pathlib import Path  # easy filesystem path handling
import argparse  # parse command-line arguments
import csv  # write CSV output when requested
import logging  # logging to file/console
import time  # sleep for repeating runs
try:
    import yaml  # optional: load YAML config if installed
except Exception:
    yaml = None  # gracefully handle missing PyYAML
import os  # access environment variables

import snowflake.connector  # Snowflake DB connector library
from openpyxl import Workbook  # create Excel workbooks for summary output

# Embedded defaults (fallback credentials - replace/remove for production)
DEFAULT_CREDENTIALS = {
    'account': 'EXITYOV-QE33421',  # Snowflake account identifier (from Snowflake)
    'user': 'DIWAKAR01P',  # Snowflake username
    'password': 'N!r68Z$e@290792',  # Snowflake password (keep for now, replace before production)
}

def load_config(path: Path):
    """Load YAML config if present.

    Returns an empty dict if no path provided, file missing, or PyYAML
    is not installed. When present, returns the parsed YAML mapping.
    """
    if not path or not path.exists():
        return {}  # no config provided
    if yaml is None:
        logging.warning('PyYAML not installed; skipping config file %s', path)
        return {}  # can't parse YAML without PyYAML
    with open(path, 'r', encoding='utf-8') as f:
        return yaml.safe_load(f) or {}  # return parsed config or empty mapping


def setup_logging(log_path: Path, verbose: bool = False):
    """Configure file+console logging.

    Creates parent folder for the logfile and sets logging level based
    on the `verbose` flag. Logs go to both the file and stdout.
    """
    log_path.parent.mkdir(parents=True, exist_ok=True)  # ensure logs folder exists
    level = logging.DEBUG if verbose else logging.INFO  # debug when verbose
    logging.basicConfig(
        level=level,
        format='%(asctime)s %(levelname)s: %(message)s',
        handlers=[
            logging.FileHandler(log_path, encoding='utf-8'),  # write to file
            logging.StreamHandler()  # also print to console
        ]
    )


def get_credentials(cli_args, cfg):
    """Resolve Snowflake credentials in order of precedence.

    Priority: CLI args > environment variables > YAML config > embedded defaults.
    Returns a tuple (account, user, password).
    If password is provided, uses password auth. Otherwise uses browser auth.
    """
    account = (
        cli_args.account
        or cfg.get('account')
        or os.getenv('SNOWFLAKE_ACCOUNT')
        or DEFAULT_CREDENTIALS.get('account')
    )  # final account value to use
    user = (
        cli_args.user
        or cfg.get('user')
        or os.getenv('SNOWFLAKE_USER')
        or DEFAULT_CREDENTIALS.get('user')
    )  # final user value to use
    password = (
        cli_args.password
        or os.getenv('SNOWFLAKE_PASSWORD')
        or cfg.get('password')
        or DEFAULT_CREDENTIALS.get('password')
    )  # final password value to use
    return account, user, password


def list_sql_files(folder: Path, pattern: str):
    """Return sorted list of SQL/text files in `folder` matching `pattern`.

    Filters out hidden files and those starting with an underscore.
    """
    files = sorted(folder.glob(pattern))  # collect matching path objects
    files = [f for f in files if f.is_file() and not f.name.startswith('_')]  # filter
    return files


def run_sql_files(conn_params, files, output_excel: Path, output_csv: Path | None, log):
    """Connect to Snowflake, run each SQL statement, and write results.

    - `conn_params` is a dict of Snowflake connection options
    - `files` is an iterable of Path objects to read SQL from
    - returns a small stats dict with counts
    """
    conn = snowflake.connector.connect(**conn_params)  # open Snowflake connection
    cursor = conn.cursor()  # get a cursor for queries

    wb = Workbook()  # create an Excel workbook for summary
    ws = wb.active  # get the active worksheet
    ws.title = 'summary'  # name the sheet
    ws.append(['OBJ', 'TEST_CASE', 'CNT', 'source_file'])  # header row

    csv_file = None
    csv_writer = None
    if output_csv:
        output_csv.parent.mkdir(parents=True, exist_ok=True)  # ensure folder exists
        csv_file = open(output_csv, 'w', newline='', encoding='utf-8')  # open CSV for writing
        csv_writer = csv.writer(csv_file)
        csv_writer.writerow(['OBJ', 'TEST_CASE', 'CNT', 'source_file'])  # CSV header

    total_files = 0  # count of files processed
    total_statements = 0  # count of SQL statements executed
    total_rows = 0  # count of result rows written

    for sql_file in files:
        total_files += 1  # increment file counter
        sql_text = sql_file.read_text(encoding='utf-8')  # load file contents
        if not sql_text.strip():
            logging.info('Skipping empty file %s', sql_file.name)  # skip empty files
            continue

        statements = [s.strip() for s in sql_text.split(';') if s.strip()]  # split by semicolon
        for idx, stmt in enumerate(statements, start=1):
            total_statements += 1  # increment statement counter
            logging.debug('Running %s statement %d', sql_file.name, idx)  # debug log
            try:
                cursor.execute(stmt)  # execute the SQL statement
            except Exception as e:
                logging.error('Error executing %s statement %d: %s', sql_file.name, idx, e)  # log error
                log.write(f'ERROR {sql_file.name} stmt {idx}: {e}\n')  # append to run log
                continue  # move to next statement

            try:
                rows = cursor.fetchall()  # attempt to fetch results
            except Exception:
                rows = []  # some statements (e.g., DDL) return no rows

            if not rows:
                ws.append(['NO RESULT', '', '', sql_file.name])  # note no result in Excel
                if csv_writer:
                    csv_writer.writerow(['NO RESULT', '', '', sql_file.name])  # and CSV
                continue

            for row in rows:
                values = [str(v) for v in row]  # stringify row values
                values += [''] * (3 - len(values))  # ensure at least 3 columns
                ws.append([values[0], values[1], values[2], sql_file.name])  # write to Excel
                if csv_writer:
                    csv_writer.writerow([values[0], values[1], values[2], sql_file.name])  # write to CSV
                total_rows += 1  # increment row counter

    output_excel.parent.mkdir(parents=True, exist_ok=True)  # ensure output folder exists
    wb.save(output_excel)  # save workbook to disk
    if csv_file:
        csv_file.close()  # close CSV if used
    cursor.close()  # close DB cursor
    conn.close()  # close DB connection

    return {'files': total_files, 'statements': total_statements, 'rows': total_rows}  # return stats


def main():
    # build CLI parser and arguments
    p = argparse.ArgumentParser(description='Run SQL files against Snowflake and save a summary')
    p.add_argument('--config', help='YAML config file', type=Path)  # optional YAML config
    p.add_argument('--folder', help='SQL folder', type=Path)  # folder containing .txt SQL files
    p.add_argument('--pattern', help='File glob pattern', default='*.txt')  # glob for filenames
    p.add_argument('--output-format', choices=['xlsx', 'csv', 'both'], default='xlsx')  # output type
    p.add_argument('--account', help='Snowflake account')  # override account
    p.add_argument('--user', help='Snowflake user')  # override user
    p.add_argument('--password', help='Snowflake password (optional; uses browser auth if not provided)')  # override password
    p.add_argument('--repeat', type=int, help='Repeat every N minutes (optional)')  # scheduler interval
    p.add_argument('--verbose', action='store_true')  # more logging output
    args = p.parse_args()

    cfg = load_config(args.config) if args.config else {}  # load YAML config if provided

    folder = args.folder or Path(cfg.get('folder') or Path.cwd())  # resolve SQL folder
    if not folder.exists():
        print('Folder does not exist:', folder)  # early exit if folder missing
        return

    files = list_sql_files(folder, args.pattern)  # find SQL files
    if not files:
        print('No SQL files found')  # nothing to do
        return

    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')  # timestamp for output files
    output_base = Path(cfg.get('output_base') or folder)  # base folder for outputs
    output_excel = output_base / f'summary_{timestamp}.xlsx'  # Excel output path
    output_csv = None
    if args.output_format in ('csv', 'both'):
        output_csv = output_base / f'summary_{timestamp}.csv'  # CSV output path when requested

    log_path = output_base / 'logs' / f'run_{timestamp}.log'  # run log path
    setup_logging(log_path, args.verbose)  # initialize logging
    logging.info('Starting run; files=%d', len(files))  # log start

    account, user, password = get_credentials(args, cfg)  # resolve credentials

    conn_params = dict(
        account=account,
        user=user,
        role=cfg.get('role') or 'ACCOUNTADMIN',  # default role
        warehouse=cfg.get('warehouse') or 'COMPUTE_WH',  # default warehouse
        database=cfg.get('database') or 'SNOWFLAKE_SAMPLE_DATA',  # default database
        schema=cfg.get('schema') or 'TPCDS_SF100TCL',  # default schema
    )
    
    # Use password auth if provided; otherwise use browser-based authentication
    if password:
        conn_params['password'] = password
        conn_params['authenticator'] = 'snowflake'  # standard password authentication
        logging.info('Using password-based authentication')
    else:
        conn_params['authenticator'] = 'externalbrowser'  # browser-based authentication
        logging.info('Using browser-based authentication (externalbrowser)')

    while True:
        with open(log_path, 'a', encoding='utf-8') as log:  # append-run log file
            log.write(f'Run started: {datetime.now()}\n')  # write run header
            log.write(f'Files: {len(files)}\n')  # write number of files
            stats = run_sql_files(conn_params, files, output_excel, output_csv, log)  # execute work
            log.write(f"Completed: files={stats['files']} statements={stats['statements']} rows={stats['rows']}\n")
        logging.info('Completed run: %s', stats)  # info-level completion message
        if not args.repeat:
            break  # exit after one run unless repeat specified
        logging.info('Sleeping for %d minutes', args.repeat)  # about to sleep
        time.sleep(args.repeat * 60)  # sleep then repeat


if __name__ == '__main__':
    main()
