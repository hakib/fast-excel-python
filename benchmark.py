import datetime
import time

from typing import IO, Iterator


import python_calamine
def iter_excel_calamine(file: IO[bytes]) -> Iterator[dict[str, object]]:
    workbook = python_calamine.CalamineWorkbook.from_filelike(file)  # type: ignore[arg-type]
    rows = iter(workbook.get_sheet_by_index(0).to_python())
    headers = list(map(str, next(rows)))
    for row in rows:
        yield dict(zip(headers, row))


import openpyxl
def iter_excel_openpyxl(file: IO[bytes]) -> Iterator[dict[str, object]]:
    workbook = openpyxl.load_workbook(file, read_only=True)
    rows = workbook.active.rows
    headers = [str(cell.value) for cell in next(rows)]
    for row in rows:
        yield dict(zip(headers, (cell.value for cell in row)))


import tablib
def iter_excel_tablib(file: IO[bytes]) -> Iterator[dict[str, object]]:
    yield from tablib.Dataset().load(file).dict


import pandas
def iter_excel_pandas(file: IO[bytes]) -> Iterator[dict[str, object]]:
    yield from pandas.read_excel(file, converters={'date': lambda ts: ts.date()}).to_dict('records')


import duckdb
def iter_excel_duckdb(file: IO[bytes]) -> Iterator[dict[str, object]]:
    duckdb.install_extension('spatial')
    duckdb.load_extension('spatial')
    rows = duckdb.sql(f"""
        SELECT * FROM st_read(
            '{file.name}',
            open_options=['HEADERS=FORCE', 'FIELD_TYPES=AUTO'])
    """)
    while row := rows.fetchone():
        yield dict(zip(rows.columns, row))


import duckdb
def iter_excel_duckdb_execute(file: IO[bytes]) -> Iterator[dict[str, object]]:
    duckdb.install_extension('spatial')
    duckdb.load_extension('spatial')
    conn = duckdb.execute(
        "SELECT * FROM st_read(?, open_options=['HEADERS=FORCE', 'FIELD_TYPES=STRING'])",
        [file.name],
    )
    assert conn.description is not None
    headers = [header for header, *rest in conn.description]
    while row := conn.fetchone():
        yield dict(zip(headers, row))


import subprocess, tempfile, csv
def iter_excel_libreoffice(file: IO[bytes]) -> Iterator[dict[str, object]]:
    with tempfile.TemporaryDirectory(prefix='excelbenchmark') as tempdir:
        subprocess.run([
            'libreoffice', '--headless', '--convert-to', 'csv',
            '--outdir', tempdir, file.name,
        ])
        with open(f'{tempdir}/{file.name.rsplit(".")[0]}.csv', 'r') as f:
            rows = csv.reader(f)
            headers = list(map(str, next(rows)))
            for row in rows:
                yield dict(zip(headers, row))


#-- Benchmark

for fn in (
    iter_excel_pandas,
    iter_excel_tablib,
    iter_excel_openpyxl,
    iter_excel_libreoffice,
    iter_excel_duckdb,
    iter_excel_duckdb_execute,
    iter_excel_calamine,
):
    print(f'\n{fn.__qualname__}')
    with open('file.xlsx', 'rb') as file:
        try:
            start = time.perf_counter()
            rows = fn(file)
            control_row = next(rows)
            for row in rows: pass
            elapsed = time.perf_counter() - start
            print(f'elapsed {elapsed:.2f}')
        except Exception as e:
            print('failed', e)
            continue

    # Test
    print(control_row)
    for key, expected_value in (
        ('number', 1),
        ('decimal', 1.1),
        ('date', datetime.date(2000, 1, 1)),
        ('boolean', True),
        ('text', 'CONTROL ROW'),
    ):
        try:
            value = control_row[key]
        except KeyError:
            print(f'🔴 "{key}" missing')
            continue
        if type(expected_value) != type(value):
            print(f'🔴 "{key}" expected type "{type(expected_value)}" received type "{type(value)}"')
        elif expected_value != value:
            print(f'🔴 "{key}" expected value "{expected_value}" received "{value}"')
        else:
            print(f'🟢 "{key}"')
