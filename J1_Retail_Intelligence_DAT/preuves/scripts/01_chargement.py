"""Charge les fichiers .dat TPCx-BB dans une base DuckDB locale pour le profilage.

Usage : python3 -I 01_chargement.py <racine_retail_project> <fichier.duckdb>
Les fichiers sources ne sont pas modifiés (lecture seule).
"""
import os
import sys
import time

import duckdb

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from schema_tpcxbb import SCHEMA  # noqa: E402

racine, db_path = sys.argv[1], sys.argv[2]
con = duckdb.connect(db_path)

for lot, prefixe in (("data", "d_"), ("data_refresh", "r_")):
    for table, cols in SCHEMA.items():
        chemin = os.path.join(racine, lot, table, f"{table}.dat")
        if not os.path.exists(chemin):
            continue
        t0 = time.time()
        noms = "[" + ",".join(f"'{c}'" for c in cols) + "]"
        con.execute(f"DROP TABLE IF EXISTS {prefixe}{table}")
        con.execute(f"""
            CREATE TABLE {prefixe}{table} AS
            SELECT * FROM read_csv('{chemin}', delim='|', header=false, quote='', escape='',
                                   names={noms}, sample_size=-1, auto_detect=true)
        """)
        n = con.execute(f"SELECT count(*) FROM {prefixe}{table}").fetchone()[0]
        print(f"{lot:13s} {table:24s} {n:>12,d} lignes  ({time.time()-t0:.1f}s)")
con.close()
