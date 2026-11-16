"""Mesure le gain de stockage du format Parquet (Snappy) par rapport aux fichiers .dat (UC-02).

Usage : python3 -I 04_test_parquet.py <fichier.duckdb> <racine_retail_project> <dossier_sortie>
"""
import os
import sys

import duckdb

db, racine, out = sys.argv[1:4]
os.makedirs(out, exist_ok=True)
con = duckdb.connect(db, read_only=True)
tot_dat = tot_pq = 0
print(f"{'table':24s} {'.dat (Mo)':>10s} {'Parquet (Mo)':>13s} {'ratio':>7s}")
for t in ["store_sales", "web_sales", "inventory", "web_clickstreams", "product_reviews",
          "customer", "customer_demographics", "item", "date_dim"]:
    p = os.path.join(out, f"{t}.parquet")
    con.execute(f"COPY d_{t} TO '{p}' (FORMAT parquet, COMPRESSION snappy)")
    a = os.path.getsize(os.path.join(racine, "data", t, f"{t}.dat"))
    b = os.path.getsize(p)
    tot_dat, tot_pq = tot_dat + a, tot_pq + b
    print(f"{t:24s} {a/1048576:10.1f} {b/1048576:13.1f} {b/a:7.1%}")
print(f"{'TOTAL':24s} {tot_dat/1048576:10.1f} {tot_pq/1048576:13.1f} {tot_pq/tot_dat:7.1%}")
