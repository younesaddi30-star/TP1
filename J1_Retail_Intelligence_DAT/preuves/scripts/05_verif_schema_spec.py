"""Compare les noms de colonnes utilisés (schema_tpcxbb.py) au DDL Hive de la spécification TPCx-BB v1.7.0
(annexe, CREATE EXTERNAL TABLE), et les tailles des fichiers au tableau 2-1 (SF1).

Usage : python3 -I 05_verif_schema_spec.py <spec.txt (pdftotext -layout)> <racine_retail_project>
"""
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from schema_tpcxbb import SCHEMA  # noqa: E402

spec_txt, racine = sys.argv[1:3]
lines = open(spec_txt, encoding="utf-8").read().split("\n")

VAR = {"customerTableName": "customer", "customerAddressTableName": "customer_address",
       "customerDemographicsTableName": "customer_demographics", "dateTableName": "date_dim",
       "householdDemographicsTableName": "household_demographics", "incomeTableName": "income_band",
       "itemTableName": "item", "promotionTableName": "promotion", "reasonTableName": "reason",
       "shipModeTableName": "ship_mode", "storeTableName": "store", "timeTableName": "time_dim",
       "warehouseTableName": "warehouse", "webSiteTableName": "web_site", "webPageTableName": "web_page",
       "inventoryTableName": "inventory", "storeSalesTableName": "store_sales",
       "storeReturnsTableName": "store_returns", "webSalesTableName": "web_sales",
       "webReturnsTableName": "web_returns", "marketPricesTableName": "item_marketprices",
       "clickstreamsTableName": "web_clickstreams", "reviewsTableName": "product_reviews"}

ddl, cur = {}, None
for i, l in enumerate(lines):
    if "CREATE EXTERNAL TABLE" in l and "${hiveconf:" not in l:
        l = l + " " + lines[i + 1]  # nom de table renvoyé à la ligne suivante dans le PDF
    m = re.search(r"CREATE EXTERNAL TABLE\s+\$\{hiveconf:(\w+)\}", l)
    if m:
        cur = VAR.get(m.group(1))
        ddl[cur] = []
        continue
    if cur:
        if "ROW FORMAT" in l:
            cur = None
            continue
        m = re.match(r"\s*[(,]\s*([a-z_]+)\s+(bigint|string|int|decimal|double|date|timestamp)", l)
        if m:
            ddl[cur].append(m.group(1))

ok = True
print("== Colonnes (ordre et noms) : schéma utilisé vs DDL spécification v1.7.0")
for t, cols in SCHEMA.items():
    s = ddl.get(t)
    if s is None:
        print(f"{t:24s} ABSENTE de la spécification"); ok = False
    elif s == cols:
        print(f"{t:24s} IDENTIQUE ({len(cols)} colonnes)")
    else:
        ok = False
        diff = [(i + 1, a, b) for i, (a, b) in enumerate(zip(cols, s)) if a != b]
        print(f"{t:24s} DIFFÉRENT : {len(cols)} vs {len(s)} colonnes ; écarts {diff[:5]}")

# Tableau 2-1 (taille en octets à SF1) : 1re colonne numérique de la ligne de chaque table
print("\n== Tailles des fichiers data/ vs tableau 2-1 (SF1)")
sf1 = {}
for l in lines:
    m = re.match(r"^([a-z _]+?)\s{2,}(\d+)\s+\d+", l)
    if m:
        name = m.group(1).replace(" ", "")
        if name in SCHEMA:
            sf1[name] = int(m.group(2))
tot_a = tot_b = 0
for t in SCHEMA:
    a = os.path.getsize(os.path.join(racine, "data", t, f"{t}.dat"))
    b = sf1.get(t)
    tot_a += a
    tot_b += b or 0
    ecart = f"{(a - b) / b:+.2%}" if b else "n/a"
    print(f"{t:24s} fichier={a:>11,d}  spec SF1={b if b is None else format(b, ',d'):>11}  écart={ecart}")
print(f"{'TOTAL':24s} fichier={tot_a:>11,d}  spec SF1={tot_b:>11,d}  écart={(tot_a - tot_b) / tot_b:+.2%}")
print("\nRÉSULTAT :", "noms de colonnes conformes à la spécification" if ok else "écarts de schéma à examiner")
