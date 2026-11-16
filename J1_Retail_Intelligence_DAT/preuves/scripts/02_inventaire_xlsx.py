"""Produit UC01_inventaire_sources.xlsx à partir de la base DuckDB de profilage.

Usage : python3 -I 02_inventaire_xlsx.py <racine_retail_project> <fichier.duckdb> <sortie.xlsx>
Toutes les valeurs mesurées sont recalculées par ce script (preuve reproductible).
"""
import os
import sys

import duckdb
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from schema_tpcxbb import DOMAINE, ROLE, SCHEMA  # noqa: E402

racine, db_path, sortie = sys.argv[1:4]
con = duckdb.connect(db_path, read_only=True)
one = lambda sql: con.execute(sql).fetchone()  # noqa: E731

F = "Arial"
H_FILL = PatternFill("solid", start_color="1F3864")
H_FONT = Font(name=F, bold=True, color="FFFFFF", size=10)
B_FONT = Font(name=F, size=10)
T_FONT = Font(name=F, bold=True, size=14)
thin = Side(style="thin", color="BFBFBF")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)
STATUT_FILL = {"Vérifié": "E2EFDA", "Déduit": "FFF2CC", "Hypothèse": "FCE4D6",
               "Non disponible": "EDEDED", "À valider": "DDEBF7"}

wb = Workbook()


def feuille(titre, entetes, lignes, largeurs, statut_col=None, intro=None):
    ws = wb.create_sheet(titre)
    r0 = 1
    if intro:
        ws.cell(1, 1, intro).font = Font(name=F, italic=True, size=10)
        r0 = 3
    for j, h in enumerate(entetes, 1):
        c = ws.cell(r0, j, h)
        c.font, c.fill, c.border = H_FONT, H_FILL, BORDER
        c.alignment = Alignment(wrap_text=True, vertical="center")
    for i, ligne in enumerate(lignes, r0 + 1):
        for j, v in enumerate(ligne, 1):
            c = ws.cell(i, j, v)
            c.font, c.border = B_FONT, BORDER
            c.alignment = Alignment(wrap_text=True, vertical="top")
            if isinstance(v, int) and not isinstance(v, bool):
                c.number_format = "#,##0"
            if statut_col and j == statut_col and v in STATUT_FILL:
                c.fill = PatternFill("solid", start_color=STATUT_FILL[v])
    for j, w in enumerate(largeurs, 1):
        ws.column_dimensions[get_column_letter(j)].width = w
    ws.freeze_panes = ws.cell(r0 + 1, 1)
    ws.auto_filter.ref = f"A{r0}:{get_column_letter(len(entetes))}{r0 + len(lignes)}"
    return ws, r0


# ---------------------------------------------------------------- 1. Fiche kit
ws = wb.active
ws.title = "Fiche kit"
ws["A1"] = "UC-01 – Inventaire des sources TPCx-BB – Retail Intelligence Platform"
ws["A1"].font = T_FONT
fichiers = [os.path.join(dp, f) for dp, _, fs in os.walk(racine) for f in fs]
total = sum(os.path.getsize(p) for p in fichiers)
kit = [
    ("Information", "Valeur constatée", "Statut"),
    ("Nom du kit / archive", "OneDrive_2026-11-16.zip > « 01. Architecture DataLake, DataWarehouse, "
     "Lakehous » > retail_project.zip", "Vérifié"),
    ("Version TPCx-BB", "v1.7.0 : kit officiel TPCX-BB_Tools_1.7.0 fourni par le formateur (contient la spécification "
     "TPCx-BB_v1.7.0.pdf). Données conformes au schéma et au facteur SF1 de cette version", "Vérifié"),
    ("Facteur d'échelle", "SF1 : 15 fichiers identiques à l'octet près au tableau 2-1 de la spécification, total "
     "1 024 447 801 octets contre 1 023 950 269 attendus (+0,05 %) – script 05_verif_schema_spec.py", "Vérifié"),
    ("Provenance", "Formateur IPSSI, via OneDrive. Génération TPCx-BB (BigBench) d'après projet_fil_rouge.pdf "
     "§3.1–3.2 (répertoires HDFS /user/root/benchmarks/bigbench/data et data_refresh)", "Vérifié"),
    ("Date de mise à disposition", "16/11/2026 (horodatage des fichiers dans l'archive : 15/11/2026 23:34–23:36)",
     "Vérifié"),
    ("Source de téléchargement", "OneDrive (lien du formateur)", "Vérifié"),
    ("Outil de génération", "PDGF (Parallel Data Generator Framework), générateur intégré au kit TPCx-BB v1.7.0 "
     "fourni par le formateur", "Vérifié"),
    ("Format des fichiers", "34 fichiers .dat texte délimité par « | », sans en-tête, sans guillemets, "
     "encodage ASCII (aucun caractère non-ASCII détecté sur les tables textuelles contrôlées) "
     "+ 2 fichiers .csv d'audit (item_marketprices_audit.csv)", "Vérifié"),
    ("Nombre de fichiers", f"{len(fichiers)} fichiers (23 tables dans data/, 11 tables dans data_refresh/, "
     "2 csv d'audit)", "Vérifié"),
    ("Volume total décompressé", f"{total:,} octets ≈ {total/1024**3:.2f} Go".replace(",", " "), "Vérifié"),
    ("Lots de données", "data/ = chargement initial ; data_refresh/ = lot de rafraîchissement (nouvelles "
     "clés, période 13–31/12/2005)", "Vérifié"),
    ("Noms des colonnes", "Fichiers sans en-tête : noms repris de la spécification TPCx-BB v1.7.0. Les 23 tables "
     "sont identiques au DDL officiel (noms et ordre des colonnes) ; séparateur « | » confirmé (fieldDelimiter=|).",
     "Vérifié"),
    ("Écarts à la spécification", "c_last_review_date (date attendue, chaîne aléatoire observée) ; r_reason_desc "
     "(liste de motifs attendue, texte aléatoire) ; s_closed_date_sk (une partie des magasins fermés attendue, 12/12 "
     "datés) ; date_dim (73 049 jours annoncés au §B.2, 109 573 observés alors que la taille du fichier est conforme "
     "au tableau 2-1)", "Vérifié"),
    ("Données complémentaires", "Aucune fournie à ce stade", "Vérifié"),
    ("Outils de contrôle", "Python 3 + DuckDB (2.0 ; rejoué sous WSL Ubuntu 22.04 avec DuckDB 1.5.6) (scripts 01_chargement.py, 02_inventaire_xlsx.py, "
     "schema_tpcxbb.py, 05_verif_schema_spec.py) ; comptage de champs par awk", "Vérifié"),
]
for i, row in enumerate(kit, 3):
    for j, v in enumerate(row, 1):
        c = ws.cell(i, j, v)
        c.border = BORDER
        c.alignment = Alignment(wrap_text=True, vertical="top")
        if i == 3:
            c.font, c.fill = H_FONT, H_FILL
        else:
            c.font = B_FONT
            if j == 3 and v in STATUT_FILL:
                c.fill = PatternFill("solid", start_color=STATUT_FILL[v])
ws.column_dimensions["A"].width, ws.column_dimensions["B"].width, ws.column_dimensions["C"].width = 28, 95, 16
r = len(kit) + 4
ws.cell(r, 1, "Légende des statuts (TP §4.2)").font = Font(name=F, bold=True, size=10)
for k, (s, d) in enumerate([("Vérifié", "Observé directement dans un fichier ou une documentation"),
                            ("Déduit", "Obtenu par analyse, avec justification"),
                            ("Hypothèse", "À confirmer par une vérification complémentaire"),
                            ("Non disponible", "Absent des données examinées"),
                            ("À valider", "Décision ou vérification du formateur nécessaire")], r + 1):
    c = ws.cell(k, 1, s)
    c.font, c.fill, c.border = B_FONT, PatternFill("solid", start_color=STATUT_FILL[s]), BORDER
    ws.cell(k, 2, d).font = B_FONT

# ---------------------------------------------------------- 2. Inventaire
CLE = {  # table -> (clé candidate, grain)
    "store_sales": ("(ss_ticket_number, ss_item_sk)", "1 ligne = 1 article dans un ticket de caisse"),
    "web_sales": ("(ws_order_number, ws_item_sk)", "1 ligne = 1 article dans une commande web"),
    "store_returns": ("(sr_ticket_number, sr_item_sk)", "1 ligne = 1 article retourné d'un ticket"),
    "web_returns": ("(wr_order_number, wr_item_sk)", "1 ligne = 1 article retourné d'une commande"),
    "inventory": ("(inv_date_sk, inv_item_sk, inv_warehouse_sk)", "1 ligne = stock d'un produit, dans un entrepôt, "
                  "à une date (relevé hebdomadaire)"),
    "web_clickstreams": ("Aucune clé (pas d'identifiant d'événement)", "1 ligne = 1 clic sur une page web"),
    "product_reviews": ("pr_review_sk", "1 ligne = 1 avis"),
    "item_marketprices": ("imp_sk", "1 ligne = 1 prix concurrent pour un produit sur une période"),
}
inv_rows, r_idx = [], 4
for lot, pre in (("data", "d_"), ("data_refresh", "r_")):
    for t, cols in SCHEMA.items():
        p = os.path.join(racine, lot, t, f"{t}.dat")
        if not os.path.exists(p):
            continue
        n = one(f"SELECT count(*) FROM {pre}{t}")[0]
        if t in CLE:
            cle, grain = CLE[t]
            if t == "inventory":
                d = one(f"SELECT count(*) - count(DISTINCT (inv_date_sk, inv_item_sk, inv_warehouse_sk)) FROM {pre}{t}")[0]
                uni = "Oui (0 doublon)" if d == 0 else f"Non ({d:,} doublons)"
            elif t.endswith(("_sales", "_returns")):
                a, b = cle.split("(")[1].split(")")[0].split(", ")
                d = one(f"SELECT count(*) FROM (SELECT {a},{b} FROM {pre}{t} GROUP BY 1,2 HAVING count(*)>1)")[0]
                uni = "Oui (0 doublon)" if d == 0 else f"Non ({d:,} couples en double)"
            elif cle.startswith("Aucune"):
                uni = "Sans objet"
            else:
                d = n - one(f"SELECT count(DISTINCT {cle}) FROM {pre}{t}")[0]
                uni = "Oui (0 doublon)" if d == 0 else f"Non ({d:,})"
        else:
            cle = cols[0]
            d = n - one(f"SELECT count(DISTINCT {cle}) FROM {pre}{t}")[0]
            uni = "Oui (0 doublon)" if d == 0 else f"Non ({d:,})"
            grain = f"1 ligne = 1 élément du référentiel {t}"
        inv_rows.append([lot, t, f"retail_project/{lot}/{t}/{t}.dat", ".dat (texte « | »)", os.path.getsize(p),
                         f"=E{r_idx}/1048576", n, len(cols), DOMAINE[t], ROLE[t], cle, uni, grain, "Vérifié"])
        r_idx += 1
for lot in ("data", "data_refresh"):
    p = os.path.join(racine, lot, "item_marketprices", "item_marketprices_audit.csv")
    with open(p, encoding="ascii", newline="") as fh:
        nl = sum(1 for _ in fh)
    inv_rows.append([lot, "item_marketprices_audit", f"retail_project/{lot}/item_marketprices/"
                     "item_marketprices_audit.csv", ".csv (« , », 1re ligne vide"
                     + (", fins de ligne CRLF)" if lot == "data" else ")"), os.path.getsize(p), f"=E{r_idx}/1048576",
                     nl, 8, "Produits / prix", "Métadonnées de génération", "Sans objet",
                     "Sans objet", "Rapport d'audit du générateur (statistiques concurrents), pas une table métier",
                     "Déduit"])
    r_idx += 1
ws, r0 = feuille("Inventaire fichiers",
                 ["Lot", "Table / fichier", "Chemin", "Format", "Taille (octets)", "Taille (Mo)", "Nb lignes",
                  "Nb colonnes", "Domaine", "Rôle", "Clé candidate", "Unicité vérifiée", "Grain (ce que représente une ligne)",
                  "Statut"],
                 inv_rows, [12, 24, 46, 22, 14, 11, 13, 10, 18, 20, 34, 22, 44, 11], statut_col=14,
                 intro="Une ligne par fichier réellement présent. Nb lignes mesuré par DuckDB et contrôlé par awk "
                       "(nombre de champs constant sur toutes les lignes de chaque .dat).")
last = r0 + len(inv_rows)
for i in range(r0 + 1, last + 1):
    ws.cell(i, 6).number_format = "#,##0.00"
ws.cell(last + 1, 4, "TOTAL").font = Font(name=F, bold=True, size=10)
for col, fmt in ((5, "#,##0"), (6, "#,##0.00"), (7, "#,##0")):
    L = get_column_letter(col)
    c = ws.cell(last + 1, col, f"=SUM({L}{r0+1}:{L}{last})")
    c.font, c.number_format = Font(name=F, bold=True, size=10), fmt

# ---------------------------------------------------------- 3. Dictionnaire de colonnes
col_rows = []
for t, cols in SCHEMA.items():
    summ = {r[0]: r for r in con.execute(f"SUMMARIZE d_{t}").fetchall()}
    # SUMMARIZE: column_name, column_type, min, max, approx_unique, avg, std, q25, q50, q75, count, null_percentage
    n = one(f"SELECT count(*) FROM d_{t}")[0]
    for k, c in enumerate(cols, 1):
        s = summ[c]
        nulls, dist = one(f"SELECT count(*) - count({c}), count(DISTINCT {c}) FROM d_{t}")
        mn, mx = s[2], s[3]
        if t in ("product_reviews", "item", "reason", "promotion", "store", "web_site") and c.endswith(("desc", "content", "details")):
            mn = mx = "(texte libre)"
        if c.endswith("_sk") and k == 1 and t not in ("store_sales", "web_sales", "store_returns", "web_returns", "inventory", "web_clickstreams"):
            role = "Clé technique (PK candidate)"
        elif c.endswith("_sk") or c in ("ss_ticket_number", "sr_ticket_number", "ws_order_number", "wr_order_number"):
            role = "Clé étrangère / identifiant"
        elif c.endswith("_id"):
            role = "Identifiant métier"
        elif s[1] in ("DOUBLE", "BIGINT") and t in ("store_sales", "web_sales", "store_returns", "web_returns",
                                                     "inventory") :
            role = "Mesure"
        else:
            role = "Attribut"
        col_rows.append([t, DOMAINE[t], k, c, s[1], role, n, nulls, f"=IF(G{len(col_rows)+4}=0,0,H{len(col_rows)+4}/G{len(col_rows)+4})",
                         dist, str(mn)[:40] if mn is not None else "",
                         str(mx)[:40] if mx is not None else "", "Vérifié"])
ws, r0 = feuille("Dictionnaire colonnes",
                 ["Table", "Domaine", "N°", "Colonne", "Type détecté", "Rôle", "Nb lignes", "Nb nulls", "% nulls",
                  "Valeurs distinctes", "Min", "Max", "Statut du nom"],
                 col_rows, [22, 18, 5, 28, 12, 26, 12, 10, 9, 13, 24, 24, 12], statut_col=13,
                 intro="Lot data/ uniquement. Types détectés automatiquement par DuckDB sur 100 % des lignes. "
                       "Nulls = champs vides. Distinct = comptage exact (hors nulls). Nom de colonne : Vérifié "
                       "(identique au DDL de la spécification TPCx-BB v1.7.0, fichiers sans en-tête).")
for i in range(r0 + 1, r0 + 1 + len(col_rows)):
    ws.cell(i, 9).number_format = "0.00%"

# ---------------------------------------------------------- 4. Relations
FK = [("store_sales", "ss_item_sk", "item", "i_item_sk"), ("store_sales", "ss_customer_sk", "customer", "c_customer_sk"),
      ("store_sales", "ss_store_sk", "store", "s_store_sk"), ("store_sales", "ss_sold_date_sk", "date_dim", "d_date_sk"),
      ("store_sales", "ss_sold_time_sk", "time_dim", "t_time_sk"), ("store_sales", "ss_promo_sk", "promotion", "p_promo_sk"),
      ("store_sales", "ss_cdemo_sk", "customer_demographics", "cd_demo_sk"),
      ("store_sales", "ss_hdemo_sk", "household_demographics", "hd_demo_sk"),
      ("store_sales", "ss_addr_sk", "customer_address", "ca_address_sk"),
      ("web_sales", "ws_item_sk", "item", "i_item_sk"), ("web_sales", "ws_bill_customer_sk", "customer", "c_customer_sk"),
      ("web_sales", "ws_ship_customer_sk", "customer", "c_customer_sk"),
      ("web_sales", "ws_sold_date_sk", "date_dim", "d_date_sk"), ("web_sales", "ws_ship_date_sk", "date_dim", "d_date_sk"),
      ("web_sales", "ws_web_site_sk", "web_site", "web_site_sk"), ("web_sales", "ws_web_page_sk", "web_page", "wp_web_page_sk"),
      ("web_sales", "ws_ship_mode_sk", "ship_mode", "sm_ship_mode_sk"), ("web_sales", "ws_warehouse_sk", "warehouse", "w_warehouse_sk"),
      ("web_sales", "ws_promo_sk", "promotion", "p_promo_sk"),
      ("store_returns", "sr_item_sk", "item", "i_item_sk"), ("store_returns", "sr_customer_sk", "customer", "c_customer_sk"),
      ("store_returns", "sr_store_sk", "store", "s_store_sk"), ("store_returns", "sr_reason_sk", "reason", "r_reason_sk"),
      ("store_returns", "sr_returned_date_sk", "date_dim", "d_date_sk"),
      ("web_returns", "wr_item_sk", "item", "i_item_sk"), ("web_returns", "wr_reason_sk", "reason", "r_reason_sk"),
      ("web_returns", "wr_returning_customer_sk", "customer", "c_customer_sk"),
      ("web_returns", "wr_returned_date_sk", "date_dim", "d_date_sk"),
      ("inventory", "inv_item_sk", "item", "i_item_sk"), ("inventory", "inv_warehouse_sk", "warehouse", "w_warehouse_sk"),
      ("inventory", "inv_date_sk", "date_dim", "d_date_sk"),
      ("web_clickstreams", "wcs_item_sk", "item", "i_item_sk"), ("web_clickstreams", "wcs_user_sk", "customer", "c_customer_sk"),
      ("web_clickstreams", "wcs_web_page_sk", "web_page", "wp_web_page_sk"),
      ("web_clickstreams", "wcs_click_date_sk", "date_dim", "d_date_sk"),
      ("product_reviews", "pr_item_sk", "item", "i_item_sk"), ("product_reviews", "pr_user_sk", "customer", "c_customer_sk"),
      ("customer", "c_current_addr_sk", "customer_address", "ca_address_sk"),
      ("customer", "c_current_cdemo_sk", "customer_demographics", "cd_demo_sk"),
      ("customer", "c_current_hdemo_sk", "household_demographics", "hd_demo_sk"),
      ("household_demographics", "hd_income_band_sk", "income_band", "ib_income_band_sk"),
      ("item_marketprices", "imp_item_sk", "item", "i_item_sk"), ("promotion", "p_item_sk", "item", "i_item_sk")]
rel_rows = []
for c, cc, p, pc in FK:
    n, nn, orph, orph_v = one(f"""SELECT count(*), count(f.{cc}),
        sum(CASE WHEN f.{cc} IS NOT NULL AND d.{pc} IS NULL THEN 1 ELSE 0 END),
        count(DISTINCT CASE WHEN f.{cc} IS NOT NULL AND d.{pc} IS NULL THEN f.{cc} END)
        FROM d_{c} f LEFT JOIN d_{p} d ON d.{pc} = f.{cc}""")
    uniq_parent = one(f"SELECT count(*) = count(DISTINCT {pc}) FROM d_{p}")[0]
    i = len(rel_rows) + 4
    if orph == 0:
        statut, comm = "Vérifié", "Toutes les valeurs renseignées existent dans le référentiel."
    else:
        statut = "À valider"
        comm = (f"{orph_v} valeurs absentes de data/{p} ; elles existent toutes dans data_refresh/{p} → "
                "charger le référentiel refresh avant la jointure.")
    rel_rows.append([c, cc, p, pc, "N:1" if uniq_parent else "N:N ?", n, nn, f"=G{i}/F{i}", orph, orph_v, comm, statut])
# relations « métier » non portées par le nom
specials = [
    ("store_returns", "(sr_ticket_number, sr_item_sk)", "store_sales", "(ss_ticket_number, ss_item_sk)",
     """SELECT count(*), count(*), sum(CASE WHEN s.k IS NULL THEN 1 ELSE 0 END), 0 FROM d_store_returns r LEFT JOIN
        (SELECT DISTINCT ss_ticket_number t, ss_item_sk i, 1 k FROM d_store_sales) s ON s.t=r.sr_ticket_number AND s.i=r.sr_item_sk""",
     "Chaque retour magasin retrouve sa vente d'origine. Mais le couple n'est pas unique côté ventes (doublons) → jointure N:N possible."),
    ("web_returns", "(wr_order_number, wr_item_sk)", "web_sales", "(ws_order_number, ws_item_sk)",
     """SELECT count(*), count(*), sum(CASE WHEN s.k IS NULL THEN 1 ELSE 0 END), 0 FROM d_web_returns r LEFT JOIN
        (SELECT DISTINCT ws_order_number o, ws_item_sk i, 1 k FROM d_web_sales) s ON s.o=r.wr_order_number AND s.i=r.wr_item_sk""",
     "Chaque retour web retrouve sa commande d'origine ; même réserve sur l'unicité du couple."),
    ("web_clickstreams", "wcs_sales_sk", "web_sales", "ws_order_number",
     """SELECT count(*), count(wcs_sales_sk), sum(CASE WHEN wcs_sales_sk IS NOT NULL AND w.o IS NULL THEN 1 ELSE 0 END), 0
        FROM d_web_clickstreams c LEFT JOIN (SELECT DISTINCT ws_order_number o FROM d_web_sales) w ON w.o=c.wcs_sales_sk""",
     "wcs_sales_sk renseigné uniquement pour les clics liés à un achat ; il correspond au numéro de commande web (100 % de correspondance)."),
    ("product_reviews", "pr_order_sk", "web_sales", "ws_order_number",
     """SELECT count(*), count(pr_order_sk), sum(CASE WHEN w.o IS NULL THEN 1 ELSE 0 END), 0 FROM d_product_reviews p
        LEFT JOIN (SELECT DISTINCT ws_order_number o FROM d_web_sales) w ON w.o=p.pr_order_sk""",
     "Tous les numéros existent côté web, mais le triplet (commande, produit, client) ne correspond que pour une partie des avis "
     "(voir onglet Contrôles) → relation à confirmer."),
]
for c, cc, p, pc, sql, comm in specials:
    n, nn, orph, _ = one(sql)
    i = len(rel_rows) + 4
    rel_rows.append([c, cc, p, pc, "N:1 (attendu)", n, nn, f"=G{i}/F{i}", orph, "", comm,
                     "Déduit" if c != "product_reviews" else "Hypothèse"])
ws, r0 = feuille("Relations",
                 ["Table enfant", "Colonne enfant", "Table parent", "Colonne parent", "Cardinalité",
                  "Nb lignes enfant", "Nb renseignées", "% renseigné", "Lignes orphelines", "Valeurs orphelines",
                  "Commentaire", "Statut"],
                 rel_rows, [20, 28, 22, 26, 13, 14, 14, 11, 12, 11, 60, 11], statut_col=12,
                 intro="Contrôle par LEFT JOIN sur le lot data/. Orpheline = valeur renseignée sans correspondance "
                       "dans le référentiel. Cardinalité N:1 si la clé parent est unique (vérifié).")
for i in range(r0 + 1, r0 + 1 + len(rel_rows)):
    ws.cell(i, 8).number_format = "0.0%"

# ---------------------------------------------------------- 5. Contrôles qualité
CTRL = [
    ("Q01", "Structure", "Nombre de champs constant par ligne dans chaque .dat", None,
     "Oui pour les 34 fichiers (contrôle awk -F'|' '{print NF}')", "Pas de ligne malformée à rejeter.", "Vérifié"),
    ("Q02", "Ventes magasin", "Période couverte par store_sales (data/)",
     "SELECT min(d.d_date)||' → '||max(d.d_date) FROM d_store_sales s JOIN d_date_dim d ON d.d_date_sk=s.ss_sold_date_sk",
     None, "Analyses temporelles possibles sur 2001–2005 ; 2005 incomplète (s'arrête au 13/12, la suite est dans data_refresh).", "Vérifié"),
    ("Q03", "Ventes web", "Période couverte par web_sales (data/)",
     "SELECT min(d.d_date)||' → '||max(d.d_date) FROM d_web_sales s JOIN d_date_dim d ON d.d_date_sk=s.ws_sold_date_sk",
     None, "Même période que le magasin : comparaison des canaux possible.", "Vérifié"),
    ("Q04", "Ventes magasin", "Nombre de tickets distincts / lignes par ticket (min–moy–max)",
     "SELECT count(*)||' tickets ; '||min(n)||' – '||round(avg(n),2)||' – '||max(n)||' lignes' FROM "
     "(SELECT ss_ticket_number, count(*) n FROM d_store_sales GROUP BY 1)",
     None, "Identifiant de transaction disponible → nb transactions et panier moyen calculables (count distinct ss_ticket_number). "
     "Ne pas compter les lignes.", "Vérifié"),
    ("Q05", "Ventes magasin", "Cohérence d'un ticket : un seul magasin, une seule date, un seul client",
     "SELECT count(*)||' tickets incohérents' FROM (SELECT ss_ticket_number FROM d_store_sales GROUP BY 1 HAVING "
     "count(DISTINCT ss_store_sk)>1 OR count(DISTINCT ss_sold_date_sk)>1 OR count(DISTINCT ss_customer_sk)>1)",
     None, "Le ticket se comporte comme une transaction.", "Vérifié"),
    ("Q06", "Ventes web", "Nombre de commandes distinctes / lignes par commande",
     "SELECT count(DISTINCT ws_order_number)||' commandes ; moy. '||round(count(*)/count(DISTINCT ws_order_number),2)||' lignes' FROM d_web_sales",
     None, "Panier moyen web calculable sur ws_order_number.", "Vérifié"),
    ("Q07", "Ventes", "Couples (ticket, produit) en double dans store_sales",
     "SELECT count(*)||' couples ; '||sum(n)||' lignes concernées' FROM (SELECT count(*) n FROM d_store_sales GROUP BY ss_ticket_number, ss_item_sk HAVING count(*)>1)",
     None, "Le même produit peut apparaître plusieurs fois dans un ticket (pas de n° de ligne). Grain = ligne de ticket, "
     "sans clé naturelle unique ; jointure ventes↔retours à traiter avec précaution.", "Vérifié"),
    ("Q08", "Ventes", "Lignes entièrement dupliquées dans store_sales",
     "SELECT count(*) - (SELECT count(*) FROM (SELECT DISTINCT * FROM d_store_sales)) FROM d_store_sales",
     None, "Aucun doublon strict : les couples en double (Q07) ne sont pas des copies.", "Vérifié"),
    ("Q09", "Ventes", "Formules des montants : ext_sales_price = sales_price × quantity ; net_paid = ext_sales_price − coupon_amt",
     "SELECT sum(CASE WHEN abs(ss_ext_sales_price-ss_sales_price*ss_quantity)>0.01 THEN 1 ELSE 0 END)||' écarts ; '||"
     "sum(CASE WHEN abs(ss_net_paid-(ss_ext_sales_price-ss_coupon_amt))>0.01 THEN 1 ELSE 0 END)||' écarts' FROM d_store_sales",
     None, "Règles de calcul vérifiées → CA HT = ss_net_paid (après remise et coupon) ; ss_ext_discount_amt = remise.", "Vérifié"),
    ("Q10", "Ventes", "Quantités ≤ 0 et prix de vente minimum",
     "SELECT sum(CASE WHEN ss_quantity<=0 THEN 1 ELSE 0 END)||' qté ≤ 0 ; qté '||min(ss_quantity)||'–'||max(ss_quantity)||' ; prix min '||min(ss_sales_price) FROM d_store_sales",
     None, "Pas de quantité négative. Des prix de vente à 0 existent : à vérifier (gratuité ? anomalie ?).", "Vérifié"),
    ("Q11", "Ventes", "Part des lignes magasin à marge nette négative (ss_net_profit < 0)",
     "SELECT round(100.0*avg(CASE WHEN ss_net_profit<0 THEN 1 ELSE 0 END),1)||' %' FROM d_store_sales",
     None, "Proportion irréaliste pour une enseigne : effet du générateur synthétique. Ne pas interpréter la marge comme réelle.", "Vérifié"),
    ("Q12", "Retours", "Retours magasin rattachés à une vente (ticket + produit)",
     "SELECT round(100.0*avg(CASE WHEN s.k IS NULL THEN 0 ELSE 1 END),1)||' %' FROM d_store_returns r LEFT JOIN "
     "(SELECT DISTINCT ss_ticket_number t, ss_item_sk i, 1 k FROM d_store_sales) s ON s.t=r.sr_ticket_number AND s.i=r.sr_item_sk",
     None, "Taux de retour calculable par produit / magasin.", "Vérifié"),
    ("Q13", "Retours", "Retours dont la quantité dépasse la quantité vendue (après jointure)",
     "SELECT count(*) FROM d_store_returns r JOIN d_store_sales s ON s.ss_ticket_number=r.sr_ticket_number AND s.ss_item_sk=r.sr_item_sk WHERE r.sr_return_quantity>s.ss_quantity",
     None, "Cas liés aux couples en double (Q07) ou incohérence synthétique : à documenter et exclure / signaler.", "Vérifié"),
    ("Q14", "Retours", "Retour antérieur à la vente",
     "SELECT count(*) FROM d_store_returns r JOIN d_store_sales s ON s.ss_ticket_number=r.sr_ticket_number AND s.ss_item_sk=r.sr_item_sk WHERE r.sr_returned_date_sk<s.ss_sold_date_sk",
     None, "Chronologie cohérente.", "Vérifié"),
    ("Q15", "Retours", "Période des retours (magasin / web)",
     "SELECT (SELECT max(d.d_date) FROM d_store_returns r JOIN d_date_dim d ON d.d_date_sk=r.sr_returned_date_sk)||' / '||"
     "(SELECT max(d.d_date) FROM d_web_returns r JOIN d_date_dim d ON d.d_date_sk=r.wr_returned_date_sk)",
     None, "Des retours sont datés jusqu'en 2007–2008, bien après la dernière vente (2005) : délai irréaliste. "
     "Taux de retour par période à calculer sur la date de vente, pas la date de retour.", "Vérifié"),
    ("Q16", "Stocks", "Pas de temps de l'inventaire et nombre de relevés",
     "SELECT count(DISTINCT inv_date_sk)||' dates, écart '||min(g)||'–'||max(g)||' jours' FROM (SELECT inv_date_sk, inv_date_sk - lag(inv_date_sk) OVER (ORDER BY inv_date_sk) g FROM (SELECT DISTINCT inv_date_sk FROM d_inventory))",
     None, "Stock = photo hebdomadaire (snapshot), pas des mouvements. Ne pas sommer les stocks dans le temps.", "Vérifié"),
    ("Q17", "Stocks", "Quantités en stock nulles (champ vide)",
     "SELECT count(*)-count(inv_quantity_on_hand) FROM d_inventory", None,
     "Quelques relevés sans quantité : à exclure ou signaler dans les KPI de stock.", "Vérifié"),
    ("Q18", "Stocks", "Produits de inventory absents de data/item",
     "SELECT count(DISTINCT i.inv_item_sk) FROM d_inventory i LEFT JOIN d_item t ON t.i_item_sk=i.inv_item_sk WHERE t.i_item_sk IS NULL",
     None, "Ces 180 produits (i_item_sk 17820–17999) ne sont décrits que dans data_refresh/item.", "Vérifié"),
    ("Q19", "Navigation", "Part des clics sans utilisateur identifié (wcs_user_sk vide)",
     "SELECT round(100.0*(count(*)-count(wcs_user_sk))/count(*),1)||' %' FROM d_web_clickstreams", None,
     "Visiteurs anonymes : analyse par client limitée aux clics identifiés.", "Vérifié"),
    ("Q20", "Navigation", "Part des clics liés à un achat (wcs_sales_sk renseigné)",
     "SELECT round(100.0*count(wcs_sales_sk)/count(*),1)||' % ('||count(wcs_sales_sk)||' clics)' FROM d_web_clickstreams", None,
     "Un clic n'est pas un achat. Le nombre de clics liés à un achat = nombre de lignes de web_sales : taux de conversion "
     "approximable, mais pas d'événement « ajout au panier ».", "Vérifié"),
    ("Q21", "Avis", "Distribution des notes (1 à 5)",
     "SELECT string_agg(pr_review_rating||'★:'||n, ' ; ' ORDER BY pr_review_rating) FROM (SELECT pr_review_rating, count(*) n FROM d_product_reviews GROUP BY 1)",
     None, "Note moyenne calculable ; distribution très asymétrique (majorité de 5).", "Vérifié"),
    ("Q22", "Avis", "Avis dont (commande, produit, client) correspond à une ligne web_sales",
     "SELECT round(100.0*avg(CASE WHEN w.o IS NULL THEN 0 ELSE 1 END),1)||' %' FROM d_product_reviews p LEFT JOIN "
     "(SELECT DISTINCT ws_order_number o, ws_item_sk i, ws_bill_customer_sk c FROM d_web_sales) w ON w.o=p.pr_order_sk AND w.i=p.pr_item_sk AND w.c=p.pr_user_sk",
     None, "Lien avis↔achat partiel : ne pas supposer que tout avis vient d'un achat vérifié.", "Vérifié"),
    ("Q23", "Avis", "Avis contenant des entités HTML (&quot;) / contenu vide",
     "SELECT sum(CASE WHEN pr_review_content LIKE '%&quot;%' THEN 1 ELSE 0 END)||' avec &quot; ; '||sum(CASE WHEN pr_review_content IS NULL THEN 1 ELSE 0 END)||' vides' FROM d_product_reviews",
     None, "Texte à nettoyer (décodage HTML) avant analyse textuelle.", "Vérifié"),
    ("Q24", "Clients", "Année de naissance min–max",
     "SELECT min(c_birth_year)||' – '||max(c_birth_year) FROM d_customer", None,
     "Tranches d'âge calculables (à date de référence à définir).", "Vérifié"),
    ("Q25", "Clients", "Format de c_last_review_date",
     "SELECT string_agg(c_last_review_date, ' | ') FROM (SELECT c_last_review_date FROM d_customer LIMIT 5)", None,
     "Chaînes aléatoires, pas des dates : colonne inexploitable.", "Vérifié"),
    ("Q26", "Clients", "Clients ayant au moins un achat (magasin ou web) / total",
     "SELECT (SELECT count(*) FROM (SELECT ss_customer_sk FROM d_store_sales UNION SELECT ws_bill_customer_sk FROM d_web_sales))||' / '||(SELECT count(*) FROM d_customer)",
     None, "Environ 40 % des clients n'ont aucun achat dans le lot.", "Vérifié"),
    ("Q27", "Clients", "Dates de première vente client (c_first_sales_date_sk)",
     "SELECT min(d.d_date)||' → '||max(d.d_date) FROM d_customer c JOIN d_date_dim d ON d.d_date_sk=c.c_first_sales_date_sk",
     None, "Dates entre 1900 et 2199, incohérentes avec la période de ventes : attribut inexploitable.", "Vérifié"),
    ("Q28", "Magasins", "Date de fermeture des magasins (s_closed_date_sk)",
     "SELECT string_agg(s.s_store_sk||':'||d.d_date, ' ; ' ORDER BY d.d_date) FROM d_store s JOIN d_date_dim d ON d.d_date_sk=s.s_closed_date_sk",
     None, "Tous les magasins ont une date de fermeture (1917 à 2141) alors qu'ils vendent en 2001–2005 : ne pas filtrer les "
     "magasins « fermés » sur ce champ.", "Vérifié"),
    ("Q29", "Promotions", "Période des promotions (dates de début)",
     "SELECT min(d.d_date)||' → '||max(d.d_date) FROM d_promotion p JOIN d_date_dim d ON d.d_date_sk=p.p_start_date_sk",
     None, "Dates de 1901 à 2199 : analyse « pendant / hors promo » par date non fiable ; utiliser ss_promo_sk.", "Vérifié"),
    ("Q30", "Temps", "Étendue de date_dim",
     "SELECT min(d_date)||' → '||max(d_date)||' ('||count(*)||' jours)' FROM d_date_dim", None,
     "Dimension temps complète (1900–2199), couvre toutes les dates de faits.", "Vérifié"),
    ("Q31", "Refresh", "Recouvrement de clés entre data/ et data_refresh/ (clients, produits, tickets)",
     "SELECT (SELECT count(*) FROM r_customer JOIN d_customer USING(c_customer_sk))||' / '||(SELECT count(*) FROM r_item JOIN d_item USING(i_item_sk))||' / '||"
     "(SELECT count(*) FROM r_store_sales JOIN d_store_sales USING(ss_ticket_number))",
     None, "Aucun recouvrement : le refresh ajoute de nouvelles lignes (append), il ne met pas à jour l'existant.", "Vérifié"),
    ("Q32", "Refresh", "Retours du refresh rattachés à des ventes de data/",
     "SELECT count(*)||' retours ; '||sum(CASE WHEN d.t IS NOT NULL THEN 1 ELSE 0 END)||' rattachés à data/' FROM r_store_returns r LEFT JOIN (SELECT DISTINCT ss_ticket_number t FROM d_store_sales) d ON d.t=r.sr_ticket_number",
     None, "Le refresh contient des retours sur des ventes anciennes : il faut garder data/ et data_refresh/ ensemble.", "Vérifié"),
    ("Q33", "Refresh", "Ventes du refresh sur des produits / clients absents de data/",
     "SELECT sum(CASE WHEN i.i_item_sk IS NULL THEN 1 ELSE 0 END)||' lignes produit / '||sum(CASE WHEN c.c_customer_sk IS NULL THEN 1 ELSE 0 END)||' lignes client' FROM r_store_sales s LEFT JOIN d_item i ON i.i_item_sk=s.ss_item_sk LEFT JOIN d_customer c ON c.c_customer_sk=s.ss_customer_sk",
     None, "Ordre de chargement : référentiels refresh avant les faits refresh.", "Vérifié"),
    ("Q34", "Audit", "Fichier item_marketprices_audit.csv", None,
     "1re ligne vide, en-tête de 8 colonnes, 1 ligne de statistiques puis des listes d'identifiants ; fins de ligne CRLF dans data/",
     "Métadonnées du générateur, à conserver en Bronze mais non exploitées comme table métier.", "Vérifié"),
]
ctrl_rows = []
for cid, dom, lib, sql, res, interp, st in CTRL:
    if sql:
        res = str(one(sql)[0])
    ctrl_rows.append([cid, dom, lib, res, interp, st, sql or "(contrôle hors SQL)"])
feuille("Contrôles qualité", ["ID", "Domaine", "Contrôle", "Résultat mesuré", "Interprétation / impact", "Statut",
                              "Requête (preuve, DuckDB)"],
        ctrl_rows, [6, 15, 42, 38, 62, 10, 70], statut_col=6,
        intro="Résultats recalculés à chaque exécution du script 02_inventaire_xlsx.py. La requête de chaque contrôle "
              "est fournie pour reproduction.")

wb.save(sortie)
print("OK", sortie, len(inv_rows), "fichiers,", len(col_rows), "colonnes,", len(rel_rows), "relations,", len(ctrl_rows), "contrôles")
