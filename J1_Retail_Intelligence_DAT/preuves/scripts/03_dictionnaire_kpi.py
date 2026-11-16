"""Produit UC01_dictionnaire_KPI.xlsx et calcule une valeur de contrôle par KPI (lot data/, 2001-2005).

Usage : python3 -I 03_dictionnaire_kpi.py <fichier.duckdb> <sortie.xlsx>
"""
import sys

import duckdb
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

db, sortie = sys.argv[1:3]
con = duckdb.connect(db, read_only=True)
val = lambda sql: con.execute(sql).fetchone()[0]  # noqa: E731

SS = "SELECT sum(ss_net_paid) FROM d_store_sales"
WS = "SELECT sum(ws_net_paid) FROM d_web_sales"
YEAR = ("SELECT sum({m}) FROM d_{t} f JOIN d_date_dim d ON d.d_date_sk=f.{c} WHERE d.d_year={y}")

# (id, nom, définition, formule, unité, périmètre, granularité, fréquence, sources, manquant, limites,
#  calculable, sql de contrôle, format, prio)
KPI = [
    ("K01", "Chiffre d'affaires net HT",
     "Montant encaissé après remises et coupons, hors taxe et hors frais de port.",
     "SUM(ss_net_paid) + SUM(ws_net_paid)", "Devise non documentée (USD supposé)",
     "Ventes magasin + web, retours non déduits (voir K06)", "Jour × produit × magasin/site × client",
     "Quotidienne, restitution mensuelle", "store_sales.ss_net_paid, web_sales.ws_net_paid, date_dim",
     "Devise et régime de taxe non documentés", "Hors retours : un CA net de retours = K01 − montant retourné.",
     "Oui", f"SELECT ({SS}) + ({WS})", "#,##0", "I"),
    ("K02", "Nombre de transactions",
     "Nombre de tickets de caisse (magasin) et de commandes (web) distincts.",
     "COUNT(DISTINCT ss_ticket_number) + COUNT(DISTINCT ws_order_number)", "Nombre",
     "Ventes magasin + web", "Ticket / commande", "Quotidienne",
     "store_sales.ss_ticket_number, web_sales.ws_order_number", "—",
     "Ne jamais compter les lignes : 1 ticket = 7,5 lignes en moyenne.",
     "Oui", "SELECT (SELECT count(DISTINCT ss_ticket_number) FROM d_store_sales) + "
            "(SELECT count(DISTINCT ws_order_number) FROM d_web_sales)", "#,##0", "I"),
    ("K03", "Panier moyen (magasin)",
     "Montant moyen dépensé par ticket de caisse.",
     "SUM(ss_net_paid) / COUNT(DISTINCT ss_ticket_number)", "Devise / ticket",
     "Ventes magasin (même calcul possible pour le web avec ws_order_number)", "Ticket", "Mensuelle",
     "store_sales.ss_net_paid, ss_ticket_number", "—",
     "Numérateur et dénominateur sur le même périmètre. Valeur très élevée (≈ 12 900) car données synthétiques.",
     "Oui", f"SELECT ({SS}) / (SELECT count(DISTINCT ss_ticket_number) FROM d_store_sales)", "#,##0.00", "I"),
    ("K04", "Quantités vendues",
     "Nombre d'unités vendues.", "SUM(ss_quantity) + SUM(ws_quantity)", "Unités",
     "Ventes magasin + web, hors retours", "Jour × produit × magasin/site", "Quotidienne",
     "store_sales.ss_quantity, web_sales.ws_quantity", "—",
     "À ne pas confondre avec le nombre de lignes ou de transactions.",
     "Oui", "SELECT (SELECT sum(ss_quantity) FROM d_store_sales) + (SELECT sum(ws_quantity) FROM d_web_sales)",
     "#,##0", "I"),
    ("K05", "Taux de remise",
     "Part du prix catalogue accordée en remise.",
     "SUM(ext_discount_amt) / SUM(ext_list_price)", "%",
     "Ventes magasin + web (calculable par canal / catégorie)", "Ligne de vente agrégée", "Mensuelle",
     "ss_ext_discount_amt, ss_ext_list_price, ws_ext_discount_amt, ws_ext_list_price", "—",
     "Coupons non inclus (ss_coupon_amt séparé). Taux ≈ 50 % : niveau irréaliste (synthétique).",
     "Oui", "SELECT ((SELECT sum(ss_ext_discount_amt) FROM d_store_sales) + (SELECT sum(ws_ext_discount_amt) FROM d_web_sales))"
            " / ((SELECT sum(ss_ext_list_price) FROM d_store_sales) + (SELECT sum(ws_ext_list_price) FROM d_web_sales))",
     "0.0%", "U"),
    ("K06", "Taux de retour (quantité, magasin)",
     "Part des unités vendues qui sont retournées.",
     "SUM(sr_return_quantity) / SUM(ss_quantity)", "%",
     "Magasin (même calcul web avec wr_/ws_) ; période rattachée à la date de VENTE", "Produit × magasin × mois de vente",
     "Mensuelle", "store_returns.sr_return_quantity, store_sales.ss_quantity (jointure ticket + produit)",
     "Motifs de retour non interprétables",
     "Retours datés jusqu'en 2008 ; couples (ticket, produit) non uniques → dédoublonner avant jointure.",
     "Oui (réserves)", "SELECT (SELECT sum(sr_return_quantity) FROM d_store_returns) / 1.0 / (SELECT sum(ss_quantity) FROM d_store_sales)",
     "0.00%", "I"),
    ("K07", "Évolution du CA (N vs N-1)",
     "Variation du CA net HT par rapport à la même période de l'année précédente.",
     "(CA année N − CA année N-1) / CA année N-1", "%",
     "Magasin + web ; années complètes uniquement", "Année (ou mois vs même mois N-1)", "Mensuelle / annuelle",
     "K01 + date_dim.d_year, d_moy", "2005 incomplète dans data/ (refresh nécessaire)",
     "Comparer des périodes de même durée. Valeur de contrôle : 2004 vs 2003.",
     "Oui", "SELECT ((" + YEAR.format(m="ss_net_paid", t="store_sales", c="ss_sold_date_sk", y=2004) + ")+(" +
            YEAR.format(m="ws_net_paid", t="web_sales", c="ws_sold_date_sk", y=2004) + ")) / ((" +
            YEAR.format(m="ss_net_paid", t="store_sales", c="ss_sold_date_sk", y=2003) + ")+(" +
            YEAR.format(m="ws_net_paid", t="web_sales", c="ws_sold_date_sk", y=2003) + ")) - 1", "0.00%", "I"),
    ("K08", "Part du CA réalisée sur le web",
     "Poids du canal web dans le CA total.", "CA web / (CA magasin + CA web)", "%",
     "Magasin + web", "Mois × canal", "Mensuelle", "store_sales.ss_net_paid, web_sales.ws_net_paid", "—",
     "Pas de canal catalogue dans le kit.", "Oui", f"SELECT ({WS}) / (({SS}) + ({WS}))", "0.0%", "I"),
    ("K09", "Stock moyen",
     "Quantité moyenne en stock par produit et entrepôt sur les relevés hebdomadaires.",
     "AVG(inv_quantity_on_hand)", "Unités",
     "5 entrepôts, relevés non vides", "Produit × entrepôt × semaine", "Hebdomadaire",
     "inventory.inv_quantity_on_hand, inv_date_sk, inv_item_sk, inv_warehouse_sk", "Mouvements de stock absents",
     "Snapshot : ne jamais sommer dans le temps. 5 999 relevés vides exclus.",
     "Oui", "SELECT avg(inv_quantity_on_hand) FROM d_inventory", "#,##0.0", "U"),
    ("K10", "Taux de rupture",
     "Part des relevés (produit × entrepôt × semaine) avec un stock nul.",
     "COUNT(qté = 0) / COUNT(qté renseignée)", "%",
     "Relevés non vides", "Produit × entrepôt × semaine", "Hebdomadaire", "inventory.inv_quantity_on_hand",
     "Demande non satisfaite inconnue", "Rupture constatée seulement à la date du relevé (pas entre deux relevés).",
     "Oui (réserves)", "SELECT avg(CASE WHEN inv_quantity_on_hand=0 THEN 1.0 ELSE 0 END) FROM d_inventory "
                       "WHERE inv_quantity_on_hand IS NOT NULL", "0.00%", "U"),
    ("K11", "Note moyenne des avis",
     "Moyenne des notes (1 à 5) laissées par les clients.", "AVG(pr_review_rating)", "Note sur 5",
     "Avis produits", "Produit × mois", "Mensuelle", "product_reviews.pr_review_rating, pr_item_sk", "—",
     "Distribution très asymétrique (61 % de 5) ; lien avis ↔ achat vérifié à 92,8 % seulement.",
     "Oui", "SELECT avg(pr_review_rating) FROM d_product_reviews", "0.00", "C"),
    ("K12", "Part des clics liés à un achat",
     "Proportion des clics web associés à une ligne de commande.",
     "COUNT(wcs_sales_sk) / COUNT(*)", "%", "Site web", "Jour × page / produit", "Quotidienne",
     "web_clickstreams.wcs_sales_sk", "Sessions, ajouts au panier, source de trafic",
     "Ce n'est PAS un taux de conversion (pas de session). 42 % des clics anonymes.",
     "Partiel", "SELECT count(wcs_sales_sk) * 1.0 / count(*) FROM d_web_clickstreams", "0.0%", "C"),
    ("K13", "Effet des promotions", "Écart de ventes entre périodes avec et sans promotion.",
     "Non défini", "%", "—", "—", "—", "promotion, ss_promo_sk", "Groupe témoin (ventes hors promo)",
     "99,98 % des lignes ont un code promo ; dates de promotion 1901–2199.", "Non", None, None, "C"),
    ("K14", "Marge nette", "Bénéfice net des ventes.", "SUM(ss_net_profit)", "Devise", "—", "—", "—",
     "store_sales.ss_net_profit", "Coûts réalistes", "74 % des lignes en perte : non interprétable.",
     "Non fiable", None, None, "C"),
]

wb = Workbook()
ws = wb.active
ws.title = "Dictionnaire KPI"
F = "Arial"
ws["A1"] = "UC-01 – Dictionnaire des indicateurs (KPI) – Retail Intelligence Platform"
ws["A1"].font = Font(name=F, bold=True, size=14)
ws["A2"] = ("Valeur de contrôle : calculée sur le lot data/ (01/01/2001 → 13/12/2005) par 03_dictionnaire_kpi.py "
            "(DuckDB) pour prouver que le KPI est calculable. Ce n'est pas un objectif métier.")
ws["A2"].font = Font(name=F, italic=True, size=10)
head = ["ID", "Nom", "Définition métier", "Formule", "Unité", "Périmètre", "Granularité", "Fréquence",
        "Sources / champs", "Données manquantes", "Limites d'interprétation", "Calculable ?",
        "Valeur de contrôle", "Priorité", "Requête de contrôle"]
thin = Side(style="thin", color="BFBFBF")
B = Border(left=thin, right=thin, top=thin, bottom=thin)
for j, h in enumerate(head, 1):
    c = ws.cell(4, j, h)
    c.font = Font(name=F, bold=True, color="FFFFFF", size=10)
    c.fill = PatternFill("solid", start_color="1F3864")
    c.alignment = Alignment(wrap_text=True, vertical="center")
    c.border = B
fills = {"Oui": "E2EFDA", "Oui (réserves)": "FFF2CC", "Partiel": "FFF2CC", "Non": "FCE4D6", "Non fiable": "FCE4D6"}
for i, k in enumerate(KPI, 5):
    (kid, nom, defi, form, unit, peri, gran, freq, src, manq, lim, calc, sql, fmt, prio) = k
    v = val(sql) if sql else "—"
    row = [kid, nom, defi, form, unit, peri, gran, freq, src, manq, lim, calc, v, prio, sql or "—"]
    for j, x in enumerate(row, 1):
        c = ws.cell(i, j, x)
        c.font = Font(name=F, size=10, bold=(j == 2))
        c.alignment = Alignment(wrap_text=True, vertical="top")
        c.border = B
    if fmt:
        ws.cell(i, 13).number_format = fmt
    ws.cell(i, 12).fill = PatternFill("solid", start_color=fills[calc])
    print(kid, nom, v)
for j, w in enumerate([6, 24, 34, 30, 14, 26, 22, 16, 34, 22, 40, 12, 16, 9, 50], 1):
    ws.column_dimensions[chr(64 + j)].width = w
ws.freeze_panes = "C5"
r = 5 + len(KPI) + 1
ws.cell(r, 1, "Priorité : I = indispensable, U = utile, C = complémentaire. Montants : devise non documentée "
              "(hypothèse USD, HT – voir UC01_limites_et_hypotheses.pdf, H02).").font = Font(name=F, italic=True, size=10)
wb.save(sortie)
