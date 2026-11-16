# TP Jour 1 – Retail Intelligence Platform

Younes Addi – IPSSI, Mastère Développement, Data & IA – 16/11/2026

| Élément | Contenu |
|---|---|
| `DAT_Retail_Intelligence_J1_Younes_Addi.pdf` | **Dossier d'Architecture Technique** : UC-01 (besoin, sources, KPI, qualité, limites) et UC-02 (architecture, flux, composants, gouvernance, Azure) |
| `annexes/UC01_inventaire_sources.xlsx` | Inventaire des 36 fichiers, dictionnaire des 344 colonnes, 47 relations, 34 contrôles qualité (avec requêtes) |
| `annexes/UC01_dictionnaire_KPI.xlsx` | 14 KPI détaillés avec requêtes et valeurs de contrôle |
| `preuves/scripts/` | Scripts Python / DuckDB de contrôle (reproductibles, voir annexe C du DAT) |
| `preuves/resultats/` | Sorties des scripts (arborescence, aperçus, test Parquet, vérification de la spécification v1.7.0) |
| `preuves/schemas/` | Sources Graphviz des schémas du DAT |
| `preuves/captures/` | 10 captures des contrôles rejoués sous WSL (Ubuntu 22.04) – reprises en annexe D du DAT |

Données : kit TPCx-BB v1.7.0 fourni par le formateur (`retail_project.zip`, facteur SF1, données synthétiques). Référence : TPCx-BB Standard Specification v1.7.0.
