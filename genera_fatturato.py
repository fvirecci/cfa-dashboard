#!/usr/bin/env python3
"""
CFA Consulenze - Generatore JSON fatturato
==========================================
Uso: python genera_fatturato.py

Metti questo script nella stessa cartella dei tuoi file Excel.
Ogni mese:
  1. Esporta il nuovo Excel dal gestionale
  2. Aggiornalo in ANNI_FILES qui sotto
  3. Esegui: python genera_fatturato.py
  4. Carica fatturato.json su GitHub / Aruba

Formato Excel richiesto: export DK SET "Fatturato Clienti" con contropartite.
"""

import openpyxl
import json
from collections import defaultdict
from pathlib import Path

# ── CONFIGURAZIONE ────────────────────────────────────────
# Aggiungi qui i tuoi file Excel per anno.
# Chiave = anno (intero), Valore = nome del file Excel
ANNI_FILES = {
    2024: "2024.xlsx",
    2025: "2025.xlsx",
    # 2026: "2026.xlsx",   # <-- aggiungi quando disponibile
}

OUTPUT_FILE = "fatturato.json"
# ─────────────────────────────────────────────────────────


def extract_from_excel(filepath, year):
    """Estrae clienti e contropartite da un export DK SET."""
    wb = openpyxl.load_workbook(filepath)
    ws = wb.active
    clients = []
    current = None

    for row in ws.iter_rows(values_only=True):
        code   = row[0]
        col3   = row[3]
        col5   = row[5]
        imp    = row[6]
        totale = row[13]

        # Riga cliente principale: codice numerico in colonna A
        if code and str(code).strip().isdigit() and col3 and totale:
            current = {
                "codice":   str(code).strip(),
                "cliente":  str(col3).strip(),
                "imponibile": round(float(imp or 0), 2),
                "totale":     round(float(totale or 0), 2),
                "contropartite": [],
            }
            clients.append(current)

        # Riga contropartita: colonna A vuota, col3=conto, col5=descrizione
        elif current and code is None and col3 and col5 and imp:
            desc = str(col5).strip().replace("Ricavi/", "").strip()
            current["contropartite"].append({
                "conto":       str(col3).strip(),
                "descrizione": desc,
                "imponibile":  round(float(imp or 0), 2),
            })

    print(f"  {year}: {len(clients)} clienti estratti")
    return clients


def build_json(anni_data):
    """Costruisce il JSON completo per la dashboard."""
    all_years = sorted(anni_data.keys())
    years_str = [str(y) for y in all_years]
    by_year   = {y: {c["codice"]: c for c in clients} for y, clients in anni_data.items()}

    # ── Totali per anno ───────────────────────────────────
    totals_by_year = {
        str(y): round(sum(c["imponibile"] for c in clients), 2)
        for y, clients in anni_data.items()
    }

    # ── Clienti multi-anno ────────────────────────────────
    all_codes = set()
    for y in all_years:
        all_codes |= set(by_year[y].keys())

    clients_timeline = []
    for code in all_codes:
        name = next((by_year[y][code]["cliente"] for y in all_years if code in by_year[y]), code)
        anni = {}
        for y in all_years:
            if code in by_year[y]:
                c = by_year[y][code]
                anni[str(y)] = {
                    "imponibile":    c["imponibile"],
                    "contropartite": c["contropartite"],
                }
            else:
                anni[str(y)] = {"imponibile": 0, "contropartite": []}

        vals = [anni[str(y)]["imponibile"] for y in all_years]
        clients_timeline.append({
            "codice":    code,
            "cliente":   name,
            "anni":      anni,
            "total_all": round(sum(vals), 2),
        })

    clients_timeline.sort(
        key=lambda x: x["anni"].get(str(max(all_years)), {}).get("imponibile", 0),
        reverse=True
    )

    # ── Categorie multi-anno ──────────────────────────────
    all_cats = set()
    for y, clients in anni_data.items():
        for c in clients:
            for cp in c["contropartite"]:
                all_cats.add(cp["descrizione"])

    cats_timeline = []
    for cat in all_cats:
        anni_cat = {}
        for y, clients in anni_data.items():
            val = sum(
                cp["imponibile"]
                for c in clients
                for cp in c["contropartite"]
                if cp["descrizione"] == cat
            )
            anni_cat[str(y)] = round(val, 2)
        cats_timeline.append({
            "categoria": cat,
            "anni":      anni_cat,
            "total_all": round(sum(anni_cat.values()), 2),
        })

    cats_timeline.sort(
        key=lambda x: x["anni"].get(str(max(all_years)), 0),
        reverse=True
    )

    # ── Waterfall nuovi/persi/cresciuti/calati ────────────
    waterfall = []
    for i, y in enumerate(all_years):
        if i == 0:
            waterfall.append({
                "anno":          y,
                "nuovi":         len(anni_data[y]),
                "persi":         0,
                "cresciuti":     0,
                "calati":        0,
                "tot_nuovi_val": round(totals_by_year[str(y)], 2),
                "tot_persi_val": 0,
                "tot_cresciuti_val": 0,
                "tot_calati_val":    0,
            })
        else:
            prev_y    = all_years[i - 1]
            prev_codes = set(by_year[prev_y].keys())
            curr_codes = set(by_year[y].keys())
            nuovi  = curr_codes - prev_codes
            persi  = prev_codes - curr_codes
            comuni = prev_codes & curr_codes
            cresciuti = [c for c in comuni if by_year[y][c]["imponibile"] > by_year[prev_y][c]["imponibile"]]
            calati    = [c for c in comuni if by_year[y][c]["imponibile"] < by_year[prev_y][c]["imponibile"]]
            waterfall.append({
                "anno":     y,
                "nuovi":    len(nuovi),
                "persi":    len(persi),
                "cresciuti": len(cresciuti),
                "calati":    len(calati),
                "tot_nuovi_val":    round(sum(by_year[y][c]["imponibile"]  for c in nuovi), 2),
                "tot_persi_val":    round(sum(by_year[prev_y][c]["imponibile"] for c in persi), 2),
                "tot_cresciuti_val": round(sum(by_year[y][c]["imponibile"] - by_year[prev_y][c]["imponibile"] for c in cresciuti), 2),
                "tot_calati_val":    round(sum(by_year[y][c]["imponibile"] - by_year[prev_y][c]["imponibile"] for c in calati), 2),
            })

    return {
        "years":          years_str,
        "totals_by_year": totals_by_year,
        "waterfall":      waterfall,
        "clients":        clients_timeline,
        "categorie":      cats_timeline,
        "top_clients_last": clients_timeline[:20],
        "top_cats_last":    cats_timeline[:20],
    }


def main():
    print("=" * 50)
    print("CFA Consulenze · Generatore JSON fatturato")
    print("=" * 50)

    # Controlla che i file esistano
    anni_data = {}
    for year, filename in ANNI_FILES.items():
        path = Path(filename)
        if not path.exists():
            print(f"  ⚠️  File non trovato: {filename} (anno {year} saltato)")
            continue
        print(f"  📂 Leggo {filename}...")
        anni_data[year] = extract_from_excel(path, year)

    if not anni_data:
        print("\n❌ Nessun file Excel trovato. Controlla ANNI_FILES.")
        return

    print(f"\n  🔧 Costruisco JSON...")
    output = build_json(anni_data)

    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(output, f, ensure_ascii=False, separators=(",", ":"))

    size_kb = Path(OUTPUT_FILE).stat().st_size // 1024
    print(f"\n✅ {OUTPUT_FILE} generato ({size_kb} KB)")
    print(f"   Anni:     {output['years']}")
    print(f"   Clienti:  {len(output['clients'])}")
    print(f"   Categorie:{len(output['categorie'])}")
    for y, tot in output["totals_by_year"].items():
        print(f"   {y}: {tot:>12,.2f} €")
    print("\n📤 Carica fatturato.json su GitHub o Aruba per aggiornare la dashboard.")


if __name__ == "__main__":
    main()
