"""Rulează tot pipeline-ul: strânge date -> exportă Excel.

Utilizare:
    python main.py
Datele se cache-uiesc în ./cache, deci re-rulările sunt rapide și nu re-lovesc serverul.
"""
import config
from collect import collect_all
from build_excel import export


def main():
    print("Strâng statisticile din calificările CM 2026 (SofaScore)...")
    matches, players = collect_all()
    if not matches:
        print("\nNiciun meci strâns. Verifică: BASE corect, "
              "căutarea echipelor, denumirile turneelor (QUAL_NAME_MATCH).")
        return
    path = export(matches, players)
    print(f"\nGata: {len(matches)} meciuri din {len(set(m['echipa'] for m in matches))} echipe.")
    print(f"Excel salvat în: {path}")


if __name__ == "__main__":
    main()
