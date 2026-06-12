# WC2026 — Predicții & Pariuri Cupa Mondială 2026

Model de predicție pentru meciurile Cupei Mondiale 2026 (48 de echipe, 104 meciuri).
Combină **Dixon-Coles Poisson**, **XGBoost calibrat** și **ELO ratings** pentru a genera
probabilități, xG și recomandări de pariuri cu valoare așteptată (EV) pozitivă.

---

## Arhitectură

```
sofascore.py / aiscore_scraper.py   →  aiscore_team_stats.json  ─┐
elo_scraper.py                      →  elo_data.json             ─┤→  model_v2.py  →  backtest_v3.xlsx
main.py + collect.py + build_excel* →  statistici_calificari_*   ─┘
```

### Componente principale

| Fișier | Rol |
|--------|-----|
| `model_v2.py` | Modelul de predicție principal → generează `backtest_v3.xlsx` |
| `main.py` | Pipeline colectare statistici din SofaScore → `statistici_calificari_cm2026.xlsx` |
| `sofascore.py` | Client HTTP cu cache și rate-limiting pentru API-ul SofaScore |
| `aiscore_scraper.py` | Scraper statistici calificări (gpg, xG, cornere, cartonașe) |
| `elo_scraper.py` | Scraper ELO ratings pentru cele 48 de echipe |
| `collect.py` | Parsează meciurile de calificare + statistici |
| `build_excel_v6.py` | Generează Excel cu statistici agregate pe echipe |

---

## Modelul de predicție (`model_v2.py`)

### Ce face

1. **Antrenare XGBoost** pe meciurile de calificare (split temporal 80/20, fără data leakage)
2. **Dixon-Coles Poisson** cu parametrul RHO optimizat — corectează probabilitățile pentru scoruri mici (0-0, 1-0, 0-1, 1-1)
3. **Matrice 8×8** de probabilități joint (scor×scor) → Over/Under, GG, scoruri exacte
4. **ELO ratings** cu avantaj teren doar pentru gazdele WC2026 (SUA, Canada, Mexic)
5. **Blend model/piață** — afișaj final: 70% model + 30% cotele bookmakerului; EV calculat pe probabilitățile pure ale modelului
6. **Expected Value (EV)** = `prob_model × cotă_piață − 1`; pozitiv = pariu cu valoare

### Factori de calibrare per confederație

Statisticile din calificări sunt inflate față de nivelul CM (adversari mai slabi):

| Conf. | Factor |
|-------|--------|
| CONMEBOL | 0.90 |
| UEFA | 0.85 |
| AFC | 0.73 |
| CONCACAF | 0.78 |
| CAF | 0.70 |
| OFC | 0.62 |

### Output — `backtest_v3.xlsx`

| Sheet | Conținut |
|-------|---------|
| `Predictii` | Probabilități W/D/L, xG, blend cotă/model, EV |
| `Pariuri_Complete` | Toate tipurile de pariu generate cu EV și cotă piață |
| `Accumulator` | Top 3 pariuri EV+ combinate per meci |
| `Cel_Mai_Bun_Pariu` | Cel mai bun pariu per meci (după EV, nu probabilitate) |

### Tipuri de pariuri generate

- Rezultat 1X2, Double chance, Draw No Bet
- Over/Under 0.5, 1.5, 2.5, 3.5 goluri
- GG / NG (ambele marchează)
- Echipă marchează / nu marchează
- Win to nil, scor exact
- Cornere O/U, cartonașe galbene O/U, faulturi
- Top marcatori per echipă

---

## Instalare

```bash
pip install -r requirements.txt
```

**Dependențe model:**
```bash
pip install xgboost scikit-learn scipy openpyxl numpy
```

**Dependențe scraper:**
```bash
pip install curl-cffi openpyxl
```

---

## Rulare

### 1. Colectare date (o singură dată sau la actualizare)

```bash
python main.py              # statistici SofaScore → statistici_calificari_cm2026.xlsx
python elo_scraper.py       # ELO ratings → elo_data.json
python aiscore_scraper.py   # statistici AiScore → aiscore_team_stats.json
```

### 2. Generare predicții

```bash
python model_v2.py          # → backtest_v3.xlsx
```

Fișierele JSON se cache-uiesc; re-rulările nu re-lovesc serverele.

---

## Logica de pariuri

### EV (Expected Value)

```
EV = probabilitate_model × cotă_piață − 1
```

- `EV > 0` → pariu cu valoare (recomandat pe termen lung)
- `EV < 0` → cotă supraestimată de piață (evitat)

### Blend model/piață

Probabilitățile **afișate** în Excel combină:
```
prob_finala = 0.70 × prob_model + 0.30 × prob_implicita_piata
```
EV se calculează pe `prob_model` pur (nu pe blend) pentru a evita raționament circular.

### Accumulator

Acumulatorul per meci selectează top 3 pariuri după EV (dacă există cote de piață),
altfel după probabilitate brută. Cotele combinate se înmulțesc.

---

## Limitări cunoscute

- **xG** din calificări are acoperire slabă la CAF/AFC/OFC — modelul compensează prin factori de calibrare
- **Cotele de piață** (`MARKET_ODDS` în cod) sunt introduse manual; lipsesc pentru meciurile din fazele eliminatorii
- Acuratețea realistă pe meciuri CM este ~48-54% (XGBoost raportează ~80% pe datele de test datorită dezechilibrului calificări vs CM)
- Modelul nu captează forme recente, accidentări sau suspendări — verifică întotdeauna înainte de a paria

---

## Avertismente

> **Pariurile implică pierderi.** EV pozitiv garantează profit *statistic pe termen lung*, nu per pariu individual.
> Modelul este un instrument de analiză, nu un oracol. Pariați responsabil și doar sume pe care vă permiteți să le pierdeți.

Scraping-ul de date folosit este pentru uz personal/educațional. Respectați ToS-ul surselor de date.
