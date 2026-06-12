"""Configurare proiect.

NOTA: denumirile câmpurilor SofaScore (STAT_ALIASES) sunt mapate defensiv.
La prima rulare, verifică un răspuns real de la endpoint-ul /statistics și
ajustează aliasurile dacă vreo coloană iese goală.
"""

# --- Endpoint-uri SofaScore (neoficiale) ---
BASE = "https://www.sofascore.com/api/v1"

# --- Fereastra calificărilor (toate confederațiile) ---
# Primele meciuri AFC au fost pe 12 oct 2023; ultimele play-off-uri în martie 2026.
QUAL_START = "2023-09-01"
QUAL_END = "2026-04-15"

# Numele turneelor de calificare conțin acest substring pe SofaScore
# (ex: "World Cup Qual. UEFA", "World Cup Qual. CONMEBOL", "... Inter-confederation play-offs")
QUAL_NAME_MATCH = ["world cup qual", "wc qual"]

# Minim de apariții ca un jucător să fie eligibil pentru "cel mai bun jucător"
MIN_APPEARANCES = 3

# ID-uri Sofascore hardcodate pentru echipe unde search-ul întoarce rezultate greșite
# (ex: "Turkey" găsește echipa de esports, "New Zealand" găsește cricket)
TEAM_ID_OVERRIDES = {
    "Türkiye": 4700,
    "New Zealand": 4784,
    "Ivory Coast": 4768,  # search "Ivory Coast" -> cricket; fotbal = Côte d'Ivoire id 4768
}

# --- Politețe ---
REQUEST_DELAY = 2.0        # secunde între cereri
REQUEST_JITTER = 1.0       # + 0..1s aleator
MAX_RETRIES = 3
TIMEOUT = 20

CACHE_DIR = "cache"
OUTPUT_XLSX = "statistici_calificari_cm2026.xlsx"

# Header de browser normal. NU folosim proxy rotation / bypass de challenge.
# Dacă primești 403, vezi README (sursa are protecție anti-bot).
HEADERS = {
    "User-Agent": ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                   "AppleWebKit/537.36 (KHTML, like Gecko) "
                   "Chrome/125.0.0.0 Safari/537.36"),
    "Accept": "application/json, text/plain, */*",
    "Accept-Language": "en-US,en;q=0.9",
    "Accept-Encoding": "gzip, deflate, br",
    "Referer": "https://www.sofascore.com/",
    "Origin": "https://www.sofascore.com",
    "Sec-Fetch-Dest": "empty",
    "Sec-Fetch-Mode": "cors",
    "Sec-Fetch-Site": "same-origin",
}

# --- Cele 48 de echipe calificate (CM 2026) ---
# Numele sunt cele folosite în căutarea SofaScore; ajustează dacă search nu le găsește
# (ex. "South Korea" vs "Korea Republic", "Ivory Coast" vs "Côte d'Ivoire").
QUALIFIED_TEAMS = [
    # Gazde (CONCACAF)
    "USA", "Mexico", "Canada",
    # UEFA (16)
    "England", "France", "Croatia", "Norway", "Portugal", "Germany",
    "Netherlands", "Switzerland", "Scotland", "Spain", "Austria", "Belgium",
    "Bosnia & Herzegovina", "Sweden", "Türkiye", "Czechia",
    # CONMEBOL (6)
    "Argentina", "Brazil", "Colombia", "Ecuador", "Paraguay", "Uruguay",
    # CAF (9)
    "Algeria", "Cape Verde", "Egypt", "Ghana", "Ivory Coast", "Morocco",
    "Senegal", "South Africa", "Tunisia",
    # CONCACAF non-gazde (3)
    "Panama", "Curacao", "Haiti",
    # AFC (8)
    "Australia", "Iran", "Japan", "Jordan", "Qatar", "Saudi Arabia",
    "South Korea", "Uzbekistan",
    # OFC (1)
    "New Zealand",
    # Play-off interconfederații (2)
    "DR Congo", "Iraq",
]

# --- Maparea statisticilor cerute -> posibile denumiri SofaScore ---
STAT_ALIASES = {
    # Atac
    "xg":                       ["expected goals", "xg"],
    "suturi_total":             ["total shots"],
    "suturi_pe_poarta":         ["shots on target", "shots on goal"],
    "suturi_pe_langa":          ["shots off target"],
    "suturi_blocate":           ["blocked shots"],
    "suturi_in_careu":          ["shots inside box"],
    "suturi_afara_careu":       ["shots outside box"],
    "sanse_mari":               ["big chances"],
    "sanse_mari_ratate":        ["big chances missed"],
    "bara":                     ["hit woodwork"],
    # Posesie / construcție
    "posesie":                  ["ball possession"],
    "pase_total":               ["passes", "total passes"],
    "pase_precise":             ["accurate passes"],
    "centrari":                 ["crosses"],
    "mingi_lungi":              ["long balls"],
    "dribling":                 ["dribbles", "successful dribbles"],
    "intrari_treime_finala":    ["final third entries"],
    "mingi_prin_aparare":       ["through balls"],
    # Set piece-uri
    "cornere":                  ["corner kicks", "corners"],
    "aruncari_de_la_margine":   ["throw-ins"],
    "lovituri_libere":          ["free kicks"],
    "lovituri_de_poarta":       ["goal kicks"],
    # Apărare
    "fault_comise":             ["fouls", "fouls committed"],
    "fault_suferite_zona":      ["fouled in final third"],
    "ofsaid":                   ["offsides"],
    "galbene":                  ["yellow cards"],
    "rosii":                    ["red cards"],
    "recuperari":               ["recoveries", "ball recoveries"],
    "interceptari":             ["interceptions"],
    "degajari":                 ["clearances"],
    "dueluri_aeriene":          ["aerial duels"],
    "dueluri_sol":              ["ground duels"],
    "tackle_reusit":            ["tackles won"],
    "tackle_total":             ["total tackles"],
    "save_uri":                 ["goalkeeper saves", "total saves"],
}

# Coloane care sunt procente/medii (nu se însumează ci se mediază)
PERCENT_STATS = {"posesie"}
