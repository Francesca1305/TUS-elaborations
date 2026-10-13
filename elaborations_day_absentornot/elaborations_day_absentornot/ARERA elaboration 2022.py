import pandas as pd
import glob, re, unicodedata
from pathlib import Path
import numpy as np

year = "2022"
ARERA_row_data_path = Path(rf"C:\Users\frenc\PycharmProjects\TUS-elaborations\2022prov\2022\dati prelievo orario per provincia3_a_4_5 anno 2022.xlsx")
out_dir = Path(r"C:\Users\frenc\PycharmProjects\TUS-elaborations\elaborations_day_absentornot\elaborations_day_absentornot\ARERA")

db_activities_path   = Path(
    r"/elaborations_day_absentornot/elaborations_day_absentornot/ARERA/Database_grafico_activities/Database_grafico_activities.xlsx")      # adatta
out_activities_path  = db_activities_path.with_name(db_activities_path.stem + f"_{year}.xlsx")

db_subactivities_path   = Path(
    r"/elaborations_day_absentornot/elaborations_day_absentornot/ARERA/Databas_grafico_subactivities/Database_grafico_subactivities.xlsx")      # adatta
out_subactivities_path  = db_subactivities_path.with_name(db_subactivities_path.stem + f"_{year}.xlsx")

MACRO = {
 "North-West": ["Alessandria","Asti","Biella","Cuneo","Novara","Torino","Verbano-Cusio-Ossola","Vercelli",
                "Valle d'Aosta/Vallée d'Aoste", "Bergamo","Brescia","Como",
                "Cremona","Lecco","Lodi","Mantova","Milano","Monza e della Brianza","Pavia","Sondrio","Varese",
                "Genova","Imperia","La Spezia","Savona"],
 "North-East": ["Bolzano/Bozen","Trento","Belluno","Padova","Rovigo","Treviso","Venezia","Verona","Vicenza",
                "Gorizia","Pordenone","Trieste","Udine","Bologna","Ferrara","Forlì-Cesena","Modena","Parma",
                "Piacenza","Ravenna","Reggio nell'Emilia","Rimini"],
 "Centre": ["Arezzo","Firenze","Grosseto","Livorno","Lucca","Massa-Carrara","Pisa","Pistoia","Prato","Siena",
            "Perugia","Terni","Ancona","Ascoli Piceno","Fermo","Macerata","Pesaro e Urbino","Frosinone","Latina",
            "Rieti","Roma","Viterbo"],
 "South": ["Chieti","L'Aquila","Pescara","Teramo","Campobasso","Isernia","Avellino","Benevento","Caserta",
           "Napoli","Salerno","Bari","Barletta-Andria-Trani","Brindisi","Foggia","Lecce","Taranto","Matera",
           "Potenza","Catanzaro","Cosenza","Crotone","Reggio di Calabria","Vibo Valentia"],
 "Islands": ["Agrigento","Caltanissetta","Catania","Enna","Messina","Palermo","Ragusa","Siracusa","Trapani",
             "Cagliari","Nuoro","Oristano","Sassari","Sud Sardegna"],
}
MACRO_CODE = {"North-West":1,"North-East":2,"Centre":3,"South":4,"Islands":5}

def norm(s):
    s = unicodedata.normalize("NFKD", str(s)).encode("ascii","ignore").decode()
    s = s.lower().replace("’","'").replace("-"," ").replace("/"," ")
    return re.sub(r"\s+"," ", re.sub(r"[^a-z' ]","", s)).strip()

PROV2MACRO = {norm(p): m for m, ps in MACRO.items() for p in ps}

ALIASES = {
 "bolzano":"Bolzano/Bozen", "bozen":"Bolzano/Bozen",
 "aosta":"Valle d'Aosta/Vallée d'Aoste", "valle d'aosta":"Valle d'Aosta/Vallée d'Aoste",
 "monza brianza":"Monza e della Brianza", "monza e brianza":"Monza e della Brianza",
 "reggio emilia":"Reggio nell'Emilia", "reggio calabria":"Reggio di Calabria",
 "forli cesena":"Forlì-Cesena", "pesaro urbino":"Pesaro e Urbino",
 "massa carrara":"Massa-Carrara", "verbano cusio ossola":"Verbano-Cusio-Ossola",
 "barletta andria trani":"Barletta-Andria-Trani",
 "carbonia iglesias":"Sud Sardegna", "medio campidano":"Sud Sardegna",
 "olbia tempio":"Sassari", "ogliastra":"Nuoro",
}
for k, v in ALIASES.items():
    if norm(k) not in PROV2MACRO:          # non sovrascrivere match diretti
        PROV2MACRO[norm(k)] = PROV2MACRO[norm(v)]

MESI = {"gen":1,"feb":2,"mar":3,"apr":4,"mag":5,"giu":6,"lug":7,"ago":8,"set":9,"ott":10,"nov":11,"dic":12}

ARERA_COLUMNS = [
    "Anno mese",
    "Provincia",
    "Tipo Mercato",
    "Classe potenza",
    "Residenza",
    "Working Day",
    "Orario",
    "Prelievo medio Orario Provinciale (kWh)",
]

INTERNAL_COLUMNS = [
    "AnnoMese",
    "Provincia",
    "TipoMercato",
    "ClassePotenza",
    "Residenza",
    "WorkingDay",
    "Orario",
    "Prelievo_kWh",
]

def load(path):
    path = Path(path)

    if path.suffix.lower() in {".xlsx", ".xlsm"}:
        df = pd.read_excel(path, engine="openpyxl")
    elif path.suffix.lower() == ".xls":
        df = pd.read_excel(path)
    elif path.suffix.lower() == ".csv":
        df = pd.read_csv(
            path,
            sep=None,
            engine="python",
            decimal=",",
            encoding="utf-8-sig"
        )
    else:
        raise ValueError(f"Formato file non supportato: {path.suffix}")

    df.columns = df.columns.astype(str).str.strip()

    missing = [col for col in ARERA_COLUMNS if col not in df.columns]
    if missing:
        raise KeyError(
            f"Colonne mancanti nel file ARERA: {missing}\n"
            f"Intestazioni trovate: {df.columns.tolist()}"
        )

    df = df[ARERA_COLUMNS].copy()
    df.columns = INTERNAL_COLUMNS

    # Converte date Excel o stringhe data, per esempio "2022-01-01 00:00:00".
    date_anno_mese = pd.to_datetime(df["AnnoMese"], errors="coerce")

    df["Mese"] = date_anno_mese.dt.month
    df["Anno"] = date_anno_mese.dt.year

    # Ripiego per valori testuali come "gen-22" o "gen/2022".
    m = df["AnnoMese"].astype(str).str.strip().str.extract(
        r"^([A-Za-z]{3})\D*(\d{2,4})$"
    )
    mese_testo = m[0].str.lower().map(MESI)
    anno_testo = pd.to_numeric(m[1], errors="coerce")

    da_testo = df["Mese"].isna() & mese_testo.notna()
    df.loc[da_testo, "Mese"] = mese_testo[da_testo]
    df.loc[da_testo, "Anno"] = anno_testo[da_testo].where(
        anno_testo[da_testo] > 99,
        2000 + anno_testo[da_testo]
    )

    invalid = df["Mese"].isna() | df["Anno"].isna()
    print(
        f"{path.name}: scartate {invalid.sum()} righe con data non valida "
        f"({df.loc[invalid, 'AnnoMese'].astype(str).unique()[:5]})"
    )

    df = df.loc[~invalid].copy()
    df["Mese"] = df["Mese"].astype(int)
    df["Anno"] = df["Anno"].astype(int)

    df["Prelievo_kWh"] = pd.to_numeric(
        df["Prelievo_kWh"].astype(str).str.replace(",", ".", regex=False),
        errors="coerce"
    )

    df["Ora"] = pd.to_numeric(
        df["Orario"].astype(str).str.extract(r"(?i)ora\s*(\d+)")[0],
        errors="coerce"
    )

    df = df.dropna(subset=["Ora"])
    df["Ora"] = df["Ora"].astype(int)

    return df.reset_index(drop=True)

df = load(ARERA_row_data_path)

df["macro_region"] = df["Provincia"].map(lambda p: PROV2MACRO.get(norm(p)))
df["macro_region_code"] = df["macro_region"].map(MACRO_CODE)

missing = df.loc[df["macro_region"].isna(), "Provincia"].unique()
print("Province non mappate:", missing)

prov_table = (df[["Provincia","macro_region","macro_region_code"]]
              .drop_duplicates().sort_values(["macro_region_code","Provincia"]))

df.to_csv(out_dir / f"arera_{year}_macroregion.csv", index=False, encoding="utf-8-sig")
prov_table.to_csv(out_dir / f"provincia_macroregion_{year}.csv", index=False, encoding="utf-8-sig")
print(prov_table.to_string(index=False))

FILTRI = {
    "TipoMercato":   "Mercato Libero",
    "ClassePotenza": "3<potenza_impegnata<=4.5",
    "Residenza":     "Residente",
}
mask = pd.Series(True, index=df.index)
for col, val in FILTRI.items():
    mask &= df[col].astype(str).str.strip().eq(val)
sub = df[mask].copy()
print("Righe dopo i filtri:", len(sub))

# Nov-Dic-Gen = 1, Feb-Apr = 2, Mag-Lug = 3, Ago-Ott = 4
sub["meseri_y"] = ((sub["Mese"] - 11) % 12) // 3 + 1
sub["rip"] = sub["macro_region_code"].astype(int)

sub["WorkingDay_norm"] = (
    sub["WorkingDay"]
    .astype(str)
    .str.strip()
    .str.upper()
)

DAYS = {
    "GIORNO_FERIALE": (1, "day_1_weekday"),
    "SAB":            (2, "day_2_saturday"),
    "DOM":            (3, "day_3_sunday"),
}

print("Valori Working Day nel file:")
print(sub["WorkingDay_norm"].value_counts(dropna=False))

out_xlsx = out_dir / f"arera_{year}_profili_macroregione.xlsx"

with pd.ExcelWriter(out_xlsx, engine="openpyxl") as writer:
    for working_day, (code, sheet) in DAYS.items():
        d = sub[sub["WorkingDay_norm"].eq(working_day)].copy()

        prof = (
            d.groupby(["rip", "meseri_y", "Ora"])["Prelievo_kWh"]
             .mean()
             .unstack("Ora")
             .reindex(columns=range(1, 25))
             .sort_index(axis=1)
        )

        prof.columns = [f"load_ora_{int(h) - 1}" for h in prof.columns]
        prof.insert(0, "day", code)
        prof = prof.reset_index().sort_values(["rip", "meseri_y"])

        prof.to_excel(writer, sheet_name=sheet, index=False)
        print(f"{sheet}: {prof.shape}")


print("Salvato:", out_xlsx)

########
profile_ARERA_path = out_xlsx

LOAD_COLS = [f"load_ora_{i}" for i in range(24)]
ORA_PROB_COLS = [f"ora_{i}" for i in range(24)]
ORA_OCC_COLS  = [f"ora_{i}_occupancy" for i in range(24)]

KEYS = ["rip", "day_type", "meseri_y"]

# Legge una sola volta i tre fogli ARERA e costruisce la tabella
# comune di 60 profili: 5 macro-regioni × 3 tipi di giorno × 4 periodi.
sheets = pd.read_excel(profile_ARERA_path, sheet_name=None)

prof = (
    pd.concat(sheets.values(), ignore_index=True)
      .rename(columns={"day": "day_type"})
      [KEYS + LOAD_COLS]
      .copy()
)

for k in KEYS:
    prof[k] = pd.to_numeric(prof[k], errors="coerce").astype("Int64")

assert not prof.duplicated(KEYS).any(), "Chiavi duplicate nei profili ARERA"
assert len(prof) == 60, (
    f"Attese 60 combinazioni ARERA (5 macro-regioni × 3 giorni × 4 periodi), "
    f"trovate {len(prof)}"
)

def pearson_r2(x, y):
    """
    Quadrato della correlazione di Pearson.

    Equivalente a Excel:
        =CORRELAZIONE(intervallo_x; intervallo_y)^2

    Restituisce NaN se:
    - almeno una delle 24 ore non è disponibile;
    - uno dei vettori è costante;
    - sono presenti valori non numerici.
    """
    x = pd.to_numeric(x, errors="coerce").to_numpy(dtype=float)
    y = pd.to_numeric(y, errors="coerce").to_numpy(dtype=float)

    valid = np.isfinite(x) & np.isfinite(y)

    # Calcolo soltanto se sono presenti tutte le 24 ore.
    if valid.sum() != 24:
        return np.nan

    x = x[valid]
    y = y[valid]

    x_centered = x - x.mean()
    y_centered = y - y.mean()

    denominator = np.sqrt(
        np.sum(x_centered ** 2) *
        np.sum(y_centered ** 2)
    )

    # Pearson non è definita se x o y sono costanti.
    if np.isclose(denominator, 0):
        return np.nan

    r = np.sum(x_centered * y_centered) / denominator

    # Piccola protezione contro arrotondamenti numerici.
    r = np.clip(r, -1.0, 1.0)

    return r ** 2


def add_load_profiles_and_r2(input_path, output_path, label):
    """
    Legge un database di attività/sotto-attività, associa il profilo ARERA
    mediante rip + day_type + meseri_y e calcola R² di Pearson, riga per riga.

    Parameters
    ----------
    input_path : Path
        File Excel sorgente, con attività oppure sotto-attività.
    output_path : Path
        File Excel finale da scrivere.
    label : str
        Nome da mostrare nei messaggi di controllo.
    """
    db = pd.read_excel(input_path)

    # Verifica preventiva delle colonne necessarie al merge.
    missing_keys = [k for k in KEYS if k not in db.columns]
    if missing_keys:
        raise KeyError(
            f"\n[{label}] Colonne chiave mancanti nel database:\n"
            + "\n".join(missing_keys)
        )

    # Il database delle sotto-attività deve contenere anch'esso le 24 colonne
    # probabilità e le 24 colonne di occupancy, se vuoi entrambi gli R².
    required_cols = ORA_PROB_COLS + ORA_OCC_COLS
    missing_profile_cols = [c for c in required_cols if c not in db.columns]

    if missing_profile_cols:
        raise KeyError(
            f"\n[{label}] Colonne orarie mancanti nel database:\n"
            + "\n".join(missing_profile_cols)
        )

    # Uniforma i tipi delle chiavi prima del merge.
    for k in KEYS:
        db[k] = pd.to_numeric(db[k], errors="coerce").astype("Int64")

    # Join: molte righe del database possono puntare allo stesso profilo ARERA.
    new = db[KEYS].merge(
        prof,
        on=KEYS,
        how="left",
        validate="many_to_one"
    )

    n_miss = new[LOAD_COLS].isna().all(axis=1).sum()

    print(f"\n[{label}] Righe senza match ARERA: {n_miss} su {len(db)}")

    if n_miss:
        print(
            f"[{label}] Combinazioni senza profilo ARERA:"
        )
        print(
            db.loc[
                new[LOAD_COLS].isna().all(axis=1).to_numpy(),
                KEYS
            ]
            .drop_duplicates()
            .sort_values(KEYS)
            .to_string(index=False)
        )

    db_upd = db.copy()

    # Inserisce / sovrascrive le 24 colonne di carico ARERA.
    db_upd[LOAD_COLS] = new[LOAD_COLS].to_numpy()

    # R² = Pearson(x, load)^2: profilo di probabilità vs carico ARERA.
    db_upd["R2_prob"] = db_upd.apply(
        lambda row: pearson_r2(row[ORA_PROB_COLS], row[LOAD_COLS]),
        axis=1
    )

    # R² = Pearson(x, load)^2: profilo occupancy vs carico ARERA.
    db_upd["R2_occ"] = db_upd.apply(
        lambda row: pearson_r2(row[ORA_OCC_COLS], row[LOAD_COLS]),
        axis=1
    )

    print(f"\n[{label}] Statistiche R² di Pearson:")
    print(
        db_upd[["R2_prob", "R2_occ"]]
        .describe()
        .round(4)
    )

    print(f"\n[{label}] Righe senza R² calcolato:")
    print(
        db_upd[["R2_prob", "R2_occ"]]
        .isna()
        .sum()
    )

    db_upd.to_excel(output_path, index=False)

    print(f"\n[{label}] Salvato: {output_path}")

    return db_upd


# 1. Database con le attività aggregate
activities_upd = add_load_profiles_and_r2(
    input_path=db_activities_path,
    output_path=out_activities_path,
    label="Activities"
)

# 2. Database con le sotto-attività
subactivities_upd = add_load_profiles_and_r2(
    input_path=db_subactivities_path,
    output_path=out_subactivities_path,
    label="Subactivities"
)