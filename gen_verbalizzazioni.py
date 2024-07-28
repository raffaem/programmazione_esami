#!/home/raffaele/.pyenv/versions/3.12.2/bin/python
# Copyright 2024 Raffaele Mancuso
# SPDX-License-Identifier: GPL-2.0-or-later

import pandas as pd
from pathlib import Path
import csv
from datetime import datetime
import argparse


def get_val(df_verb_row, df_val, verb_col, data_sostenimento, df_rifiuti):
    """
    Processa una riga del DataFrame verbalizzazioni

    :param df_verb_row: Riga del file verbalizzazioni
    :type df_veb_row: pandas.Series
    :param df_val: DataFrame delle valutazioni
    :type df_val: pandas.DataFrame
    :param verb_col: Header della colonna del DataFrame valutazioni
                     che contiene le valutazioni da pubblicare su AlmaEsami
    :type verb_col: str
    :param dat_sostenimento: La data dell'esame
    :type data_sostenimento: datetime.date
    :param df_rifiuti: DataFrame del file rifiuti
    :type df_rifiuti: pandas.DataFrame

    """
    # Prende esito dal file valutazioni
    mat = df_verb_row["Matricola"]
    mask = df_val["Matricola"] == mat
    assert (sum(mask) == 1)
    df_val_sub = df_val[mask]
    assert (isinstance(df_val_sub, pd.DataFrame))
    assert (df_val_sub.shape[0] == 1)
    val = df_val_sub.iloc[0, :][verb_col]
    # Cerca nel dataframe rifiuti
    if "Email" not in df_val.columns:
        raise Exception("Il file delle valutazioni non contiene la colonna Email")
    email = df_val_sub.iloc[0, :]["Email"]
    if "Email" not in df_rifiuti.columns:
        raise Exception("Il file dei rifiuti non contiene la colonna Email")
    mask = df_rifiuti["Email"] == email
    if sum(mask) >= 2:
        raise Exception(f"L'email {email} ha compilato il form rifiuti più di una volta")
    if sum(mask) == 1:
        print(f"L'email {email} ha rifiutato la propria valutazione di {val}")
        val = "Rifiutato"
    # Memorizza esito nel file verbalizzazioni
    df_verb_row["Esito"] = val
    # Date must be in italian format
    # DD/MM/YYYY
    # otherwise it won't work
    df_verb_row["Data sostenimento"] = data_sostenimento.strftime("%d/%m/%Y")
    # row["Superata"] = "SI" if is_numeric else "NO"
    df_verb_row["Quesito 1"] = "Prova scritta"
    return df_verb_row


def proc_verb(infile, df_val, data_sostenimento, df_rifiuti):
    """
    Processa il file verbalizzazioni

    :param infile: Percorso al file verbalizzazioni
    :type infile: Path
    :param df_val: DataFrame delle valutazioni
    :type df_val: pandas.DataFrame
    :data_sostenimento: La data dell'esame
    :type data_sostenimento: datetime.datetime
    :param df_rifiuti: DataFrame al foglio rifiuti
    :type df_rifiuti: pandas.DataFrame

    """
    df_verb = pd.read_csv(
        infile,
        sep=";",
        index_col=False,
        dtype={"Matricola": str},
    )
    verb_col = infile.name[:-4]
    if verb_col not in df_val.columns:
        raise Exception(f"ERRORE: {verb_col} non è una colonna di df_val")
    df_verb_2 = df_verb.apply(
        lambda df_verb_row: get_val(
            df_verb_row,
            df_val,
            verb_col,
            data_sostenimento,
            df_rifiuti
        ),
        axis=1,
    )
    df_verb_2.to_csv(
        infile.stem+"_compilato.csv",
        index=False,
        sep=";",
        quoting=csv.QUOTE_ALL
    )


# MAIN

# Parse command line arguments
parser = argparse.ArgumentParser(
    prog='gen_verbalizzazioni',
    description='Genera le verbalizzazioni a partire dal foglio valutazioni'
)
parser.add_argument('--rifiuti',
                    help="Processa il foglio di rifiuti",
                    action='store_true',
                    default=False)
args = parser.parse_args()

# Elenca i file di "verbalizzazioni" da compilare
verb_files = Path(".").iterdir()
verb_files = [
    x
    for x in verb_files
    if x.stem.startswith("verbalizzazioni") and
    not x.stem.endswith("_compilato")
]
if len(verb_files) == 0:
    raise Exception("ERRORE: Non ci sono file di verbalizzazioni da compilare")

# Leggi file "valutazioni" - foglio "valutazioni"
val_file = Path("valutazioni.ods")
if not val_file.is_file():
    raise Exception(f"ERRORE: file {val_file} non esiste")
df_val = pd.read_excel(
    val_file,
    sheet_name="valutazioni",
    dtype={"Matricola": str}
)

# Leggi file "valutazioni" - foglio "var"
df_var = pd.read_excel(
    val_file,
    sheet_name="var"
)

# Prendi percorso a file rifiuti
df_rifiuti = None
if args.rifiuti:
    rifiuti_path = Path(df_var["file_rifiuti"][0])
    assert (rifiuti_path.is_file())
    df_rifiuti = pd.read_excel(rifiuti_path)

# Prendi data dell'esame dal nome della cartella genitore
data_sostenimento = Path(".").resolve().parent.name[:10]
data_sostenimento = datetime.fromisoformat(data_sostenimento).date()
print(f"data_sostenimento={data_sostenimento.isoformat()}")

# Prendi numero dell'esame dal nome della cartella genitore
n_app = Path(".").resolve().parent.name[-1]
n_app = int(n_app)
assert (n_app >= 1)
assert (n_app <= 6)
print(f"n_app={n_app}")

# Itera sul file di verbalizzazione e processali
for verb_file in verb_files:
    proc_verb(verb_file, df_val, data_sostenimento, df_rifiuti)

#
