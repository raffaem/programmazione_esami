#!/home/raffaele/.pyenv/versions/3.12.2/bin/python
# Copyright 2024 Raffaele Mancuso
# SPDX-License-Identifier: GPL-2.0-or-later

import pandas as pd
from pathlib import Path
import csv
from datetime import datetime


def get_val(row, df_val, verb_col, dat_sostenimento):
    mat = row["Matricola"]
    mask = df_val["Matricola"] == mat
    assert (sum(mask) == 1)
    df_val_sub = df_val[mask]
    assert (isinstance(df_val_sub, pd.DataFrame))
    assert (df_val_sub.shape[0] == 1)
    val = df_val_sub.iloc[0, :][verb_col]
    is_numeric = isinstance(val, int) or \
        (isinstance(val, str) and (val.isnumeric() or val == "30L"))
    row["Esito"] = val
    # Date must be in italian format
    # DD/MM/YYYY
    # otherwise it won't work
    row["Data sostenimento"] = data_sostenimento.date().strftime("%d/%m/%Y")
    row["Superata"] = "SI" if is_numeric else "NO"
    return row


def proc_verb(infile, df_val, data_sostenimento):
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
        lambda row: get_val(
            row,
            df_val,
            verb_col,
            data_sostenimento
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

verb_files = Path(".").iterdir()
verb_files = [
    x
    for x in verb_files
    if x.stem.startswith("verbalizzazioni") and
    not x.stem.endswith("_compilato")
]

val_file = Path("valutazioni.ods")
if not val_file.is_file():
    raise Exception(f"ERRORE: file {val_file} non esiste")
df_val = pd.read_excel(
    val_file,
    sheet_name="valutazioni",
    dtype={"Matricola": str},
)

data_sostenimento = Path(".").resolve().parent.name[:10]
data_sostenimento = datetime.fromisoformat(data_sostenimento)

for verb_file in verb_files:
    proc_verb(verb_file, df_val, data_sostenimento)

#
