#!/home/raffaele/.pyenv/versions/3.12.2/bin/python
# Copyright 2024 Raffaele Mancuso
# SPDX-License-Identifier: GPL-2.0-or-later

import pandas as pd
from pathlib import Path
import csv
from datetime import datetime
import argparse
import shutil


def my_is_numeric(val: str):
    if isinstance(val, int):
        return True
    elif isinstance(val, str):
        return val.isnumeric() or val == "30L"
    else:
        raise Exception(f"Unmanaged type")


def get_val(
    df_verb_row,
    df_val,
    verb_col,
    data_sostenimento,
    n_app,
    df_rifiuti,
    df_sospesi_in,
    df_sospesi_out
):
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
    :param n_app: Numero dell'appello
    :type n_app: int
    :param df_rifiuti: DataFrame del file rifiuti
    :type df_rifiuti: pandas.DataFrame
    :param df_sospesi_in: DataFrame del file sospesi
    :type df_sospesi_in: pandas.DataFrame
    :param df_sospesi_out: DataFrame del file sospesi in uscita
    :type df_sospesi_out: pandas.DataFrame

    """
    assert (1 <= n_app <= 6)
    # Prende esito dal file valutazioni
    mat = df_verb_row["Matricola"]
    mask = df_val["Matricola"] == mat
    assert (sum(mask) == 1)
    df_val_row = df_val[mask].iloc[0]
    val = df_val_row[verb_col]
    # Cerca nel dataframe rifiuti
    if "Email" not in df_val.columns:
        raise Exception(
            "Il file delle valutazioni non contiene la colonna Email")
    email = df_val_row["Email"]
    if "Email" not in df_rifiuti.columns:
        raise Exception("Il file dei rifiuti non contiene la colonna Email")
    mask = df_rifiuti["Email"] == email
    if sum(mask) >= 2:
        raise Exception(f"L'email {email} ha compilato il form rifiuti "
                        "più di una volta")
    elif sum(mask) == 1:
        df_rifiuti_row = df_rifiuti[mask].iloc[0]
        opzione = df_rifiuti_row["Opzione"]
        opzione = int(opzione)
        assert (1 <= opzione <= 3)
        # Opzione 1: “Ho conseguito una valutazione ≥ 18. Verbalizza ‘Rifiutato’.” # noqa: E501
        if opzione == 1:
            if my_is_numeric(val):
                print(
                    f"L'email {email} ha rifiutato "
                    f"la propria valutazione di {val}")
                val = "Rifiutato"
            else:
                print(f"ERRORE: L'email {email} ha selezionato l'opzione 1 "
                      "ma non poteva farlo "
                      f"perchè la sua valutazione di {val} non è numerica")
        # Opzione 2: “Ho conseguito una valutazione ≥ 18
        #             in un appello antecedente all’appello n. 6.
        #             Verbalizza ‘Ritirato’ e sospendi valutazione.”
        elif opzione == 2:
            if not my_is_numeric(val):
                print(f"ERRORE: L'email {email} ha selezionato "
                      "l'opzione 2 ma non poteva farlo perchè "
                      f"la sua valutazione di {val} non è numerica")
            elif n_app >= 6:
                print(f"ERRORE: L'email {email} ha selezionato "
                      "l'opzione 2 ma non poteva farlo perchè "
                      f"siamo all'appello n. {n_app} >= 6.")
            else:
                df_sospesi_out.loc[email, "esito"] = val
                df_sospesi_out.loc[email,
                                   "data_sostenimento"] = data_sostenimento
                val = "Ritirato"
                print(f"L'email {email} ha sospeso "
                      f"la propria valutazione di {val}")
        # Opzione 3: “Mi sono ritirato in un appello antecedente all’appello n. 6. "
        #            "Mantieni valutazione in sospeso.”
        elif opzione == 3:
            presente = int(df_val_row["Presente"])
            assert (presente in [0, 1])
            ritirato = int(df_val_row["Ritirato"])
            assert (ritirato in [0, 1])
            if presente != 1:
                print(f"ERRORE: L'email {email} ha selezionato "
                      "l'opzione 3 ma non poteva farlo perchè "
                      f"non è presente.")
            elif ritirato != 1:
                print(f"ERRORE: L'email {email} ha selezionato "
                      "l'opzione 3 ma non poteva farlo perchè "
                      f"non si è ritirato.")
            elif n_app >= 6:
                print(f"ERRORE: L'email {email} ha selezionato "
                      "l'opzione 3 ma non poteva farlo perchè "
                      f"siamo all'appello n. {n_app} >= 6.")
            else:
                mask = df_sospesi_in.loc[email]
                if sum(mask) >= 2:
                    raise Exception(f"ERRORE: L'email {email} "
                                    "compare più volte nel file sospesi "
                                    "per qualche motivo")
                elif sum(mask) == 0:
                    raise Exception(f"ERRORE: L'email {email} "
                                    "ha selezionato l'opzione 3 ma non è nel "
                                    "file sospesi")
                else:
                    df_sospesi_in_row = df_sospesi_in[mask].iloc[0]
                    for col in list(df_sospesi_in_row.index):
                        df_sospesi_out.loc[email, col] = df_sospesi_in_row[col]
        # Opzione 4: “Mi sono ritirato. Verbalizza valutazione in sospeso.”
        elif opzione == 4:
            presente = int(df_val_row["Presente"])
            assert (presente in [0, 1])
            ritirato = int(df_val_row["Ritirato"])
            assert (ritirato in [0, 1])
            if presente != 1:
                print(f"ERRORE: L'email {email} ha selezionato "
                      "l'opzione 4 ma non poteva farlo perchè "
                      f"non è presente.")
            elif ritirato != 1:
                print(f"ERRORE: L'email {email} ha selezionato "
                      "l'opzione 4 ma non poteva farlo perchè "
                      "non si è ritirato.")
            else:
                mask = df_sospesi_in.loc[email]
                if sum(mask) >= 2:
                    raise Exception(f"ERRORE: L'email {email} "
                                    "compare più volte nel file sospesi "
                                    "per qualche motivo")
                elif sum(mask) == 0:
                    raise Exception(f"ERRORE: L'email {email} "
                                    "ha selezionato l'opzione 4 ma non è nel "
                                    "file sospesi")
                else:
                    val = df_sospesi_in.loc[email, "esito"]
                    print(f"INFO: L'email {email} ha verbalizzato "
                          f"la propria valutazione in sospeso di {val}")
        # Altre opzioni
        else:
            raise Exception(f"Opzione {opzione} non implementata")

    # Memorizza esito nel file verbalizzazioni
    df_verb_row["Esito"] = val
    # Date must be in italian format
    # DD/MM/YYYY
    # otherwise it won't work
    df_verb_row["Data sostenimento"] = data_sostenimento.strftime("%d/%m/%Y")
    # row["Superata"] = "SI" if is_numeric else "NO"
    df_verb_row["Quesito 1"] = "Prova scritta"
    return df_verb_row


def proc_verb(
    fp_verb,
    df_val,
    data_sostenimento,
    n_app,
    df_rifiuti,
    df_sospesi_in,
    df_sospesi_out
):
    """
    Processa il file verbalizzazioni

    :param fp_verb: Percorso al file verbalizzazioni
    :type fp_verb: Path
    :param df_val: DataFrame delle valutazioni
    :type df_val: pandas.DataFrame
    :param data_sostenimento: La data dell'esame
    :type data_sostenimento: datetime.datetime
    :param n_app: Numero dell'appello
    :type n_app: int
    :param df_rifiuti: DataFrame al foglio rifiuti
    :type df_rifiuti: pandas.DataFrame
    :param df_sospesi_in: DataFrame del file sospesi
    :type df_sospesi_in: pandas.DataFrame
    :param df_sospesi_out: DataFrame del file sospesi in uscita
    :type df_sospesi_out: pandas.DataFrame

    """
    df_verb = pd.read_csv(
        fp_verb,
        sep=";",
        index_col=False,
        dtype={"Matricola": str},
    )
    verb_col = fp_verb.name[:-4]
    if verb_col not in df_val.columns:
        raise Exception(f"ERRORE: {verb_col} non è una colonna di df_val")
    df_verb_2 = df_verb.apply(
        lambda df_verb_row: get_val(
            df_verb_row,
            df_val,
            verb_col,
            data_sostenimento,
            n_app,
            df_rifiuti,
            df_sospesi_in
        ),
        axis=1,
    )
    df_verb_2.to_csv(
        fp_verb.stem+"_compilato.csv",
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
    if "file_rifiuti" not in df_var:
        print("ERRORE: `file_rifiuti` non è "
              "una colonna del foglio `var` del file valutazioni")
    rifiuti_path = Path(df_var["file_rifiuti"][0])
    assert (rifiuti_path.is_file())
    rifiuti_local_path = Path("rifiuti.xlsx")
    if rifiuti_local_path.is_file():
        print("ATTENZIONE: Il file locale dei rifiuti esiste già. Uso quello")
    else:
        shutil.copy(rifiuti_path, rifiuti_local_path)
    df_rifiuti = pd.read_excel(rifiuti_local_path)

# Apri file dei sospesi


def gen_df_sospesi():
    df_sospesi = pd.DataFrame(
        {
            "esito": [],
            "data_sostenimento": [],
        },
        index=pd.Series([], name="Email",),
        dtype=str
    )
    return df_sospesi


df_sospesi_in = None
df_sospesi_out = None
if args.rifiuti:
    filefp = Path("sospesi_in.csv")
    if filefp.is_file():
        df_sospesi_in = pd.read_csv(filefp, index_col="Email")
    else:
        print("Il file dei sospesi non esiste. Lo creo.")
        df_sospesi_in = gen_df_sospesi()
    df_sospesi_out = gen_df_sospesi()

# Prendi data dell'esame dal nome della cartella genitore
data_sostenimento = Path(".").resolve().parent.name[:10]
data_sostenimento = datetime.fromisoformat(data_sostenimento).date()
print(f"data_sostenimento={data_sostenimento.isoformat()}")

# Prendi numero dell'esame dal nome della cartella genitore
n_app = Path(".").resolve().parent.name[-1]
n_app = int(n_app)
assert (1 <= n_app <= 6)
print(f"n_app={n_app}")

# Itera sul file di verbalizzazione e processali
for verb_file in verb_files:
    proc_verb(
        verb_file,
        df_val,
        data_sostenimento,
        n_app,
        df_rifiuti,
        df_sospesi_in,
        df_sospesi_out,
    )

df_sospesi_out.to_csv("sospesi_out.csv")

#
