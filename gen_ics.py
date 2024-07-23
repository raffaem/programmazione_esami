#!/usr/bin/env python3
# Copyright 2023 Raffaele Mancuso
# SPDX-License-Identifier: GPL-2.0-or-later

import pandas as pd
from ics import Calendar, Event
from pathlib import Path

from datetime import datetime, date
from zoneinfo import ZoneInfo

outdir = Path("./out")
outdir.mkdir(exist_ok=True)


def rome2utc(intime):
    # Get timezone we're trying to convert from
    local_tz = ZoneInfo("Europe/Rome")
    # UTC timezone
    utc_tz = ZoneInfo("UTC")
    #
    intime = str(intime)
    dt = datetime.strptime(intime, "%Y-%m-%d %H:%M:%S")
    dt = dt.replace(tzinfo=local_tz)
    dt_utc = dt.astimezone(utc_tz)
    return dt_utc


def timestamp(s):
    s = str(s)
    # Remove seconds
    return s[:-3]


def checkweekend(row):
    for col, item in row.items():
        if col in ["n", "esame_aula"]:
            continue
        item = item.date()
        #  date.isoweekday()
        # Return the day of the week as an integer, where Monday is 1 and Sunday is 7.
        weekday = item.isoweekday()
        if weekday in [6, 7]:
            print(f"\t\tERROR: {col} has a date of "
                  f"{item} which falls on a {item.strftime('%A')}!")


def procfile(infile):
    print(f"Processing {infile}")
    df = pd.read_excel(infile, sheet_name="5_final")
    #
    dfdata = pd.read_excel(infile, sheet_name="1_data")
    course = dfdata.loc[0, "sigla_corso"]
    course
    #
    mymd = open(outdir/f"{course}.md", "w", encoding="utf8")
    #
    c = Calendar()
    #
    for index, row in df.iterrows():
        #
        n = row["n"]
        print(f"\tProcessing exam {n}")
        checkweekend(row)
        # mymd.write(f"**Appello n. {n}** \n\n")
        table = pd.DataFrame()
        #
        # Inizio iscrizioni
        # e = Event()
        # e.name = f"[{course}] Appello {n} - Inizio Iscrizioni"
        # d = row["iscrizioni_inizio"]
        # e.begin = rome2utc(d)
        # e.make_all_day()
        # c.events.add(e)
        # table["Inizio iscrizioni"] = [timestamp(d)]
        #
        # Fine iscrizioni
        # e = Event()
        # e.name = f"[{course}] Appello {n} - Fine Iscrizioni"
        # d = row["iscrizioni_fine"]
        # e.begin = rome2utc(d)
        # e.make_all_day()
        # c.events.add(e)
        # table["Fine iscrizioni"] = [timestamp(d)]
        #
        # Iscrizioni
        start = row["iscrizioni_inizio"].date()
        e = Event()
        e.name = f"[{course}] Appello {n} - Inizio Iscrizioni"
        e.begin = start
        e.make_all_day()
        c.events.add(e)
        end = row["iscrizioni_fine"].date()
        e = Event()
        e.name = f"[{course}] Appello {n} - Fine Iscrizioni"
        e.begin = end
        e.make_all_day()
        c.events.add(e)
        table["Iscrizioni"] = [f"Dal {start.isoformat()} al {end.isoformat()}"]
        #
        # Esame
        e = Event()
        e.name = f"[{course}] Appello {n} - Esame"
        start = row["esame_inizio"]
        end = row["esame_fine"]
        room = row["esame_aula"]
        e.begin = rome2utc(start)
        e.end = rome2utc(end)
        e.location = str(room)
        c.events.add(e)
        table["Esame"] = [timestamp(start) + " " + room]
        #
        # Pubblicazione
        e = Event()
        e.name = f"[{course}] Appello {n} - Pubblicazione"
        d = row["pubblicazione"]
        d = d.date()
        e.begin = d
        e.make_all_day()
        c.events.add(e)
        table["Pubblicazione"] = [d.isoformat()]
        #
        # Visione
        e = Event()
        e.name = f"[{course}] Appello {n} - Visione"
        start = row["visione_inizio"]
        end = row["visione_fine"]
        e.begin = rome2utc(start)
        e.end = rome2utc(end)
        c.events.add(e)
        table["Visione"] = [timestamp(start)]
        #
        # Verbalizzazione
        e = Event()
        e.name = f"[{course}] Appello {n} - Verbalizzazione"
        d = row["verbalizzazione"]
        d = d.date()
        e.begin = d
        e.make_all_day()
        c.events.add(e)
        table["Verbalizzazione"] = [d.isoformat()]
        #
        # Write table
        table = table.transpose()
        table.index.name = f"**Appello n. {n}**"
        table.columns = [""]
        mymd.write(table.to_markdown())
        mymd.write("\n\n")
    #
    # print(c.events)
    with open(outdir/f'{course}.ics', 'w') as my_file:
        my_file.writelines(c.serialize_iter())
    #
    mymd.close()


infile = Path("./in/programmazione_esami_EOA.ods")
procfile(infile)
infile = Path("./in/programmazione_esami_TESADA.ods")
procfile(infile)
