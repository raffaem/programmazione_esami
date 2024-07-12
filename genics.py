#!/usr/bin/env python3
# Copyright 2023 Raffaele Mancuso
# SPDX-License-Identifier: GPL-2.0-or-later

import pandas as pd
from ics import Calendar, Event
from pathlib import Path

from datetime import datetime
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


def dateonly(s):
    s = str(s)
    return s.split(" ")[0]


def timestamp(s):
    s = str(s)
    # Remove seconds
    return s[:-3]


df = pd.read_excel("programmazione_esami.ods",  sheet_name="calgen")
df

dfdata = pd.read_excel("programmazione_esami.ods",  sheet_name="data")
course = dfdata.loc[0, "sigla_corso"]
course

mymd = open(outdir/f"{course}.md", "w", encoding="utf8")

c = Calendar()

for index, row in df.iterrows():
    #
    n = row["n"]
    print(f"Processing exam {n}")
    mymd.write(f"**Appello n. {n}** \\\n")
    #
    e = Event()
    e.name = f"[{course}] Appello {n} - Inizio Iscrizioni"
    e.begin = rome2utc(row["iscrizioni_inizio"])
    e.make_all_day()
    c.events.add(e)
    mymd.write(f"Inizio iscrizioni: {
               timestamp(row['iscrizioni_inizio'])} \\\n")
    #
    e = Event()
    e.name = f"[{course}] Appello {n} - Fine Iscrizioni"
    e.begin = rome2utc(row["iscrizioni_fine"])
    e.make_all_day()
    c.events.add(e)
    mymd.write(f"Fine iscrizioni: {timestamp(row['iscrizioni_fine'])} \\\n")
    #
    e = Event()
    e.name = f"[{course}] Appello {n} - Esame"
    e.begin = rome2utc(row["esame_inizio"])
    e.end = rome2utc(row["esame_fine"])
    e.location = str(row["esame_aula"])
    c.events.add(e)
    mymd.write(f"Esame: {timestamp(row['esame_inizio'])} {
               str(row['esame_aula'])} \\\n")
    #
    e = Event()
    e.name = f"[{course}] Appello {n} - Pubblicazione"
    e.begin = rome2utc(row["pubblicazione"])
    e.make_all_day()
    c.events.add(e)
    mymd.write(f"Pubblicazione: {dateonly(row['pubblicazione'])} \\\n")
    #
    e = Event()
    e.name = f"[{course}] Appello {n} - Visione"
    e.begin = rome2utc(row["visione_inizio"])
    e.end = rome2utc(row["visione_fine"])
    e.location = str(row["visione_link"])
    c.events.add(e)
    mymd.write(f"Visione: {timestamp(row['visione_inizio'])} "
               f"[link]({str(row['visione_link'])}) \\\n")
    #
    e = Event()
    e.name = f"[{course}] Appello {n} - Verbalizzazione"
    e.begin = rome2utc(row["verbalizzazione"])
    e.make_all_day()
    c.events.add(e)
    mymd.write(f"Verbalizzazione: {dateonly(row['verbalizzazione'])} \n")
    #
    mymd.write("\n")

print(c.events)
with open(outdir/f'{course}.ics', 'w') as my_file:
    my_file.writelines(c.serialize_iter())

mymd.close()
