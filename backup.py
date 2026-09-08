# -*- coding: utf-8 -*-
"""
Backup-rutine for Supportsystem-filen (.xlsx/.xlsm/.ods).

Kopierer Excel-filen til en lokal backupmappe og valgfritt en nettverksdisk,
med tidsstempel i filnavnet. Beholder de N nyeste kopiene per mappe og
roterer eldre bort. Alle hendelser logges til backup.log.

Kan kjøres:
  - Én gang (manuelt/test):     python backup.py enkelt
  - I løkke med intervall:      python backup.py loop
  - Som Windows-tjeneste:       via backup_service.py
  - Som planlagt oppgave:       installer med installer_oppgave.bat

Innstillinger hentes fra backup_config.json i samme mappe.
"""
import json
import os
import shutil
import sys
import time
from datetime import datetime

MAPPE = os.path.dirname(os.path.abspath(__file__))
CONFIG_FIL = os.path.join(MAPPE, "backup_config.json")

STANDARD_CONFIG = {
    # Standard kildefil er Supportsystem.xlsx. Kan også peke på Supportsystem.xlsm
    # (makro) eller Supportsystem.ods (LibreOffice/OpenOffice) – backupen bevarer
    # alltid kildefilens egen endelse.
    "kildefil": "Supportsystem.xlsx",
    "lokal_backupmappe": "backups",
    "nettverk_backupmappe": "",          # f.eks. "\\\\server\\felles\\backup" – tom = deaktivert
    "intervall_minutter": 60,
    "behold_antall": 10,
    "loggfil": "backups\\backup.log",
}


def last_config():
    """Leser config, oppretter standard ved behov, fyller ut manglende nøkler."""
    if not os.path.exists(CONFIG_FIL):
        with open(CONFIG_FIL, "w", encoding="utf-8") as f:
            json.dump(STANDARD_CONFIG, f, indent=2, ensure_ascii=False)
    with open(CONFIG_FIL, "r", encoding="utf-8-sig") as f:
        cfg = json.load(f)
    for nokkel, verdi in STANDARD_CONFIG.items():
        cfg.setdefault(nokkel, verdi)
    return cfg


def logg(melding, cfg=None):
    """Skriver tidsstemplet linje til loggfil og konsoll/Event Log."""
    linje = f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {melding}"
    print(linje)
    try:
        loggfil = (cfg or {}).get("loggfil") or STANDARD_CONFIG["loggfil"]
        loggsti = loggfil if os.path.isabs(loggfil) else os.path.join(MAPPE, loggfil)
        os.makedirs(os.path.dirname(loggsti), exist_ok=True)
        with open(loggsti, "a", encoding="utf-8") as f:
            f.write(linje + "\n")
    except OSError:
        pass  # logging skal aldri knekke backupen


def absolutt(sti):
    return sti if os.path.isabs(sti) else os.path.join(MAPPE, sti)


def kjor_backup(cfg):
    """Utfører én backup-runde. Returnerer True hvis minst én kopi ble laget."""
    kilde = absolutt(cfg["kildefil"])
    if not os.path.exists(kilde):
        logg(f"FEIL: Finner ikke kildefilen: {kilde}", cfg)
        return False

    stempel = datetime.now().strftime("%Y%m%d_%H%M%S")
    navn_grunn = os.path.splitext(os.path.basename(kilde))[0]
    ekst = os.path.splitext(os.path.basename(kilde))[1]  # .xlsx / .xlsm / .ods
    backupnavn = f"{navn_grunn}_{stempel}{ekst}"

    mal = [absolutt(cfg["lokal_backupmappe"])]
    if cfg.get("nettverk_backupmappe"):
        mal.append(cfg["nettverk_backupmappe"])

    minst_en_ok = False
    for mappe in mal:
        try:
            os.makedirs(mappe, exist_ok=True)
            destinasjon = os.path.join(mappe, backupnavn)
            shutil.copy2(kilde, destinasjon)
            logg(f"OK: Backup laget -> {destinasjon}", cfg)
            minst_en_ok = True
        except OSError as e:
            logg(f"FEIL ved backup til {mappe}: {e}", cfg)
        roter(mappe, navn_grunn, int(cfg["behold_antall"]), cfg, ekst)
    return minst_en_ok


def roter(mappe, navn_grunn, behold, cfg, ekst):
    """Sletter eldste kopier slik at kun 'behold' nyeste gjenstår.

    'ekst' er kildefilens endelse (f.eks. ".xlsm") – kun kopier med samme
    endelse roteres, så .xlsx- og .xlsm-backuper ikke blandes.
    """
    try:
        kopier = sorted(
            (f for f in os.listdir(mappe)
             if f.startswith(navn_grunn + "_")
             and os.path.splitext(f)[1].lower() == ekst.lower()),
            key=lambda f: os.path.getmtime(os.path.join(mappe, f)),
        )
    except OSError:
        return
    for gammel in kopier[:-behold] if behold > 0 else kopier:
        try:
            os.remove(os.path.join(mappe, gammel))
            logg(f"Rotert bort gammel kopi: {gammel}", cfg)
        except OSError as e:
            logg(f"FEIL ved sletting av {gammel}: {e}", cfg)


def loop(cfg):
    intervall = max(1, int(cfg["intervall_minutter"])) * 60
    logg(f"Backup-loop startet. Intervall: {intervall // 60} min. "
         f"Beholder: {cfg['behold_antall']} kopier.", cfg)
    while True:
        kjor_backup(cfg)
        time.sleep(intervall)


if __name__ == "__main__":
    konfig = last_config()
    modus = sys.argv[1].lower() if len(sys.argv) > 1 else "enkelt"
    if modus == "loop":
        loop(konfig)
    else:
        sys.exit(0 if kjor_backup(konfig) else 1)
