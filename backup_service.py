# -*- coding: utf-8 -*-
"""
Windows-tjeneste som kjører backup.py i et intervall-definert loop.

Installasjon (kjør i PowerShell/Terminal som administrator):
    pip install pywin32
    python backup_service.py install
    python backup_service.py start

Administrasjon:
    python backup_service.py stop / start / restart
    python backup_service.py remove      (avinstaller)

Tjenesten heter "SupportsystemBackup" og logger både til Windows Event Log
og til backups\\backup.log.
"""
import os
import sys
import time
import servicemanager
import win32event
import win32service
import win32serviceutil

MAPPE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, MAPPE)  # sørg for at backup.py kan importeres

import backup  # noqa: E402


class SupportBackupTjeneste(win32serviceutil.ServiceFramework):
    _svc_name_ = "SupportsystemBackup"
    _svc_display_name_ = "Supportsystem Backup"
    _svc_description_ = (
        "Tar periodisk backup av Supportsystem.xlsx til lokal mappe og "
        "valgfri nettverksdisk, med automatisk rotasjon av gamle kopier. "
        "Innstillinger: backup_config.json."
    )

    def __init__(self, args):
        win32serviceutil.ServiceFramework.__init__(self, args)
        self.wait_handle = win32event.CreateEvent(None, 0, 0, None)
        self.running = True

    def SvcStop(self):
        self.ReportStatus(win32service.SERVICE_STOP_PENDING)
        self.running = False
        win32event.SetEvent(self.wait_handle)

    def SvcDoRun(self):
        servicemanager.LogMsg(
            servicemanager.EVENTLOG_INFORMATION_TYPE,
            servicemanager.PYS_SERVICE_STARTED,
            (self._svc_name_, None),
        )
        cfg = backup.last_config()
        intervall_sek = max(1, int(cfg["intervall_minutter"])) * 60
        backup.logg(
            f"Tjenesten startet. Intervall: {intervall_sek // 60} min.", cfg
        )
        while self.running:
            backup.kjor_backup(cfg)
            # Vent i inntil 'intervall' sekunder, men stopp umiddelbart ved SvcStop
            win32event.WaitForSingleObject(self.wait_handle, intervall_sek * 1000)
        servicemanager.LogMsg(
            servicemanager.EVENTLOG_INFORMATION_TYPE,
            servicemanager.PYS_SERVICE_STOPPED,
            (self._svc_name_, None),
        )


if __name__ == "__main__":
    if len(sys.argv) == 1:
        # Service Control Manager kaller uten argumenter
        servicemanager.Initialize()
        servicemanager.PrepareToHostSingle(SupportBackupTjeneste)
        servicemanager.StartServiceCtrlDispatcher()
    else:
        win32serviceutil.HandleCommandLine(SupportBackupTjeneste)
