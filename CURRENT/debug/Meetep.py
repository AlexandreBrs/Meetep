#!/usr/bin/env python
# -*- coding: utf-8 -*-

## =================================================================================================
## MODULES IMPORT
## =================================================================================================
## ------------
""" Importation des dlls natives de Python """
import os
import sys
import getpass
from PyQt4 import QtGui, QtCore
import win32api, win32net, win32netcon
import win32com.client as win32Module
# ...]
## ------------
# [...
""" Importation des modules internes """
shell = win32Module.Dispatch("WScript.Shell")
realPath = r"D:\Users\%s\Desktop\Meetep.lnk" % getpass.getuser()
if os.path.exists(realPath):
    shortcut = shell.CreateShortCut(realPath)
    sys.path.append(os.path.dirname(r"%s" % shortcut.Targetpath))
else:
    sys.path.append(r"%s" % os.getcwd())
from config.env import USER_ENV
from utils.ulib import createShorcut, networkConnection
from gui.wdwlib import preloadWindow, mainWindow
# ...]

## =================================================================================================
## MODULE DESCRIPTION
## =================================================================================================
""" Program Module:
    Version initiale du module - No changes
"""

## =================================================================================================
## HEADER
## =================================================================================================
# [...
__author__      = "Alexandre Brosse"
__copyright__   = "Copyright 2017 (C)"
__license__     = "None"
__maintainer__  = "Alexandre Brosse"
__status__      = "Production/Oper"
__date__        = "Mon, 21/08/2017"
__description__ = "Easy Meeting Planner"
__comment__     = "Main source"
__version__     = "V01R00"
# ...]

## =================================================================================================
## MODULE DOCUMENTATION
## =================================================================================================

## =================================================================================================
## MAIN PROGRAM
## =================================================================================================
# [...
if __name__ == "__main__":
    """ Main Program
    """
    ## | Appel de la class GUI application |
    app = QtGui.QApplication(sys.argv)
    app.setStyle(QtGui.QStyleFactory.create("Cleanlooks"))
    ## | Création d'un raccourci |
    createShorcut(sName="Meetep",
                  sPath=r"D:\Users\%s\Desktop" % getpass.getuser(),
                  targetPath=r"%s\Meetep.exe" % os.path.dirname(os.path.abspath(sys.argv[0])),
                  hotKey="ctrl+alt+G")
    ## | Internet connection checking |
    networkConnection(domain="xxxx")
    ## | Création de l'environnement en local |
    for key in list(USER_ENV.keys()):
        if not os.path.isdir(USER_ENV[key]):
            os.makedirs(USER_ENV[key], mode=0777)
    ## | Appel du COM Outlook |
    oAppl = win32Module.Dispatch("Outlook.Application")
    session = oAppl.GetNamespace("MAPI")
    ## | Pré-chargement des salles |
    preLoading = preloadWindow(session,
                               lenght=int(400),
                               hight=int(100),
                               timeStep=int(15),
                               yearsGap=float(1./12.),
                               parent=None)
    ## | UI de recherche |
    mainWindow(session,
               olRooms=preLoading.roomsDico,
               olDateList=preLoading.dateCounter,
               timeStep=int(15),
               parent=None)
    ## | Arrêt de l'application et du processus Python |
    sys.exit(app.exec_())
# ...]