#!/usr/bin/env python
# -*- coding: utf-8 -*-

## =================================================================================================
## MODULES IMPORT
## =================================================================================================
## ------------
# [...
""" Importation des dlls natives de Python """
import os
import sys
import getpass
import time
import datetime
from PyQt4 import QtGui, QtCore
import win32api, win32net, win32netcon
import win32com.client as win32Module
# ...]
## ------------
# [...
""" Importation des modules internes """
from config.env import TOOL_ENV
from gui.wlib import criticalBox
# ...]

## =================================================================================================
## MODULE DESCRIPTION
## =================================================================================================
""" Utilities Library Module:
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
__comment__     = "Fonctions annexes"
__version__     = "V01R00"
# ...]

## =================================================================================================
## MODULE DOCUMENTATION
## =================================================================================================

## =================================================================================================
## METHODS AND CLASSES
## =================================================================================================
## ------------
# [...
def createShorcut(sName="Meetep",
                  sPath=r"D:\Users\%s\Desktop" % getpass.getuser(),
                  targetPath=r"%s\Meetep.exe" % TOOL_ENV["ROOT"],
                  hotKey="ctrl+alt+G"):
    """ Creation d'un raccourci bureau
        @param: sName = Nom du raccourci
        @param: sPath = Chemin d'ecriture du raccourci
        @param: targetPath = Lien cible du raccourci
        @param: hotKey = Raccourci clavier
    """
    ident = "createShorcut"
    shell = win32Module.Dispatch("WScript.Shell")
    if not os.path.isdir(sPath):
        os.makedirs(sPath)
    shortcut = shell.CreateShortCut(r"%s\%s.lnk" % (sPath, sName))
    shortcut.Targetpath = targetPath
    shortcut.Description = "Easy Meeting Planner"
    shortcut.HotKey = hotKey
    shortcut.save()
# ...]
## ------------
# [...
def networkConnection(domain="SNM"):
    """ Verifie la connexion reseau
        @param: domain = Nom du domaine reseau
    """
    ident = "networkConnection"
    try:
        win32net.NetGetDCName(None, domain)
    except Exception as e:
        criticalBox(wdwTitle="Program Error - Network Connection Failed",
                    wdwIcon="lock.png",
                    Txt=u"Aucune connection internet ne peut être établie.\
                          \nMerci de vous connecter via wifi ou câble éthernet avant de relancer l'application.",
                    stdButtons=QtGui.QMessageBox.Ok)
        sys.exit()
# ...]
## ------------
# [...
def getTimeDelta(timeStep=int(15),
                 index=int(0)):
    """ Get seconds from now
        @param: timeStep = Time step (en minutes)
        @param: index = Index de la liste
    """
    ident = "getTimeDelta"
    secondsFromToday = datetime.timedelta(hours=datetime.datetime.now().hour,
                                          minutes=datetime.datetime.now().minute,
                                          seconds=datetime.datetime.now().second,
                                          microseconds=datetime.datetime.now().microsecond).total_seconds()
    return index * timeStep * 60 - secondsFromToday
# ...]