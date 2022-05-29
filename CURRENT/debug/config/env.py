#!/usr/bin/env python
# -*- coding: utf-8 -*-

## =================================================================================================
## MODULES IMPORT
## =================================================================================================
## ------------
# [...
""" Importation des dlls natives de Python """
import os
import getpass
from collections import OrderedDict
# ...]

## =================================================================================================
## MODULE DESCRIPTION
## =================================================================================================
""" Environnement Module:
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
__comment__     = "Allocation de l'environnement"
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
def rooter(keyWord="Meetep",
           path=os.path.dirname(os.path.abspath(__file__))):
    """ Return tool's server root
    """
    ident = "serverRoot"
    while (os.path.basename(path) != keyWord):
        path = os.path.abspath(r"%s\.." % path)
    return path
# ...]

## =================================================================================================
## DECLARATION DE L'ENVIRONNEMENT
## =================================================================================================
## ------------
# [... 
""" Tool's maintainer """
GARANT           = OrderedDict()
GARANT["PRENOM"] = __maintainer__.split(" ")[0]
GARANT["NOM"]    = __maintainer__.split(" ")[-1]
GARANT["TEL"]    = "xxxx"
GARANT["EMAIL"]  = "%s.%s@xxxx" % (GARANT["PRENOM"].lower(), GARANT["NOM"].lower())
# ...]
## ------------
# [...
""" UI Windows configurations """
WINDOWS           = OrderedDict()
WINDOWS["TITLES"] = {"LOAD_INSTALL": u"[MEETEP] Chargement en cours (please wait ...)",
                     "PREFERENCES": u"[MEETEP] Edition des préférences",
                     "MAIN": u"Bienvenue sur l'application MEETEP"}
# ...]
## ------------
# [...
""" Configuration de l'environnement serveur """
TOOL_ENV          = OrderedDict()
TOOL_ENV["ROOT"]  = os.path.abspath(r"%s\.." % os.path.dirname(os.path.abspath(__file__)))
TOOL_ENV["ICONS"] = r"%s\icons" % TOOL_ENV["ROOT"]
TOOL_ENV["DOCS"]  = r"%s\docs" % TOOL_ENV["ROOT"]
# ...]
## ------------
# [...
""" Configuration de l'environnement utilisateur en local """
USER_ENV              = OrderedDict()
USER_ENV["ROOT"]      = r"D:\Users\%s\MEETEP" % getpass.getuser()
USER_ENV["VERSIONS"]  = r"%s\Versions" % USER_ENV["ROOT"]
USER_ENV["WKDIR"]     = r"%s\Workdir" % USER_ENV["ROOT"]
USER_ENV["PREF"]      = r"%s\Preferences" % USER_ENV["ROOT"]
# ...]