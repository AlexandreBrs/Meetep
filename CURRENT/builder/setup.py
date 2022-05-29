#!/usr/bin/env python
# -*- coding: utf-8 -*-

## =================================================================================================
## MODULES IMPORT
## =================================================================================================
## ------------
# [...
""" Importation des dlls natives de Python """
import sys 
import os 
import shutil
import glob 
import os.path
from distutils.core import setup
import py2exe 
# ...]

## =================================================================================================
## MODULE DESCRIPTION
## =================================================================================================
""" py2exe setup:
    Modules de compilation du program (encoding en .exe)
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
__comment__     = "Py2exe Setup"
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
def listdirectory(path):
    """ Fonctions de listing des fichiers
    """
    return filter(os.path.isfile, glob.glob(path + os.sep + "*"))
# ...]

## =================================================================================================
## SETUP
## =================================================================================================
TOOL_PATH = os.path.abspath(r"%s\.." % os.path.dirname(os.path.abspath(__file__)))
## ------------
# [...
""" Compilation des data files """
data_files = [("", filter(os.path.isfile, glob.glob(r"C:\Appl\Python\Anaconda2_py27\Library\bin" + os.sep + "*.dll"))),
              ("imageformats", listdirectory("C:\\Appl\\Python\\Anaconda2_py27\\Lib\\site-packages\\PyQt4\\plugins\\imageformats")),
              ("utils", listdirectory(r"%s\debug\utils" % TOOL_PATH)),
              ("templates", listdirectory(r"%s\debug\templates" % TOOL_PATH)),
              ("config", listdirectory(r"%s\debug\config" % TOOL_PATH)),
              ("gui", listdirectory(r"%s\debug\gui" % TOOL_PATH)),
              ("docs", listdirectory(r"%s\debug\docs" % TOOL_PATH)),
              ("icons", listdirectory(r"%s\debug\icons" % TOOL_PATH))]
# ...]
## ------------
# [...
""" Mise en place du setup """
setup(name = "MEETEP",
      version = "V01R00",
      description = "Easy Meeting Planner",
      author = "A. Brosse",
      license = "license GPL v3.0",
      url = "http://python.jpvweb.com",
      windows = [{"script": r"%s\debug\Meetep.py" % TOOL_PATH}],
      options = {"py2exe":{"includes":["sip", "numpy", "PyQt4.QtGui", "PyQt4.QtCore"],
                           "dist_dir":"released",
                           "bundle_files":1}
                           # "dll_excludes":["QtCore4.dll", "QtGui4.dll", "QtNetwork4.dll"]}
                },
      zipfile = None,
      # zipfile = "library.zip",
      data_files = data_files)
# ...]
## ------------
# [...
""" Effacement des anciens répertoires et copie du 'released new' """
try:
    shutil.rmtree(r"%s\Meetep_%s_released" % (TOOL_PATH, __version__))
except:
    pass
shutil.rmtree("build")
shutil.copytree("released", r"%s\Meetep_%s_released" % (TOOL_PATH, __version__), symlinks=True)
shutil.rmtree("released")