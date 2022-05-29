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
import __builtin__
import time
import datetime
from collections import OrderedDict
from PyQt4 import QtGui, QtCore
import win32com.client as win32Module
# ...]
## ------------
# [...
""" Importation des modules internes """
from config.env import WINDOWS, TOOL_ENV, USER_ENV
import gui.wlib
from utils.ulib import getTimeDelta
# ...]
## ------------
# [...
""" Variables Outlook """
OUTLOOK_APPOINTMENT_ITEM = 1
OUTLOOK_MEETING = 1
OUTLOOK_ORGANIZER = 0
OUTLOOK_OPTIONAL_ATTENDEE = 2
OUTLOOK_RESOURCE_ATTENDEE = 3
olRecursDaily = 0
olRecursWeekly = 1
olRecursMonthly = 2
olRecursMonthNth = 3
olRecursYearly = 5
olRecursYearNth = 6
# ...]

## =================================================================================================
## MODULE DESCRIPTION
## =================================================================================================
""" UIWindows Library Module:
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
__comment__     = "Instanciation des dialogWindows"
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
class preloadWindow(QtGui.QProgressDialog):
    """ Window de chargement de l'application
        @param: session = Session Outlook connectee
        @param: lenght = Largeur de la fenetre
        @param: hight = Hauteur de la fenetre
        @param: timeStep = Time Step
        @param: yearsGap = Plage pour la peridiodicite
    """
    ident = "loadWindow"
    roomsDico = {"FBINFOS": [],
                 "LISTE": [],
                 "CAPACITY": []}
    dateCounter = []
    # ____________________________________________
    def __init__(self,
                 session,
                 lenght=int(400),
                 hight=int(100),
                 timeStep=int(15),
                 yearsGap=int(2),
                 parent=None):
        """ Initialisation de la classe
        """
        QtGui.QProgressDialog.__init__(self)
        self.session = session
        self.timeStep = timeStep
        self.yearsGap = yearsGap
        QtGui.QToolTip.setFont(QtGui.QFont("SansSerif", int(10)))
        self.setFixedSize(lenght, hight)
        self.setWindowTitle(WINDOWS["TITLES"]["LOAD_INSTALL"])
        self.setWindowIcon(QtGui.QIcon(r"%s\outlook.png" % TOOL_ENV["ICONS"]))
        self.displayStatus()
    # ____________________________________________
    def displayStatus(self):
        """ Affichage de la fenetre
        """
        ## | Chargement de l'annuaire des salles sous Outlook | 
        self.carnet = self.session.AddressLists.Item("All Rooms")
        ## | Mise en place du label |
        self.statusLabel = QtGui.QLabel()
        self.statusLabel.setAlignment(QtCore.Qt.AlignLeft)
        self.setLabel(self.statusLabel)
        ## | Mise en place de la barre de progression |
        self.setRange(int(0), self.carnet.AddressEntries.Count)
        self.value = int(0) 
        self.center()
        self.open()
        self.startLoading()
    # ____________________________________________
    def startLoading(self):
        """ Suivi du chargement
        """
        ## | Définition de la date butoire (from today) |
        endDate = datetime.datetime.now() + datetime.timedelta(weeks=52*self.yearsGap)
        ## | Itérations pour récupérer les infos de disponibilité et de capacité des salles |
        for i in range(self.carnet.AddressEntries.Count):
            if (self.wasCanceled()):
                gui.wlib.infoBox(wdwTitle="Information - Annulation du chargement",
                                 wdwIcon="outlook.png",
                                 Txt=u"Arrêt de l'application suite à annulation du chargement.",
                                 stdButtons=QtGui.QMessageBox.Ok)
                sys.exit()
            else:
                room = self.carnet.AddressEntries.Item(i+1)
                ## | Filtre sur la localisation @Villaroche |
                if "xxxx" in room.GetExchangeUser().OfficeLocation:
                    counter = int(0)
                    currentDate = datetime.datetime.now()
                    while currentDate <= endDate:
                        try:
                            recip = self.session.CreateRecipient("%s" % room.Name.split(" (XXXX)")[0]) 
                            myFBInfo = recip.FreeBusy(currentDate,
                                                    self.timeStep,
                                                    True)
                        except Exception as e:
                            break
                        else: 
                            try:        
                                self.roomsDico["CAPACITY"].append(int(room.GetExchangeUser().JobTitle.split(" ")[1]))
                            except Exception as e:
                                self.roomsDico["CAPACITY"].append("#N/A")
                            finally: 
                                if counter == int(0):
                                    self.roomsDico["LISTE"].append(room.Name)
                                    self.roomsDico["FBINFOS"].append([myFBInfo])
                                    self.dateCounter.append([])
                                elif counter != int(0):
                                    self.roomsDico["FBINFOS"][-1].append(myFBInfo)
                                self.dateCounter[-1].append([])
                                for id in range(len(myFBInfo)):
                                    self.dateCounter[-1][-1].append(currentDate + datetime.timedelta(seconds=getTimeDelta(timeStep=self.timeStep,
                                                                                                                          index=id)))
                                currentDate += datetime.timedelta(seconds=getTimeDelta(timeStep=self.timeStep,
                                                                                       index=len(myFBInfo)))
                                counter += 1
                self.setValue(i)
                self.setLabelText(u"FreeBusy process en cours ...\
                                    \nAnalyse de la salle: %s" % room.Name)
        self.close()
    # ____________________________________________
    def center(self):
        """ Recentrage de l'interface
        """
        qr = self.frameGeometry()
        cp = QtGui.QDesktopWidget().availableGeometry().center()
        qr.moveCenter(cp)
        self.move(qr.topLeft())
    # ____________________________________________
    def closeEvent(self,
                   event=None):
        """ Fermeture de la window
        """
        QtCore.QCoreApplication.instance().quit()
# ...]
## ------------
# [...
class editPreferences(QtGui.QDialog):
    """ Edition de preferences
        @param: lenght = Largeur de la fenetre
        @param: hight = Hauteur de la fenetre
    """
    ident = "editPreferences"
    Dico = {"REQUIRED_PATH": "",
            "OPTIONAL_PATH": "",
            "ROOMS_PATH": ""}
    dayTime = {"MIN": int(8*3600),
               "MAX": int(19*3600)}
    PREFENCES_FILEPATH = r"%s\Preferences_%s.py" % (USER_ENV["PREF"], __version__)
    # ____________________________________________
    def __init__(self,
                 show=True,
                 lenght=int(300),
                 hight=int(150),
                 parent=None):
        """ Initialisation de la classe
        """
        QtGui.QDialog.__init__(self)
        QtGui.QToolTip.setFont(QtGui.QFont("SansSerif", int(10)))
        # self.setFixedSize(lenght, hight)
        self.setWindowTitle(WINDOWS["TITLES"]["PREFERENCES"])
        self.setWindowIcon(QtGui.QIcon(r"%s\preferences.jpg" % TOOL_ENV["ICONS"]))
        if os.path.isfile(self.PREFENCES_FILEPATH):
            sys.path.append(os.path.dirname(self.PREFENCES_FILEPATH))
            prefData = __builtin__.__import__(os.path.basename(os.path.splitext(str(self.PREFENCES_FILEPATH))[0]),
                                              globals(),
                                              locals(),
                                              [],
                                              -1)
            if os.path.isfile(prefData.rAttendeePath):
                self.Dico["REQUIRED_PATH"] = prefData.rAttendeePath
            elif not os.path.isfile(prefData.rAttendeePath):
                self.Dico["REQUIRED_PATH"] = ""
            if os.path.isfile(prefData.oAttendeePath):
                self.Dico["OPTIONAL_PATH"] = prefData.oAttendeePath
            elif not os.path.isfile(prefData.oAttendeePath):
                self.Dico["OPTIONAL_PATH"] = ""
            if os.path.isfile(prefData.favoriteRoomsPath):
                self.Dico["ROOMS_PATH"] = prefData.favoriteRoomsPath
            elif not os.path.isfile(prefData.favoriteRoomsPath):
                self.Dico["ROOMS_PATH"] = ""
            self.dayTime["MIN"] = int(prefData.minDayTime)
            self.dayTime["MAX"] = int(prefData.maxDayTime)
        self.displayWdw(show)
    # ____________________________________________
    def displayWdw(self,
                   show=True):
        """ Affichage de la fenetre
        """        
        ## | Appel des widgets |
        self.rLnkEdit = gui.wlib.lineEdit(Txt=self.Dico["REQUIRED_PATH"],
                                          keyWord="required",
                                          toolTip=u"Lien vers les invités requis par défaut",
                                          acceptDrops=True,
                                          setEnabled=True)
        self.rLnkFind = gui.wlib.fileExplorer(dim=int(15),
                                              icon="folder.png",
                                              Txt=u"Invités requis",
                                              openPath=str(self.rLnkEdit.text()),
                                              zoneTxt=self.rLnkEdit,
                                              toolTip=u"Sélectionner un fichier ...",
                                              setEnabled=True)
        self.oLnkEdit = gui.wlib.lineEdit(Txt=self.Dico["OPTIONAL_PATH"],
                                          keyWord="optional",
                                          toolTip=u"Lien vers les invités optionnels par défaut",
                                          acceptDrops=True,
                                          setEnabled=True)
        self.oLnkFind = gui.wlib.fileExplorer(dim=int(15),
                                              icon="folder.png",
                                              Txt=u"Invités optionnels",
                                              openPath=str(self.oLnkEdit.text()),
                                              zoneTxt=self.oLnkEdit,
                                              toolTip=u"Sélectionner un fichier ...",
                                              setEnabled=True)
        self.roomLnkEdit = gui.wlib.lineEdit(Txt=self.Dico["ROOMS_PATH"],
                                             keyWord="rooms",
                                             toolTip=u"Lien vers les invités requis par défaut",
                                             acceptDrops=True,
                                             setEnabled=True)
        self.roomLnkFind = gui.wlib.fileExplorer(dim=int(15),
                                                 icon="folder.png",
                                                 Txt=u"Salles favorites",
                                                 openPath=str(self.roomLnkEdit.text()),
                                                 zoneTxt=self.roomLnkEdit,
                                                 toolTip=u"Sélectionner un fichier ...",
                                                 setEnabled=True)
        self.minDayTime = gui.wlib.timeEditer(time=self.dayTime["MIN"],
                                              timeStep=int(15))
        self.maxDayTime = gui.wlib.timeEditer(time=self.dayTime["MAX"],
                                              timeStep=int(15))
        self.validate = gui.wlib.pushButton(dim=int(15),
                                            icon="validation.jpg",
                                            Txt="Valider",
                                            toolTip=None,
                                            setDisabled=False)
        self.cancel = gui.wlib.pushButton(dim=int(15),
                                          icon="annulation.jpg",
                                          Txt="Annuler",
                                          toolTip=None,
                                          setDisabled=False)
        ## | Layout definition |
        wLayout = QtGui.QGridLayout()
        wLayout.addWidget(self.rLnkFind, 0, 0, 1, 1)
        wLayout.addWidget(self.rLnkEdit, 0, 1, 1, 2)
        wLayout.addWidget(self.oLnkFind, 1, 0, 1, 1)
        wLayout.addWidget(self.oLnkEdit, 1, 1, 1, 2)
        wLayout.addWidget(self.roomLnkFind, 2, 0, 1, 1)
        wLayout.addWidget(self.roomLnkEdit, 2, 1, 1, 2)
        wLayout.addWidget(QtGui.QLabel(u"Heure de début"), 3, 0, 1, 1)
        wLayout.addWidget(self.minDayTime, 3, 1, 1, 2)
        wLayout.addWidget(QtGui.QLabel(u"Heure de fin"), 4, 0, 1, 1)
        wLayout.addWidget(self.maxDayTime, 4, 1, 1, 2)
        wLayout.addWidget(self.validate, 5, 1, 1, 1)
        wLayout.addWidget(self.cancel, 5, 2, 1, 1)
        wLayout.setRowStretch(5, 0)
        self.setLayout(wLayout)
        self.center()
        if (show):
            self.show()
        ## | Connection |
        QtCore.QObject.connect(self.rLnkFind, QtCore.SIGNAL("clicked()"), lambda who=[self.rLnkFind, self.rLnkEdit]: self.select(who))
        QtCore.QObject.connect(self.oLnkFind, QtCore.SIGNAL("clicked()"), lambda who=[self.oLnkFind, self.oLnkEdit]: self.select(who))
        QtCore.QObject.connect(self.roomLnkFind, QtCore.SIGNAL("clicked()"), lambda who=[self.roomLnkFind, self.roomLnkEdit]: self.select(who))
        QtCore.QObject.connect(self.validate, QtCore.SIGNAL("clicked()"), lambda who=[self.rLnkEdit, self.oLnkEdit, self.roomLnkEdit, self.minDayTime, self.maxDayTime]: self.writePreferences(who))
        QtCore.QObject.connect(self.cancel, QtCore.SIGNAL("clicked()"), self.close)
    # ____________________________________________
    def writePreferences(self,
                         arg):
        """ Ecriture du fichier de preferences
        """
        error, flags = None, []
        keys = ["required",
                "optional",
                "rooms"]
        for item in arg:
            Path = r"%s" % str(item.text())
            if (os.path.isfile(Path) and ".py" in os.path.splitext(Path)):
                sys.path.append(os.path.dirname(Path))
                tmp = __builtin__.__import__(os.path.basename(os.path.splitext(Path)[0]),
                                             globals(),
                                             locals(),
                                             [],
                                             -1)
                if (tmp.__KeyWord__).lower() != keys[arg.index(item)]:
                    error = True
                    flags.append(arg.index(item))
        if error is not None:
            gui.wlib.criticalBox(wdwTitle="Program Error - Paths failure",
                                 wdwIcon="outlook.png",
                                 Txt=u"Incompatibilité des chemins:\
                                       \nIndices n° ... %s." % flags,
                                 stdButtons=QtGui.QMessageBox.Ok)
        else:
            File = open(self.PREFENCES_FILEPATH, "w")
            Preferences = """# -*- coding: utf-8 -*-
## *******************************************
## MY PREFERENCES:
__Application__ = "Meetep"
__Version__     = "__toolVersion__"
__KeyWord__     = "Preferences"
## *******************************************

rAttendeePath = r"__rAttendeePath__"
oAttendeePath = r"__oAttendeePath__"
favoriteRoomsPath = r"__favoriteRoomsPath__"
minDayTime = __minDayTime__
maxDayTime = __maxDayTime__
"""
            Preferences = Preferences.replace("__toolVersion__", __version__)
            Preferences = Preferences.replace("__rAttendeePath__", str(arg[0].text()))
            Preferences = Preferences.replace("__oAttendeePath__", str(arg[1].text()))
            Preferences = Preferences.replace("__favoriteRoomsPath__", str(arg[2].text()))
            Preferences = Preferences.replace("__minDayTime__", str(sum([int(arg[3].time().hour()) * 3600,
                                                                         int(arg[3].time().minute()) * 60])))
            Preferences = Preferences.replace("__maxDayTime__", str(sum([int(arg[4].time().hour()) * 3600,
                                                                         int(arg[4].time().minute()) * 60])))
            File.write(Preferences)
            File.close()
            self.close()
    # ____________________________________________
    def select(self,
               arg):
        """ Selection d'un fichier de donnees via l'explorateur Windows
        """
        if str(arg[1].text()) != "":
            selectFile = QtGui.QFileDialog.getOpenFileName(arg[0],
                                                           arg[0].tr(u"Sélectionner un fichier de préférences"),
                                                           str(arg[1].text()),
                                                           arg[0].tr("*.py"))
        else:
            selectFile = QtGui.QFileDialog.getOpenFileName(arg[0],
                                                           arg[0].tr(u"Sélectionner un fichier de préférences"),
                                                           USER_ENV["WKDIR"],
                                                           arg[0].tr("*.py"))
        try:
            selectFile = r"%s" % selectFile.replace("/", os.sep)
        except Exception as e:
            pass
        else:
            if selectFile != "":
                sys.path.append(os.path.dirname(str(selectFile)))
                prefData = __builtin__.__import__(os.path.basename(os.path.splitext(str(selectFile))[0]),
                                                  globals(),
                                                  locals(),
                                                  [],
                                                  -1)
                if arg[0] == self.rLnkFind: key = "required"
                if arg[0] == self.oLnkFind: key = "optional"
                if arg[0] == self.roomLnkFind: key = "rooms"
                try:
                    if (prefData.__KeyWord__).lower() != key:
                        gui.wlib.criticalBox(wdwTitle="Program Error - Reading file failure",
                                             wdwIcon="outlook.png",
                                             Txt=u"Incompatibilité du lien sélectionné.",
                                             stdButtons=QtGui.QMessageBox.Ok)  
                    else:
                        arg[0].zoneTxt.setText(selectFile)
                except Exception as e:
                    gui.wlib.criticalBox(wdwTitle="Program Error - Reading file failure",
                                         wdwIcon="outlook.png",
                                         Txt=u"Incompatibilité du lien sélectionné\
                                               \n(ne correspond pas au format d'un fichier de préférences).",
                                         stdButtons=QtGui.QMessageBox.Ok)
    # ____________________________________________
    def center(self):
        """ Recentrage de l'interface
        """
        qr = self.frameGeometry()
        cp = QtGui.QDesktopWidget().availableGeometry().center()
        qr.moveCenter(cp)
        self.move(qr.topLeft())
    # # ____________________________________________
    # def closeEvent(self,
    #                event=None):
    #     """ Fermeture de la window
    #     """
    #     QtCore.QCoreApplication.instance().quit()
# ...]
## ------------
# [...
class tmpWindow(QtGui.QDialog):
    """ Window de modification des salles lors d'un appel periodique
        @param: available_rooms = Liste des salles disponibles
        @param: booleanList = Liste de booleens
        @param: dateList = Listes des dates cibles
    """
    ident = "tmpWindow"
    # ____________________________________________
    def __init__(self,
                 attendees=[],
                 olRoomList=[],
                 capaList=[],
                 available_rooms=[],
                 booleanList=[],
                 dateList=[],
                 parent=None):
        """ Initialisation de la classe
        """
        QtGui.QDialog.__init__(self)
        QtGui.QToolTip.setFont(QtGui.QFont("SansSerif", int(10)))
        self.setWindowTitle(u"[MEETEP] Correction de la périodicité")
        self.setWindowIcon(QtGui.QIcon(r"%s\outlook.png" % TOOL_ENV["ICONS"]))
        self.booleanList = booleanList
        self.attendees = attendees
        self.olRoomList = olRoomList
        self.capaList = capaList
        self.dateList = dateList
        self.available_rooms = available_rooms
        self.addContxtMenu()
        self.changeRoom()
        self.showMaximized()
    # ____________________________________________
    def addContxtMenu(self):
        """ Insertion d'un menu contextuel
        """
        self.contextMenu = QtGui.QMenu(self)
        self.setContextMenuPolicy(QtCore.Qt.ActionsContextMenu)
        self.checkAll = QtGui.QAction(QtGui.QIcon(""),
                                      "Check/Uncheck all elements",
                                      self.contextMenu)
        self.checkAll.setCheckable(True)
        actionList = {"WIDGETS": [self.checkAll],
                      "SLOTS": [self.checking]}
        for i in range(len(actionList["WIDGETS"])):
            self.addAction(actionList["WIDGETS"][i])
            self.connect(actionList["WIDGETS"][i], QtCore.SIGNAL("triggered()"), actionList["SLOTS"][i])
    # ____________________________________________
    def checking(self):
        """ Ajouter un element
        """
        if (self.checkAll.isChecked()):
            for widget in self.widgets:
                widget[0].setCheckState(QtCore.Qt.Checked)
        else:
            for widget in self.widgets:
                widget[0].setCheckState(QtCore.Qt.Unchecked)
    # ____________________________________________
    def changeRoom(self):
        """ Modification des salles
        """
        tmpLayout = QtGui.QGridLayout()
        self.widgets = []
        self.filterBox = gui.wlib.filterBox()
        self.validate = gui.wlib.pushButton(dim=int(15),
                                            icon="validation.jpg",
                                            Txt="Valider",
                                            toolTip=None,
                                            setDisabled=False)
        self.cancel = gui.wlib.pushButton(dim=int(15),
                                          icon="annulation.jpg",
                                          Txt="Annuler",
                                          toolTip=None,
                                          setDisabled=False)
        QtCore.QObject.connect(self.cancel, QtCore.SIGNAL("clicked()"), self.close)
        tmpLayout.addWidget(self.filterBox, 0, 1, 1, 3)
        for i in range(len(self.dateList)):
            if self.booleanList[i] == "0":
                self.widgets.append((gui.wlib.checkBox(Txt="",
                                                       toolTip=None,
                                                       setDisabled=False),
                                     gui.wlib.datetimeEditer(dateTime=self.dateList[i],
                                                             calendar=None,
                                                             setDisabled=True),
                                     gui.wlib.comboBox(comboList=self.available_rooms[i],
                                                       TxtToSelect="",
                                                       setEditable=False,
                                                       adjustSizeActive=False,
                                                       setEnabled=True)))
                coloringBox(attendeeList=self.attendees,
                            roomList=self.available_rooms[i],
                            olRoomList=self.olRoomList,
                            capaList=self.capaList,
                            box=self.widgets[-1][2])
                tmpLayout.addWidget(self.widgets[-1][0], i+1, 0, 1, 1)
                tmpLayout.addWidget(self.widgets[-1][1], i+1, 1, 1, 1)
                tmpLayout.addWidget(self.widgets[-1][2], i+1, 2, 1, 2)
        tmpLayout.addWidget(self.validate, i+2, 2, 1, 1)
        tmpLayout.addWidget(self.cancel, i+2, 3, 1, 1)
        self.setLayout(tmpLayout)
    # ____________________________________________
    def center(self):
        """ Recentrage de l'interface
        """
        qr = self.frameGeometry()
        cp = QtGui.QDesktopWidget().availableGeometry().center()
        qr.moveCenter(cp)
        self.move(qr.topLeft())
    # # ____________________________________________
    # def closeEvent(self,
    #                event=None):
    #     """ Fermeture de la window
    #     """
    #     QtCore.QCoreApplication.instance().quit()
# ...]
## ------------
# [...
class filterWindow(QtGui.QDialog):
    """ Window de filtre pour ajouter une salle aux favoris
        @param: rooms = Liste des salles disponibles
    """
    ident = "filterWindow"
    # ____________________________________________
    def __init__(self,
                 attendees=[],
                 rooms=[],
                 olRoomList=[],
                 capaList=[],
                 parent=None):
        """ Initialisation de la classe
        """
        QtGui.QDialog.__init__(self)
        QtGui.QToolTip.setFont(QtGui.QFont("SansSerif", int(10)))
        self.setWindowTitle(u"[MEETEP] Filtre sur les salles")
        self.setWindowIcon(QtGui.QIcon(r"%s\outlook.png" % TOOL_ENV["ICONS"]))
        self.attendees = attendees
        self.olRoomList = olRoomList
        self.capaList = capaList
        self.rooms = rooms
        self.searchRoom()
    # ____________________________________________
    def searchRoom(self):
        """ Modification des salles
        """
        tmpLayout = QtGui.QGridLayout()
        self.filterBox = gui.wlib.filterBox()
        QtCore.QObject.connect(self.filterBox, QtCore.SIGNAL("textChanged(QString)"), self.roomFilter)
        self.roomsListing = gui.wlib.listWidget(itemsList=self.rooms,
                                                keyWord="",
                                                acceptDrops=False)
        coloringBox(attendeeList=self.attendees,
                    roomList=self.rooms,
                    olRoomList=self.olRoomList,
                    capaList=self.capaList,
                    box=self.roomsListing)
        self.validate = gui.wlib.pushButton(dim=int(15),
                                            icon="validation.jpg",
                                            Txt="Valider",
                                            toolTip=None,
                                            setDisabled=False)
        self.cancel = gui.wlib.pushButton(dim=int(15),
                                          icon="annulation.jpg",
                                          Txt="Annuler",
                                          toolTip=None,
                                          setDisabled=False)
        QtCore.QObject.connect(self.cancel, QtCore.SIGNAL("clicked()"), self.close)
        tmpLayout.addWidget(self.filterBox, 0, 0, 1, 2)
        tmpLayout.addWidget(self.roomsListing, 1, 0, 1, 2)
        tmpLayout.addWidget(self.validate, 2, 0, 1, 1)
        tmpLayout.addWidget(self.cancel, 2, 1, 1, 1)
        self.setLayout(tmpLayout)
    # ____________________________________________
    def roomFilter(self,
                   txt):
        """ Fonction de filtre des salles
        """
        ## | Affectation des arguments |
        rooms = {"DEFAULT": self.rooms,
                 "FILTERED": []}
        ## | Récupération du filtre renseigné |
        val = u"%s" % str(txt)
        ## | Listing des salles après application du filtre |
        if (val == u"" or val == u"*" or val == u"**"):
            for room in rooms["DEFAULT"]:
                rooms["FILTERED"].append(room)
        else:
            for room in rooms["DEFAULT"]:
                if (val[0] == u"*" and val[-1] == u"*"):
                    if val[1:-1].lower() in room.lower():
                        rooms["FILTERED"].append(room)
                elif (val[0] == u"*"):
                    if val[1:].lower() == room[1:len(val)-1].lower():
                        rooms["FILTERED"].append(room)
                elif (val[-1] == u"*"):
                    if val[:-1].lower() == room[len(val)-1:-1].lower():
                        rooms["FILTERED"].append(room)
                else:
                    if val.lower() == room.lower():
                        rooms["FILTERED"].append(room)
        ## | Mise à jour des comboBox |
        self.roomsListing.clear()
        self.roomsListing.addItems(QtCore.QStringList(rooms["FILTERED"]))
        self.roomsListing.setCurrentRow(int(0))
    # ____________________________________________
    def center(self):
        """ Recentrage de l'interface
        """
        qr = self.frameGeometry()
        cp = QtGui.QDesktopWidget().availableGeometry().center()
        qr.moveCenter(cp)
        self.move(qr.topLeft())
    # # ____________________________________________
    # def closeEvent(self,
    #                event=None):
    #     """ Fermeture de la window
    #     """
    #     QtCore.QCoreApplication.instance().quit()
# ...]
## ------------
# [...
class coloringBox():
    """ Mise en evidence des salles compatibles suivant leur capacite de remplissage
        @param: attendeeList = Liste des participants
        @param: roomList = Liste des salles disponibles
        @param: olRoomList = Liste complete des salles sur Villaroche
        @param: capaList = Liste des capacites
        @param: box = Combobox cible
    """
    ident = "coloringBox"
    # ____________________________________________
    def __init__(self,
                 attendeeList=[],
                 roomList=[],
                 olRoomList=[],
                 capaList=[],
                 box=None):
        nb = len(attendeeList)
        for i in range(len(roomList)):
            room = roomList[i]
            for j in range(len(olRoomList)):
                if room in olRoomList[j]:
                    id = j
                    break
            if capaList[id] != "#N/A":
                if (capaList[id] < nb or capaList[id] > nb + 1):
                    if box.ident == "comboBox":
                        box.setItemData(i, QtGui.QBrush(QtCore.Qt.red), QtCore.Qt.TextColorRole)
                    elif box.ident == "listWidget":
                        box.item(i).setData(QtCore.Qt.ForegroundRole, QtGui.QBrush(QtCore.Qt.red))
                elif (nb <= capaList[id] <= nb + 1):
                    if box.ident == "comboBox":
                        box.setItemData(i, QtGui.QBrush(QtCore.Qt.green), QtCore.Qt.TextColorRole)
                    elif box.ident == "listWidget":
                        box.item(i).setData(QtCore.Qt.ForegroundRole, QtGui.QBrush(QtCore.Qt.green))
# ...]
## ------------
# [...
class mainWindow(QtGui.QMainWindow):
    """ Edition de l'UI
        @param: session = Session Outlook
        @param: olRooms = Infos sur les salles
        @param: olDateList = Compilation des horaires
        @param: timeStep = Time step
    """
    ## | Initialisation des variables globales de la classe |
    ident = "mainWindow"
    savingName = None
    errorFlag = int(0)
    maxPriority = int(0)
    indexation, startId = int(0), int(0)
    itemsList = []
    modification = False
    searchData = OrderedDict([("requiredAttendee", OrderedDict()),
                              ("optionalAttendee", []),
                              ("favoriteRooms", []),
                              ("duration", int(0))])
    ouputs = {"MEETING_SLOTS": OrderedDict(),
              "PRIORITIES": []}
    oFilter = {"MEETING_SLOTS": OrderedDict(),
               "PRIORITIES": None}
    # ____________________________________________
    def __init__(self,
                 session,
                 olRooms=OrderedDict(),
                 olDateList=[],
                 timeStep=int(15),
                 parent=None):
        """ Initialisation de la classe
        """
        QtGui.QMainWindow.__init__(self)
        self.olRooms = olRooms
        self.olDateList = olDateList
        self.session = session
        self.timeStep = timeStep
        self.pref = editPreferences(show=False)
        QtGui.QToolTip.setFont(QtGui.QFont("SansSerif", int(10)))
        self.setWindowTitle(WINDOWS["TITLES"]["MAIN"])
        self.setWindowIcon(QtGui.QIcon(r"%s\outlook.png" % TOOL_ENV["ICONS"]))
        self.setOptionToolbar()
        self.queryUI()
    # ____________________________________________
    def setOptionToolbar(self):
        """ Edition de la toolbar contenant les options
        """
        ## | toolBar item |
        self.optionBar = QtGui.QToolBar("Options Toolbar")
        self.addToolBar(QtCore.Qt.TopToolBarArea, self.optionBar)
        ## | Création des actions |
        self.load = QtGui.QAction(QtGui.QIcon(r"%s\load.png" % TOOL_ENV["ICONS"]),
                                  u"Charger un fichier de données {Ctrl+O})",
                                  self,
                                  shortcut="CTRL+O")
        self.save = QtGui.QAction(QtGui.QIcon(r"%s\save.png" % TOOL_ENV["ICONS"]),
                                  u"Sauvegarder {Ctrl+S})",
                                  self,
                                  shortcut="CTRL+S")
        self.saveAs = QtGui.QAction(QtGui.QIcon(r"%s\saveAs.png" % TOOL_ENV["ICONS"]),
                                    u"Sauvegarder sous (press {Ctrl+Alt+S})",
                                    self,
                                    shortcut="CTRL+ALT+S")
        self.preferences = QtGui.QAction(QtGui.QIcon(r"%s\preferences.jpg" % TOOL_ENV["ICONS"]),
                                         u"Edition des préférences (press {Ctrl+P})",
                                         self,
                                         shortcut="CTRL+P")
        self.clean = QtGui.QAction(QtGui.QIcon(r"%s\clean.png" % TOOL_ENV["ICONS"]),
                                   u"Reset (press {F1})",
                                   self,
                                   shortcut="F1")
        self.help = QtGui.QAction(QtGui.QIcon(r"%s\help.png" % TOOL_ENV["ICONS"]),
                                  "Aide (press {F2})",
                                  self,
                                  shortcut="F2")
        ## | Association des actions à leurs fonctionnalités respectives |
        actionList = {"WIDGETS": [self.load, self.save, self.saveAs,
                                  self.preferences, self.clean, self.help],
                      "SLOTS": [self.loadData, self.saving, self.savingAs,
                                editPreferences, self.resetData, self.openManual]}
        for i in range(len(actionList["WIDGETS"])):
            self.optionBar.addAction(actionList["WIDGETS"][i])
            self.connect(actionList["WIDGETS"][i], QtCore.SIGNAL("triggered()"), actionList["SLOTS"][i])
        self.optionBar.insertSeparator(self.clean)
    # ____________________________________________
    def loadData(self):
        """ Chargement d'un fichier
        """
        dataFile = QtGui.QFileDialog.getOpenFileName(self,
                                                     self.tr(u"Sélectionner un fichier de données"),
                                                     USER_ENV["WKDIR"],
                                                     self.tr("*.py"))
        if dataFile != "":
            sys.path.append(os.path.dirname(str(dataFile)))
            dropInfos = __builtin__.__import__(os.path.basename(os.path.splitext(str(dataFile))[0]),
                                               globals(),
                                               locals(),
                                               [],
                                               -1)
            if (dropInfos.__KeyWord__).lower() == "sauvegarde":
                self.meetingDuration.setTime(QtCore.QTime().addSecs(int(dropInfos.meetingDuration * 60)))
                self.meetingDuration.applyCorrection()
                self.editRequiredAttendee.clearContents()
                self.editRequiredAttendee.setRowCount(len(dropInfos.requiredAttendeeList))
                for elem in dropInfos.requiredAttendeeList:
                    self.editRequiredAttendee.setItem(dropInfos.requiredAttendeeList.index(elem), 0, QtGui.QTableWidgetItem(QtCore.QString(elem[0])))
                    self.editRequiredAttendee.setItem(dropInfos.requiredAttendeeList.index(elem), 1, QtGui.QTableWidgetItem(QtCore.QString(str(elem[1]))))
                self.editOptionalAttendee.clear()
                self.editOptionalAttendee.addItems(QtCore.QStringList(dropInfos.optionalAttendeeList))
                self.editFavoriteRooms.clear()
                self.editFavoriteRooms.addItems(QtCore.QStringList(dropInfos.favoriteRoomsList))
                gui.wlib.infoBox(wdwTitle="Program Info - Loading file",
                                 wdwIcon="outlook.png",
                                 Txt=u"Fichier de sauvegarde chargé avec succès.",
                                 stdButtons=QtGui.QMessageBox.Ok)
            else:
                gui.wlib.criticalBox(wdwTitle="Program Error - Reading file failure",
                                     wdwIcon="outlook.png",
                                     Txt=u"Incompatibilité du lien sélectionné\
                                           \n(ne correspond pas au format d'un fichier de données).",
                                     stdButtons=QtGui.QMessageBox.Ok)
    # ____________________________________________
    def saving(self):
        """ Sauvegarde de la recherche
        """
        if self.savingName is not None:
            self.writeSavingFile(r"%s\%s" % (USER_ENV["WKDIR"], self.savingName))
        else:
            self.savingAs()
    # ____________________________________________
    def savingAs(self):
        """ Sauvegarde sous de la recherche
        """
        savingFile = QtGui.QFileDialog.getSaveFileName(self,
                                                       self.tr(u"Sauvegarde sous"),
                                                       r"%s\autoSave" % USER_ENV["WKDIR"],
                                                       self.tr("*.py"))
        if savingFile != "":
            self.savingName = os.path.basename(str(savingFile))
            self.writeSavingFile(str(savingFile))
    # ____________________________________________
    def writeSavingFile(self,
                        filePath=r"%s\autoSave.py" % USER_ENV["WKDIR"]):
        """ Ecriture du fichier de sauvegarde
        """
        ## | Récupération des invités requis |
        requiredAttendees = []
        if self.editRequiredAttendee.rowCount() != int(0):
            for i in range(self.editRequiredAttendee.rowCount()):
                name = str(self.editRequiredAttendee.item(i, 0).text())
                priority = int(self.editRequiredAttendee.item(i, 1).text())
                requiredAttendees.append((name, priority))
        ## | Récupération des invités optionnels |
        optionalAttendees = []
        if self.editOptionalAttendee.count() != int(0):
            for i in range(self.editOptionalAttendee.count()):
                optionalAttendees.append(str(self.editOptionalAttendee.item(i).text()))
        ## | Récupération des salles favorites |
        favoriteRooms = []
        if self.editFavoriteRooms.count() != int(0):
            for i in range(self.editFavoriteRooms.count()):
                favoriteRooms.append(str(self.editFavoriteRooms.item(i).text()))
        ## | Ecriture du fichier de sauvegarde |
        File = open(filePath, "w")
        uData = """#!/usr/bin/env python
# -*- coding: utf-8 -*-

## =================================================================================================
## HEADER
## =================================================================================================
# [...
__author__      = "Alexandre Brosse"
__copyright__   = "Copyright 2017 (C)"
__license__     = "None"
__maintainer__  = "Alexandre Brosse"
__status__      = "Production/Oper"
__date__        = "Wed, 21/06/2017"
__Application__ = "Meetep"
__comment__     = "Preferences"
__version__     = "V01R00"
__KeyWord__     = "Sauvegarde"
# ...]

## =================================================================================================
## MY SAVING FILE:
## =================================================================================================

## -- Durée de la réunion (en minutes)
meetingDuration = __meetingDuration__
## -- Liste des invités requis
requiredAttendeeList = __rAttendeeList__
## --  Liste des invités optionnels
optionalAttendeeList = __oAttendeeList__
## -- Liste des salles favorites
favoriteRoomsList = __favoriteRoomsList__
"""
        uData = uData.replace("__meetingDuration__", '%i' % sum([int(self.meetingDuration.time().hour()) * 60,
                                                                 int(self.meetingDuration.time().minute())]))
        uData = uData.replace("__rAttendeeList__", '%s' % requiredAttendees)
        uData = uData.replace("__oAttendeeList__", '%s' % optionalAttendees)
        uData = uData.replace("__favoriteRoomsList__", '%s' % favoriteRooms)
        File.write(uData)
        File.close()
        gui.wlib.infoBox(wdwTitle="Program Info - Saving file",
                         wdwIcon="outlook.png",
                         Txt=u"Lien vers fichier de sauvegarde:\
                               \n%s" % filePath,
                         stdButtons=QtGui.QMessageBox.Ok)
    # ____________________________________________
    def queryUI(self):
        """ Definition de l'UI
        """
        ## | UI de paramètrage de la recherche |
        self.executionWdw()
        self.searchWdw()
        self.resultsWdw()
        self.switch = QtGui.QStackedWidget()
        ## | Layout definition |
        self.mainLayout = QtGui.QGridLayout()
        self.mainLayout.addWidget(self.executionDetails, 0, 0, 1, 2)
        self.mainLayout.addWidget(self.searchDetails, 1, 0, 1, 1)
        self.mainLayout.addWidget(self.switch, 1, 1, 1, 1)
        ## | Display window |
        self.centralWidget = QtGui.QWidget(self)
        self.centralWidget.setLayout(self.mainLayout)
        self.setCentralWidget(self.centralWidget)
        self.center()
        self.show()
    # ____________________________________________
    def executionWdw(self):
        """ Definition de la fenetre d'execution
        """  
        ## | Appel des widgets |
        self.executionDetails = gui.wlib.groupBox(Title="",
                                                  setFlat=True)
        self.computing = gui.wlib.pushButton(dim=int(30),
                                             icon="compute.jpg",
                                             Txt="Compute",
                                             toolTip="Get common slots",
                                             setDisabled=False)
        self.seePreviousDay = gui.wlib.pushButton(dim=int(30),
                                             icon="previous.jpg",
                                             Txt="Back day",
                                             toolTip="Back to previous day",
                                             setDisabled=True)
        self.seeNextDay = gui.wlib.pushButton(dim=int(30),
                                             icon="next.jpg",
                                             Txt="Next day",
                                             toolTip="Go to next day",
                                             setDisabled=True)
        ## | Layout definition |
        hbox = QtGui.QHBoxLayout()
        hbox.addWidget(self.computing)
        hbox.addWidget(self.seePreviousDay)
        hbox.addWidget(self.seeNextDay)
        hbox.addStretch(0)
        self.executionDetails.setLayout(hbox)
        ## | Connection |
        QtCore.QObject.connect(self.computing, QtCore.SIGNAL("clicked()"),self.compute)
        QtCore.QObject.connect(self.seePreviousDay, QtCore.SIGNAL("clicked()"), self.gotoBack)
        QtCore.QObject.connect(self.seeNextDay, QtCore.SIGNAL("clicked()"), self.gotoNext)
    # ____________________________________________
    def searchWdw(self):
        """ Definition de la fenetre de recherche
        """
        ## | Initialisation des variables temporaires |
        requiredAttendees, optionalAttendees, favoriteRooms = [], [], []
        tList = ["rAttendee",
                 "oAttendee",
                 "Rooms"]
        ## | Chargement des préférences |
        if os.path.isfile(editPreferences.PREFENCES_FILEPATH):
            sys.path.append(os.path.dirname(editPreferences.PREFENCES_FILEPATH))
            readInfos = __builtin__.__import__(os.path.basename(os.path.splitext(editPreferences.PREFENCES_FILEPATH)[0]),
                                               globals(),
                                               locals(),
                                               [],
                                               -1)
            if readInfos.rAttendeePath != "":
                sys.path.append(os.path.dirname(readInfos.rAttendeePath))
                requiredInfos = __builtin__.__import__(os.path.basename(os.path.splitext(readInfos.rAttendeePath)[0]),
                                                       globals(),
                                                       locals(),
                                                       [],
                                                       -1)
                requiredAttendees = requiredInfos.requiredAttendeeList
            if readInfos.oAttendeePath != "":
                sys.path.append(os.path.dirname(readInfos.oAttendeePath))
                optionalInfos = __builtin__.__import__(os.path.basename(os.path.splitext(readInfos.oAttendeePath)[0]),
                                                       globals(),
                                                       locals(),
                                                       [],
                                                       -1)
                optionalAttendees = optionalInfos.optionalAttendeeList
            if readInfos.favoriteRoomsPath != "":
                sys.path.append(os.path.dirname(readInfos.favoriteRoomsPath))
                roomInfos = __builtin__.__import__(os.path.basename(os.path.splitext(readInfos.favoriteRoomsPath)[0]),
                                                   globals(),
                                                   locals(),
                                                   [],
                                                   -1)
                favoriteRooms = roomInfos.favoriteRoomsList
        ## | Appel des widgets | 
        self.searchDetails = gui.wlib.groupBox(Title="SEARCH DETAILS",
                                               setFlat=False)
        self.meetingDuration = gui.wlib.timeEditer()
        self.editRequiredAttendee = gui.wlib.tableWidget(nRows=len(requiredAttendees),
                                                         headersList=[u"Nom Prénom", u"Priorité"],
                                                         itemsList=requiredAttendees,
                                                         keyWord="required",
                                                         acceptDrops=True)
        self.editOptionalAttendee = gui.wlib.listWidget(itemsList=optionalAttendees,
                                                        keyWord="optional",
                                                        acceptDrops=True)
        self.editFavoriteRooms = gui.wlib.listWidget(itemsList=favoriteRooms,
                                                     keyWord="rooms",
                                                     acceptDrops=True)
        self.subTools = OrderedDict()
        for key in tList:
            self.subTools[key] = [gui.wlib.pushButton(dim=int(15),
                                                      icon="add.png",
                                                      Txt="",
                                                      toolTip=None,
                                                      setDisabled=False),
                                  gui.wlib.pushButton(dim=int(15),
                                                      icon="delete.png",
                                                      Txt="",
                                                      toolTip=None,
                                                      setDisabled=False),
                                  gui.wlib.pushButton(dim=int(15),
                                                      icon="clean.png",
                                                      Txt="",
                                                      toolTip=None,
                                                      setDisabled=False)]
        ## | Layout definition |
        self.searchLayout = QtGui.QGridLayout()
        self.searchLayout.addWidget(QtGui.QLabel(u"Durée de la réunion (en h:min:s) :"), 0, 0, 1, 1)
        self.searchLayout.addWidget(self.meetingDuration, 0, 1, 1, 3)
        self.searchLayout.addWidget(QtGui.QLabel(u"Liste des participants obligatoires :"), 1, 0, 1, 1)
        self.searchLayout.addWidget(self.editRequiredAttendee, 2, 0, 1, len(self.subTools["rAttendee"])+1)
        self.searchLayout.addWidget(QtGui.QLabel(u"Liste des participants facultatifs :"), 3, 0, 1, 1)
        self.searchLayout.addWidget(self.editOptionalAttendee, 4, 0, 1, len(self.subTools["oAttendee"])+1)
        self.searchLayout.addWidget(QtGui.QLabel(u"Liste des salles favorites :"), 5, 0, 1, 1)
        self.searchLayout.addWidget(self.editFavoriteRooms, 6, 0, 1, len(self.subTools["Rooms"])+1)
        for key in tList:
            i = tList.index(key)
            for widget in self.subTools[key]:
                j = self.subTools[key].index(widget)
                self.searchLayout.addWidget(widget, 2*i+1, j+1, 1, 1)
        self.searchDetails.setLayout(self.searchLayout)
        ## | Connection |
        argList = [self.editRequiredAttendee,
                   self.editOptionalAttendee,
                   self.editFavoriteRooms]
        for key in tList:
            QtCore.QObject.connect(self.subTools[key][0], QtCore.SIGNAL("clicked()"), lambda who=argList[tList.index(key)]: self.add(who))
            QtCore.QObject.connect(self.subTools[key][1], QtCore.SIGNAL("clicked()"), lambda who=argList[tList.index(key)]: self.delete(who))
            QtCore.QObject.connect(self.subTools[key][2], QtCore.SIGNAL("clicked()"), lambda who=argList[tList.index(key)]: self.resetWidget(who))
    # ____________________________________________
    def resultsWdw(self):
        """ Definition de la fenetre affichant les resultats
        """  
        ## | Appel des widgets |
        self.meetingDetails = gui.wlib.groupBox(Title="MEETING DETAILS",
                                                setFlat=False)
        self.calendar_view = gui.wlib.calendar(displayNavBar=True,
                                               enableSelection=False)
        self.slotSelection = gui.wlib.treeWidget(column=int(0),
                                                 columnNb=int(1),
                                                 headerLabel=u"Listing des timeslots",
                                                 itemList=[])
        self.timeMeeting = gui.wlib.datetimeEditer(dateTime=datetime.datetime.now(),
                                                   calendar=self.calendar_view,
                                                   setDisabled=True)
        self.selectRoom = gui.wlib.comboBox(comboList=[],
                                            TxtToSelect="",
                                            setEditable=False,
                                            adjustSizeActive=False,
                                            setEnabled=True)
        self.filterRoom = gui.wlib.filterBox()
        self.nModelList = self.filterRoom.modelList
        self.activatePeriodicity = gui.wlib.checkBox(Txt=u"Activation du mode 'périodique'",
                                                     toolTip=None,
                                                     setDisabled=False)
        self.periodicityChoice = gui.wlib.comboBox(comboList=["Quotidienne", "Hebdo", "Mensuelle", "Annuelle"],
                                                   TxtToSelect="Quotidienne",
                                                   setEditable=False,
                                                   adjustSizeActive=False,
                                                   setEnabled=False)
        self.periodicityGo = gui.wlib.pushButton(dim=int(15),
                                                 icon="find.jpg",
                                                 Txt="Check Dispos",
                                                 toolTip=None,
                                                 setDisabled=True)
        self.attendeeSummary = gui.wlib.listWidget(itemsList=[],
                                                   keyWord="",
                                                   acceptDrops=False)
        self.sendMeeting = gui.wlib.pushButton(dim=None,
                                               icon=None,
                                               Txt="Envoyer l'invitation",
                                               toolTip=None,
                                               setDisabled=True)
        self.displayMeeting = gui.wlib.pushButton(dim=None,
                                                  icon=None,
                                                  Txt="Afficher l'invitation",
                                                  toolTip=None,
                                                  setDisabled=False)
        ## | Insertion d'un splitter |
        hbox = QtGui.QVBoxLayout()
        Frames = [gui.wlib.frame(),
                  gui.wlib.frame()]
        Splitter = QtGui.QSplitter(QtCore.Qt.Horizontal)
        for i in range(len(Frames)):
            Splitter.addWidget(Frames[i])
        hbox.addWidget(Splitter)
        self.meetingDetails.setLayout(hbox)
        ## | Layout definition |
        self.resultLayout = [QtGui.QGridLayout(), QtGui.QGridLayout()]
        self.resultLayout[0].addWidget(self.calendar_view, 0, 0, 5, 1)
        self.resultLayout[0].addWidget(self.slotSelection, 6, 0, 1, 1)
        self.resultLayout[1].addWidget(QtGui.QLabel(u"Créneau sélectionné :"), 0, 0, 1, 1)
        self.resultLayout[1].addWidget(self.timeMeeting, 0, 1, 1, 4)
        self.resultLayout[1].addWidget(QtGui.QLabel(u"Salle de réunion :"), 1, 0, 1, 1)
        self.resultLayout[1].addWidget(self.selectRoom, 1, 1, 1, 2)
        self.resultLayout[1].addWidget(self.filterRoom, 1, 3, 1, 2)
        self.resultLayout[1].addWidget(QtGui.QLabel(u"Périodicité :"), 2, 0, 1, 1)
        self.resultLayout[1].addWidget(self.activatePeriodicity, 2, 1, 1, 2)
        self.resultLayout[1].addWidget(self.periodicityChoice, 2, 3, 1, 1)
        self.resultLayout[1].addWidget(self.periodicityGo, 2, 4, 1, 1)
        self.resultLayout[1].addWidget(QtGui.QLabel(u"Rappel des invités :"), 3, 0, 1, 1)
        self.resultLayout[1].addWidget(self.attendeeSummary, 4, 0, 1, 5)
        self.resultLayout[1].addWidget(self.displayMeeting, 5, 1, 1, 2)
        self.resultLayout[1].addWidget(self.sendMeeting, 5, 3, 1, 2)
        for i in range(len(Frames)):
            Frames[i].setLayout(self.resultLayout[i])
        ## | Connection |
        QtCore.QObject.connect(self.slotSelection, QtCore.SIGNAL("itemSelectionChanged()"),self.showSlot)
        QtCore.QObject.connect(self.activatePeriodicity, QtCore.SIGNAL("stateChanged(int)"), self.enablePeriod)
        QtCore.QObject.connect(self.periodicityGo, QtCore.SIGNAL("clicked()"), self.getRecursiveRooms)
        QtCore.QObject.connect(self.displayMeeting, QtCore.SIGNAL("clicked()"), lambda who=True: self.oMeeting(who))
        QtCore.QObject.connect(self.sendMeeting, QtCore.SIGNAL("clicked()"), lambda who=False: self.oMeeting(who))
    # ____________________________________________
    def add(self,
            widget):
        """ Ajout d'un item
        """
        if widget.ident == "tableWidget":
            widget.insertRow(0)
            widget.setItem(0, 0, QtGui.QTableWidgetItem(QtCore.QString("new attendee")))
            widget.setItem(0, 1, QtGui.QTableWidgetItem(QtCore.QString("0")))
        elif widget.ident == "listWidget":
            if widget == self.editFavoriteRooms:
                attendees = []
                for i in range(int(self.editRequiredAttendee.rowCount())):
                    attendees.append(str(self.editRequiredAttendee.item(i, 0).text()))
                self.tmpSearch = filterWindow(attendees=attendees,
                                              rooms=self.olRooms["LISTE"],
                                              olRoomList=self.olRooms["LISTE"],
                                              capaList=self.olRooms["CAPACITY"])
                QtCore.QObject.connect(self.tmpSearch.validate, QtCore.SIGNAL("clicked()"), self.addRoom)
                self.tmpSearch.exec_()
            else:
                widget.insertItem(0, QtCore.QString("new attendee"))
                widget.item(0).setFlags(widget.item(0).flags() | QtCore.Qt.ItemIsEditable)
                widget.editItem(widget.item(0))
    # ____________________________________________
    def addRoom(self):
        """ Ajout d'une salle
        """
        selectedItem = self.tmpSearch.roomsListing.currentItem()
        txt = str(selectedItem.text())
        self.editFavoriteRooms.insertItem(0, QtCore.QString(txt))
        self.tmpSearch.close()
    # ____________________________________________
    def delete(self,
               widget):
        """ Suppression d'un item
        """
        if widget.ident == "tableWidget":
            widget.removeRow(widget.currentRow())
        elif widget.ident == "listWidget":
            widget.takeItem(widget.currentRow())
    # ____________________________________________
    def resetWidget(self,
                    widget):
        """ Reset du widget
        """
        if widget.ident == "tableWidget":
            widget.clearContents()
        elif widget.ident == "listWidget":
            widget.clear()
    # ____________________________________________
    def resetData(self):
        """ Reset de l'UI
        """
        self.meetingDuration.setTime(QtCore.QTime().addSecs(int(0)))
        self.editRequiredAttendee.clearContents()
        self.editOptionalAttendee.clear()
        self.editFavoriteRooms.clear()
    # ____________________________________________
    def openManual(self,
                   helpFile="UserManual.pptx"):
        """ Ouverture du manuel utilisateur
        """
        cmd = r"%s\%s" % (TOOL_ENV["DOCS"], helpFile)
        subprocess.call("start %s" % cmd, shell=True)
    # ____________________________________________
    def gotoBack(self):
        """ Back to previous day
        """
        self.modification = False
        ## | Mise à jour des indices et du visuel graphique |
        self.startId -= len(self.itemsList)
        if int(0) < self.indexation:
            self.indexation -= int(1)
        if self.indexation != len(self.oFilter["MEETING_SLOTS"].keys()) - 1:
            self.seeNextDay.setDisabled(False)  
        if self.indexation == int(0):
            self.seePreviousDay.setDisabled(True)
        ## | Mise à jour des slots du jour sélectionné |
        self.udpateTimeslotUI()
    # ____________________________________________
    def gotoNext(self):
        """ Go to next day
        """    
        self.modification = False
        ## | Mise à jour des indices et du visuel graphique |
        self.startId += len(self.itemsList)
        if self.indexation < len(self.oFilter["MEETING_SLOTS"].keys()) - 1:
            self.indexation += int(1)
        if self.indexation != int(0):
            self.seePreviousDay.setDisabled(False)
        if self.indexation == len(self.oFilter["MEETING_SLOTS"].keys()) - 1:
            self.seeNextDay.setDisabled(True)
        ## | Mise à jour des slots du jour sélectionné |
        self.udpateTimeslotUI()
    # ____________________________________________ 
    def getSearchData(self):
        """ Lecture des donnees de la recherche
            & Compilation brutes des slots disponibles
        """
        ## | Plage de steps pour couvrir la durée de réunion spécifiée |
        self.searchData["duration"] = sum([int(self.meetingDuration.time().hour()) * 60,
                                           int(self.meetingDuration.time().minute())])
        self.plage = self.searchData["duration"] / self.timeStep
        ## | Récupération des invités et salles |
        for i in range(int(self.editRequiredAttendee.rowCount())):
            name = str(self.editRequiredAttendee.item(i, 0).text())
            priority = int(self.editRequiredAttendee.item(i, 1).text())
            if priority >= self.maxPriority:
                self.maxPriority = priority
            if not "priority_%i" % priority in self.searchData["requiredAttendee"].keys():
                self.searchData["requiredAttendee"]["priority_%i" % priority] = []
            self.searchData["requiredAttendee"]["priority_%i" % priority].append(name)
        if (len(self.searchData["requiredAttendee"].keys()) == int(0) or self.searchData["duration"] == int(0)):
            self.errorFlag = int(1)
            gui.wlib.criticalBox(wdwTitle="Program Error - Incomplete data",
                                 wdwIcon="outlook.png",
                                 Txt=u"Mise en données incomplète.\
                                       \nVérifiez que la durée renseignée est bien non nulle et que la liste des participants obligatoire n'est pas vide.",
                                 stdButtons=QtGui.QMessageBox.Ok)
        else:
            ## | Affectation des variables de la recherche |
            for i in range(int(1), len(self.searchData["requiredAttendee"].keys())):
                self.searchData["requiredAttendee"]["priority_%i" % i].extend(self.searchData["requiredAttendee"]["priority_%i" % int(i-1)])
            for i in range(int(self.editOptionalAttendee.count())):
                self.searchData["optionalAttendee"].append(str(self.editOptionalAttendee.item(i).text()))
            for i in range(int(self.editFavoriteRooms.count())):
                self.searchData["favoriteRooms"].append(str(self.editFavoriteRooms.item(i).text()))
            ## | Compilation brute des timeslots disponibles |
            for key in self.searchData["requiredAttendee"].keys():
                list2find = self.searchData["requiredAttendee"][key]
                self.ouputs["MEETING_SLOTS"][key] = self.getXDispo(list2find)
    # ____________________________________________
    def getXDispo(self,
                  userList=[]):
        """ Recuperation des disponibilites croisees
        """
        ## | Initialisation des variables locales de la fonction |
        dispos, errorUserFB = [], []
        ## | Lecture des disponibilités individuelles via la fonction FreeBusy sur le mois à venir |
        for user in userList:
            try:
                recip = self.session.CreateRecipient(user) 
                myFBInfo = recip.FreeBusy(datetime.date.today(),
                                          self.timeStep,
                                          True)
            except Exception as e:
                self.errorFlag = int(2)
                errorUserFB.append(user)
            else:
                motif = "".join(["0"] * self.plage) #/Indication d'une disponibilité pendant la durée de réunion spécifiée
                startId, tmp = int(0), int(0)
                dispos.append([]) 
                while startId < len(myFBInfo) - self.plage: #/Extraction des dates et heures disponibles sur le mois à venir
                    tmp_old = tmp
                    tmp = myFBInfo.find(motif, startId)
                    if tmp - tmp_old < int(0):
                        break
                    else:
                        startId = tmp + int(1)
                        dispos[-1].append(tmp)
        if self.errorFlag != int(0):
            gui.wlib.criticalBox(wdwTitle="Program Error - FreeBusy process failed",
                                 wdwIcon="outlook.png",
                                 Txt=u"Lecture en erreur pour les users suivants:\
                                       \n%s" % errorUserFB,
                                 stdButtons=QtGui.QMessageBox.Ok)
            return None
        else:
            ## | Retrait des heures antérieures à l'heure actuelle |
            timeDelta = int(-1)
            while timeDelta < 0:
                result = min(set.intersection(*(set(x) for x in dispos)))
                timeDelta = getTimeDelta(timeStep=self.timeStep,
                                         index=result)
                if timeDelta < int(0):
                    for elem in dispos:
                        elem.remove(result)
            ## | Croisement des disponiblités et récupération des créneaux communs |
            return sorted(set.intersection(*(set(x) for x in dispos)))
    # ____________________________________________
    def sortSlots(self,
                  slotList=[],
                  priorityList=[]):
        """ Trie des timeslots
            @Retrait des samedis et dimanches
            @Retrait des horaires peu orthodoxes (avant 8h, pause midi, après 19h)
        """
        ## | Compilation des timeslots non compatibles |
        tmpSlots, tmpIndex = [], []
        minTime = sum([int(self.pref.minDayTime.time().hour()) * 3600,
                       int(self.pref.minDayTime.time().minute()) * 60])
        maxTime = sum([int(self.pref.maxDayTime.time().hour()) * 3600,
                       int(self.pref.maxDayTime.time().minute()) * 60])
        for i in range(len(slotList)):
            timeDelta = getTimeDelta(timeStep=self.timeStep,
                                     index=slotList[i])
            ndate = datetime.datetime.now() + datetime.timedelta(seconds=timeDelta)
            secDeltaRef = datetime.timedelta(hours=ndate.hour,
                                             minutes=ndate.minute,
                                             seconds=ndate.second,
                                             microseconds=ndate.microsecond).total_seconds()
            secDeltaNew = secDeltaRef + self.searchData["duration"] * 60
            iterLoop = int(1)
            while secDeltaNew > 24 * 3600:
                secDeltaNew = secDeltaNew - iterLoop * 24 * 3600
                iterLoop += int(1)
            startOfSlot = sum([time.mktime(time.localtime()),
                               timeDelta])
            day = time.strftime("%a", time.localtime(startOfSlot))
            if (day in ["Sat", "Sun"]):
                tmpSlots.append(slotList[i])
                tmpIndex.append(i)
            if (secDeltaRef < minTime or secDeltaNew < minTime or secDeltaNew > maxTime):
                tmpSlots.append(slotList[i])
                tmpIndex.append(i)
            if (secDeltaRef >= 12 * 3600 and secDeltaRef < 13 * 3600):
                tmpSlots.append(slotList[i])
                tmpIndex.append(i)
        ## | Retrait des éléments multiples |
        koIndex = []
        for item in set(tmpIndex):
            koIndex.append(item)
        ## | Exclusion des timeslots non compatibles |
        sortedSlots, sortedPriorities = [], []
        for i in range(len(slotList)):
            if not i in koIndex:
                sortedSlots.append(slotList[i])
                sortedPriorities.append(priorityList[i])
        if len(sortedSlots) == int(0):
            self.errorFlag == int(3)
            gui.wlib.warningBox(wdwTitle="Program Warning - No matching found",
                                wdwIcon="outlook.png",
                                Txt=u"Aucun créneau commun n'a pu être trouvé. ",
                                stdButtons=QtGui.QMessageBox.Ok)
        return sortedSlots, sortedPriorities
    # ____________________________________________
    def filterOuputs(self):
        """ Filtrage des outputs
            @les rendre compatibles de l'affichage graphique
        """
        try:
            ## | Affectation de la variable |
            self.oFilter["PRIORITIES"] = [[self.ouputs["PRIORITIES"][0]]]
            ## | Mise en place du filtre per day |
            for i in range(len(self.ALL)):
                timeDelta = getTimeDelta(timeStep=self.timeStep,
                                         index=self.ALL[i])
                day = time.strftime("%d", time.localtime(time.mktime(time.localtime()) + timeDelta))
                if i == int(0): #/Initialisation
                    self.oFilter["MEETING_SLOTS"][day] = [self.ALL[i]]
                    pday = day
                else: #/Poursuite de la compilation
                    if day == pday:
                        self.oFilter["MEETING_SLOTS"][day].append(self.ALL[i])
                        self.oFilter["PRIORITIES"][-1].append(self.ouputs["PRIORITIES"][i])
                    else:
                        self.oFilter["MEETING_SLOTS"][day] = [self.ALL[i]]
                        self.oFilter["PRIORITIES"].append([self.ouputs["PRIORITIES"][i]])
                        pday = day
        except Exception as e:
            self.errorFlag == int(4)
            gui.wlib.criticalBox(wdwTitle="Program Error - No matching slots",
                                 wdwIcon="outlook.png",
                                 Txt=u"Aucun créneau commun trouvé pour le mois à venir.",
                                 stdButtons=QtGui.QMessageBox.Ok)
    # ____________________________________________
    def udpateTimeslotUI(self):
        """ Update de la zone d'affichage des timeslots
        """     
        ## | Reset de la variable |
        self.itemsList = []
        ## | Recompilation de la liste des slots du jour sélectionné |
        for item in self.oFilter["MEETING_SLOTS"].values()[self.indexation]:
            timeDelta = getTimeDelta(timeStep=self.timeStep,
                                     index=item)
            nDateTime = datetime.datetime.now() + datetime.timedelta(seconds=timeDelta) 
            i = self.oFilter["MEETING_SLOTS"].values()[self.indexation].index(item)
            self.itemsList.append(gui.wlib.treeItem(column=int(0),
                                                    Title="(priority %s) | %s" % (self.oFilter["PRIORITIES"][self.indexation][i].split("_")[-1], nDateTime)))
        ## | Mise à jour de l'UI |
        self.slotSelection.clear()
        self.slotSelection.addTopLevelItems(self.itemsList)
        for item in self.itemsList:
            if self.maxPriority == int(0):
                item.setIcon(int(0), QtGui.QIcon(r"%s\redFlag.ico" % TOOL_ENV["ICONS"]))
            else:
                if "priority %i" % self.maxPriority in str(item.text(0)):
                    item.setIcon(int(0), QtGui.QIcon(r"%s\greenFlag.ico" % TOOL_ENV["ICONS"]))
                elif "priority 0" in str(item.text(0)):
                    item.setIcon(int(0), QtGui.QIcon(r"%s\redFlag.ico" % TOOL_ENV["ICONS"]))
                else:
                    item.setIcon(int(0), QtGui.QIcon(r"%s\orangeFlag.jpg" % TOOL_ENV["ICONS"]))
        self.calendar_view.setSelectedDate(QtCore.QDate(int(nDateTime.year), int(nDateTime.month), int(nDateTime.day)))
        self.modification = True
        self.slotSelection.setCurrentItem(self.itemsList[0], int(0), QtGui.QItemSelectionModel.SelectCurrent)
    # ____________________________________________
    def getRooms(self,
                 item=int(0)):
        """ Recuperation des salles
        """
        available_rooms = []
        ## | Check parmi les salles favorites |
        for room in self.searchData["favoriteRooms"]:
            for i in range(len(self.olRooms["LISTE"])):
                if room in self.olRooms["LISTE"][i]:
                    if not " xxxx" in room:
                        room = "%s xxx" % room
                    myFBInfo = self.olRooms["FBINFOS"][i][0]
                    tp = []
                    for j in range(self.plage):
                        tp.append(myFBInfo[item+j])
                    dispo = "".join(tp)
                    if dispo == "".join(["0"] * self.plage):
                        available_rooms.append(room)
        ## | Sinon, listing des autres salles disponibles |
        if available_rooms == []:
            gui.wlib.warningBox(wdwTitle="Program Warning - No match room",
                                wdwIcon="outlook.png",
                                Txt=u"Aucune salle disponible parmis les favoris sur ce créneau.\
                                      \nRecherche parmis toutes les salles.",
                                stdButtons=QtGui.QMessageBox.Ok)
            for j in range(len(self.olRooms["FBINFOS"])):
                list = self.olRooms["FBINFOS"][j]
                elem = list[0]
                tp = []
                for i in range(self.plage):
                    tp.append(elem[item+i])
                dispo = "".join(tp)
                if dispo == "".join(["0"] * self.plage):
                    available_rooms.append(self.olRooms["LISTE"][j])
        return available_rooms
    # ____________________________________________  
    def showSlot(self):
        """ Affichage des donnees associees au slot selectionne
        """
        ## | Affichage de la date et heure du slot sélectionné |
        if (self.modification):
            self.itemID = self.itemsList.index(self.slotSelection.selectedItems()[0])
            item = self.oFilter["MEETING_SLOTS"].values()[self.indexation][self.itemID]
            timeDelta = getTimeDelta(timeStep=self.timeStep,
                                     index=item)
            nDateTime = datetime.datetime.now() + datetime.timedelta(seconds=timeDelta)
            self.timeMeeting.setDateTime(nDateTime)
            ## | Affichage des salles disponibles |
            self.roomsList = self.getRooms(item=item)
            self.selectRoom.clear()
            self.selectRoom.addItems(self.roomsList)
            coloringBox(attendeeList=self.searchData["requiredAttendee"][self.oFilter["PRIORITIES"][self.indexation][self.itemID]],
                        roomList=self.roomsList,
                        olRoomList=self.olRooms["LISTE"],
                        capaList=self.olRooms["CAPACITY"],
                        box=self.selectRoom)
            QtCore.QObject.connect(self.filterRoom, QtCore.SIGNAL("returnPressed()"), lambda who=[self.filterRoom, [self.roomsList], [self.selectRoom], False] : self.roomFilter(who))
            ## | Affichage des invités |         
            self.attendeeSummary.clear()
            for elem in self.searchData["requiredAttendee"][self.oFilter["PRIORITIES"][self.indexation][self.itemID]]:
                self.attendeeSummary.addItem(QtCore.QString(elem))
            for elem in self.searchData["optionalAttendee"]:
                self.attendeeSummary.addItem(QtCore.QString(elem))
    # ____________________________________________
    def roomFilter(self,
                   arg):
        """ Fonction de filtre des salles
        """
        ## | Affectation des arguments |
        rooms = {"DEFAULT": arg[1],
                 "FILTERED": []}
        searchBox, comboboxList, period = arg[0], arg[2], arg[3]
        ## | Récupération du filtre renseigné |
        val = u"%s" % str(searchBox.text())
        ## | Listing des salles après application du filtre |
        if (val == u"" or val == u"*" or val == u"**"):
            for list in rooms["DEFAULT"]:
                rooms["FILTERED"].append([])
                for room in list:
                    rooms["FILTERED"][-1].append(room)
        else:
            for list in rooms["DEFAULT"]:
                rooms["FILTERED"].append([])
                for room in list:
                    if (val[0] == u"*" and val[-1] == u"*"):
                        if val[1:-1].lower() in room.lower():
                            rooms["FILTERED"][-1].append(room)
                    elif (val[0] == u"*"):
                        if val[1:].lower() == room[1:len(val)-1].lower():
                            rooms["FILTERED"][-1].append(room)
                    elif (val[-1] == u"*"):
                        if val[:-1].lower() == room[len(val)-1:-1].lower():
                            rooms["FILTERED"][-1].append(room)
                    else:
                        if val.lower() == room.lower():
                            rooms["FILTERED"][-1].append(room)
        ## | Mise à jour des comboBox |
        for i in range(len(comboboxList)):
            if (period):
                if self.periodWdw.widgets[i][0].checkState() == QtCore.Qt.Checked:
                    comboboxList[i].clear()
                    comboboxList[i].addItems(QtCore.QStringList(rooms["FILTERED"][i]))
                    coloringBox(attendeeList=self.searchData["requiredAttendee"][self.oFilter["PRIORITIES"][self.indexation][self.itemID]],
                                roomList=rooms["FILTERED"][i],
                                olRoomList=self.olRooms["LISTE"],
                                capaList=self.olRooms["CAPACITY"],
                                box=comboboxList[i])
            else:
                comboboxList[i].clear()
                comboboxList[i].addItems(QtCore.QStringList(rooms["FILTERED"][i]))
                coloringBox(attendeeList=self.searchData["requiredAttendee"][self.oFilter["PRIORITIES"][self.indexation][self.itemID]],
                            roomList=rooms["FILTERED"][i],
                            olRoomList=self.olRooms["LISTE"],
                            capaList=self.olRooms["CAPACITY"],
                            box=comboboxList[i])
    # ____________________________________________
    def enablePeriod(self,
                     state=int(0)):
        """ Active ou non le mode 'periodicity"
        """
        if state == int(2):
            self.periodicityChoice.setEnabled(True)
            self.periodicityGo.setDisabled(False)
            self.displayMeeting.setDisabled(True)
            # self.sendMeeting.setDisabled(True)
        else:
            self.periodicityChoice.setEnabled(False)
            self.periodicityGo.setDisabled(True)
            self.displayMeeting.setDisabled(False)
            # self.sendMeeting.setDisabled(False)
    # ____________________________________________  
    def getRecursiveRooms(self):
        """ Recuperation des salles en mode periodique
        """
        self.itemID = self.itemsList.index(self.slotSelection.currentItem())
        item = self.oFilter["MEETING_SLOTS"].values()[self.indexation][self.itemID]
        timeDelta = getTimeDelta(timeStep=self.timeStep,
                                 index=item)
        ndate = datetime.datetime.now() + datetime.timedelta(seconds=timeDelta)
        self.listingDate, self.boolList, self.available_rooms = [], [], []
        delta = datetime.timedelta(microseconds=0)
        if str(self.periodicityChoice.currentText()) == "Quotidienne":
            delta = datetime.timedelta(days=1)
        if str(self.periodicityChoice.currentText()) == "Mensuelle":
            delta = datetime.timedelta(weeks=4)
        if str(self.periodicityChoice.currentText()) == "Hebdo":
            delta = datetime.timedelta(weeks=1)
        if str(self.periodicityChoice.currentText()) == "Annuelle":
            delta = datetime.timedelta(weeks=52)
        while ndate + delta < self.olDateList[-1][-1][-1]:
            ndate += delta
            self.listingDate.append(ndate)
        for date in self.listingDate:
            self.boolList.append("0")
            self.available_rooms.append([])
            endloop = False
            for i in range(len(self.olRooms["LISTE"])):
                if u"%s" % str(self.selectRoom.currentText()) in self.olRooms["LISTE"][i]:
                    for j in range(len(self.olDateList[i])):
                        for k in range(len(self.olDateList[i][j])):
                            if self.olDateList[i][j][k] == date: 
                                tp = []
                                for m in range(self.plage):
                                    tp.append(self.olRooms["FBINFOS"][i][j][k+m])
                                dispo = "".join(tp) 
                                if dispo == "".join(["0"] * self.plage):
                                    self.boolList[-1] = "1"
                                endloop = True
                                break
                        if (endloop):
                            break
                if (endloop):
                    break
            if "0" in self.boolList[-1]:
                for i in range(len(self.olRooms["LISTE"])):
                    for j in range(len(self.olDateList[i])):
                        for k in range(len(self.olDateList[i][j])):
                            if self.olDateList[i][j][k] == date:
                                tp = []
                                for m in range(self.plage):
                                    tp.append(self.olRooms["FBINFOS"][i][j][k+m])
                                dispo = "".join(tp)
                                if dispo == "".join(["0"] * self.plage):
                                    self.available_rooms[-1].append(self.olRooms["LISTE"][i])
                                endloop = True
                                break
                        if (endloop):
                            break
        if "0" in self.boolList:
            self.periodWdw = tmpWindow(attendees=self.searchData["requiredAttendee"][self.oFilter["PRIORITIES"][self.indexation][self.itemID]],
                                       olRoomList=self.olRooms["LISTE"],
                                       capaList=self.olRooms["CAPACITY"],
                                       available_rooms=self.available_rooms,
                                       booleanList=self.boolList,
                                       dateList=self.listingDate)
            comboList = []
            for widget in self.periodWdw.widgets:
                comboList.append(widget[2])
            QtCore.QObject.connect(self.periodWdw.filterBox, QtCore.SIGNAL("returnPressed()"), lambda who=[self.periodWdw.filterBox, self.available_rooms, comboList, True] : self.roomFilter(who))
            QtCore.QObject.connect(self.periodWdw.validate, QtCore.SIGNAL("clicked()"), self.periodicitySettings)
            self.periodWdw.exec_()
    # ____________________________________________
    def periodicitySettings(self):
        """ Sauvegarde des parametres de periodicite
        """
        self.dates = [self.timeMeeting.dateTime()]
        self.rooms = [u"%s" % str(self.selectRoom.currentText())]
        for triplet in self.periodWdw.widgets:
            self.dates.append(triplet[1].dateTime())
            self.rooms.append(u"%s" % str(triplet[2].currentText()))
        self.displayMeeting.setDisabled(False)
        # self.sendMeeting.setDisabled(False)
        self.periodWdw.close()
    # ____________________________________________
    def compute(self):
        """ Lancement du processus de recherche
        """
        ## | Reset des variables |
        self.savingName = None
        self.errorFlag = int(0)
        self.maxPriority = int(0)
        self.indexation, self.startId = int(0), int(0)
        self.itemsList = []
        self.modification = False
        self.searchData = OrderedDict([("requiredAttendee", OrderedDict()),
                                       ("optionalAttendee", []),
                                       ("favoriteRooms", []),
                                       ("duration", int(0))])
        self.ouputs = {"MEETING_SLOTS": OrderedDict(),
                       "PRIORITIES": []}
        self.oFilter = {"MEETING_SLOTS": OrderedDict(),
                        "PRIORITIES": None}
        ## | Maximisation de la taille de la fenêtre |
        self.showMaximized()
        ## | Récupération des inputs |
        self.getSearchData()
        ## | Pré-traitement  des inputs |
        if self.errorFlag == int(0):
            ## | Si slots en conflits, affection du niveau de priorité le plus bas |
            self.ALL = self.ouputs["MEETING_SLOTS"].values()[0]
            self.ouputs["PRIORITIES"] = [self.searchData["requiredAttendee"].keys()[0]] * len(self.ouputs["MEETING_SLOTS"].values()[0])
            for key in (x for x in self.searchData["requiredAttendee"].keys() if x != self.searchData["requiredAttendee"].keys()[0]):
                for slot in self.ouputs["MEETING_SLOTS"][key]:
                    if slot in self.ALL:
                        i = self.ALL.index(slot)
                        self.ouputs["PRIORITIES"][i] = key
            ## | Exclusion des slots incompatibles |
            self.ALL, self.ouputs["PRIORITIES"] = self.sortSlots(slotList=self.ALL,
                                                                 priorityList=self.ouputs["PRIORITIES"])
        ## | Filtre des données pour adaptation à l'UI et affichage du visuel | 
        if self.errorFlag == int(0):
            self.filterOuputs()
            self.switch.addWidget(self.meetingDetails)
            self.switch.setCurrentWidget(self.meetingDetails)
            self.udpateTimeslotUI()
            self.seeNextDay.setEnabled(True)
    # ____________________________________________
    def oMeeting(self,
                 display=True):
        """ Envoi du message automatique """
        outlookAppl = win32Module.Dispatch("Outlook.Application")
        format = '%d/%m/%Y %I:%M %p'
        if self.activatePeriodicity.checkState() == QtCore.Qt.Checked:
            for i in range(len(self.dates)):
                oItem = outlookAppl.CreateItem(OUTLOOK_APPOINTMENT_ITEM)
                oItem.MeetingStatus = OUTLOOK_MEETING
                oItem.Subject = "TBD"
                room = str(self.selectRoom.currentText())
                self.itemID = self.itemsList.index(self.slotSelection.currentItem())
                for pers in self.searchData["requiredAttendee"][self.oFilter["PRIORITIES"][self.indexation][self.itemID]]:
                    myRequiredAttendee = oItem.Recipients.Add(pers)
                for pers in self.searchData["optionalAttendee"]:
                    myOptionalAttendee = oItem.Recipients.Add(pers)
                    myOptionalAttendee.Type = OUTLOOK_OPTIONAL_ATTENDEE
                pyTimer = self.dates[i].toPyDateTime()
                oItem.Start = time.strftime(format, time.localtime(time.mktime(pyTimer.timetuple())))
                oItem.Duration = self.searchData["duration"]
                oItem.ReminderMinutesBeforeStart = 15
                myResourceAttendee = oItem.Recipients.Add(self.rooms[i])
                myResourceAttendee.Type = OUTLOOK_RESOURCE_ATTENDEE
                oPattern = oItem.GetRecurrencePattern()
                if i < len(self.dates) - 1:
                    pyTimer = self.dates[i+1].toPyDateTime() - datetime.timedelta(days=1)
                    oPattern.PatternEndDate = time.strftime(format, time.localtime(time.mktime(pyTimer.timetuple())))
                else:
                    pyTimer = self.olDateList[-1][-1][-1]
                    oPattern.PatternEndDate = time.strftime(format, time.localtime(time.mktime(pyTimer.timetuple())))
                oPattern.Interval = 1
                if str(self.periodicityChoice.currentText()) == "Quotidienne":
                    oPattern.RecurrenceType = olRecursDaily
                if str(self.periodicityChoice.currentText()) == "Mensuelle":
                    oPattern.RecurrenceType = olRecursMonthly 
                if str(self.periodicityChoice.currentText()) == "Hebdo":
                    oPattern.RecurrenceType = olRecursWeekly 
                if str(self.periodicityChoice.currentText()) == "Annuelle":
                    oPattern.RecurrenceType = olRecursYearly
                # oItem.Save()
                if (display): oItem.Display()
                elif not (display): oItem.Send()
        else:
            oItem = outlookAppl.CreateItem(OUTLOOK_APPOINTMENT_ITEM)
            oItem.MeetingStatus = OUTLOOK_MEETING
            oItem.Subject = "TBD"
            room = str(self.selectRoom.currentText())
            self.itemID = self.itemsList.index(self.slotSelection.currentItem())
            for pers in self.searchData["requiredAttendee"][self.oFilter["PRIORITIES"][self.indexation][self.itemID]]:
                myRequiredAttendee = oItem.Recipients.Add(pers)
            for pers in self.searchData["optionalAttendee"]:
                myOptionalAttendee = oItem.Recipients.Add(pers)
                myOptionalAttendee.Type = OUTLOOK_OPTIONAL_ATTENDEE
            pyTimer = self.timeMeeting.dateTime()
            pyTimer = pyTimer.toPyDateTime()
            oItem.Start = time.strftime(format, time.localtime(time.mktime(pyTimer.timetuple())))
            oItem.Duration = self.searchData["duration"]
            myResourceAttendee = oItem.Recipients.Add(room)
            myResourceAttendee.Type = OUTLOOK_RESOURCE_ATTENDEE
            oItem.ReminderMinutesBeforeStart = 15
            # oItem.Save()
            if (display): oItem.Display()
            elif not (display): oItem.Send()
    # ____________________________________________
    def center(self):
        """ Recentrage de l'interface
        """
        qr = self.frameGeometry()
        cp = QtGui.QDesktopWidget().availableGeometry().center()
        qr.moveCenter(cp)
        self.move(qr.topLeft())
    # ____________________________________________
    def closeEvent(self,
                   event=None):
        """ Fermeture de la fenetre
        """
        QtCore.QCoreApplication.instance().quit()
# ...]