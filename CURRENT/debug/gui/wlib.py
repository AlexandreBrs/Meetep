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
import cPickle
import datetime
from PyQt4 import QtGui, QtCore
# ...]
## ------------
# [...
""" Importation des modules internes """
from config.env import TOOL_ENV, USER_ENV
# ...]

## =================================================================================================
## MODULE DESCRIPTION
## =================================================================================================
""" Widgets Library Module:
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
__comment__     = "Instanciation des widgets"
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
class criticalBox(QtGui.QMessageBox):
    """ Messages d'erreur
        @param: wdwTitle = Intitule de la fenetre
        @param: wdwIcon = Icone de la fenetre
        @param: Txt = Message affiche dans la fenetre
        @param: stdButtons = Liste des boutons declares
    """
    ident = "criticalBox"
    errorIcon = int(3)
    # ____________________________________________
    def __init__(self,
                 wdwTitle="Error",
                 wdwIcon="outlook.png",
                 Txt="PROGRAM ERROR",
                 stdButtons=QtGui.QMessageBox.Ok):
        """ Initialisation de la classe
        """
        QtGui.QMessageBox.__init__(self)
        QtGui.QToolTip.setFont(QtGui.QFont("SansSerif", int(10)))
        self.setWindowTitle(wdwTitle)
        self.setWindowIcon(QtGui.QIcon(r"%s\%s" % (TOOL_ENV["ICONS"], wdwIcon)))
        self.setIcon(self.errorIcon)
        self.setText(Txt)
        self.setStandardButtons(stdButtons)
        self.exec_()
# ...]
## ------------
# [...
class warningBox(QtGui.QMessageBox):
    """ Messages d'attention
        @param: wdwTitle = Intitule de la fenetre
        @param: wdwIcon = Icone de la fenetre
        @param: Txt = Message affiche dans la fenetre
        @param: stdButtons = Liste des boutons declares
    """
    ident = "warningBox"
    warningIcon = int(2)
    # ____________________________________________
    def __init__(self,
                 wdwTitle="Warning",
                 wdwIcon="outlook.png",
                 Txt="PROGRAM WARNING",
                 stdButtons=QtGui.QMessageBox.Ok):
        """ Initialisation de la classe
        """
        QtGui.QMessageBox.__init__(self)
        QtGui.QToolTip.setFont(QtGui.QFont("SansSerif", int(10)))
        self.setWindowTitle(wdwTitle)
        self.setWindowIcon(QtGui.QIcon(r"%s\%s" % (TOOL_ENV["ICONS"], wdwIcon)))
        self.setIcon(self.warningIcon)
        self.setText(Txt)
        self.setStandardButtons(stdButtons)
        self.exec_()
# ...]
## ------------
# [...
class infoBox(QtGui.QMessageBox):
    """ Messages d'information
        @param: wdwTitle = Intitule de la fenetre
        @param: wdwIcon = Icone de la fenetre
        @param: Txt = Message affiche dans la fenetre
        @param: stdButtons = Liste des boutons declares
    """
    ident = "infoBox"
    infoIcon = int(1)
    # ____________________________________________
    def __init__(self,
                 wdwTitle="Information",
                 wdwIcon="outlook.png",
                 Txt="PROGRAM INFORMATION",
                 stdButtons=QtGui.QMessageBox.Ok):
        """ Initialisation de la classe
        """
        QtGui.QMessageBox.__init__(self)
        QtGui.QToolTip.setFont(QtGui.QFont("SansSerif", int(10)))
        self.setWindowTitle(wdwTitle)
        self.setWindowIcon(QtGui.QIcon(r"%s\%s" % (TOOL_ENV["ICONS"], wdwIcon)))
        self.setIcon(self.infoIcon)
        self.setText(Txt)
        self.setStandardButtons(stdButtons)
        self.exec_()
# ...]
## ------------
# [...
class questionBox(QtGui.QMessageBox):
    """ Questions pour l'utilisateur
        @param: wdwTitle = Intitule de la fenetre
        @param: wdwIcon = Icone de la fenetre
        @param: Txt = Message affiche dans la fenetre
        @param: stdButtons = Liste des boutons declares
    """
    ident = "questionBox"
    questionIcon = int(4)
    # ____________________________________________
    def __init__(self,
                 wdwTitle="Question",
                 wdwIcon="outlook.png",
                 Txt="PROGRAM INTERROGATION",
                 stdButtons=QtGui.QMessageBox.Yes|QtGui.QMessageBox.No):
        """ Initialisation de la classe
        """
        QtGui.QMessageBox.__init__(self)
        QtGui.QToolTip.setFont(QtGui.QFont("SansSerif", int(10)))
        self.setWindowTitle(wdwTitle)
        self.setWindowIcon(QtGui.QIcon(r"%s\%s" % (TOOL_ENV["ICONS"], wdwIcon)))
        self.setIcon(self.questionIcon)
        self.setText(Txt)
        self.setStandardButtons(stdButtons)
        self.exec_()
# ...]
## ------------
# [...
class progressBar(QtGui.QProgressBar):
    """ Statut du suivi de progression
        @param: min/max = Bornes mini/maxi de progression
    """
    ident = "progressBar"
    # ____________________________________________
    def __init__(self,
                 min=int(0),
                 max=int(0)):
        """ Initialisation de la classe
        """
        QtGui.QProgressBar.__init__(self)
        self.setRange(min, max)
        self.setValue(min)
# ...]
## ------------
# [...
class pushButton(QtGui.QPushButton):
    """ Push Button
        @param: dim = Dimensions de l'icone
        @param: icon = Lien vers l'icone a inserer
        @param: Txt = Intitule du button
        @param: toolTip = Etiquette descriptive
        @param: setDisabled = Booleen d'activation du widget
    """
    ident = "pushButton"
    # ____________________________________________
    def __init__(self,
                 dim=None,
                 icon=None,
                 Txt="",
                 toolTip=None,
                 setDisabled=False):
        """ Initialisation de la classe
        """
        QtGui.QPushButton.__init__(self)
        self.setText(Txt)
        if icon is not None:
            self.setIcon(QtGui.QIcon(r"%s\%s" % (TOOL_ENV["ICONS"], icon)))
            self.setIconSize(QtCore.QSize(dim, dim))
        if toolTip is not None:
            self.setToolTip(toolTip)
        if (setDisabled):
            self.setDisabled(True)
# ...]
## ------------
# [...
class comboBox(QtGui.QComboBox):
    """ Menu deroulant
        @param: comboList = Liste a faire figurer
        @param: TxtToSelect = Texte a selectionner
        @param: setEditable = Booleen rendant le widget editable
        @param: adjustSizeActive = Taille adaptable au contenu
        @param: setEnabled = Booleen d'activation du widget
    """
    ident = "comboBox"
    # ____________________________________________
    def __init__(self,
                 comboList=[],
                 TxtToSelect="",
                 setEditable=False,
                 adjustSizeActive=False,
                 setEnabled=True):
        """ Initialisation de la classe
        """
        QtGui.QComboBox.__init__(self)
        self.addItems(comboList)
        if (setEditable):
            self.setEditable(True)
            if (TxtToSelect != "" and not TxtToSelect in comboList):
                self.insertItem(self.count(), TxtToSelect)
            self.setEditText(TxtToSelect)
            self.setCurrentIndex(self.findText(TxtToSelect))
        elif not (setEditable):
            if (TxtToSelect in comboList or TxtToSelect == ""):
                self.setCurrentIndex(self.findText(TxtToSelect))
            else:
                self.setCurrentIndex(int(0))
                warningBox(wdwTitle=u"Program Warning - Donnée incompatible",
                           wdwIcon="outlook.png",
                           Txt=u"Choix défini inexistant dans le menu déroulant (1er item sélectionné par défaut).",
                           stdButtons=QtGui.QMessageBox.Ok)
        if (adjustSizeActive):
            self.setSizeAdjustPolicy(self.AdjustToContents)
        if not (setEnabled):
            self.setEnabled(False)
# ...]
## ------------
# [...
class frame(QtGui.QFrame):
    """ Edition d'une Frame
    """
    ident = "frame"
    # ____________________________________________
    def __init__(self):
        """ Initialisation de la classe
        """	
        QtGui.QFrame.__init__(self)
        self.setFrameShape(QtGui.QFrame.StyledPanel)
# ...]
## ------------
# [...
class calendar(QtGui.QCalendarWidget):
    """ Edition d'un calendrier
        @param: displayNavBar = Booleen affichant la barre de navigation
        @param: enableSelection = Booleen autorisant la selection d'une date dans le calendrier
    """
    ident = "calendar"
    # ____________________________________________
    def __init__(self,
                 displayNavBar=True,
                 enableSelection=False):
        """ Initialisation de la classe
        """
        QtGui.QCalendarWidget.__init__(self)
        if not (enableSelection):
            self.setSelectionMode(QtGui.QCalendarWidget.NoSelection)
        if not (displayNavBar):
            self.setNavigationBarVisible(False)
# ...]
## ------------
# [...
class timeEditer(QtGui.QTimeEdit):
    """ Timing edition
        @param: defaultStep = Step par defaut
        @param: timeStep = Duree mini (en minutes) de la reunion
    """
    ident = "timeEditer"
    # ____________________________________________
    def __init__(self,
                 time=int(0),
                 timeStep=int(30)):
        """ Initialisation de la classe
        """
        QtGui.QTimeEdit.__init__(self)
        self.timeStep = timeStep
        self.setTime(QtCore.QTime().addSecs(time))
        self.connect(self, QtCore.SIGNAL("timeChanged()"), self.stepBy)
        self.connect(self, QtCore.SIGNAL("editingFinished()"), self.applyCorrection)
    # ____________________________________________
    def stepBy(self,
               defaultStep=int(1)):
        """ Update de la duree par session de 'timeStep' minutes
        """
        if defaultStep == int(1):
            self.updateTime = self.time().addSecs(defaultStep * self.timeStep * int(60))
        elif defaultStep == int(-1):
            self.updateTime = self.time().addSecs(defaultStep * self.timeStep * int(60))
        self.setTime(self.updateTime)
    # ____________________________________________
    def applyCorrection(self):
        """ Correction de la duree renseignee au clavier
        """
        ## | Récupération de la durée actuellement renseignée dans le widget |
        currentTime = {"MINUTES": int(self.sectionText(QtGui.QDateTimeEdit.MinuteSection)),
                       "SECONDES": int(self.sectionText(QtGui.QDateTimeEdit.SecondSection))}
        ## | Reset des secondes |
        if currentTime["SECONDES"] != int(0):
            self.updateTime = self.time().addSecs(int(-1) * currentTime["SECONDES"])
            self.setTime(self.updateTime)
        ## | Si nécessaire, calcul du correctif pour assurer la cohérence avec la valeur de 'timeStep' définie |
        if (currentTime["MINUTES"] % self.timeStep) != int(0):
            if (currentTime["MINUTES"] % self.timeStep - self.timeStep / int(2)) >= int(0):
                coeff = currentTime["MINUTES"] // self.timeStep + int(1)
            else:
                coeff = currentTime["MINUTES"] // self.timeStep - int(1)
            newTime = (int(coeff * self.timeStep) - currentTime["MINUTES"]) * int(60)
            self.updateTime = self.time().addSecs(newTime)
            self.setTime(self.updateTime)
# ...]
## ------------
# [...
class datetimeEditer(QtGui.QDateTimeEdit):
    """ Datetime edition
        @param: dateTime = Valeur a rentrer dans le widget
        @param: calendar = Calendrier associe
        @param: setDisabled = Booleen d'activation du widget
    """
    ident = "datetimeEditer"
    # ____________________________________________
    def __init__(self,
                 dateTime=datetime.datetime.now(),
                 calendar=None,
                 setDisabled=True):
        """ Initialisation de la classe
        """
        QtGui.QDateTimeEdit.__init__(self)
        self.setDateTime(dateTime)
        self.setCalendarPopup(True)
        if calendar is not None:
            self.setCalendarWidget(calendar)
        if (setDisabled):
            self.setDisabled(True)
# ...]
## ------------
# [...
class checkBox(QtGui.QCheckBox):
    """ Edition d'un checkBox
        @param: Txt = Intitule du widget
        @param: toolTip = Etiquette descriptive
        @param: setDisabled = Booleen d'activation du widget
    """
    ident = "checkBox"
    # ____________________________________________
    def __init__(self,
                 Txt="",
                 toolTip=None,
                 setDisabled=False):
        """ Initialisation de la classe
        """
        QtGui.QCheckBox.__init__(self)
        self.setText(Txt)
        if toolTip is not None:
            self.setToolTip(toolTip)
        if (setDisabled):
            self.setDisabled(True)
# ...]
## ------------
# [...
class groupBox(QtGui.QGroupBox):
    """ Edition d'un groupbox
        @param: Title = Intitule du groupBox
        @param: setFlat = Booleen activant la visualisation en relief du widget
    """
    ident = "groupBox"
    # ____________________________________________
    def __init__(self,
                 Title="",
                 setFlat=False):
        """ Initialisation de la classe
        """
        QtGui.QGroupBox.__init__(self)
        self.setTitle(Title)
        if (setFlat):
            self.setFlat(True)
# ...]
## ------------
# [...
class fileExplorer(QtGui.QPushButton):
    """ Explorateur windows de fichiers
        @param: dim = Dimensions de l'icone
        @param: icon = Lien vers l'icone a inserer
        @param: Txt = Intitule du button
        @param: openPath = Chemin vers lequel pointe l'explorateur
        @param: zoneTxt = Objet receveur du chemin selectionne
        @param: toolTip = Etiquette descriptive
        @param: setEnabled = Booleen d'activation du widget
    """
    ident = "fileExplorer"
    # ____________________________________________
    def __init__(self,
                 dim=int(15),
                 icon="folder.png",
                 Txt="",
                 openPath=USER_ENV["ROOT"],
                 zoneTxt=None,
                 toolTip=u"Sélectionner un fichier ...",
                 setEnabled=True):
        """ Initialisation de la classe
        """
        QtGui.QPushButton.__init__(self)
        self.setText(Txt)
        self.openPath = openPath
        self.zoneTxt = zoneTxt
        if toolTip is not None:
            self.setToolTip(toolTip)
        if icon is not None:
            self.setIcon(QtGui.QIcon(r"%s\%s" % (TOOL_ENV["ICONS"], icon)))
            self.setIconSize(QtCore.QSize(dim, dim))
        if not (setEnabled):
            self.setEnabled(False)
# ...]
## ------------
# [...
class lineEdit(QtGui.QLineEdit):
    """ Ligne d'edition
        @param: icon = Lien vers l'icone a inserer
        @param: Txt = Intitule du lineEdit
        @param: keyWord = Mot cle d'identification
        @param: acceptDrops = Booleen d'activation drag & drop
        @param: toolTip = Etiquette descriptive
        @param: setEnabled = Booleen d'activation du widget
    """
    ident = "lineEdit"
    # ____________________________________________
    def __init__(self,
                 Txt="",
                 keyWord="",
                 toolTip=None,
                 acceptDrops=True,
                 setEnabled=True):
        """ Initialisation de la classe
        """
        QtGui.QLineEdit.__init__(self)
        self.key = keyWord
        if (os.path.isfile(Txt) and ".py" in os.path.splitext(Txt)):
            sys.path.append(os.path.dirname(Txt))
            tmp = __builtin__.__import__(os.path.basename(os.path.splitext(Txt)[0]),
                                         globals(),
                                         locals(),
                                         [],
                                         -1)
            if (tmp.__KeyWord__).lower() == self.key:
                self.setText(Txt)
            else:
                criticalBox(wdwTitle="Program Error - Reading file failure",
                            wdwIcon="outlook.png",
                            Txt=u"Incompatibilité du lien sélectionné.",
                            stdButtons=QtGui.QMessageBox.Ok)
        elif (os.path.isfile(Txt)):
            criticalBox(wdwTitle="Program Error - Reading file failure",
                        wdwIcon="meoutlook.png",
                        Txt=u"Incompatibilité du lien sélectionné\
                              \n(ne correspond pas au format python d'un fichier de données).",
                        stdButtons=QtGui.QMessageBox.Ok)
        if toolTip is not None:
            self.setToolTip(toolTip)
        if (acceptDrops):
            self.setAcceptDrops(True)
        if not (setEnabled):
            self.setEnabled(False)
    # ____________________________________________
    def dragMoveEvent(self,
                      event):
        if (event):
            event.acceptProposedAction()
    # ____________________________________________
    def dragEnterEvent(self,
                       event):
        if (event):
            self.setFrame(False)
            event.acceptProposedAction()
    # ____________________________________________
    def dragLeaveEvent(self,
                       event):
        if (event):
            self.setFrame(True)
    # ____________________________________________
    def dropEvent(self,
                  event):
        Data = event.mimeData()
        Url = Data.urls()[0]
        self.setFrame(True)
        try:
            Path = r"%s" % Url.toLocalFile().replace("/", os.sep)
        except Exception as e:
            pass
        else:
            if (os.path.isfile(Path) and ".py" in os.path.splitext(Path)):
                sys.path.append(os.path.dirname(Path))
                tmp = __builtin__.__import__(os.path.basename(os.path.splitext(Path)[0]),
                                             globals(),
                                             locals(),
                                             [],
                                             -1)
                if (tmp.__KeyWord__).lower() == self.key:
                    self.selectAll()
                    self.del_()
                    self.setText(Path)
                else:
                    criticalBox(wdwTitle="Program Error - Reading file failure",
                                wdwIcon="outlook.png",
                                Txt=u"Incompatibilité du lien sélectionné.",
                                stdButtons=QtGui.QMessageBox.Ok)
            elif (os.path.isfile(Path)):
                criticalBox(wdwTitle="Program Error - Reading file failure",
                            wdwIcon="outlook.png",
                            Txt=u"Incompatibilité du lien sélectionné\
                                  \n(ne correspond pas au format python d'un fichier de données).",
                            stdButtons=QtGui.QMessageBox.Ok)
# ...]
## ------------
# [...
class treeItem(QtGui.QTreeWidgetItem):
    """ Edition d'un item d'arborescence
        @param: column = Position de l'item
        @param: Title = Intitule de l'item
    """
    ident = "treeItem"
    # ____________________________________________
    def __init__(self,
                 column=int(0),
                 Title=""):
        """ Initialisation de la classe
        """
        QtGui.QTreeWidgetItem.__init__(self)
        self.setText(column, Title)
# ...]
## ------------
# [...
class treeWidget(QtGui.QTreeWidget):
    """ Edition d'une arborescence
        @param: column = Position de l'item
        @param: Title = Intitule de l'item
    """
    ident = "treeWidget"
    # ____________________________________________
    def __init__(self,
                 column=int(0),
                 columnNb=int(1),
                 headerLabel="",
                 itemList=[]):
        """ Initialisation de la classe
        """
        QtGui.QTreeWidget.__init__(self)
        self.setColumnCount(columnNb)
        self.setHeaderItem(treeItem(column, headerLabel))
        self.headerItem().setFont(int(0), QtGui.QFont("SansSerif", int(8), italic=True))
        self.addTopLevelItems(itemList)
# ...]
## ------------
# [...
class listWidget(QtGui.QListWidget):
    """ Edition d'un listing
        @param: itemsList = Liste des items
        @param: keyWord = Mot cle d'identification
        @param: acceptDrops = Booleen d'activation drag & drop
    """
    ident = "listWidget"
    # ____________________________________________
    def __init__(self,
                 itemsList=[],
                 keyWord="",
                 acceptDrops=True):
        """ Initialisation de la classe
        """
        QtGui.QListWidget.__init__(self)
        self.key = keyWord
        self.addItems(itemsList)
        self.addContxtMenu()
        self.itemDoubleClicked.connect(self.modifyItem)
        if (acceptDrops):
            self.setAcceptDrops(True)
    # ____________________________________________
    def addContxtMenu(self):
        """ Insertion d'un menu contextuel
        """
        self.contextMenu = QtGui.QMenu(self)
        self.setContextMenuPolicy(QtCore.Qt.ActionsContextMenu)
        self.addElement = QtGui.QAction(QtGui.QIcon(r"%s\add.png" % TOOL_ENV["ICONS"]),
                                        "Add Element",
                                        self.contextMenu)
        self.deleteElement = QtGui.QAction(QtGui.QIcon(r"%s\delete.png" % TOOL_ENV["ICONS"]),
                                           "Delete Element",
                                           self.contextMenu)
        self.cleanAll = QtGui.QAction(QtGui.QIcon(r"%s\clean.png" % TOOL_ENV["ICONS"]),
                                      "Clean all Element",
                                      self.contextMenu)
        actionList = {"WIDGETS": [self.addElement, self.deleteElement, self.cleanAll],
                      "SLOTS": [self.add, self.delete, self.clear]}
        for i in range(len(actionList["WIDGETS"])):
            self.addAction(actionList["WIDGETS"][i])
            self.connect(actionList["WIDGETS"][i], QtCore.SIGNAL("triggered()"), actionList["SLOTS"][i])
    # ____________________________________________
    def add(self,
            defaultRow=int(0)):
        """ Ajouter un element
        """
        self.insertItem(defaultRow, QtCore.QString("new element"))
        self.item(defaultRow).setFlags(self.item(defaultRow).flags() | QtCore.Qt.ItemIsEditable)
        self.editItem(self.item(defaultRow))
    # ____________________________________________
    def delete(self):
        """ Supprimer l'element selectionne
        """
        self.takeItem(self.currentRow())
    # ____________________________________________  
    def modifyItem(self):
        """ Donne la main a l'utilisateur pour modifier l'item
        """
        self.currentItem().setFlags(self.currentItem().flags() | QtCore.Qt.ItemIsEditable)
        self.editItem(self.currentItem())
    # ____________________________________________
    def dragMoveEvent(self,
                      event):
        if (event): event.acceptProposedAction()
    # ____________________________________________
    def dragEnterEvent(self,
                       event):
        if (event):
            event.acceptProposedAction()
    # ____________________________________________
    def dropEvent(self,
                  event):
        Data = event.mimeData()
        Url = Data.urls()[0]
        try:
            Path = r"%s" % Url.toLocalFile().replace("/", os.sep)
        except Exception as e:
            pass
        else:
            if (os.path.isfile(Path) and ".py" in os.path.splitext(Path)):
                sys.path.append(os.path.dirname(Path))
                dropInfos = __builtin__.__import__(os.path.basename(os.path.splitext(Path)[0]),
                                                   globals(),
                                                   locals(),
                                                   [],
                                                   -1)
                if (dropInfos.__KeyWord__).lower() == self.key:
                    self.clear()
                    self.addItems(QtCore.QStringList(dropInfos.list))
                else:
                    criticalBox(wdwTitle="Program Error - Reading file failure",
                                wdwIcon="outlook.png",
                                Txt=u"Incompatibilité du lien sélectionné.",
                                stdButtons=QtGui.QMessageBox.Ok)
            elif (os.path.isfile(Path)):
                criticalBox(wdwTitle="Program Error - Reading file failure",
                            wdwIcon="outlook.png",
                            Txt=u"Incompatibilité du lien sélectionné\
                                  \n(ne correspond pas au format python d'un fichier de données).",
                            stdButtons=QtGui.QMessageBox.Ok)
# ...]
## ------------
# [...
class tableWidget(QtGui.QTableWidget):
    """ Edition d'une table
        @param: nRows = Nombre de lignes
        @param: headersList = Intitules des colonnes
        @param: itemsList = Liste des items
        @param: keyWord = Mot cle d'identification
        @param: acceptDrops = Booleen d'activation drag & drop
    """
    ident = "tableWidget"
    # ____________________________________________
    def __init__(self,
                 nRows=int(0),
                 headersList=[u"Nom Prénom", u"Priorité"],
                 itemsList=[],
                 keyWord="required",
                 acceptDrops=True):
        """ Initialisation de la classe
        """
        QtGui.QTableWidget.__init__(self)
        self.adjustSize()
        self.key = keyWord
        self.setRowCount(nRows)
        self.setColumnCount(len(headersList))
        for i in range(len(headersList)):
            self.setHorizontalHeaderItem(i, QtGui.QTableWidgetItem(QtCore.QString(headersList[i])))
        for i in range(nRows):
            self.setItem(i, int(0), QtGui.QTableWidgetItem(QtCore.QString(str(itemsList[i][0]))))
            self.setItem(i, int(1), QtGui.QTableWidgetItem(QtCore.QString(str(itemsList[i][1]))))
        self.addContxtMenu()
        if (acceptDrops):
            self.setAcceptDrops(True)
    # ____________________________________________
    def addContxtMenu(self):
        """ Insertion d'un menu contextuel
        """
        self.contextMenu = QtGui.QMenu(self)
        self.setContextMenuPolicy(QtCore.Qt.ActionsContextMenu)
        self.addElement = QtGui.QAction(QtGui.QIcon(r"%s\add.png" % TOOL_ENV["ICONS"]),
                                        "Add Element",
                                        self.contextMenu)
        self.deleteElement = QtGui.QAction(QtGui.QIcon(r"%s\delete.png" % TOOL_ENV["ICONS"]),
                                           "Delete Element",
                                           self.contextMenu)
        self.cleanAll = QtGui.QAction(QtGui.QIcon(r"%s\clean.png" % TOOL_ENV["ICONS"]),
                                      "Clean all Element",
                                      self.contextMenu)
        actionList = {"WIDGETS": [self.addElement, self.deleteElement, self.cleanAll],
                      "SLOTS": [self.add, self.delete, self.clear]}
        for i in range(len(actionList["WIDGETS"])):
            self.addAction(actionList["WIDGETS"][i])
            self.connect(actionList["WIDGETS"][i], QtCore.SIGNAL("triggered()"), actionList["SLOTS"][i])
    # ____________________________________________
    def add(self,
            defaultRow=int(0)):
        """ Ajouter un element
        """
        self.insertRow(defaultRow)
        self.setItem(int(0), int(1), QtGui.QTableWidgetItem(QtCore.QString("0")))
    # ____________________________________________
    def delete(self):
        """ Supprimer l'element selectionne
        """
        self.removeRow(self.currentRow())
    # ____________________________________________
    def dragMoveEvent(self,
                      event):
        if (event):
            event.acceptProposedAction()
    # ____________________________________________
    def dragEnterEvent(self,
                       event):
        if (event):
            event.acceptProposedAction()
    # ____________________________________________
    def dropEvent(self,
                  event):
        Data = event.mimeData()
        Url = Data.urls()[0]
        try:
            Path = r"%s" % Url.toLocalFile().replace("/", os.sep)
        except Exception as e:
            pass
        else:
            if (os.path.isfile(Path) and ".py" in os.path.splitext(Path)):
                sys.path.append(os.path.dirname(Path))
                dropInfos = __builtin__.__import__(os.path.basename(os.path.splitext(Path)[0]),
                                                   globals(),
                                                   locals(),
                                                   [],
                                                   -1)
                if (dropInfos.__KeyWord__).lower() == self.key:
                    self.clearContents()
                    self.setRowCount(len(dropInfos.list))
                    for elem in dropInfos.list:
                        self.setItem(dropInfos.list.index(elem), int(0), QtGui.QTableWidgetItem(QtCore.QString(str(elem[0]))))
                        self.setItem(dropInfos.list.index(elem), int(1), QtGui.QTableWidgetItem(QtCore.QString(str(elem[1]))))
                else:
                    criticalBox(wdwTitle="Program Error - Reading file failure",
                                wdwIcon="outlook.png",
                                Txt=u"Incompatibilité du lien sélectionné.",
                                stdButtons=QtGui.QMessageBox.Ok)
            elif (os.path.isfile(Path)):
                criticalBox(wdwTitle="Program Error - Reading file failure",
                            wdwIcon="outlook.png",
                            Txt=u"Incompatibilité du lien sélectionné\
                                  \n(ne correspond pas au format python d'un fichier de données).",
                            stdButtons=QtGui.QMessageBox.Ok)
# ...]
## ------------
# [...
class completer(QtGui.QCompleter):
    """ Edition d'un completer
        @param: maxItems = Nombre maximaux d'items visibles lors de l'auto completion
    """
    ident = "completer"
    # ____________________________________________
    def __init__(self,
                 maxItems=int(15)):
        """ Initialisation de la classe
        """
        QtGui.QCompleter.__init__(self)
        self.setMaxVisibleItems(maxItems)
## ------------
# [...
class filterBox(QtGui.QLineEdit):
    """ Edition d'un filtre
    """
    ident = "filterBox"
    path = r"%s\autoSearchCompleter.p" % USER_ENV["PREF"]
    # ____________________________________________
    def __init__(self):
        """ Initialisation de la classe
        """
        QtGui.QLineEdit.__init__(self)
        self.customizeWidget()
        self.modelList = self.loadModel()
        self.completer = completer(maxItems=int(15))
        self.applyAutoCompletion()
        self.connect(self, QtCore.SIGNAL("returnPressed()"), self.updateAutoCompleter)
    # ____________________________________________
    def customizeWidget(self,
                        dim=int(15),
                        toolTip=u"Filtre sur les salles de réunion",
                        icon="search.png"):
        """ Customisation de la search bar
        """
        Pixmap = QtGui.QLabel()
        Pixmap.setPixmap(QtGui.QPixmap(r"%s\%s" % (TOOL_ENV["ICONS"], icon)).scaled(QtCore.QSize(dim, dim)))
        layout = QtGui.QHBoxLayout(self)
        layout.setAlignment(QtCore.Qt.AlignVCenter)
        layout.setContentsMargins(5, 0, 0, 0)
        layout.addWidget(Pixmap, 0, QtCore.Qt.AlignLeft)
        self.setTextMargins(20, 0, 0, 0)
        if toolTip is not None:
            self.setToolTip(toolTip)
    # ____________________________________________
    def loadModel(self):
        """ Chargement des parametres initiaux de l'auto completion
        """
        if os.path.isfile(self.path):
            fic = open(self.path, "rb")
            modelList = cPickle.load(fic)
            fic.close()
        else:
            modelList = []
        return modelList
    # ____________________________________________
    def applyAutoCompletion(self):
        """ Ajout d'un module d'auto completion
        """
        model = QtGui.QStringListModel()
        model.setStringList(self.modelList)
        self.completer.setModel(model)
        self.setCompleter(self.completer)
    # ____________________________________________
    def updateAutoCompleter(self):
        ## | Update de l'auto complétion |
        if not self.text() in self.modelList:
            self.modelList.append(self.text())
        model = QtGui.QStringListModel()
        model.setStringList(self.modelList)
        self.completer.setModel(model)
        self.setCompleter(self.completer)
        ## | Sauvegarde de la liste à jour |
        fic = open(self.path, "wb")
        cPickle.dump(self.modelList, fic)
        fic.close()
# ...]