#!/usr/bin/env python

'''
--
TO RUN FROM COMMAND LINE:
>> python EchoviewExport.py
--

Echoview Exporter - A GUI based tool for batch exporting of acoustic results
                    from Echoview using COM scripting.  The scripting module
                    for Echoview is required for the operation of the app.

*****************************************************************************
* THIS VERSION OF ECHOVIEW EXPORTER IS DESIGNED FOR USE WITH NOAA/AFSC/MACE *
* MACEBASE2 DATABASE AND THEIR EXPECTED EXPORT FORMATTING.  OTHER USES OF   *
* THIS EXPORT TOOL FOR GENERAL PROCESSING OF ECHOVIEW EXPORTS REQUIRE USER  *
* MODIFICATION.  THIS VERSION IS FOR MB2 ONLY!                              *
*****************************************************************************

The following is the script to run the Echoview exporting application using the
Qt designed UI file (currently 'ui_EchoviewExporter.py').  The
export defaults are set in the library for the UI file for each field season based around
the parameters used aboard the FSV Oscar Dyson during the Exporter EBS/GOA Summer Pollock
Surveys.  Future iterations will rely on an external configuration file to provide the
defaults, or direct communication with the information stored within the database.

The GUI allows for export of single variable, or 'multifrequency export' commonly
used to export the data for krill analyses.

An in depth description of the operation is available in the accompanying manual
'Echoview Exporting GUI Procedures and Operation' (R. Levine, 2014)

--
Required accompanying python scripts:
- ui_EchoviewExporter.py  (contains the GUI interface)

Required python modules:
PyQt4, numpy, glob, sys, os, win32com(pywin32)
--

created: 10 Jul 2014 robert.levine
edited: 14 Oct 2016 rick.towler
last edited: 20 Oct 2017 nathan.lauffenburger- removed dependency on Lib file, drawing most fields from database, added multiple zone thicknesses
'''

from PyQt6 import QtCore,  QtGui,  QtWidgets 
from ui import ui_EchoviewExporter
from MaceFunctions import connectdlg, dbConnection
import numpy
import glob
import win32com.client
import sys, os

class Exporter(QtWidgets.QDialog, ui_EchoviewExporter.Ui_ExportDialog):

    def __init__(self, odbcSource, username, password,  acoustic_schema, bio_schema, parent=None):
        super(Exporter, self).__init__(parent)
        self.setupUi(self)

        #  store the variables we passed into our init method
        self.odbc = odbcSource
        self.dbUser = username
        self.dbPassword = password
        self.bioSchema = bio_schema
        self.acousticSchema = acoustic_schema

        # Find last saved settings - place these in the appropriate places
        # This is an update to using a LIB file for loading all these parameters
        self.appSettings = QtCore.QSettings('afsc.noaa.gov', 'Exporter')
        size = self.appSettings.value('winsize', QtCore.QSize(948,646))
        position = self.appSettings.value('winposition', QtCore.QPoint(10,10))
        self.latestShip = self.appSettings.value('latestShip','')
        self.latestSurvey = self.appSettings.value('latestSurvey','')
        self.latestDataSet = self.appSettings.value('latestDataSet','')
        self.input_dir.insert(self.appSettings.value('latestInDir',''))
        self.output_dir_mb2.insert(self.appSettings.value('latestOutDir',''))
        self.latestRawDir = self.appSettings.value('latestRawDir','')
        if self.latestRawDir!='':
            self.rawFilesDir.insert(self.latestRawDir)
            self.setRawFiles.setChecked(True)
        self.cal_file.insert(self.appSettings.value('latestCalFile',''))
        self.latestFileSet = self.appSettings.value('latestFileSet','')
        if self.latestFileSet!='':
            self.fileset_name.insert(self.latestFileSet)
        else:
            self.fileset_name.insert('Fileset1')

        #  check the current position and size to make sure the app is on the screen
        position, size = self.checkWindowLocation(position, size)

        #  now move and resize the window
        self.move(position)
        self.resize(size)

        #  set the application icon
        try:
            iconImage = QtGui.QIcon('./resources/export-icon.png')
            self.setWindowIcon(iconImage)
        except:
            pass

        #Assign button controls.  Buttons simply assign text to fields.  text is assigned to variable in export function.
        self.input_button.clicked.connect(self.getInputDirectory)
        self.output_button_mb2.clicked.connect(self.getOutputDirectoryMb2)
        self.cal_button.clicked.connect(self.getCalFile)
        self.Cancel.clicked.connect(self.quit)
        self.Export.clicked.connect(self.export)
        self.setRawFiles.stateChanged[int].connect(self.selectRaw)
        self.rawfiles_button.clicked.connect(self.getRawFilesDirectory)
        self.maxThresholdCheck.stateChanged[int].connect(self.threshOnOff)
        self.minThresholdCheck.stateChanged[int].connect(self.threshOnOff)
        self.shipBox.activated[int].connect(self.getSurveys)
        self.surveyBox.activated[int].connect(self.getDataSets)
        self.dataSetBox.activated[int].connect(self.getExportParameters)
        self.editAllBox.stateChanged[int].connect(self.updateEdits)
        self.mfCheckBox.stateChanged[int].connect(self.activateMF)

        self.zoneCheckBoxes=[self.z0ex, self.z1ex, self.z2ex, self.z3ex, self.z4ex, self.z5ex,  self.z6ex, self.z7ex]
        self.layerThicknessList=[self.lineEdit_0, self.lineEdit_1, self.lineEdit_2, self.lineEdit_3, self.lineEdit_4, self.lineEdit_5, self.lineEdit_6, self.lineEdit_7]
        self.upNamesList=[self.upper_0, self.upper_1, self.upper_2, self.upper_3, self.upper_4, self.upper_5, self.upper_6, self.upper_7]
        self.lowNamesList=[self.lower_0, self.lower_1, self.lower_2, self.lower_3, self.lower_4, self.lower_5, self.lower_6, self.lower_7]
        for box in self.zoneCheckBoxes:
            box.stateChanged[int].connect(self.checkZones)
        # Export command executed when 'Export' button on either 'Echoview Export' or 'Echoview Export Options'
        # is clicked.  Export for single variable.
        self.selectRaw()

        # Set up fonts for later
        self.myboldFont=QtGui.QFont()
        self.mynormalFont=QtGui.QFont()
        self.myboldFont.setBold(True)
        self.mynormalFont.setBold(False)

        timer = QtCore.QTimer(self)
        timer.setSingleShot(True)
        timer.timeout.connect(self.applicationInit)
        timer.start(0)

    def applicationInit(self):
        self.db=None

        #  check if we're missing any of our required connection parameters
        if ((self.odbc == None) or (self.dbUser == None) or
            (self.dbPassword == None)):

            #  we're missing at least one - display the connect dialog to get the rest of the args.
            #  Note the use of the new createConnection argument which keeps ConnectDlg from creating
            #  an instance of dbConnection. We'll do that below.
            #  Also note the new enableBioschema argument which will disable the bioschema combobox
            connectDlg = connectdlg.ConnectDlg(self.odbc, self.dbUser, self.dbPassword, label='Exporter',
                    bioSchema=self.bioSchema, enableBioschema=False, createConnection=False, parent=self)

            if not connectDlg.exec():
                #  user hit cancel so we exit this example program
                self.close()
                return

            #  update our connection credentials
            self.odbc = connectDlg.getSource()
            self.dbUser = connectDlg.getUsername()
            self.dbPassword = connectDlg.getPassword()
            self.bioSchema = connectDlg.getBioSchema()

        #  create the database connection
        self.db = dbConnection.dbConnection(self.odbc, self.dbUser,
                self.dbPassword, 'EchoviewExport')

        #  store the bioSchema and acousticSchema in the db object
        self.db.bioSchema=self.bioSchema
        self.db.acousticSchema = self.acousticSchema

        try:
            self.db.dbOpen()
        except dbConnection.DBError as e:
            #  ooops, there was a problem
            errorMsg = ('Unable to connect to ' + self.dbUser+ '@' +
                    self.odbc + '\n' + e.error)
            QtWidgets.QMessageBox.critical(self, "Database Login Error", errorMsg)
            self.close()
            return

        if self.db !=None:
            #populate ship
            query = self.db.dbQuery("SELECT ship FROM "+self.db.bioSchema+".ships WHERE ship <> 999 ORDER BY ship")
            for ship,  in query:
                self.shipBox.addItem(ship)
            if self.latestShip =='':
                self.shipBox.setCurrentIndex(self.shipBox.findText('157', QtCore.Qt.MatchFlag.MatchExactly))
            else:
                self.shipBox.setCurrentIndex(self.shipBox.findText(self.latestShip, QtCore.Qt.MatchFlag.MatchExactly))
                
            self.getAllIntervalTypes()
            self.activateMF()
            self.getSurveys()

    def getSurveys(self):
        self.ship=self.shipBox.currentText()
        if self.ship != '' and self.ship is not None:
            query = self.db.dbQuery("SELECT survey FROM "+self.db.bioSchema+".surveys WHERE survey>200000 and " +
                    "survey<209900 and ship=" + self.ship+" ORDER BY survey DESC")
            self.surveyBox.clear()
            for survey, in query:
                self.surveyBox.addItem(survey)
            self.surveyBox.setCurrentIndex(self.surveyBox.findText(self.latestSurvey,
                    QtCore.Qt.MatchFlag.MatchExactly))
            self.getDataSets()
            

    def getDataSets(self):
        self.survey=self.surveyBox.currentText()
        if self.survey != '' and self.survey is not None:
            query = self.db.dbQuery("SELECT data_set_id FROM "+self.db.acousticSchema+".data_sets" +
               " WHERE (ship = "+self.ship+" ) AND"+
               " (survey = "+self.survey+" )" )
            self.dataSetBox.clear()
            for data_set_id  in query:
                self.dataSetBox.addItem(data_set_id[0])
            self.dataSetBox.setCurrentIndex(self.dataSetBox.findText(self.latestDataSet,
                    QtCore.Qt.MatchFlag.MatchExactly))
            self.getExportParameters()


    def getExportParameters(self):
        self.dataSet = self.dataSetBox.currentText()
        if self.dataSet != '' and self.dataSet is not None:
            self.getExportVariable()
            self.getIntervalType()
            self.getLayerReference()
            self.getThresholds()
            self.getZones()
            self.editAllBox.setChecked(False)
            self.updateEdits()


    def getExportVariable(self):
        query = self.db.dbQuery("SELECT source_name FROM "+self.db.acousticSchema+".data_sets" +
               " INNER JOIN "+self.db.acousticSchema+".acoustic_data_sources ON "+self.db.acousticSchema+".data_sets.source_id = "+self.db.acousticSchema+".acoustic_data_sources.source_id" +
               " WHERE ship = "+self.ship+" AND survey = "+self.survey+" AND data_set_id = "+self.dataSet)
        self.exportVariable=query.first()[0]
        # Right now, source_id in data_sets table can be set to null, so we will catch that case here.  Maybe this shouldn't be allow int he database though.
        if self.exportVariable!='' and self.exportVariable is not None:
            self.export_variable.clear()
            self.export_variable.insert(self.exportVariable)
            self.export_variable.setFont(self.mynormalFont)
        else:
            self.export_variable.clear()
            self.export_variable.insert('NO DATA')
            self.export_variable.setFont(self.myboldFont)
            self.exportVariable='NO DATA'


    def getLayerReference(self):
        query = self.db.dbQuery("SELECT layer_reference FROM "+self.db.acousticSchema+".data_sets" +
               " WHERE ship = "+self.ship+" AND survey = "+self.survey+" AND data_set_id = "+self.dataSet)
        self.layerReference=query.first()[0]
        # Do not need to catch the condition that there is no layer reference because it cannot be null in the database
        self.reference_label.clear()
        self.reference_label.insert(self.layerReference)
        self.reference_label.setFont(self.mynormalFont)
        query = self.db.dbQuery("SELECT layer_reference_name FROM "+self.db.acousticSchema+".data_sets" +
               " WHERE ship = "+self.ship+" AND survey = "+self.survey+" AND data_set_id = "+self.dataSet)

        self.referenceOffset = str(0)
        self.reference_offset.clear()
        self.reference_offset.insert(self.referenceOffset)
        self.layerReferenceName=query.first()[0]
        if (self.layerReferenceName=='' or self.layerReferenceName is None) and self.layerReference=='Surface':
            # For past surveys and data sets that have not filled in this entry, assume the surface reference will be "Surface (depth of zero)"
            sql=("UPDATE "+self.db.acousticSchema+".data_sets SET layer_reference_name = 'Surface (depth of zero)' "
                                " WHERE survey="+self.survey+" and ship="+self.ship+" and data_set_id="+self.dataSet)
            self.db.dbExec(sql)
            self.reference_label_name.clear()
            self.reference_label_name.insert('Surface (depth of zero)')
            self.reference_label_name.setFont(self.mynormalFont)
        elif self.layerReferenceName=='' or self.layerReferenceName is None:
            self.reference_label_name.clear()
            self.reference_label_name.insert('NO DATA')
            self.reference_label_name.setFont(self.myboldFont)
            self.layerReferenceName='NO DATA'
        else:
            self.reference_label_name.clear()
            self.reference_label_name.insert(self.layerReferenceName)
            self.reference_label_name.setFont(self.mynormalFont)
            if self.layerReference != 'Surface':
                # For now assume that bottom referenced data sets have one zone
                query = self.db.dbQuery("SELECT lower_exclusion_line FROM "+self.db.acousticSchema+".zones" +
               " WHERE ship = "+self.ship+" AND survey = "+self.survey+" AND data_set_id = "+self.dataSet+" AND lower_exclusion_name = '"+self.layerReferenceName+"'")
                excl_num = query.first()[0]
                if excl_num != '' and excl_num is not None:
                    query = self.db.dbQuery("SELECT exclusion_line_offset FROM "+self.db.acousticSchema+".exclusion_lines" +
                    " WHERE exclusion_line_id = "+excl_num)
                    self.reference_offset.clear()
                    self.referenceOffset = str(-float(query.first()[0]))
                    self.reference_offset.insert(self.referenceOffset)


    def getOffset(self, line_name, line_type):
        query = self.db.dbQuery("SELECT layer_reference, exclusion_line_offset FROM "+self.db.acousticSchema+".exclusion_lines a" +
                    " JOIN zones b ON a.exclusion_line_id = b."+line_type+"_exclusion_line WHERE b."+line_type+"_exclusion_name = '"+line_name+"'" + 
                    " AND b.ship = "+self.ship+" AND b.survey = "+self.survey+" AND b.data_set_id = "+self.dataSet)
        return query.first()
        
        
    def getThresholds(self):
        query = self.db.dbQuery("SELECT minimum_threshold_applied as min_bool, minimum_threshold as min_val, maximum_threshold_applied as max_bool, maximum_threshold as max_val FROM "+self.db.acousticSchema+".data_sets" +
               " WHERE ship = "+self.ship+" AND survey = "+self.survey+" AND data_set_id = "+self.dataSet)
        # Catch cases when these are null in the database
        for min_bool, min_val, max_bool, max_val in query:
            if min_bool=='1':
                self.minThresholdCheck.setChecked(True)
                self.applyMinThresh = 1
                self.startMinThresh = 1
                self.int_threshold_min.setEnabled(True)
            elif min_bool=='0':
                self.minThresholdCheck.setChecked(False)
                self.applyMinThresh = 0
                self.startMinThresh = 0
                self.int_threshold_min.setEnabled(False)
            else:
                self.minThresholdCheck.setChecked(False)
                self.applyMinThresh = -1
                self.startMinThresh = -1
                self.int_threshold_min.setEnabled(False)

            if min_val:
                self.int_threshold_min.clear()
                self.int_threshold_min.insert(min_val)
                self.int_threshold_min.setFont(self.mynormalFont)
                self.MinThreshold=min_val
            else:
                self.int_threshold_min.clear()
                self.int_threshold_min.insert('NO DATA')
                self.int_threshold_min.setFont(self.myboldFont)
                self.MinThreshold='NO DATA'

            if max_bool=='1':
                self.maxThresholdCheck.setChecked(True)
                self.applyMaxThresh = 1
                self.startMaxThresh = 1
                self.int_threshold_max.setEnabled(True)
            elif max_bool=='0':
                self.maxThresholdCheck.setChecked(False)
                self.applyMaxThresh = 0
                self.startMaxThresh = 0
                self.int_threshold_max.setEnabled(False)
            else:
                self.maxThresholdCheck.setChecked(False)
                self.applyMaxThresh = -1
                self.startMaxThresh = -1
                self.int_threshold_max.setEnabled(False)

            if max_val:
                self.int_threshold_max.clear()
                self.int_threshold_max.insert(max_val)
                self.int_threshold_max.setFont(self.mynormalFont)
                self.MaxThreshold=max_val
            else:
                self.int_threshold_max.clear()
                self.int_threshold_max.insert('NO DATA')
                self.int_threshold_max.setFont(self.myboldFont)
                self.MaxThreshold='NO DATA'


    def getZones(self):
        if self.dataSet != '' and self.dataSet is not None:
            
            #  disconnect zoneCheckBoxes signals while we manipulate the zoneCheckBoxes
            for box in self.zoneCheckBoxes:
                box.stateChanged[int].disconnect()
            
            query = self.db.dbQuery("SELECT zone, lower_exclusion_name as low_name, upper_exclusion_name as " +
                    "up_name, layer_thickness as thickness FROM "+self.db.acousticSchema+".zones" +
                   " WHERE ship = "+self.ship+" AND survey = "+self.survey+" AND data_set_id = "+self.dataSet)

            # hide all boxes to reset
            self.zonesChecked=[]
            for box in self.zoneCheckBoxes:
                box.hide()
            for ledit in self.layerThicknessList:
                ledit.hide()
            for label in self.lowNamesList:
                label.hide()
            for label in self.upNamesList:
                label.hide()
            count=0
            # Save list of zones, upper and lower exclusion line names for exporting later
            self.zonesChecked=[]
            self.zonesAvailable=[]
            self.lowNamesAvailable=[]
            self.upNamesAvailable=[]
            self.thicknessAvailable=[]
            for zone, low_name,  up_name, thickness  in query:
                self.zonesChecked.append(True)
                self.zoneCheckBoxes[count].setChecked(True)
                self.zoneCheckBoxes[count].show()
                self.zoneCheckBoxes[count].setText(zone)
                self.zonesAvailable.append(zone)
                self.layerThicknessList[count].show()
                if thickness is not None:
                    self.thicknessAvailable.append(thickness)
                    self.layerThicknessList[count].setText(thickness)
                    self.layerThicknessList[count].setFont(self.mynormalFont)
                else:
                    self.thicknessAvailable.append('NO DATA')
                    self.layerThicknessList[count].setText('NO DATA')
                    self.layerThicknessList[count].setFont(self.myboldFont)
                if low_name is not None:
                    self.lowNamesAvailable.append(low_name)
                    self.lowNamesList[count].show()
                    self.lowNamesList[count].setText(low_name)
                    self.lowNamesList[count].setFont(self.mynormalFont)
                else:
                    self.lowNamesAvailable.append('NO DATA')
                    self.lowNamesList[count].show()
                    self.lowNamesList[count].setText('NO DATA')
                    self.lowNamesList[count].setFont(self.myboldFont)

                if up_name is not None:
                    self.upNamesAvailable.append(up_name)
                    self.upNamesList[count].show()
                    self.upNamesList[count].setText(up_name)
                    self.upNamesList[count].setFont(self.mynormalFont)
                else:
                    self.upNamesAvailable.append('NO DATA')
                    self.upNamesList[count].show()
                    self.upNamesList[count].setText('NO DATA')
                    self.upNamesList[count].setFont(self.myboldFont)
                count=count+1
            
            #  now reconnect the zoneCheckBoxes signals
            for box in self.zoneCheckBoxes:
                box.stateChanged[int].connect(self.checkZones)


    def getAllIntervalTypes(self):
        queryType = self.db.dbQuery("SELECT interval_type FROM "+self.db.acousticSchema+".interval_types")
        self.intervalTypeOptions=[]
        for type in queryType:
            self.intervalTypeBox.addItem(type[0])

        queryUnit = self.db.dbQuery("SELECT interval_units FROM "+self.db.acousticSchema+".interval_units")
        self.intervalUnitOptions=[]
        for unit in queryUnit:
            self.intervalUnitBox.addItem(unit[0])


    def getIntervalType(self):
        query = self.db.dbQuery("SELECT interval_type, interval_units, interval_length FROM "+self.db.acousticSchema+".data_sets" +
               " WHERE ship = "+self.ship+" AND survey = "+self.survey+" AND data_set_id = "+self.dataSet)
        for type, unit, length in query:
            self.type=type
            self.unit=unit
            self.length=float(length)
            # Do not need to catch cases of nothing in query because they have to be set in the database- not nullable
            self.intervalTypeBox.setCurrentIndex(self.intervalTypeBox.findText(type, QtCore.Qt.MatchFlag.MatchExactly))
            self.intervalUnitBox.setCurrentIndex(self.intervalUnitBox.findText(unit, QtCore.Qt.MatchFlag.MatchExactly))
            self.EDSU_length.setValue(float(length))


    @QtCore.pyqtSlot(int)
    def updateEdits(self):
        if self.editAllBox.isChecked():
            self.label_8.setEnabled(True)
            self.export_variable.setEnabled(True)
            self.intervalTypeBox.setEnabled(True)
            self.intervalUnitBox.setEnabled(True)
            self.EDSU_length.setEnabled(True)
            self.label_7.setEnabled(True)
            self.reference_label.setEnabled(True)
            self.reference_label_name.setEnabled(True)
            self.reference_offset.setEnabled(True)
            self.label_13.setEnabled(True)
            self.label_10.setEnabled(True)
            self.label_17.setEnabled(True)
            self.label_14.setEnabled(True)
            self.label_9.setEnabled(True)
            self.label_15.setEnabled(True)
            # Deal with zones in a loop
            for zone_ind in range(len(self.zonesAvailable)):
                if self.zonesChecked[zone_ind]:
                    self.zoneCheckBoxes[zone_ind].setEnabled(True)
                    self.layerThicknessList[zone_ind].setEnabled(True)
                    self.lowNamesList[zone_ind].setEnabled(True)
                    self.upNamesList[zone_ind].setEnabled(True)
                else:
                    self.zoneCheckBoxes[zone_ind].setEnabled(True)
                    self.layerThicknessList[zone_ind].setEnabled(False)
                    self.lowNamesList[zone_ind].setEnabled(False)
                    self.upNamesList[zone_ind].setEnabled(False)
            # Deal with min and max thresholds- apply the preserved enabled status of these from when everything was turned off.
            self.minThresholdCheck.setEnabled(True)
            self.maxThresholdCheck.setEnabled(True)
            if self.applyMinThresh==1:
                self.int_threshold_min.setEnabled(True)
            else:
                self.int_threshold_min.setEnabled(False)
            if self.applyMaxThresh==1:
                self.int_threshold_max.setEnabled(True)
            else:
                self.int_threshold_max.setEnabled(False)

        else:
            self.label_8.setEnabled(False)
            self.export_variable.setEnabled(False)
            self.intervalTypeBox.setEnabled(False)
            self.intervalUnitBox.setEnabled(False)
            self.EDSU_length.setEnabled(False)
            self.label_7.setEnabled(False)
            self.reference_label.setEnabled(False)
            self.reference_label_name.setEnabled(False)
            self.reference_offset.setEnabled(False)
            self.label_13.setEnabled(False)
            self.label_10.setEnabled(False)
            self.label_17.setEnabled(False)
            self.label_14.setEnabled(False)
            self.label_9.setEnabled(False)
            self.label_15.setEnabled(False)
            # Deal with zones in a loop
            for box in self.zoneCheckBoxes:
                box.setEnabled(False)
            for ledit in self.layerThicknessList:
                ledit.setEnabled(False)
            for label in self.lowNamesList:
                label.setEnabled(False)
            for label in self.upNamesList:
                label.setEnabled(False)
            # Deal with min and max thresholds- remember which were true and which were false so you can preserve and apply when editable again.
            self.minThresholdCheck.setEnabled(False)
            self.maxThresholdCheck.setEnabled(False)
            self.int_threshold_min.setEnabled(False)
            self.int_threshold_max.setEnabled(False)

    def activateMF(self):
        if self.mfCheckBox.isChecked():
            self.Export.setText('Multi-frequency Export')
            self.exportType=1
            self.tabWidget.setTabEnabled(2, True)
        else:
            self.Export.setText('Export')
            self.exportType=0
            self.tabWidget.setTabEnabled(2, False)

    def checkZones(self):
        # Adjust the status of the check box array
        ind=self.zoneCheckBoxes.index(self.sender())
        if self.zoneCheckBoxes[ind].isChecked():
            self.zoneCheckBoxes[ind].setChecked(True)
            self.zonesChecked[ind]=True
            self.layerThicknessList[ind].setEnabled(True)
            self.upNamesList[ind].setEnabled(True)
            self.lowNamesList[ind].setEnabled(True)

        else:
            self.zoneCheckBoxes[ind].setChecked(False)
            self.zonesChecked[ind]=False
            self.layerThicknessList[ind].setEnabled(False)
            self.upNamesList[ind].setEnabled(False)
            self.lowNamesList[ind].setEnabled(False)


    def selectRaw(self):
        if self.setRawFiles.isChecked():
            self.rawfiles_button.setEnabled(True)
            self.rawFilesDir.setEnabled(True)
            self.setRawFilesi = 1
        else:
            self.rawfiles_button.setEnabled(False)
            self.rawFilesDir.setEnabled(False)
            self.setRawFilesi = 0


    def threshOnOff(self):
        if self.maxThresholdCheck.isChecked():
            self.int_threshold_max.setEnabled(True)
            self.applyMaxThresh = 1
        else:
            self.int_threshold_max.setEnabled(False)
            self.applyMaxThresh = 0
        if self.minThresholdCheck.isChecked():
            self.int_threshold_min.setEnabled(True)
            self.applyMinThresh = 1
        else:
            self.int_threshold_min.setEnabled(False)
            self.applyMinThresh = 0


    def checksAndSetup(self):

        #initiate params class from export tools
        params = parameterSetup()

        # Extract parameters from the GUI fields, but catch if there is nothing entered
        # A lot of error handling in here for cases when things are changed from what is in the database
        # and allowing user to enter in data to the database when it is not already there.
        params.input_dir = self.input_dir.text()
        if params.input_dir == '':
            QtWidgets.QMessageBox.critical(self, "Error", "No input directory defined. Export aborted.")
            return False
        params.output_dir_mb2 = self.output_dir_mb2.text()
        if params.output_dir_mb2 == '':
            QtWidgets.QMessageBox.critical(self, "Error", "No output directory defined. Export aborted.")
            return False
        params.transect_name  = self.transect_no.text()
        if params.transect_name == '':
            QtWidgets.QMessageBox.critical(self, "Error", "No transect(s) specified. Export aborted.")
            return False
        params.ECSfilename = self.cal_file.text()
        if params.ECSfilename == '':
            QtWidgets.QMessageBox.critical(self, "Error", "No calibration file specified. Export aborted.")
            return False
        params.survey_no = self.surveyBox.currentText()
        if params.survey_no == '':
            QtWidgets.QMessageBox.critical(self, "Error", "No survey number specified. Export aborted.")
            return False
        params.Fileset = self.fileset_name.text()
        if params.Fileset == '':
            QtWidgets.QMessageBox.critical(self, "Error", "No fileset specified. Export aborted.")
            return False

        params.Variable_for_export = self.export_variable.text() # set variable to export (38kHz default)
        if params.Variable_for_export == '' or params.Variable_for_export =='NO DATA':
            QtWidgets.QMessageBox.critical(self, "Error", "No export variable specified. Export aborted.")
            return False
        elif self.exportVariable!=params.Variable_for_export and self.exportVariable=='NO DATA':
            self.refresh_text_box('Warning- The export variable was changed in the parameters section but ' +
                                    'it is not defined in the database.  For a more permanent fix, make sure the source_id ' +
                                    'is set up in the data_sets table and tied to the appropriate source name (export variable) '+
                                    'in the acoustic_data_sources table.')
        elif self.exportVariable!=params.Variable_for_export and self.exportVariable!='NO DATA':
                self.refresh_text_box('Warning- The export variable has been changed from that which is specified in the database.')

        type=self.intervalTypeBox.currentText()
        unit=self.intervalUnitBox.currentText()
        length=self.EDSU_length.value()
        if type!=self.type:
            self.refresh_text_box('Warning- The interval type has been changed from that which is specified in the database.')
        if unit!=self.unit:
            self.refresh_text_box('Warning- The interval unit has been changed from that which is specified in the database.')
        if length!=self.length:
            self.refresh_text_box('Warning- The interval length has been changed from that which is specified in the database.')

        # There are 6 options for interval type and unit in exporting, ignoring for now the no time/distance grid option
        # Catch cases of erroneous picks by the user, such as a time unit with a distance type
        # Establish the time ditsance grid mode identifier for echoview and save it in params for use later
        # 1 is time (in minutes)
        # 2 is GPS distance (nmi)
        # 3 is vessel log (nmi)
        # 4 is distance (pings)
        # 5 is GPS distance (m)
        # 6 is vessel log (m)
        params.EDSU_length=length
        if type=='Time':
            params.int_class=1
            if unit=='minutes':
                pass
            elif unit=='hours':
                params.EDSU_length=length*60
            elif unit=='days':
                params.EDSU_length=length*24*60
            else:
                QtWidgets.QMessageBox.critical(self, "Error", "Interval units do not fit with the interval type.")
                return False
        if type=='GPS distance':
            if unit=='m':
                params.int_class=5
            elif unit=='nmi':
                params.int_class=2
            else:
                QtWidgets.QMessageBox.critical(self, "Error", "Interval units do not fit with the interval type.")
                return False
        if type=='Vessel log distance':
            if unit=='m':
                params.int_class=6
            elif unit=='nmi':
                params.int_class=3
            else:
                QtWidgets.QMessageBox.critical(self, "Error", "Interval units do not fit with the interval type.")
                return False
        if type=='Ping number':
            if unit=='pings':
                params.int_class=4
            else:
                QtWidgets.QMessageBox.critical(self, "Error", "Interval units do not fit with the interval type.")
                return False

        # Handle cases for the minimum and maximum integration threshold
        if self.startMinThresh!=self.applyMinThresh:
            self.refresh_text_box('Warning- The apply minimum threshold flag (yes/no) has been changed from that which is specified in the database.')

        if self.startMaxThresh!=self.applyMaxThresh:
            self.refresh_text_box('Warning- The apply maximum threshold flag (yes/no) has been changed from that which is specified in the database.')

        if self.minThresholdCheck.isChecked():
            t_min = self.int_threshold_min.text() #set min
            if t_min == '' or t_min=='NO DATA':
                QtWidgets.QMessageBox.critical(self, "Error", "No minimum integration threshold specified. Export aborted.")
                return False
            elif t_min!=self.MinThreshold and self.MinThreshold=='NO DATA' and t_min!='':
                try:
                    params.min_int_threshold=float(t_min)
                except:
                    QtWidgets.QMessageBox.critical(self, "Error", "Minimum integration threshold isn't a number.  Export aborted")
                    return False
                yes = QtWidgets.QMessageBox.question(self, "Commit?","The minimum integration threshold is unspecified in database for this data set. " +
                            "Do you want to insert "+t_min+" as minimum threshold value " +
                            "into database for survey " +self.survey+ ", ship "+self.ship+", data set id "+self.dataSet+"?",
                            QtWidgets.QMessageBox.StandardButton.Yes|QtWidgets.QMessageBox.StandardButton.No)
                if (yes == QtWidgets.QMessageBox.StandardButton.Yes):
                    # This will insure min threshold is in database and min threshold applied is set to true in the database
                    sql=("UPDATE "+self.db.acousticSchema+".data_sets SET minimum_threshold = " +t_min+ ", minimum_threshold_applied = 1 "
                                " WHERE survey="+self.survey+" and ship="+self.ship+" and data_set_id="+self.dataSet)
                    self.db.dbExec(sql)
            elif t_min!=self.MinThreshold and self.MinThreshold!='NO DATA' and t_min!='':
                self.refresh_text_box('Warning- The minimum integration threshold has been changed from that which is specified in the database.')
                try:
                    params.min_int_threshold=float(t_min)
                except:
                    QtWidgets.QMessageBox.critical(self, "Error", "Minimum integration threshold probably is not a number.  Export aborted")
                    return False
            elif t_min==self.MinThreshold:
                try:
                    params.min_int_threshold=float(t_min)
                except:
                    QtWidgets.QMessageBox.critical(self, "Error", "Minimum integration threshold probably is not a number.  Export aborted")
                    return False

        if self.maxThresholdCheck.isChecked():
            t_max = self.int_threshold_max.text() #set min
            if t_max == '' or t_max=='NO DATA':
                QtWidgets.QMessageBox.critical(self, "Error", "No maximum integration threshold specified. Export aborted.")
                return False
            elif t_max!=self.MaxThreshold and self.MaxThreshold=='NO DATA' and t_max!='':
                try:
                    params.max_int_threshold=float(t_max)
                except:
                    QtWidgets.QMessageBox.critical(self, "Error", "Maximum integration threshold isn't a number.  Export aborted")
                    return False
                yes = QtWidgets.QMessageBox.question(self, "Commit?","The maximum integration threshold is unspecified in database for this data set. " +
                            "Do you want to insert "+t_max+" as maximum threshold value " +
                            "into database for survey " +self.survey+ ", ship "+self.ship+", data set id "+self.dataSet+"?",
                            QtWidgets.QMessageBox.StandardButton.Yes|QtWidgets.QMessageBox.StandardButton.No)
                if (yes == QtWidgets.QMessageBox.StandardButton.Yes):
                    # This will insure min threshold is in database and min threshold applied is set to true in the database
                    sql=("UPDATE "+self.db.acousticSchema+".data_sets SET maximum_threshold = " +t_max+ ", maximum_threshold_applied = 1 "
                                " WHERE survey="+self.survey+" and ship="+self.ship+" and data_set_id="+self.dataSet)
                    self.db.dbExec(sql)
            elif t_max!=self.MaxThreshold and self.MaxThreshold!='NO DATA' and t_max!='':
                self.refresh_text_box('Warning- The maximum integration threshold has been changed from that which is specified in the database.')
                try:
                    params.max_int_threshold=float(t_max)
                except:
                    QtWidgets.QMessageBox.critical(self, "Error", "Maximum integration threshold isn't a number.  Export aborted")
                    return False
            elif t_max==self.MaxThreshold:
                try:
                    params.max_int_threshold=float(t_max)
                except:
                    QtWidgets.QMessageBox.critical(self, "Error", "Maximum integration threshold isn't a number.  Export aborted")
                    return False

        # Deal with any change in gui input for layer reference
        if self.layerReference==self.reference_label.text():
            params.reference_label=self.reference_label.text()
        elif self.reference_label.text()=='':
            QtWidgets.QMessageBox.critical(self, "Error", "No layer reference specified. Export aborted.")
            return False
        else:
            params.reference_label=self.reference_label.text()
            self.refresh_text_box('Warning- The layer reference has been changed from that which is specified in the database.')

        name=self.reference_label_name.text()
        if name=='' or name=='NO DATA':
            QtWidgets.QMessageBox.critical(self, "Error", "No layer reference name specified. Export aborted.")
            return False
        elif name!=self.layerReferenceName and self.layerReferenceName=='NO DATA':
            params.layerReferenceName=name
            yes = QtWidgets.QMessageBox.question(self, "Commit?","The layer reference name is unspecified in database for this data set. " +
                            "Do you want to insert "+name+" as layer reference name " +
                            "into database for survey " +self.survey+ ", ship "+self.ship+", data set id "+self.dataSet+"?",
                            QtWidgets.QMessageBox.StandardButton.Yes|QtWidgets.QMessageBox.StandardButton.No)
            if (yes == QtWidgets.QMessageBox.StandardButton.Yes):
                # This will insure min threshold is in database and min threshold applied is set to true in the database
                sql=("UPDATE "+self.db.acousticSchema+".data_sets SET layer_reference_name = '" +name+ "' "
                                " WHERE survey="+self.survey+" and ship="+self.ship+" and data_set_id="+self.dataSet)
                self.db.dbExec(sql)
        elif name!=self.layerReferenceName and self.layerReference!='NO DATA':
            params.layerReferenceName=name
            self.refresh_text_box('Warning- The layer reference name has been changed from that which is specified in the database.')
        elif name==self.layerReferenceName:
            params.layerReferenceName=name

        reference_offset = self.reference_offset.text()
        if reference_offset=='' or reference_offset=='NO DATA':
            QtWidgets.QMessageBox.critical(self, "Error", "No layer offset specified. Offset set to 0.")
            params.reference_offset = 0
        elif self.referenceOffset != reference_offset:
            params.reference_offset = float(reference_offset)
            self.refresh_text_box('Warning- The layer reference offset has been changed from that which is specified in the database.')
        elif reference_offset == self.referenceOffset:
            params.reference_offset = float(reference_offset)
            
        # Set zones.  First see if box is checked.  If checked, assign text to variable.  Line names are retrieved from
        # 'Default Zones' tab and are set by default to those used by the Exporter.
        params.zone = []
        params.exclude_above_line = []
        params.exclude_below_line = []
        params.layer_thickness=[]

        # Set up zone specific parameters for exporting...
        for zone_ind in range(len(self.zonesChecked)):
            if self.zonesChecked[zone_ind]:
                # Place zone parameters into params data structure
                zone=self.zonesAvailable[zone_ind]
                low_name=self.lowNamesList[zone_ind].text()
                up_name=self.upNamesList[zone_ind].text()
                thickness=self.layerThicknessList[zone_ind].text()
                params.zone.append(zone)
                self.zoneCheckBoxes[zone_ind].setEnabled(True)
                params.exclude_below_line.append(low_name)
                params.exclude_above_line.append(up_name)
                # Catch case when thickness cannot be converted to a float because non-numeric values were entered
                try:
                    params.layer_thickness.append(float(thickness))
                except:
                    QtWidgets.QMessageBox.critical(self, "Error", "No suitable layer thickness specified for zone "+zone+". Export aborted.")
                    return False

                # Perform checks for each parameter in each zone.
                # Can be one of the following:
                # a) Parameter was changed and isn't set in database- Ask if parameter should be inserted into the database
                # b) Parameter was changed from that which is in the database- Provide warning in text box but continue
                # c) Parameter was changed to empty OR parameter isn't set in the database and wasn't changed- Provide error message and abort exporter

                # Lower exclusion line name
                if self.lowNamesAvailable[zone_ind]!=low_name and self.lowNamesAvailable[zone_ind]=='NO DATA' and low_name!='':
                    yes = QtWidgets.QMessageBox.question(self, "Commit?","The lower exclusion line name is unspecified in database for zone " + zone + ". " +
                        "Do you want to insert "+low_name+" as lower exclusion line " +
                        "into database for survey " +self.survey+ ", ship "+self.ship+", data set id "+self.dataSet+" and zone "+zone+ "?",
                        QtWidgets.QMessageBox.StandardButton.Yes|QtWidgets.QMessageBox.StandardButton.No)
                    if (yes == QtWidgets.QMessageBox.StandardButton.Yes):
                        sql=("UPDATE "+self.db.acousticSchema+".zones SET lower_exclusion_name = '" +low_name+
                            "' WHERE survey="+self.survey+" and ship="+self.ship+" and data_set_id="+self.dataSet+" and zone="+zone)
                        self.db.dbExec(sql)
                elif self.lowNamesAvailable[zone_ind]!=low_name and low_name!='NO DATA' and low_name!='':
                    self.refresh_text_box('Warning- The lower exclusion line name for zone '+zone+' has been changed from that which is specified in the database.')
                elif low_name=='' or low_name=='NO DATA':
                    QtWidgets.QMessageBox.critical(self, "Error", "No lower exlusion line name specified. Export aborted.")
                    return False

                # Upper exclusion line name
                if self.upNamesAvailable[zone_ind]!=up_name and self.upNamesAvailable[zone_ind]=='NO DATA' and up_name!='':
                    yes = QtWidgets.QMessageBox.question(self, "Commit?","The upper exclusion line name is unspecified in database for zone " + zone + ". " +
                        "Do you want to insert "+up_name+" as upper exclusion line " +
                        "into database for survey " +self.survey+ ", ship "+self.ship+", data set id "+self.dataSet+" and zone "+zone+ "?" +
                        " If no, that's okay- the export will still continue without updating database.",
                        QtWidgets.QMessageBox.StandardButton.Yes|QtWidgets.QMessageBox.StandardButton.No)
                    if (yes == QtWidgets.QMessageBox.StandardButton.Yes):
                        sql=("UPDATE "+self.db.acousticSchema+".zones SET upper_exclusion_name = '" +up_name+
                            "' WHERE survey="+self.survey+" and ship="+self.ship+" and data_set_id="+self.dataSet+" and zone="+zone)
                        self.db.dbExec(sql)
                elif self.upNamesAvailable[zone_ind]!=up_name and up_name!='NO DATA' and up_name!='':
                    self.refresh_text_box('Warning- The upper exclusion line name for zone '+zone+' has been changed from that which is specified in the database.')
                elif up_name=='' or up_name=='NO DATA':
                    QtWidgets.QMessageBox.critical(self, "Error", "No upper exlusion line name specified. Export aborted.")
                    return False

                # Layer thickness
                if self.thicknessAvailable[zone_ind]!=thickness and self.thicknessAvailable[zone_ind]=='NO DATA':
                    yes = QtWidgets.QMessageBox.question(self, "Commit?","The layer thickness is unspecified in database for zone " + zone + ". " +
                        "Do you want to insert "+thickness+" as layer thickness " +
                        "into database for survey " +self.survey+ ", ship "+self.ship+", data set id "+self.dataSet+" and zone "+zone+ "?" +
                        " If no, that's okay- the export will still continue without updating database.",
                        QtWidgets.QMessageBox.StandardButton.Yes|QtWidgets.QMessageBox.StandardButton.No)
                    if yes == QtWidgets.QMessageBox.StandardButton.Yes:
                        sql=("UPDATE "+self.db.acousticSchema+".zones SET layer_thickness = " +thickness+
                            " WHERE survey="+self.survey+" and ship="+self.ship+" and data_set_id="+self.dataSet+" and zone="+zone)
                        self.db.dbExec(sql)
                elif self.thicknessAvailable[zone_ind]!=thickness:
                    self.refresh_text_box('Warning- The layer thickness for zone '+zone+' has been changed from that which is specified in the database.')
        return params


    def setupMF(self, params):
        params.variable_export_list = []
        if self.select_mf1.isChecked():
            if self.variable_mf1.text()!='':
                params.variable_export_list.append(str(self.variable_mf1.text()))
                params.v38min = self.min_mf1.value()
                params.v38max = self.max_mf1.value()
        if self.select_mf2.isChecked():
            if self.variable_mf2.text()!='':
                params.variable_export_list.append(str(self.variable_mf2.text()))
                params.v120min = self.min_mf2.value()
                params.v120max = self.max_mf2.value()
        if self.select_mf3.isChecked():
            if self.variable_mf3.text()!='':
                params.variable_export_list.append(str(self.variable_mf3.text()))
                params.autokrillmin = self.min_mf3.value()
                params.autokrillmax = self.max_mf3.value()
        if self.select_mf4.isChecked():
            if self.variable_mf4.text()!='':
                params.variable_export_list.append(str(self.variable_mf4.text()))
        if self.select_mf5.isChecked():
            if self.variable_mf4.text()!='':
                params.variable_export_list.append(str(self.variable_mf5.text()))
                params.autopollockmin = self.min_mf5.value()
                params.autopollockmax = self.max_mf5.value()
        if self.select_mf6.isChecked():
            if self.variable_mf4.text()!='':
                params.variable_export_list.append(str(self.variable_mf6.text()))
        return params


    def export(self):

        # do checks on variables and store in params structure for use in exporting
        params=self.checksAndSetup()
        if params==False:
            return

        # if multi-frequency has been checked, do checks on variables and store in params structure
        if self.exportType==1:
            params=self.setupMF(params)

        self.refresh_text_box('Starting New Export')

        #If multiple transects, separate into list
        transect_names=[]
        all = 0
        if ',' in params.transect_name:
            transect_names = params.transect_name.split(',')
        elif ' ' in params.transect_name:
            transect_names = params.transect_name.split(' ')
        elif '-' in params.transect_name:
            transect_names = params.transect_name.split('-')
            transect_names = range(int(transect_names[0]), int(transect_names[1])+1)
            transect_names = [str(i) for i in transect_names]
        elif params.transect_name == 'ALL':
            allfiles = numpy.asarray(glob.glob(str(params.input_dir) + '\\'  + '*.EV'))
            if len(allfiles) == 0:
                QtWidgets.QMessageBox.critical(self, "Error", "No .EV files found. Export aborted.")
                return
            self.refresh_text_box('\nFound '+ str(len(allfiles)) +' files\n')
            transect_names = allfiles
            all = 1
        else:
            transect_names.append(params.transect_name)

        #How many transects are there?
        transect_ct = len(transect_names)
        if transect_ct > 1:
            for k in (range(transect_ct)): #iterate through each transect
                if all == 1:
                    transect_name = allfiles[k][allfiles[k].find("-t")+2:allfiles[k].find("-z")]
                else:
                    transect_name = transect_string(transect_names[k])
                #  find all .EV files containing the transect name in that input directory
                filelist = numpy.asarray(glob.glob(str(params.input_dir) + '\\' + '*'+str(params.survey_no)+
                        '*' + '*'+str(transect_name)+'*' + '*.EV'))
                 # make sure that the file for the correct transect existed in the folder you chose
                num_loops = len(filelist)
                if num_loops == 0:
                    QtWidgets.QMessageBox.critical(self, "Error", "No .EV files found for transect "  + transect_name[1:] +
                            ". This transect will be skipped.")
                elif [] in filelist:
                    #  there could be 2 transects of .EV files, but maybe not the two transect specified by the user
                    QtWidgets.QMessageBox.critical(self, "Error", "Not all if the .EV files could be found for transect " +
                            transect_name[1:] +  ". This transect will be skipped.")
                else:
                    self.refresh_text_box('Beginning Export of Transect ' + transect_name[1:] + '...')
                    successMB2 =  self.export_py_MB2(filelist, params)

                    self.refresh_text_box('For Transect ' + transect_name[1:] + '...')
                    total_zones_exported=sum(successMB2)
                    if self.exportType==0:
                        if total_zones_exported ==len(params.zone) and self.exportType==0:
                            self.refresh_text_box('All zones exported \n')
                        else:
                            self.refresh_text_box(str(total_zones_exported)+' zone(s) exported out of '+str(len(params.zones))+' zone(s)')

        else: #If there's only one transect
            if all == 1:
                transect_name = allfiles[0][allfiles[0].find("-t")+2:allfiles[0].find("-z")]
            else:
                transect_name = transect_string(transect_names[0])
            filelist =  numpy.asarray(glob.glob(str(params.input_dir) + '\\' + '*'+str(params.survey_no)+
                    '*' + '*'+str(transect_name)+'*' +'*.EV'))
            num_loops = len(filelist)
            if num_loops == 0:
                QtWidgets.QMessageBox.critical(self, "Error", "No .EV files found for transect "  + transect_name[1:] +
                            ". This transect will be skipped.")
            else:
                self.refresh_text_box('Beginning Export of Transect ' + transect_name[1:] + '...')
                successMB2 =self.export_py_MB2(filelist, params)
                #  if the export is successful
                total_zones_exported=sum(successMB2)
                if self.exportType==0:
                    if total_zones_exported ==len(params.zone):
                        self.refresh_text_box('All Files Done \n')
                    else:
                        self.refresh_text_box(str(total_zones_exported)+' zone(s) exported out of '+str(len(params.zone))+' zone(s)')


    # Button function for input directory dialog button.  assign directory to input directory text field.
    def getInputDirectory(self):
        self.input_dir.clear()
        input_dirname = str(QtWidgets.QFileDialog.getExistingDirectory(self, "Select Input Directory"))
        self.input_dir.insert(input_dirname)

    # Button function for output directory dialog button.  assign directory to output directory text field.
    def getOutputDirectoryMb2(self):
        self.output_dir_mb2.clear()
        output_dirname_mb2 = str(QtWidgets.QFileDialog.getExistingDirectory(self, "Select Macebase2 Output Directory"))
        self.output_dir_mb2.insert(output_dirname_mb2)

    def getRawFilesDirectory(self):
        self.rawFilesDir.clear()
        raw_file_dir = str(QtWidgets.QFileDialog.getExistingDirectory(self, "Select Raw Files Directory"))
        self.rawFilesDir.insert(raw_file_dir)

    # Button function for calibration file dialog button.  assign directory to calibration file text field.
    def getCalFile(self):
        self.cal_file.clear()
        input_dirname = QtWidgets.QFileDialog.getOpenFileName(self, "Select Calibration File")
        input_dirname = input_dirname[0]
        self.cal_file.insert(input_dirname)

    # cancel button
    def refresh_text_box(self, MyString):
        self.textBrowser.append(MyString)
        QtWidgets.QApplication.processEvents()


    def quit(self):
        sys.exit()


    def export_py_MB2(self, files, params):
        #Open up Echoview
        EvApp = win32com.client.Dispatch("EchoviewCom.EvApplication")
        license = EvApp.IsLicensed()
        if license == 0:
            self.refresh_text_box('No Scripting Module Found')
            EvApp.Quit()
            pass
        EvApp.Minimize()
        EvFileName = str(files[0]) #pick the file
        filename = os.path.basename(EvFileName) #filename
        EvExportName = filename[:filename.find('-z')] #chop off the .EV
        self.refresh_text_box('\nExporting for Macebase 2...')
        self.refresh_text_box('Working on '+ str(EvFileName))
        if self.setRawFilesi == 1:
        # The following 7 line should be part of an if statement based on a flag for resetting the raw data directory.
            rawDir = self.rawFilesDir.text()
            if rawDir == '':
                QtWidgets.QMessageBox.about(self, "Warning", "Raw Files Directory is Blank")
            self.refresh_text_box('Setting new raw file directory')
            EvFile = EvApp.OpenFile(EvFileName) #Open up the file
            EvFile.PreReadDataFiles #pre-read just in case
            EvFile.Properties.DataPaths.Add(rawDir);
            EvFile.SaveAs(EvFileName)
            EvApp.CloseFile(EvFile)
        self.refresh_text_box('Loading raw files...')
        EvFile = EvApp.OpenFile(EvFileName) #Open up the file
        Evfileset = EvFile.Filesets.FindByName(params.Fileset)
        EvVar =  EvFile.Variables.FindByName(params.Variable_for_export)
        EvFile.PreReadDataFiles #pre-read just in case
        # Set up cal file
        calfiletest = Evfileset.SetCalibrationFile(params.ECSfilename)
        if calfiletest != 1:
            self.refresh_text_box('Failed to set .ecs file')
            self.refresh_text_box(EvExportName)

        # set grid settings- params.int_class is set above using combination of types and units-
        # 1 is time (in minutes), 2 is GPS distance (nmi), 3 is vessel log (nmi), 4 is distance (pings), 5 is GPS distance (m), 6 is vessel log (m)
        EvVar.Properties.Grid.SetTimeDistanceGrid(params.int_class, params.EDSU_length)


        # Single variable export
        if self.exportType==0:
            Date_E=EvFile.Properties.Export.Variables.Item('Date_E');
            Date_E.Enabled=1;
            Lat_E=EvFile.Properties.Export.Variables.Item('Lat_E');
            Lat_E.Enabled=1;
            Lon_E=EvFile.Properties.Export.Variables.Item('Lon_E');
            Lon_E.Enabled=1;
            Time_E=EvFile.Properties.Export.Variables.Item('Time_E');
            Time_E.Enabled=1;
            Region_notes=EvFile.Properties.Export.Variables.Item('Region_notes');
            Region_notes.Enabled=1;
            Grid_reference_line=EvFile.Properties.Export.Variables.Item('Grid_reference_line');
            Grid_reference_line.Enabled=1;
            Layer_btrld=EvFile.Properties.Export.Variables.Item('Layer_bottom_to_reference_line_depth');
            Layer_btrld.Enabled=1;
            Layer_ttrld=EvFile.Properties.Export.Variables.Item('Layer_top_to_reference_line_depth');
            Layer_ttrld.Enabled=1;
            Samples_In_Domain=EvFile.Properties.Export.Variables.Item('Samples_In_Domain');
            Samples_In_Domain.Enabled=1;
            Good_samples=EvFile.Properties.Export.Variables.Item('Good_samples');
            Good_samples.Enabled=1;
            No_data_samples=EvFile.Properties.Export.Variables.Item('No_data_samples');
            No_data_samples.Enabled=1;
            Sv_max=EvFile.Properties.Export.Variables.Item('Sv_max');
            Sv_max.Enabled=1;

            EvVar.Properties.Data.ApplyMinimumThreshold= self.applyMinThresh
            if self.applyMinThresh ==1:
                EvVar.Properties.Data.MinimumThreshold= params.min_int_threshold
            EvVar.Properties.Data.ApplyMaximumThreshold= self.applyMaxThresh
            if self.applyMaxThresh ==1:
                EvVar.Properties.Data.MaximumThreshold= params.max_int_threshold

            ExportFileName = params.output_dir_mb2 + '\\' + EvExportName + '- (regions).csv' #output .csv filename
            exporttest1 = EvVar.ExportRegionsLogAll(ExportFileName);
            if exporttest1 != 1:
                self.refresh_text_box('Error: Unable to make regions logbook \n')

            # Save calibration ecs file in export directory
            f=open(params.ECSfilename,'r')
            contents=f.read()
            f.close()
            h=open(params.output_dir_mb2 + '\\' + EvExportName + '-calibration-.ecs','w+')
            h.write(contents)
            h.close()

            # Create a subfolder called 'Regions'
            regionOutDir = params.output_dir_mb2 + '\\Regions'
            dirExist = os.path.exists(regionOutDir)
            if not dirExist:
                os.mkdir(regionOutDir)
            # Export Regions file
            ExportFileName = regionOutDir + '\\' + EvExportName + '-regions.evr'
            exporttest = EvFile.Regions.ExportDefinitionsAll(ExportFileName)
            
            
            # Create a subfolder called 'Lines'
            lineOutDir = params.output_dir_mb2 + '\\Lines'
            dirExist = os.path.exists(lineOutDir)
            if not dirExist:
                os.mkdir(lineOutDir)

            self.exporttestMB2=[]
            exported_line_names = []
            for z in range(len(params.zone)): #for each zone
                EvVar.Properties.Grid.SetDepthRangeGrid(1,params.layer_thickness[z])
                cur_zone=params.zone[z]
                try:
                    # Reference line
                    if params.layerReferenceName!='Surface (depth of zero)':
                        # Set an offset for non-surface referenced exports.
                        EvLine = EvFile.Lines.FindByName(str(params.layerReferenceName))
                        NewEvLine = EvFile.Lines.CreateOffsetLinear(EvLine, 1, params.reference_offset)
                        NewEvLine.Name = str(params.layerReferenceName+"-offset"+str(params.reference_offset))
                        EvVar.Properties.Grid.DepthRangeReferenceLine = NewEvLine
                    self.refresh_text_box('Exporting Zone '+str(cur_zone)+'...')
                    # Deal with lines:
                    # Set exclude above line
                    cur_line = str(params.exclude_above_line[z])
                    exported_line_names.append(cur_line)
                    EvVar.Properties.Analysis.ExcludeAboveLine = cur_line # set exclude above line
                    # Export exclude above line
                    line_ref = EvFile.Lines.FindByName(cur_line)
                    ref,  offset = self.getOffset(cur_line, 'upper')
                    if float(offset)<=0:
                        ref_string = str(-float(offset))+' above '+ ref.lower()
                    else:
                        ref_string = str(float(offset))+' below '+ ref.lower()
                    test = EvVar.ExportLine(line_ref, lineOutDir+'\\'+EvExportName+'-'+cur_line+'-'+ref_string+'-z'+str(cur_zone)+'-upper'+'.evl', -1, -1)
                    if not test:
                        self.refresh_text_box('There was a problem exporting the exclude above line file for zone'+str(cur_zone))
                    # Set exclude below line
                    cur_line = str(params.exclude_below_line[z])
                    exported_line_names.append(cur_line)
                    EvVar.Properties.Analysis.ExcludeBelowLine = cur_line # set exclude below line
                    line_ref = EvFile.Lines.FindByName(cur_line)
                    ref,  offset = self.getOffset(cur_line, 'lower')
                    if float(offset)<0:
                        ref_string = str(-float(offset))+' above '+ ref.lower()
                    else:
                        ref_string = str(-float(offset))+' below '+ ref.lower()
                    # Export exclude below line
                    test = EvVar.ExportLine(line_ref, lineOutDir+'\\'+EvExportName+'-'+cur_line+'-'+ref_string+'-z'+str(cur_zone)+'-lower'+'.evl', -1, -1)
                    if not test:
                        self.refresh_text_box('There was a problem exporting the exclude above line file for zone'+str(cur_zone))
                    
                    # Now complete the final export
                    ExportFileName = params.output_dir_mb2 + '\\' + EvExportName + '-z' + str(cur_zone) +'-' +'.csv' #output .csv filename
                    self.exporttest = EvVar.ExportIntegrationByRegionsByCellsAll(ExportFileName)
                except:
                    self.exporttest=False
                    self.refresh_text_box('There is no exclude_above and/or exclude_below line associated with zone '+str(cur_zone)+' specified or it does not match a line in the EV file' )

                if self.exporttest != True:
                    self.refresh_text_box('The export has failed for zone '+str(cur_zone))
                    self.exporttestMB2.append(0)
                else:
                    self.refresh_text_box('Zone '+ str(cur_zone) +' Export Complete')
                    self.exporttestMB2.append(1)
            # Export the rest of the lines
            N = EvFile.Lines.count
            for ind in range(0, N):
                EvLine = EvFile.Lines(ind)
                EvName = EvLine.Name
                # For now, we will skip the 'Fileset1: line data...' lines since these should be included with the raw file and the colon is causing issues
                isReject = EvName.find(':')
                if EvName not in exported_line_names and isReject==-1:
                    EvVar.ExportLine(EvLine, lineOutDir+'\\'+EvExportName+'-'+EvName+'.evl', -1, -1)            

        # Multi-frequency export setup and execution
        else:
            self.exporttestMB2=[] # Fill this in because it will be returned at the end but not used for multi-frequency
            ExportSamplesStatus=EvFile.Properties.Export.Variables.Item('Good_samples')  #use Item method to get handle to status of export variable samples
            ExportSamplesStatus.Enabled=1  #set the status to enabled
            ExportKurtosisStatus=EvFile.Properties.Export.Variables.Item('Kurtosis')
            ExportKurtosisStatus.Enabled=1
            ExportSkewnessStatus=EvFile.Properties.Export.Variables.Item('Skewness')
            ExportSkewnessStatus.Enabled=1
            ExportSv_meanStatus=EvFile.Properties.Export.Variables.Item('Sv_mean')
            ExportSv_meanStatus.Enabled=1
            ExportStandard_deviationStatus=EvFile.Properties.Export.Variables.Item('Standard_deviation')
            ExportStandard_deviationStatus.Enabled=1

            # Create a subfolder called 'Regions'
            regionOutDir = params.output_dir_mb2 + '\\Regions'
            dirExist = os.path.exists(regionOutDir)
            if not dirExist:
                os.mkdir(regionOutDir)
            # Export Regions file
            ExportFileName = regionOutDir + '\\' + EvExportName + '-regions.evr'
            exporttest = EvFile.Regions.ExportDefinitionsAll(ExportFileName)
            
            
            # Create a subfolder called 'Lines'
            lineOutDir = params.output_dir_mb2 + '\\Lines'
            dirExist = os.path.exists(lineOutDir)
            if not dirExist:
                os.mkdir(lineOutDir)

            #The following sections are for the export of individual variables.  Each variable is exported if it is found within the list set within
            #the parameters, assuming it was checked on the GUI.
            if '38 kHz for survey' in params.variable_export_list:
                variable_for_export = '38 kHz for survey'
                EvVar =  EvFile.Variables.FindByName(variable_for_export)
                EvVar.Properties.Data.ApplyMinimumThreshold= 1
                EvVar.Properties.Data.MinimumThreshold= params.v38min
                EvVar.Properties.Data.ApplyMaximumThreshold= 1
                EvVar.Properties.Data.MaximumThreshold= params.v38max
                for k in range(len(params.zone)):
                    EvVar.Properties.Grid.SetDepthRangeGrid(1,params.layer_thickness[k])
                    zone = params.zone[k]
                    # Reference line
                    if params.layerReferenceName!='Surface (depth of zero)':
                        EvLine = EvFile.Lines.FindByName(str(params.layerReferenceName))
                        EvVar.Properties.Grid.DepthRangeReferenceLine = EvLine
                    self.refresh_text_box('Exporting 38 kHz for survey from zone '+ str(zone))
                    # Deal with lines:
                    # Set exclude above line
                    cur_line = str(params.exclude_above_line[k])
                    EvVar.Properties.Analysis.ExcludeAboveLine = cur_line # set exclude above line
                    # Export exclude above line
                    line_ref = EvFile.Lines.FindByName(cur_line)
                    ref,  offset = self.getOffset(cur_line, 'upper')
                    if float(offset)<=0:
                        ref_string = str(-float(offset))+' above '+ ref.lower()
                    else:
                        ref_string = str(float(offset))+' below '+ ref.lower()
                    test = EvVar.ExportLine(line_ref, lineOutDir+'\\'+EvExportName+'-'+cur_line+'-'+ref_string+'-z'+str(zone)+'-upper'+'.evl', -1, -1)
                    if not test:
                        self.refresh_text_box('There was a problem exporting the exclude above line file for zone'+str(zone))
                    # Set exclude below line
                    cur_line = str(params.exclude_below_line[k])
                    EvVar.Properties.Analysis.ExcludeBelowLine = cur_line # set exclude below line
                    line_ref = EvFile.Lines.FindByName(cur_line)
                    ref,  offset = self.getOffset(cur_line, 'lower')
                    if float(offset)<0:
                        ref_string = str(-float(offset))+' above '+ ref.lower()
                    else:
                        ref_string = str(-float(offset))+' below '+ ref.lower()
                    # Export exclude below line
                    test = EvVar.ExportLine(line_ref, lineOutDir+'\\'+EvExportName+'-'+cur_line+'-'+ref_string+'-z'+str(zone)+'-lower'+'.evl', -1, -1)
                    if not test:
                        self.refresh_text_box('There was a problem exporting the exclude above line file for zone'+str(zone))
                        
                    ExportFileName = params.output_dir_mb2 + '\\' + EvExportName + 'z' + str(zone) +'.csv' #output .csv filename- edited name on 7/3/2016 by nel requested by patrick
                    exporttest = EvVar.ExportIntegrationByRegionsByCellsAll(ExportFileName)
                    if exporttest != 1:
                        self.refresh_text_box('The export has failed for zone '+str(zone))
                        self.refresh_text_box(ExportFileName)
                    else:
                        self.refresh_text_box('Zone '+ str(zone) +' Export Complete')

            if '120 kHz for survey' in params.variable_export_list:
                variable_for_export = '120 kHz for survey'
                EvVar = EvFile.Variables.FindByName(variable_for_export)
                EvVar.Properties.Data.ApplyMinimumThreshold= 1 #this is an example of implicit syntax for setting the property ApplyMinimumThreshold of COM object
                EvVar.Properties.Data.MinimumThreshold= params.v120min
                EvVar.Properties.Data.ApplyMaximumThreshold= 1
                EvVar.Properties.Data.MaximumThreshold= params.v120max
                for k in range(len(params.zone)):
                    EvVar.Properties.Grid.SetDepthRangeGrid(1,params.layer_thickness[k])
                    zone = params.zone[k]
                    # Reference line
                    if params.layerReferenceName!='Surface (depth of zero)':
                        EvLine = EvFile.Lines.FindByName(str(params.layerReferenceName))
                        EvVar.Properties.Grid.DepthRangeReferenceLine = EvLine
                    self.refresh_text_box('Exporting 120 kHz for survey from zone '+ str(zone))
                    EvVar.Properties.Analysis.ExcludeAboveLine = str(params.exclude_above_line[k])  #this is working even though it spits gibberish to the screen
                    EvVar.Properties.Analysis.ExcludeBelowLine = str(params.exclude_below_line[k])
                    ExportFileName = params.output_dir_mb2 + '\\' + EvExportName + 'z' + str(zone) +'.csv' #output .csv filename- edited name on 7/3/2016 by nel requested by patrick
                    ExportFileName.replace("x2-f38", "x4-f120")
                    exporttest = EvVar.ExportIntegrationByRegionsByCellsAll(ExportFileName)
                    if exporttest != 1:
                        self.refresh_text_box('The export has failed for zone '+str(zone))
                        self.refresh_text_box(ExportFileName)
                    else:
                        self.refresh_text_box('Zone '+ str(zone) +' Export Complete')

            if 'Autokrill for export' in params.variable_export_list:
                variable_for_export = 'Autokrill for export'
                EvVar =  EvFile.Variables.FindByName(variable_for_export)
                EvVar.Properties.Data.ApplyMinimumThreshold= 1
                EvVar.Properties.Data.MinimumThreshold= params.autokrillmin
                EvVar.Properties.Data.ApplyMaximumThreshold= 1
                EvVar.Properties.Data.MaximumThreshold= params.autokrillmax
                for k in range(len(params.zone)):
                    EvVar.Properties.Grid.SetDepthRangeGrid(1,params.layer_thickness[k])
                    zone = params.zone[k]
                    # Reference line
                    if params.layerReferenceName!='Surface (depth of zero)':
                        EvLine = EvFile.Lines.FindByName(str(params.layerReferenceName))
                        EvVar.Properties.Grid.DepthRangeReferenceLine = EvLine
                    self.refresh_text_box('Exporting Autokrill from zone '+ str(zone))
                    EvVar.Properties.Analysis.ExcludeAboveLine = str(params.exclude_above_line[k])  #this is working even though it spits gibberish to the screen
                    EvVar.Properties.Analysis.ExcludeBelowLine = str(params.exclude_below_line[k])
                    ExportFileName = params.output_dir_mb2 + '\\' + EvExportName + 'k1' +'.csv' #output .csv filename- edited name on 7/3/2016 by nel requested by patrick
                    ExportFileName.replace("x2-f38", "x4-f120")
                    exporttest = EvVar.ExportIntegrationByRegionsByCellsAll(ExportFileName)
                    if exporttest != 1:
                        self.refresh_text_box('The export has failed for zone '+str(zone))
                        self.refresh_text_box(ExportFileName)
                    else:
                        self.refresh_text_box('Zone '+ str(zone) +' Export Complete')

            if 'Autokrill mean z for export' in params.variable_export_list:
                variable_for_export = 'Autokrill mean z for export'
                EvVar =  EvFile.Variables.FindByName(variable_for_export)
                EvVar.Properties.Data.ApplyMinimumThreshold= 0
                EvVar.Properties.Data.ApplyMaximumThreshold= 0
                for k in range(len(params.zone)):
                    EvVar.Properties.Grid.SetDepthRangeGrid(1,params.layer_thickness[k])
                    zone = params.zone[k]
                    # Reference line
                    if params.layerReferenceName!='Surface (depth of zero)':
                        EvLine = EvFile.Lines.FindByName(str(params.layerReferenceName))
                        EvVar.Properties.Grid.DepthRangeReferenceLine = EvLine
                    self.refresh_text_box('Exporting Autokrill mean z from zone '+ str(zone))
                    EvVar.Properties.Analysis.ExcludeAboveLine = str(params.exclude_above_line[k])  #this is working even though it spits gibberish to the screen
                    EvVar.Properties.Analysis.ExcludeBelowLine = str(params.exclude_below_line[k])
                    ExportFileName = params.output_dir_mb2 + '\\' + EvExportName + 'k2' +'.csv' #output .csv filename- edited name on 7/3/2016 by nel requested by patrick
                    ExportFileName.replace("x2-f38", "x4-f120")
                    exporttest = EvVar.ExportIntegrationByRegionsByCellsAll(ExportFileName)
                    if exporttest != 1:
                        self.refresh_text_box('The export has failed for zone '+str(zone))
                        self.refresh_text_box(ExportFileName)
                    else:
                        self.refresh_text_box('Zone '+ str(zone) +' Export Complete')

            if 'Autopollock for export' in params.variable_export_list:
                variable_for_export = 'Autopollock for export'
                EvVar =  EvFile.Variables.FindByName(variable_for_export)
                EvVar.Properties.Data.ApplyMinimumThreshold= 1
                EvVar.Properties.Data.MinimumThreshold = params.autopollockmin
                EvVar.Properties.Data.ApplyMaximumThreshold= 1
                EvVar.Properties.Data.MaximumThreshold = params.autopollockmax
                for k in range(len(params.zone)):
                    EvVar.Properties.Grid.SetDepthRangeGrid(1,params.layer_thickness[k])
                    zone = params.zone[k]
                    # Reference line
                    if params.layerReferenceName!='Surface (depth of zero)':
                        EvLine = EvFile.Lines.FindByName(str(params.layerReferenceName))
                        EvVar.Properties.Grid.DepthRangeReferenceLine = EvLine
                    self.refresh_text_box('Exporting Autopollock from zone '+ str(zone))
                    EvVar.Properties.Analysis.ExcludeAboveLine = str(params.exclude_above_line[k])  #this is working even though it spits gibberish to the screen
                    EvVar.Properties.Analysis.ExcludeBelowLine = str(params.exclude_below_line[k])
                    ExportFileName = params.output_dir_mb2 + '\\' + EvExportName + 'p1' +'.csv' #output .csv filename- edited name on 7/3/2016 by nel requested by patrick
                    exporttest = EvVar.ExportIntegrationByRegionsByCellsAll(ExportFileName)
                    if exporttest != 1:
                        self.refresh_text_box( 'The export has failed for zone '+str(zone))
                        self.refresh_text_box(ExportFileName)
                    else:
                        self.refresh_text_box('Zone '+ str(zone) +' Export Complete')

            if 'Autopollock mean z for export' in params.variable_export_list:
                variable_for_export = 'Autopollock mean z for export'
                EvVar =  EvFile.Variables.FindByName(variable_for_export)
                EvVar.Properties.Data.ApplyMinimumThreshold= 0
                EvVar.Properties.Data.ApplyMaximumThreshold= 0
                for k in range(len(params.zone)):
                    EvVar.Properties.Grid.SetDepthRangeGrid(1,params.layer_thickness[k])
                    zone = params.zone[k]
                    # Reference line
                    if params.layerReferenceName!='Surface (depth of zero)':
                        EvLine = EvFile.Lines.FindByName(str(params.layerReferenceName))
                        EvVar.Properties.Grid.DepthRangeReferenceLine = EvLine
                    self.refresh_text_box( 'Exporting Autopollock mean z from zone '+ str(zone))
                    EvVar.Properties.Analysis.ExcludeAboveLine = str(params.exclude_above_line[k])  #this is working even though it spits gibberish to the screen
                    EvVar.Properties.Analysis.ExcludeBelowLine = str(params.exclude_below_line[k])
                    ExportFileName = params.output_dir_mb2 + '\\' + EvExportName + 'p2' +'.csv' #output .csv filename- edited name on 7/3/2016 by nel requested by patrick
                    exporttest = EvVar.ExportIntegrationByRegionsByCellsAll(ExportFileName)
                    if exporttest != 1:
                        self.refresh_text_box('The export has failed for zone '+str(zone))
                        self.refresh_text_box(ExportFileName)
                    else:
                        self.refresh_text_box('Zone '+ str(zone) +' Export Complete')
        
        EvApp.CloseFile(EvFile) #close .ev file
        EvApp.Quit() #quit echoview to refresh for next .EV file, just in case
        return self.exporttestMB2

    def closeEvent(self, event=None):
        """
          Clean up when the main window is closed.
        """
        self.appSettings.setValue('winposition', self.pos())
        self.appSettings.setValue('winsize', self.size())
        self.appSettings.setValue('latestShip',self.shipBox.currentText())
        self.appSettings.setValue('latestSurvey',self.surveyBox.currentText())
        self.appSettings.setValue('latestDataSet',self.dataSetBox.currentText())
        self.appSettings.setValue('latestInDir',self.input_dir.text())
        self.appSettings.setValue('latestOutDir',self.output_dir_mb2.text())
        self.appSettings.setValue('latestRawDir',self.rawFilesDir.text())
        self.appSettings.setValue('latestCalFile',self.cal_file.text())
        self.appSettings.setValue('latestFileSet',self.fileset_name.text())
        #  close our connection to the database
        try:
            self.db.close()
        except:
            pass


    def checkWindowLocation(self, position, size, padding=[5, 25]):
        '''
        checkWindowLocation accepts a window position (QPoint) and size (QSize)
        and returns a potentially new position and size if the window is currently
        positioned off the screen.

        This function uses QScreen.availableVirtualGeometry() which returns the full
        available desktop space *not* including taskbar. For all single and "typical"
        multi-monitor setups this should work reasonably well. But for multi-monitor
        setups where the monitors may be different resolutions, have different
        orientations or different scaling factors, the app may still fall partially
        or totally offscreen. A more thorough check gets complicated, so hopefully
        those cases are very rare.

        If the user is holding the <shift> key while this method is run, the
        application will be forced to the primary monitor.
        '''

        #  create a QRect that represents the app window
        appRect = QtCore.QRect(position, size)

        #  check for the shift key which we use to force a move to the primary screem
        resetPosition = QtGui.QGuiApplication.queryKeyboardModifiers() == QtCore.Qt.KeyboardModifier.ShiftModifier
        if resetPosition:
            position = QtCore.QPoint(padding[0], padding[0])

        #  get a reference to the primary system screen - If the app is off the screen, we
        #  will restore it to the primary screen
        primaryScreen =QtGui. QGuiApplication.primaryScreen()

        #  assume the new and old positions are the same
        newPosition = position
        newSize = size

        #  Get the desktop geometry. We'll use availableVirtualGeometry to get the full
        #  desktop rect but note that if the monitors are different resolutions or have
        #  different scaling, some parts of this rect can still be offscreen.
        screenGeometry = primaryScreen.availableVirtualGeometry()

        #  if the app is partially or totally off screen or we're force resetting
        if resetPosition or not screenGeometry.contains(appRect):

            #  check if the upper left corner of the window is off the left side of the screen
            if position.x() < screenGeometry.x():
                newPosition.setX(screenGeometry.x() + padding[0])
            #  check if the upper right is off the right side of the screen
            if position.x() + size.width() >= screenGeometry.width():
                p = screenGeometry.width() - size.width() - padding[0]
                if p < padding[0]:
                    p = padding[0]
                newPosition.setX(p)
            #  check if the top of the window is off the top/bottom of the screen
            if position.y() < screenGeometry.y():
                newPosition.setY(screenGeometry.y() + padding[0])
            if position.y() + size.height() >= screenGeometry.height():
                p = screenGeometry.height() - size.height() - padding[1]
                if p < padding[0]:
                    p = padding[0]
                newPosition.setY(p)

            #  now make sure the lower right (resize handle) is on the screen
            if (newPosition.x() + newSize.width()) > screenGeometry.width():
                newSize.setWidth(screenGeometry.width() - newPosition.x() - padding[0])
            if (newPosition.y() + newSize.height()) > screenGeometry.height():
                newSize.setHeight(screenGeometry.height() - newPosition.y() - padding[1])

        return [newPosition, newSize]


class parameterSetup:
    setup = True

# global function.  used to change from a transect number to the 't###' format commonly used in the .EV file names.
def transect_string(transect_names):
    if '.' in transect_names:
        transect_1 = transect_names.split('.')[0]
        transect_2 = transect_names.split('.')[1]
        if len(transect_1) == 1:
            transect_name = 't00' + transect_1
        elif len(transect_1) ==2:
            transect_name = 't0' + transect_1
        elif len(transect_1) ==3:
            transect_name = 't' + transect_1
        transect_name = transect_name + '.'+transect_2
    else:
        if len(transect_names) == 1:
            transect_name = 't00' + transect_names
        elif len(transect_names) ==2:
            transect_name = 't0' + transect_names
        elif len(transect_names) ==3:
            transect_name = 't' + transect_names
    return transect_name


# main, runs all from command line
if __name__ == "__main__":
    '''
    PARSE THE COMMAND LINE ARGS
    '''
    import argparse

    #  specify the default credential and schema values
    acoustic_schema = "macebase2"
    bio_schema = "clamsbase2"
    odbc_connection = None
    username = None
    password = None

     #  create the argument parser. Set the application description.
    parser = argparse.ArgumentParser(description='Exporter')

    #  specify the positional arguments: ODBC connection, username, password
    parser.add_argument("odbc_connection", nargs='?', help="The name of the ODBC connection used to connect to the database.")
    parser.add_argument("username", nargs='?', help="The username used to log into the database.")
    parser.add_argument("password", nargs='?', help="The password for the specified username.")

    #  specify optional keyword arguments
    parser.add_argument("-a", "--acoustic_schema", help="Specify the acoustic database schema to use.")
    parser.add_argument("-b", "--bio_schema", help="Specify the biological database schema to use.")

    #  parse our arguments
    args = parser.parse_args()

    #  and assign to our vars (and convert from unicode to standard strings)
    if (args.acoustic_schema):
        #  strip off the leading space (if any)
        acoustic_schema = str(args.acoustic_schema).strip()
    if (args.bio_schema):
        #  strip off the leading space (if any)
        bio_schema = str(args.bio_schema).strip()
    if (args.odbc_connection):
        odbc_connection = str(args.odbc_connection)
    if (args.username):
        username = str(args.username)
    if (args.password):
        password = str(args.password)

    app = QtWidgets.QApplication(sys.argv)
    form = Exporter(odbc_connection, username, password,  acoustic_schema, bio_schema)
    form.show()
    app.exec()
