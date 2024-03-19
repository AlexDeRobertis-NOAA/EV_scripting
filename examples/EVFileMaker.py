
import os
import time
import re
from PyQt6.QtCore import *
from PyQt6.QtGui import *
from PyQt6.QtWidgets import *
from ui import ui_EVFileMaker
import sys, traceback, glob
import win32com.client
import SelectSurveyDlg
from MaceFunctions import connectdlg,  dbConnection

class EVFileMaker(QMainWindow, ui_EVFileMaker.Ui_MainWindow):

    #  set the maximum time in seconds that we will wait for EV to index the raw files
    #  we add to our EV file.
    EV_INDEXING_TIMEOUT = 60

    #  define the window size in seconds after the start time of a file that we
    #  will use to consider adding the file to the EV file to ensure at least one
    #  partial interval before the start of our transect interval.
    #  At 10 knts 0.5nmi = 3 minutes
    JUSTMISSEDTHRESH = 5 * 60


    def __init__(self, odbc_connection, username, password, bio_schema, parent=None):
        super(EVFileMaker, self).__init__(parent)
        self.setupUi(self)

        self.odbc = odbc_connection
        self.dbUser = username
        self.dbPassword = password
        self.bioSchema = bio_schema

        #  get the application settings
        self.appSettings = QSettings('afsc.noaa.gov', 'EVFileMaker')
        size = self.appSettings.value('winsize', QSize(760,350))
        position = self.appSettings.value('winposition', QPoint(10,10))
        ek_dir = self.appSettings.value('ek_dir', QDir.home().path())
        self.EKFilePathEdit.setText(ek_dir)
        dest_dir = self.appSettings.value('dest_dir', QDir.home().path())
        self.destinationEdit.setText(dest_dir)
        templ_file= self.appSettings.value('templ_file', QDir.home().path())
        self.templateEvFileEdit.setText(templ_file)
        temp2_file= self.appSettings.value('temp2_file', QDir.home().path())
        self.ECSFileEdit.setText(temp2_file)
        lineregion_dir = self.appSettings.value('lineregion_dir', QDir.home().path())
        self.lineregionPath.setText(lineregion_dir)

        #  check the current position and size to make sure the app is on the screen
        position, size = self.checkWindowLocation(position, size)

        #  now move and resize the window
        self.move(position)
        self.resize(size)

        #  connect the signals
        self.actionExit.triggered.connect(self.close)
        self.actionChange_Survey.triggered.connect(self.changeSurvey)
        self.makeFileBtn.clicked.connect(self.makeFileSetup)
        self.pbPickTemplate.clicked.connect(self.pickFile)
        self.pbPickRaw.clicked.connect(self.pickFile)
        self.pbPickDest.clicked.connect(self.pickFile)
        self.pbPickECS.clicked.connect(self.pickFile)
        self.lineregionCheck.clicked.connect(self.enableLineRegion)
        self.lineregionButton.clicked.connect(self.pickFile)
        self.reloadBtn.clicked.connect(self.getTransects)


        #  create a label on the status bar for feedback
        self.statusLabel = QLabel('')
        self.statusBar.addPermanentWidget(self.statusLabel)

        #  set the application icon
        try:
            self.setWindowIcon(QIcon('./resources/unicorn5.png'))
            self.reloadBtn.setIcon(QIcon('./resources/refresh.png'))
            self.appIcon.setPixmap(QPixmap('./resources/unicorn5.png').scaled(90,90,
                    transformMode = Qt.TransformationMode.SmoothTransformation))
        except:
            pass

        #  connect to the db in a separate init method
        timer = QTimer(self)
        timer.setSingleShot(True)
        timer.timeout.connect(self.applicationInit)
        timer.start(1)


    def applicationInit(self):

        #  check if we're missing any of our required connection parameters
        if ((self.odbc == None) or (self.dbUser == None) or
            (self.dbPassword == None)):

            #  we're missing at least one - display the connect dialog to get the rest of the args.
            #  Note the use of the new createConnection argument which keeps ConnectDlg from creating
            #  an instance of dbConnection. We'll do that below.
            #  Also note the new enableBioschema argument which will disable the bioschema combobox
            connectDlg = connectdlg.ConnectDlg(self.odbc, self.dbUser, self.dbPassword, label='EVFileMaker',
                    enableBioschema=False, createConnection=False, parent=self)

            if not connectDlg.exec():
                #  user hit cancel so we exit this example program
                self.close()
                return

            #  update our connection credentials
            self.odbc = connectDlg.getSource()
            self.dbUser = connectDlg.getUsername()
            self.dbPassword = connectDlg.getPassword()

        #  create the database connection
        self.db = dbConnection.dbConnection(self.odbc, self.dbUser,
                self.dbPassword, 'EVFileMaker')

        try:
            #  attempt to connect to the database
            self.db.dbOpen()
        except dbConnection.DBError as e:
            #  ooops, there was a problem
            errorMsg = ('Unable to connect to ' + self.dbUser+ '@' +
                    self.odbc + '\n' + e.error)
            QMessageBox.critical(self, "Databse Login Error", errorMsg)
            self.close()
            return

        #  query CLAMS to determine the current active ship and survey
        sql = ("SELECT parameter_value FROM " + self.bioSchema + ".application_configuration " +
                "WHERE parameter='ActiveShip'")
        query = self.db.dbQuery(sql)
        self.ship, = query.first()
        sql = ("SELECT parameter_value FROM " + self.bioSchema + ".application_configuration " +
                "WHERE parameter='ActiveSurvey'")
        query = self.db.dbQuery(sql)
        self.survey, = query.first()

        #  set the dataset
        sql = ("SELECT data_set_id FROM macebase2.data_sets WHERE ship=" + self.ship +
                " AND survey=" + self.survey + " ORDER BY data_set_id ASC")
        query = self.db.dbQuery(sql)
        self.dataset, = query.first()

        #  and update the labels
        self.shipLabel.setText(self.ship)
        self.surveyLabel.setText(self.survey)
        self.datasetLabel.setText(self.dataset)

        #  update the transects combobox
        self.getTransects()


    def changeSurvey(self):
        '''
        changeSurvey presents the survey selection dialog and if a survey is selected
        sets it as active for the session and gets that surveys transects.
        '''
        #  show the dialog
        surveyDlg = SelectSurveyDlg.SelectSurveyDlg(self.db, self.bioSchema, self.ship, self.survey, self.dataset)
        surveyDlg.exec()
        if (surveyDlg.ship):
            #  something was selected, update
            self.ship = surveyDlg.ship
            self.survey = surveyDlg.survey
            self.dataset = surveyDlg.dataset
            self.shipLabel.setText(self.ship)
            self.surveyLabel.setText(self.survey)
            self.datasetLabel.setText(self.dataset)

            #  update the transects combobox
            self.getTransects()


    def enableLineRegion(self):
        '''
        enableLineRegion allows the user to insert saved lines and regions into the EV file
        The option is enabled or disabled based on the check mark
        '''

        if self.lineregionCheck.isChecked():
            self.label_10.setEnabled(True)
            self.lineregionPath.setEnabled(True)
            self.lineregionButton.setEnabled(True)

            self.pickFile()

        else:
            self.label_10.setEnabled(False)
            self.lineregionPath.setEnabled(False)
            self.lineregionButton.setEnabled(False)

    def pickFile(self):
        '''
        pickFile handles picking the template file and the .raw and output directories
        '''

        button = self.sender()
        if (button == self.pbPickTemplate):
            fileName = QFileDialog().getOpenFileName(self, "Select a template file", self.templateEvFileEdit.text(),
                    "Echoview Files (*.ev)")
            fileName = fileName[0]
            if (fileName == ''):
                return
            self.templateEvFileEdit.setText(fileName)
            self.appSettings.setValue('templ_file',self.templateEvFileEdit.text())

        elif (button == self.pbPickRaw) :
            #  set the raw directory
            path = QFileDialog.getExistingDirectory(self, "Select raw file directory", self.EKFilePathEdit.text())
            if (path == ''):
                return
            self.EKFilePathEdit.setText(path)
            self.appSettings.setValue('ek_dir',self.EKFilePathEdit.text())

        elif (button == self.pbPickDest):
            #  set the destination directory
            path = QFileDialog.getExistingDirectory(self, "Select output directory", self.destinationEdit.text())
            if (path == ''):
                return
            self.destinationEdit.setText(path)
            self.appSettings.setValue('dest_dir',self.destinationEdit.text())
            
        button = self.sender()
        if (button == self.pbPickECS):
            fileName = QFileDialog().getOpenFileName(self, "Select an ECS file", self.ECSFileEdit.text(),
                    "ECS Files (*.ecs)")
            fileName = fileName[0]
            if (fileName == ''):
                return
            self.ECSFileEdit.setText(fileName)
            self.appSettings.setValue('temp2_file',self.ECSFileEdit.text())

        elif (button == self.lineregionButton):
            # set the path to the lines and regions files
            path = QFileDialog.getExistingDirectory(self, "Select output directory", self.lineregionPath.text())
            if (path == ''):
                return
            self.lineregionPath.setText(path)
            self.appSettings.setValue('lineregion_dir',self.lineregionPath.text())


    def getTransects(self):
        '''
        getTransects updates the transect combobox
        '''

        #  clear the combobox
        self.cbTransects.clear()
        self.transect_list = []

        #  get a list of the completed transects
        sql = ("SELECT transect FROM transect_events WHERE transect_event_type='ET' AND " +
                "ship=" + self.ship + " AND survey=" + self.survey + " GROUP BY transect " +
                "ORDER BY transect DESC")
        query = self.db.dbQuery(sql)

        #  add them to the combobox
        for transect, in query:
            self.cbTransects.addItem(transect)
            self.transect_list.append(transect)
        self.cbTransects.setCurrentIndex(-1)


    def makeFileSetup(self):
        if self.doallCheck.isChecked():
            for ind in reversed(range(0, len(self.transect_list))):
                self.cbTransects.setCurrentIndex(ind)
                self.makeFile()
        else:
            self.makeFile()


    def makeFile(self):


        # check that all of our inputs are complete
        if (self.cbTransects.currentText() == ''):
            QMessageBox.critical(self, "Error", "Please select a transect number.")
            return
        if not QDir(self.EKFilePathEdit.text()).exists():
            QMessageBox.critical(self, "Error", "EK raw file directory does not exist.")
            return
        if not QDir(self.destinationEdit.text()).exists():
            QMessageBox.critical(self, "Error", "File destination directory does not exist.")
            return
        if not QFile(self.templateEvFileEdit.text()).exists():
            QMessageBox.critical(self, "Error", "Template file doesn't exist.")
            return

        #  get the dataset properties
        self.updateStatusBar('Getting dataset parameters...')
        sql = ("SELECT b.source_name,a.layer_reference,a.interval_type," +
                "a.interval_units,a.interval_length FROM macebase2.data_sets a," +
                "macebase2.acoustic_data_sources b WHERE ship=" + self.ship +
                " AND survey=" + self.survey + " AND a.data_set_id=" + self.dataset +
                " AND a.source_id=b.source_id")
        query = self.db.dbQuery(sql)
        sourceName, layerReference, intervalType, intervalUnits, intervalLength = query.first()

        #  get the surface exclusion line depth
        sql = ("SELECT b.exclusion_line_offset from zones a, exclusion_lines b " +
                "WHERE ship=" + self.ship + " AND survey=" + self.survey +
                " AND a.data_set_id=" + self.dataset + " AND " +
                "a.upper_exclusion_name='surface_exclusion' AND " +
                "a.upper_exclusion_line=b.exclusion_line_id")
        query = self.db.dbQuery(sql)
        surface_exclusion_depth, = query.first()
        if surface_exclusion_depth is None:
            QMessageBox.critical(self, "Error", "Unable to find the surface exclusion line depth. " +
                    "Have you created your zone(s) for this dataset and is the upper_exclusion_name " +
                    "for your upper most zone set to 'surface_exclusion'?")
            self.updateStatusBar('')
            return
        try:
            surface_exclusion_depth = float(surface_exclusion_depth)
        except:
            QMessageBox.critical(self, "Error", "Invalid (non-numeric) surface exclusion line depth found. " +
                "Please correct your zone(s) for this data set and try again.")
            self.updateStatusBar('')
            return

        #  get the bottom offset.
        sql = ("SELECT b.exclusion_line_offset from zones a, exclusion_lines b " +
                "WHERE ship=" + self.ship + " AND survey=" + self.survey +
                " AND a.data_set_id=" + self.dataset + " AND " +
                "a.lower_exclusion_name='bottom_exclusion' AND " +
                "a.lower_exclusion_line=b.exclusion_line_id")
        query = self.db.dbQuery(sql)
        botom_line_offset, = query.first()
        if botom_line_offset is None:
            QMessageBox.critical(self, "Error", "Unable to find the bottom exclusion line offset. " +
                    "Have you created your zone(s) for this dataset and is the lower_exclusion_name " +
                    "for your deepest zone set to 'bottom_exclusion'?")
            self.updateStatusBar('')
            return
        try:
            botom_line_offset = float(botom_line_offset)
        except:
            QMessageBox.critical(self, "Error", "Invalid (non-numeric) bottom exclusion line offset found. " +
                "Please correct your zone(s) for this data set and try again.")
            self.updateStatusBar('')
            return

        #  generate the EV filename
        transect = '%03i' % float(self.cbTransects.currentText())
        EvFileName = 'v' + self.ship + '-s' + self.survey + '-x2-f38-t' + transect + '-z0.ev'
        self.EvFileName = os.path.normpath(str(self.destinationEdit.text())) + os.sep + EvFileName

        # check to see if file exists
        if QFile(self.EvFileName).exists():
            reply = QMessageBox.warning(self, "WARNING", "This EV file already exists. Do you want to " +
                    "replace it?",  QMessageBox.StandardButton.Yes, QMessageBox.StandardButton.No)
            if (reply == QMessageBox.StandardButton.No):
                return

        #  get the events and times for this transect
        self.updateStatusBar('Determining time spans for this transect...')
        transect = self.cbTransects.currentText()
        event_times=[]
        events=[]
        sql = ("SELECT transect_event_type, TO_CHAR(time) FROM transect_events WHERE transect=" +
            transect + " AND ship=" + self.ship + " AND survey=" + self.survey + " ORDER BY time ASC")
        query = self.db.dbQuery(sql)
        for event_type, evtime in query:
            event_times.append(QDateTime().fromString(evtime,'MM/dd/yyyy hh:mm:ss.zzz'))
            events.append(event_type)

        #  Create lists of the starting and ending times of our transect segments
        #  to use to build our .raw file list
        start_times = []
        end_times = []
        if not 'BT' in events:
            # this is a non broken transect
            start_times.append(event_times[events.index('ST')])
            end_times.append(event_times[events.index('ET')])
        elif events.count('BT')==1:
            # one break
            start_times.append(event_times[events.index('ST')])
            end_times.append(event_times[events.index('BT')])
            start_times.append(event_times[events.index('RT')])
            end_times.append(event_times[events.index('ET')])
        else:
            # many breaks
            start_times.append(event_times[events.index('ST')])
            cnt=events.count('BT')
            for i in range(cnt):
                idx = events.index('BT')
                end_times.append(event_times.pop(idx))
                events.pop(idx)
                idx = events.index('RT')
                start_times.append(event_times.pop(idx))
                events.pop(idx)
            end_times.append(event_times[events.index('ET')])

        try:

            #  get a listing of all of the raw files in the raw file firectory
            self.updateStatusBar('Finding the files associated with timespans...')
            EKfilelist = sorted(glob.glob(str(os.path.normpath(self.EKFilePathEdit.text())) +
                    os.sep + '*.raw'))
            if (not EKfilelist):
                QMessageBox.critical(self, "Error", "No .raw files found in raw file directory.")
                return

            #  work through the raw file list to determine which files are within our transect events
            keepFiles = []
            file_index = range(len(EKfilelist) - 1)
            for i in range(len(start_times)):
                keep_ind=[]
                keep_files = False
                for j in file_index:
                    #  get the current and next file file names
                    fullPath = EKfilelist[j]
                    filename = fullPath.split(os.sep)[-1]
                    nextPath = EKfilelist[j + 1]
                    nextName = nextPath.split(os.sep)[-1]

                    #  extract the data files' date/time string
                    #  2/19/21 - this method was extended to us regular expressions
                    #            to extract the date/time to allow for more flexibility.
                    try:
                        fileDate = re.findall('D[0-9]{8}-T[0-9]{6}', filename)[0]
                        fileDate =  QDateTime().fromString(fileDate,'DyyyyMMdd-Thhmmss')
                    except:
                        QMessageBox.critical(self, "Error", "The raw file " + filename +
                            " is misnamed. Raw files must have the date and time in the name " +
                            "in the form DYYYYMMDD-Thhmmss.")
                        return
                    try:
                        nextFileDate = re.findall('D[0-9]{8}-T[0-9]{6}', nextName)[0]
                        nextFileDate =  QDateTime().fromString(nextFileDate,'DyyyyMMdd-Thhmmss')
                    except:
                        QMessageBox.critical(self, "Error", "The raw file " + nextName +
                            " is misnamed. Raw files must have the date and time in the name " +
                            "in the form DYYYYMMDD-Thhmmss.")
                        return

                    #print(start_times[i],end_times[i],fileDate)

                    if (keep_files):
                        keepFiles.append(fullPath)
                        keep_ind.append(j)
                    else:
                        #  check if this file falls within
                        if ((fileDate <= start_times[i] <= nextFileDate) or
                                abs(fileDate.secsTo(start_times[i])) < self.JUSTMISSEDTHRESH):

                            keepFiles.append(fullPath)
                            keep_ind.append(j)
                            keep_files = True

                    if (fileDate <= end_times[i] <= nextFileDate):
                        keep_files = False

#                    if (((fileDate >=start_times[i]) and (fileDate<=end_times[i])) or
#                            abs(fileDate.secsTo(start_times[i])) < self.JUSTMISSEDTHRESH):
#                        keepFiles.append(fullPath)
#                        keep_ind.append(j)


                if (not keep_ind):
                    QMessageBox.critical(self, "Error", "There are no data files for your " +
                            "transect segment that starts at " + str(start_times[i]) +
                            ". This usually means the data hasn't been copied into your " +
                            "EK80 raw data directory yet.")
                    return

                #start_file = min(keep_ind)
                #keepFiles.append(EKfilelist[start_file - 1])

            #print(keepFiles)

            #Open up Echoview
            self.updateStatusBar('Opening echoview...')
            QApplication.setOverrideCursor(Qt.CursorShape.WaitCursor)
            EvApp = win32com.client.Dispatch("EchoviewCom.EvApplication")
            license = EvApp.IsLicensed()
            if (license == 0):
                self.updateStatusBar('ERROR: No dongle or no licensed scripting module.')
                QMessageBox.warning(self, "ERROR", 'No Scripting Module Found')
                EvApp.Quit()
                self.updateStatusBar('')
                QApplication.restoreOverrideCursor()
                return

            #  minimize
            EvApp.Minimize()

            #  create the new EV file
            self.updateStatusBar('Loading template...')
            EvFile = EvApp.NewFile(self.templateEvFileEdit.text())
            # add the ECS file
            Evfileset = EvFile.Filesets.FindByName('Fileset 1')
            Evfileset.SetCalibrationFile(self.ECSFileEdit.text())
            #  add the .raw files
            self.updateStatusBar('Adding .raw files...')
            for file in keepFiles:
                EvFile.Filesets.Item(0).DataFiles.Add(file)

            #  we must wait for EV to index all of the raw files before proceeding since
            #  our line created below will not be complete if some files haven't been indexed
            #  this shows up as the bottom_exclusion line being incomplete or "flat" for
            #  whole raw file segments.
            self.updateStatusBar('Waiting for echoview to index .raw files...')

            #  the EvFile.PreRead method doesn't do squat here - we have to wait
            #  for the files to be indexed

            #  this was my initial attempt at doing this before I knew about
            #  PreReadDataFiles. We'll try
            indexing = True
            waitTime = 0
            while (indexing):
                time.sleep(3)
                allIndexed = True
                for file in keepFiles:
                    #  check if the .evi file exists
                    allIndexed &= QFile(file+'.evi').exists()
                waitTime += 3

                #  check to see if EV's done or we've timed out
                if (allIndexed) or (waitTime > self.EV_INDEXING_TIMEOUT):
                    indexing = False

            #  give EV just a bit more time after indexing all of the files
            time.sleep(3)

            #  At this time (EV 8.0.x) we cannot create time based regions so we cannot
            #  directly script the creation of our marker regions. Instead we have to
            #  create an EVR file, import it, then delete it
            if not self.lineregionCheck.isChecked():
                self.updateStatusBar('Importing regions...')
                tempFilePath = os.path.normpath(str(self.destinationEdit.text()))
                evrFile = self.createEVRFile(transect, tempFilePath)
                EvFile.Import(evrFile)
                os.remove(evrFile)

            #  create the new bottom_exclusion line based on the mean of all sounder detected bottom lines
            self.updateStatusBar('Creating new bottom_exclusion line...')
            EvLine = EvFile.Lines.FindByName('Mean of all sounder-detected bottom lines')
            EvNewLine = EvFile.Lines.CreateOffsetLinear(EvLine,1, botom_line_offset,1)
            EvLineOld = EvFile.Lines.FindByName('bottom_exclusion')
            EvLineOld.OverwriteWith(EvNewLine)
            EvFile.Lines.Delete(EvNewLine)

            #  create the surface exclusion line
            self.updateStatusBar('Creating new surface_exclusion line...')
            EvNewLine = EvFile.Lines.CreateFixedDepth(surface_exclusion_depth)
            EvLineOld = EvFile.Lines.FindByName('surface_exclusion')
            EvLineOld.OverwriteWith(EvNewLine)
            EvFile.Lines.Delete(EvNewLine)
            
            # Future work can be to set the 0.5/2/3 m off bottom lines to be based off the depth of the bottom exclusion
            # This changes in winter (-0.5) and EBS summer (-0.25)
            # Need to figure out how to use scripting to create a virtual line with an offset.
            #  create the 0.5 m off bottom line
#            self.updateStatusBar('Creating new 0.5 m off bottom line...')
#            EvLine = EvFile.Lines.FindByName('Mean of all sounder-detected bottom lines')
#            half_offset = -0.5 - botom_line_offset
#            EvNewLine = EvFile.Lines.CreateOffsetLinear(EvLine,1, int(half_offset),1)
#            EvLineOld = EvFile.Lines.FindByName('0.5m off bottom')
#            EvLineOld.OverwriteWith(EvNewLine)
#            EvFile.Lines.Delete(EvNewLine)
#            
#            #  create the 2 m off bottom line
#            self.updateStatusBar('Creating new 2 m off bottom line...')
#            two_offset = -2 - botom_line_offset
#            EvNewLine = EvFile.Lines.CreateOffsetLinear(EvLine,1, two_offset,1)
#            EvLineOld = EvFile.Lines.FindByName('2m off botom')
#            EvLineOld.OverwriteWith(EvNewLine)
#            EvFile.Lines.Delete(EvNewLine)
#            
#            #  create the surface exclusion line
#            self.updateStatusBar('Creating new 3 m off bottom line...')
#            three_offset = -3 - botom_line_offset
#            EvNewLine = EvFile.Lines.CreateOffsetLinear(EvLine,1, three_offset,1)
#            EvLineOld = EvFile.Lines.FindByName('3m off bottom')
#            EvLineOld.OverwriteWith(EvNewLine)
#            EvFile.Lines.Delete(EvNewLine)

            # Import lines and regions
            if self.lineregionCheck.isChecked():
                # Find all the line and region files for this transect and load them
                # Get line file names
                t = 't%03i' % float(self.cbTransects.currentText())
                lineFiles = glob.glob(str(os.path.normpath(self.lineregionPath.text()))+os.sep+'Lines'+os.sep+'*'+t+'*')
                for file in lineFiles:

                    splits = file.split('-')
                    # Get the line name that this line should replace.  It is embedded into the file name, after the transect number
                    lineName = splits[splits.index(t)+1]
                    needsFix = lineName.find('.evl')
                    if needsFix > -1:
                        lineName = lineName[:needsFix]
                    EvLineOld = EvFile.Lines.FindByName(lineName)

                    if EvLineOld and EvLineOld.AsLineEditable():
                        EvFile.Import(file)
                        # Get the line name that was inserted into EV.  Strip off path and the end '.evl'
                        newlineName = file.split('\\')[-1][:-4]
                        # If there is a decimal in the name, only the string before the decimal is used by EV as
                        #  the line name (it assumes that is the '.evl' part)
                        newlineName = newlineName.split('.')[0]
                        EvLineNew = EvFile.Lines.FindByName(newlineName)
                        # If the evl file is empty, which apparently happens, then there will be no line to find, so skip it
                        if EvLineNew:
                            # This line already exists, so replace it
                            EvLineOld.OverwriteWith(EvLineNew)
                            EvFile.Lines.Delete(EvLineNew)
                            # This line doesn't exist, so create it/rename the new one as the embedded name
                    elif not EvLineOld:
                        EvFile.Import(file)
                        # Get the line name that was inserted into EV.  Strip off path and the end '.evl'
                        newlineName = file.split('\\')[-1][:-4]
                        # If there is a decimal in the name, only the string before the decimal is used by
                        #  EV as the line name (it assumes that is the '.evl' part)
                        newlineName = newlineName.split('.')[0]
                        EvLineNew = EvFile.Lines.FindByName(newlineName)
                        # If the evl file is empty, which apparently happens, then there will be no line to find, so skip it
                        if EvLineNew:
                            EvLineNew.Name = lineName

                regionFiles = glob.glob(str(os.path.normpath(str(self.lineregionPath.text())))+os.sep+'Regions'+os.sep+'*'+t+'*')
                for region in regionFiles:
                    EvFile.Import(region)
            #  save the changes
            self.updateStatusBar('Saving file...')
            EvFile.SaveAs(self.EvFileName)
            EvApp.CloseFile(EvFile)
            EvApp.Quit()

            #  give EV some time to clean up
            time.sleep(3)

            #  update the GUI and inform the user we're done
            QApplication.restoreOverrideCursor()
            self.statusLabel.setText('')
            if not self.doallCheck.isChecked():
                QMessageBox.information(self, "Congratulations!", "EV file has been created.")

        except:
            #  there was an error - give the user a wee bit of feedback
            self.sendError()


    def sendError(self):
        t=traceback.format_exception(*sys.exc_info())
        error=''
        for line in t:
            error=error+line+'\n'
        QMessageBox.warning(self, "ERROR", error)


    def closeEvent(self,  event=None):

        #  update the app settings
        self.appSettings.setValue('winposition', self.pos())
        self.appSettings.setValue('winsize', self.size())
        self.appSettings.setValue('ek_dir', self.EKFilePathEdit.text())
        self.appSettings.setValue('dest_dir',self.destinationEdit.text())
        self.appSettings.setValue('templ_file',self.templateEvFileEdit.text())

        event.accept()


    def createEVRFile(self, transect, path):

        #  determine the number of events for this transect
        sql = ("SELECT count(*) FROM macebase2.transect_events WHERE transect=" +
                transect + " AND " + "ship=" + self.ship + " AND survey=" + self.survey)
        query = self.db.dbQuery(sql)
        nEvents, = query.first()

        #  ensure we have a proper path
        pathText = os.path.normpath(str(path))
        pathText = pathText + os.sep + 'Transect_' + transect + '.evr'

        #  open te output file and write the header
        evrFile = QFile(pathText)
        evrFile.open(QIODevice.OpenModeFlag.ReadWrite|QIODevice.OpenModeFlag.Truncate)
        evrStream = QTextStream(evrFile)
        evrStream << 'EVRG 7 7.1.34.30284' << '\r\n'
        evrStream << nEvents << '\r\n'
        evrStream << '\r\n'

        #get the data
        sql = ("SELECT transect_event_type, TO_CHAR(time) FROM macebase2.transect_events " +
                "WHERE transect=" + transect + " AND ship=" + self.ship + " AND survey=" + self.survey)
        query = self.db.dbQuery(sql)

        #  loop thru the events for this transect
        cnt = 1
        for event_type, evtime in query:

            #  convert the time to a QDateTime
            t = QDateTime().fromString(evtime,'MM/dd/yyyy hh:mm:ss.zzz')
            #  get the start date and time components
            d1 = t.toString('yyyyMMdd')
            t1 = t.toString('hhmmsszzz0')
            #  now generate the end date and time
            tf = t.addMSecs(1003)
            d2 = tf.toString('yyyyMMdd')
            t2 = t.toString('hhmmsszzz0')

            #  write the region data for this event
            evrStream << '13 4 ' << str(cnt) << ' 0 6 -1 1 ' << d1 << ' ' << t1 << \
                    '  -9999.99 ' << d2 << ' ' << t2 << '  9999.99\r\n'
            evrStream << '1\r\n'
            evrStream << event_type << '_' << transect << '\r\n'
            evrStream << '0\r\n'
            evrStream << 'Unclassified\r\n'
            evrStream << d1 << ' ' << t1 << ' -9999.9900000000 ' << d1 << ' ' << t1 << \
                    '  9999.9900000000 ' << d2 << ' ' << t2 << ' 9999.9900000000 ' << d2 << \
                    ' ' << t2 << ' -9999.9900000000 2 \r\n'
            evrStream << event_type << '_' << transect << '\r\n'
            evrStream<<'\r\n'

            #  increment the counter
            cnt += 1

        #  close the file
        evrFile.close()

        return str(pathText)


    def updateStatusBar(self, text, color='0030FF'):
        '''
        updateStatusBar simply formats text for our status bar and because these updates
        usually happen in the middle of long running methods we force Qt to process the event
        queue to update the gui and actually draw our text
        '''
        self.statusLabel.setText('<span style=" color:#' + color + ';">' + text + '</span>')
        QApplication.processEvents()


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
        appRect = QRect(position, size)

        #  check for the shift key which we use to force a move to the primary screem
        resetPosition = QGuiApplication.queryKeyboardModifiers() == Qt.KeyboardModifier.ShiftModifier
        if resetPosition:
            position = QPoint(padding[0], padding[0])

        #  get a reference to the primary system screen - If the app is off the screen, we
        #  will restore it to the primary screen
        primaryScreen = QGuiApplication.primaryScreen()

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


if __name__ == "__main__":

    import argparse

    #  specify the default credential and schema values
    bio_schema = "clamsbase2"
    odbc_connection = None
    username = None
    password = None

    #  create the argument parser. Set the application description.
    parser = argparse.ArgumentParser(description='EVFileMaker')

    #  specify the positional arguments: ODBC connection, username, password
    parser.add_argument("odbc_connection", nargs='?', help="The name of the ODBC connection used to connect to the database.")
    parser.add_argument("username", nargs='?', help="The username used to log into the database.")
    parser.add_argument("password", nargs='?', help="The password for the specified username.")

    #  specify optional keyword arguments
    parser.add_argument("-b", "--bio_schema", help="Specify the biological database schema to use.")

    #  parse our arguments
    args = parser.parse_args()

    #  and assign to our vars (and convert from unicode to standard strings)
    if (args.bio_schema):
        #  strip off the leading space (if any)
        bio_schema = str(args.bio_schema).strip()
    if (args.odbc_connection):
        odbc_connection = str(args.odbc_connection)
    if (args.username):
        username = str(args.username)
    if (args.password):
        password = str(args.password)

    app = QApplication(sys.argv)
    form = EVFileMaker(odbc_connection, username, password, bio_schema)
    form.show()
    app.exec()
