'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'
'Copyright 2016 Signalworks Pty Ltd, ABN 26 066 681 598

'Licensed under the Apache License, Version 2.0 (the "License");
'you may not use this file except in compliance with the License.
'You may obtain a copy of the License at
'
'http://www.apache.org/licenses/LICENSE-2.0
'
'Unless required by applicable law or agreed to in writing, software
'distributed under the License is distributed on an "AS IS" BASIS,
''WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
'See the License for the specific language governing permissions and
'limitations under the License.
'
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Public Class Main
    'The ADVL_Shares_1 application stores historical data for publicly traded shares and performs processing and analysis techniques aimed at optimizing share trading returns.

#Region " Coding Notes - Notes on the code used in this class." '------------------------------------------------------------------------------------------------------------------------------

    'ADD THE SYSTEM UTILITIES REFERENCE: ==========================================================================================
    'The following references are required by this software: 
    'Project \ Add Reference... \ ADVL_Utilities_Library_1.dll
    'The Utilities Library is used for Project Management, Archive file management, running XSequence files and running XMessage files.
    'If there are problems with a reference, try deleting it from the references list and adding it again.

    'ADD THE SERVICE REFERENCE: ===================================================================================================
    'A service reference to the Message Service must be added to the source code before this service can be used.
    'This is used to connect to the Application Network.

    'Adding the service reference to a project that includes the WcfMsgServiceLib project: -----------------------------------------
    'Project \ Add Service Reference
    'Press the Discover button.
    'Expand the items in the Services window and select IMsgService.
    'Press OK.
    '------------------------------------------------------------------------------------------------------------------------------
    '------------------------------------------------------------------------------------------------------------------------------
    'Adding the service reference to other projects that dont include the WcfMsgServiceLib project: -------------------------------
    'Run the ADVL_Application_Network_1 application to start the Application Network message service.
    'In Microsoft Visual Studio select: Project \ Add Service Reference
    'Enter the address: http://localhost:8733/ADVLService
    'Press the Go button.
    'MsgService is found.
    'Press OK to add ServiceReference1 to the project.
    '------------------------------------------------------------------------------------------------------------------------------
    '
    'ADD THE MsgServiceCallback CODE: =============================================================================================
    'This is used to connect to the Application Network.
    'In Microsoft Visual Studio select: Project \ Add Class
    'MsgServiceCallback.vb
    'Add the following code to the class:
    'Imports System.ServiceModel
    'Public Class MsgServiceCallback
    '    Implements ServiceReference1.IMsgServiceCallback
    '    Public Sub OnSendMessage(message As String) Implements ServiceReference1.IMsgServiceCallback.OnSendMessage
    '        'A message has been received.
    '        'Set the InstrReceived property value to the message (usually in XMessage format). This will also apply the instructions in the XMessage.
    '        Main.InstrReceived = message
    '    End Sub
    'End Class
    '------------------------------------------------------------------------------------------------------------------------------

#End Region 'Coding Notes ---------------------------------------------------------------------------------------------------------------------------------------------------------------------

#Region " Variable Declarations - All the variables used in this form and this application." '-------------------------------------------------------------------------------------------------

    'Declare Utility objects used to store information on the application, project, usage and display messages.
    Public WithEvents ApplicationInfo As New ADVL_Utilities_Library_1.ApplicationInfo 'This object is used to store application information.
    Public WithEvents Project As New ADVL_Utilities_Library_1.Project 'This object is used to store Project information.
    Public WithEvents Message As New ADVL_Utilities_Library_1.Message 'This object is used to display messages in the Messages window.
    Public WithEvents ApplicationUsage As New ADVL_Utilities_Library_1.Usage 'This object stores application usage information.

    'Declare objects used to connect to the Application Network:
    Public client As ServiceReference1.MsgServiceClient
    Public WithEvents XMsg As New ADVL_Utilities_Library_1.XMessage
    Dim XDoc As New System.Xml.XmlDocument
    Public Status As New System.Collections.Specialized.StringCollection
    Dim ClientAppName As String = "" 'The name of the client requesting service
    Dim ClientAppLocn As String = "" 'The location in the Client application requesting service
    Dim MessageXDoc As System.Xml.Linq.XDocument
    Dim xmessage As XElement 'This will contain the message. It will be added to MessageXDoc.
    Dim xlocns As New List(Of XElement) 'A list of locations. Each location forms part of the reply message. The information in the reply message will be sent to the specified location in the client application.
    Dim MessageText As String = "" 'The text of a message sent through the Application Network.

    'Declare Forms used by the application:
    Public WithEvents SharePrices As frmSharePrices
    Public SharePricesFormList As New ArrayList 'Used for displaying multiple SharePrices forms.
    Public SharePricesSettings As New DataViewSettingsList 'Stores a list of settings used to display data views on the Share Prices forms.
    Public SharePriceSettingsChanged As Boolean = False 'If True then the Share Prices Settings List file needs to be updated.

    Public WithEvents Financials As frmFinancials
    Public FinancialsFormList As New ArrayList 'Used for displaying multiple Financials forms.
    Public FinancialsSettings As New DataViewSettingsList 'Stores a list of settings used to display data views on the Financials forms.
    Public FinancialsSettingsChanged As Boolean = False 'If True then the Financials Settings List file needs to be updated.

    Public WithEvents Calculations As frmCalculations
    Public CalculationsFormList As New ArrayList 'Used for displaying multiple Calculations forms.
    Public CalculationsSettings As New DataViewSettingsList 'Stores a list of settings used to display data views on the Calculations forms.
    Public CalculationsSettingsChanged As Boolean = False 'If True then the Calculations Settings List file needs to be updated.

    'Public WithEvents Sequence As frmImportSequence
    Public WithEvents Sequence As frmSequence

    'NOTE: THE ViewTable FORM WAS DESIGNED TO REPLACE ALL OTHER DATA VIEW FORMS SUCH AS SharePrices and Financials.
    '      IT WAS DECIDED TO RETAIN THE SEPARATE FORMS.
    'Public WithEvents ViewTable As frmViewTable
    'Public ViewTableList As New ArrayList 'Used for displaying multiple ViewTable forms.

    Dim WithEvents Zip As ADVL_Utilities_Library_1.ZipComp

    Dim dsInput As DataSet = New DataSet 'The input dataset for calculations.
    Dim dsOutput As DataSet = New DataSet 'The output dataset for calculations.
    Dim outputQuery As String
    Dim outputConnString As String
    Dim outputConnection As OleDb.OleDbConnection = New OleDb.OleDbConnection
    Dim outputDa As OleDb.OleDbDataAdapter
    'Dim outputCmdBuilder As New OleDb.OleDbCommandBuilder(outputDa)
    'Dim outputCmdBuilder As OleDb.OleDbCommandBuilder()

    Public WithEvents XSeq As New ADVL_Utilities_Library_1.XSequence 'This is used to run a set of XML Sequence statements. These are used for data processing.


#End Region 'Variable Declarations ------------------------------------------------------------------------------------------------------------------------------------------------------------

#Region " Properties - All the properties used in this form and this application" '------------------------------------------------------------------------------------------------------------

    Private _connectionHashcode As Integer 'The Application Network connection hashcode. This is used to identify a connection in the Application Netowrk when reconnecting.
    Property ConnectionHashcode As Integer
        Get
            Return _connectionHashcode
        End Get
        Set(value As Integer)
            _connectionHashcode = value
        End Set
    End Property

    Private _connectedToAppNet As Boolean = False  'True if the application is connected to the Application Network.
    Property ConnectedToAppnet As Boolean
        Get
            Return _connectedToAppNet
        End Get
        Set(value As Boolean)
            _connectedToAppNet = value
        End Set
    End Property

    Private _instrReceived As String = "" 'Contains Instructions received from the Application Network message service.
    Property InstrReceived As String
        Get
            Return _instrReceived
        End Get
        Set(value As String)
            If value = Nothing Then
                Message.Add("Empty message received!")
            Else
                _instrReceived = value

                'Add the message to the XMessages window:
                Message.Color = Color.Blue
                Message.FontStyle = FontStyle.Bold
                Message.XAdd("Message received: " & vbCrLf)
                Message.SetNormalStyle()
                Message.XAdd(_instrReceived & vbCrLf & vbCrLf)

                If _instrReceived.StartsWith("<XMsg>") Then 'This is an XMessage set of instructions.
                    Try
                        'Inititalise the reply message:
                        Dim Decl As New XDeclaration("1.0", "utf-8", "yes")
                        MessageXDoc = New XDocument(Decl, Nothing) 'Reply message - this will be sent to the Client App.
                        xmessage = New XElement("XMsg")
                        xlocns.Add(New XElement("Main")) 'Initially set the location in the Client App to Main.

                        'Run the received message:
                        Dim XmlHeader As String = "<?xml version=""1.0"" encoding=""utf-8"" standalone=""yes""?>"
                        XDoc.LoadXml(XmlHeader & vbCrLf & _instrReceived)
                        XMsg.Run(XDoc, Status)
                    Catch ex As Exception
                        Message.Add("Error running XMsg: " & ex.Message & vbCrLf)
                    End Try

                    'XMessage has been run.
                    'Reply to this message:
                    'Add the message reply to the XMessages window:
                    'Complete the MessageXDoc:
                    xmessage.Add(xlocns(xlocns.Count - 1)) 'Add the last location reply instructions to the message.
                    MessageXDoc.Add(xmessage)
                    MessageText = MessageXDoc.ToString

                    If ClientAppName = "" Then
                        'No client to send a message to!
                    Else
                        Message.Color = Color.Red
                        Message.FontStyle = FontStyle.Bold
                        Message.XAdd("Message sent to " & ClientAppName & ":" & vbCrLf)
                        Message.SetNormalStyle()
                        Message.XAdd(MessageText & vbCrLf & vbCrLf)
                        'SendMessage sends the contents of MessageText to MessageDest.
                        SendMessage() 'This subroutine triggers the timer to send the message after a short delay.
                    End If
                Else

                End If
            End If

        End Set
    End Property

    Private _sharePriceDbPath As String = "" 'The path of the share price database.
    Property SharePriceDbPath As String
        Get
            Return _sharePriceDbPath
        End Get
        Set(value As String)
            _sharePriceDbPath = value
            txtSharePriceDatabase.Text = _sharePriceDbPath
        End Set
    End Property

    Private _sharePriceDataViewList As String = "" 'The name of the list of views of data in the Share Price database. These files have the extension .SPViewList.
    Property SharePriceDataViewList As String
        Get
            Return _sharePriceDataViewList
        End Get
        Set(value As String)
            _sharePriceDataViewList = value
        End Set
    End Property

    Private _financialsDbPath As String = "" 'The path of the historical financials database.
    Property FinancialsDbPath As String
        Get
            Return _financialsDbPath
        End Get
        Set(value As String)
            _financialsDbPath = value
            txtFinancialsDatabase.Text = _financialsDbPath
        End Set
    End Property

    Private _financialsDataViewList As String = "" 'Name of the list of views of data in the Financials database. These files have the extension .FinViewList.
    Property FinancialsDataViewList As String
        Get
            Return _financialsDataViewList
        End Get
        Set(value As String)
            _financialsDataViewList = value
        End Set
    End Property

    Private _calculationsDbPath As String = "" 'The path to the calculations database.
    Property CalculationsDbPath As String
        Get
            Return _calculationsDbPath
        End Get
        Set(value As String)
            _calculationsDbPath = value
            txtCalcsDatabase.Text = _calculationsDbPath
        End Set
    End Property

    Private _calculationsDataViewList As String = "" 'The name of the list of views of data in the Calculations database. These files have the extension .CalcViewList.
    Property CalculationsDataViewList As String
        Get
            Return _calculationsDataViewList
        End Get
        Set(value As String)
            _calculationsDataViewList = value
        End Set
    End Property

    Private _newsDbPath As String = "" 'The path of the News database.
    Property NewsDbPath As String
        Get
            Return _newsDbPath
        End Get
        Set(value As String)
            _newsDbPath = value
        End Set
    End Property

    Private _newsDataViewList As String = "" 'The name of the list of views of data in the News database. These files have the extension .NewsViewList.
    Property NewsDataViewList As String
        Get
            Return _newsDataViewList
        End Get
        Set(value As String)
            _newsDataViewList = value
        End Set
    End Property

    Private _newsDirectory As String = "" 'The path of the News directory. Structure: News\'A-Z'\'Year' eg: News\A\2015\ News\C\2010\
    Property NewsDirectory As String
        Get
            Return _newsDirectory
        End Get
        Set(value As String)
            _newsDirectory = value
        End Set
    End Property

    Private _otherDbPath As String = "" 'The path of another database selected on the Design Query form.
    Property OtherDbPath As String
        Get
            Return _otherDbPath
        End Get
        Set(value As String)
            _otherDbPath = value
        End Set
    End Property

    Private _copyDataSettingsFile As String = "" 'The name of the file used to store the current Copy Data settings.
    Property CopyDataSettingsFile As String
        Get
            Return _copyDataSettingsFile
        End Get
        Set(value As String)
            _copyDataSettingsFile = value
        End Set
    End Property

    Private _selectDataSettingsFile As String = "" 'The name of the file used to store the current Select Data settings.
    Property SelectDataSettingsFile As String
        Get
            Return _selectDataSettingsFile
        End Get
        Set(value As String)
            _selectDataSettingsFile = value
        End Set
    End Property

    Private _simpleCalcsSettingsFile As String = "" 'The name of the file used to store the current Simple Calculations settings.
    Property SimpleCalcsSettingsFile As String
        Get
            Return _simpleCalcsSettingsFile
        End Get
        Set(value As String)
            _simpleCalcsSettingsFile = value
        End Set
    End Property

    Private _dateCalcSettingsFile As String = "" 'The name of the file used to store the current Date Calculations settings.
    Property DateCalcSettingsFile As String
        Get
            Return _dateCalcSettingsFile
        End Get
        Set(value As String)
            _dateCalcSettingsFile = value
        End Set
    End Property

    Private _dateSelectSettingsFile As String = "" 'The name of the file used to store the current Date Selections settings.
    Property DateSelectSettingsFile As String
        Get
            Return _dateSelectSettingsFile
        End Get
        Set(value As String)
            _dateSelectSettingsFile = value
        End Set
    End Property

    Private _closedFormNo As Integer 'Temporarily holds the number of the form that is being closed. 
    Property ClosedFormNo As Integer
        Get
            Return _closedFormNo
        End Get
        Set(value As Integer)
            _closedFormNo = value
        End Set
    End Property

    Private _recordSequence As Boolean 'If True then processing sequences manually applied using the forms will be recorded in the processing sequence.
    Property RecordSequence As Boolean
        Get
            Return _recordSequence
        End Get
        Set(value As Boolean)
            _recordSequence = value
        End Set
    End Property

#End Region 'Properties -----------------------------------------------------------------------------------------------------------------------------------------------------------------------

#Region " Process XML Files - Read and write XML files." '-------------------------------------------------------------------------------------------------------------------------------------

    Private Sub SaveFormSettings()
        'Save the form settings in an XML document.
        'Try
        Dim settingsData = <?xml version="1.0" encoding="utf-8"?>
                           <!---->
                           <!--Form settings for Main form.-->
                           <FormSettings>
                               <Left><%= Me.Left %></Left>
                               <Top><%= Me.Top %></Top>
                               <Width><%= Me.Width %></Width>
                               <Height><%= Me.Height %></Height>
                               <!---->
                               <SelectedMainTabIndex><%= TabControl1.SelectedIndex %></SelectedMainTabIndex>
                               <SelectedViewDataTabIndex><%= TabControl3.SelectedIndex %></SelectedViewDataTabIndex>
                               <SharePriceDbPath><%= SharePriceDbPath %></SharePriceDbPath>
                               <SharePriceDataViewList><%= SharePriceDataViewList %></SharePriceDataViewList>
                               <FinancialsDbPath><%= FinancialsDbPath %></FinancialsDbPath>
                               <FinancialsDataViewList><%= FinancialsDataViewList %></FinancialsDataViewList>
                               <CalculationsDbPath><%= CalculationsDbPath %></CalculationsDbPath>
                               <CalculationsDataViewList><%= CalculationsDataViewList %></CalculationsDataViewList>
                               <NewsDbPath><%= NewsDbPath %></NewsDbPath>
                               <NewsDataViewList><%= NewsDataViewList %></NewsDataViewList>
                               <OtherDbPath><%= OtherDbPath %></OtherDbPath>
                               <!--Calculations Settings-->
                               <SelectedCalculationsTabIndex><%= TabControl2.SelectedIndex %></SelectedCalculationsTabIndex>
                               <!--Copy Data-->
                               <CopyDataInputDb><%= cmbCopyDataInputDb.SelectedItem.ToString %></CopyDataInputDb>
                               <CopyDataInputData><%= cmbCopyDataInputData.SelectedItem.ToString %></CopyDataInputData>
                               <CopyDataOutputDb><%= cmbCopyDataOutputDb.SelectedItem.ToString %></CopyDataOutputDb>
                               <CopyDataOutputTable><%= cmbCopyDataOutputData.SelectedItem.ToString %></CopyDataOutputTable>
                               <CopyDataSettingsFile><%= CopyDataSettingsFile %></CopyDataSettingsFile>
                               <!--Select Data-->
                               <SelectDataInputDb><%= cmbSelectDataInputDb.SelectedItem.ToString %></SelectDataInputDb>
                               <SelectDataInputData><%= cmbSelectDataInputData.SelectedItem.ToString %></SelectDataInputData>
                               <SelectDataOutputDb><%= cmbSelectDataOutputDb.SelectedItem.ToString %></SelectDataOutputDb>
                               <SelectDataOutputTable><%= cmbSelectDataOutputData.SelectedItem.ToString %></SelectDataOutputTable>
                               <SelectDataSettingsFile><%= SelectDataSettingsFile %></SelectDataSettingsFile>
                               <!--Simple Calculations-->
                               <SimpleCalcsSettingsFile><%= SimpleCalcsSettingsFile %></SimpleCalcsSettingsFile>
                               <SimpleCalcsVerticalSplitterDist><%= SplitContainer2.SplitterDistance %></SimpleCalcsVerticalSplitterDist>
                               <SimpleCalcsLhsHorSplitterDist><%= SplitContainer3.SplitterDistance %></SimpleCalcsLhsHorSplitterDist>
                               <SimpleCalcsRhsHorSplitterDist><%= SplitContainer4.SplitterDistance %></SimpleCalcsRhsHorSplitterDist>
                               <!--Date Calculations-->
                               <DateCalcsSettingsFile><%= DateCalcSettingsFile %></DateCalcsSettingsFile>
                               <DateCalcsFixedDate><%= txtFixedDate.Text %></DateCalcsFixedDate>
                               <DateCalcsDateFormatString><%= txtDateFormatString.Text %></DateCalcsDateFormatString>
                               <!--Date Selections-->
                               <DateSelectionSettingsFile><%= DateSelectSettingsFile %></DateSelectionSettingsFile>
                               <!--Daily Prices Calculations-->
                               <DailyPricesInputDb><%= cmbDailyPriceInputDb.SelectedItem.ToString %></DailyPricesInputDb>
                               <DailyPricesInputTable><%= cmbDailyPriceInputTable.SelectedItem.ToString %></DailyPricesInputTable>
                               <DailyPricesOutoutDb><%= cmbDailyPriceOutputDb.SelectedItem.ToString %></DailyPricesOutoutDb>
                               <DailyPricesOutputTable><%= cmbDailyPriceOutputTable.SelectedItem.ToString %></DailyPricesOutputTable>
                               <DailyPricesCalculationType><%= cmbDailyPriceCalcType.SelectedItem.ToString %></DailyPricesCalculationType>
                           </FormSettings>

        'NOTE: THIS DATA IS NOW STORED IN SharePricesSettingsList
        '                  <ViewSharePricesItems><%= From item In lstSharePrices.Items
        '                                  Select
        '                          <Description><%= item %></Description>
        '                              %>
        '                  </ViewSharePricesItems>
        'NOTE: THIS DATA IS NOW STROED IN FinancialsSettingsList
        '                  <ViewFinancialsItems><%= From item In lstFinancials.Items
        '                         Select
        '                         <Description><%= item %></Description>
        '                             %>
        '                  </ViewFinancialsItems>

        'NOTE: Separate tabs now used for different calculation types.
        '<SimpleCalcType><%= cmbCalcType.SelectedItem.ToString %></SimpleCalcType>

        'Add code to include other settings to save after the comment line <!---->



        Dim SettingsFileName As String = "FormSettings_" & ApplicationInfo.Name & "_" & Me.Text & ".xml"
            Debug.Print("Writing settings file: " & SettingsFileName)
            Project.SaveXmlSettings(SettingsFileName, settingsData)
        'Catch ex As Exception
        'Message.AddWarning("Error saving main form settings: " & ex.Message & vbCrLf)
        'Debug.Print("Error saving main form settings: " & ex.Message)
        'End Try
    End Sub

    Private Sub RestoreFormSettings()
        'Read the form settings from an XML document.

        Dim SettingsFileName As String = "FormSettings_" & ApplicationInfo.Name & "_" & Me.Text & ".xml"
        Debug.Print("Reading settings file: " & SettingsFileName)

        If Project.SettingsFileExists(SettingsFileName) Then
            Debug.Print("Settings file found.")
            Dim Settings As System.Xml.Linq.XDocument
            Project.ReadXmlSettings(SettingsFileName, Settings)

            If IsNothing(Settings) Then 'There is no Settings XML data.
                Debug.Print("Settings file is blank.")
                Exit Sub
            End If

            'Restore form position and size:
            If Settings.<FormSettings>.<Left>.Value <> Nothing Then Me.Left = Settings.<FormSettings>.<Left>.Value
            If Settings.<FormSettings>.<Top>.Value <> Nothing Then Me.Top = Settings.<FormSettings>.<Top>.Value
            If Settings.<FormSettings>.<Height>.Value <> Nothing Then Me.Height = Settings.<FormSettings>.<Height>.Value
            If Settings.<FormSettings>.<Width>.Value <> Nothing Then Me.Width = Settings.<FormSettings>.<Width>.Value

            'Add code to read other saved setting here:

            'Restore Main Tab selection (Project Information - Settings - View Data - Calculations - Charts)
            If Settings.<FormSettings>.<SelectedMainTabIndex>.Value <> Nothing Then TabControl1.SelectedIndex = Settings.<FormSettings>.<SelectedMainTabIndex>.Value

            'Restore View Data Sub Tab selection (Share Prices - Financials - Calculations - News)
            If Settings.<FormSettings>.<SelectedViewDataTabIndex>.Value <> Nothing Then TabControl3.SelectedIndex = Settings.<FormSettings>.<SelectedViewDataTabIndex>.Value

            'Restore View Data - Share Prices settings
            If Settings.<FormSettings>.<SharePriceDbPath>.Value <> Nothing Then SharePriceDbPath = Settings.<FormSettings>.<SharePriceDbPath>.Value
            If Settings.<FormSettings>.<SharePriceDataViewList>.Value <> Nothing Then
                SharePriceDataViewList = Settings.<FormSettings>.<SharePriceDataViewList>.Value
                SharePricesSettings.ListFileName = SharePriceDataViewList 'Set the file name in SharePricesSettings List.
                txtSharePricesDataList.Text = SharePriceDataViewList
                SharePricesSettings.LoadFile() 'Load the settings list in SharePricesSettings.
                DisplaySharePricesList()
            End If

            'Restore View Data - Financials settings
            If Settings.<FormSettings>.<FinancialsDbPath>.Value <> Nothing Then FinancialsDbPath = Settings.<FormSettings>.<FinancialsDbPath>.Value
            If Settings.<FormSettings>.<FinancialsDataViewList>.Value <> Nothing Then
                FinancialsDataViewList = Settings.<FormSettings>.<FinancialsDataViewList>.Value
                FinancialsSettings.ListFileName = FinancialsDataViewList 'Set the file name in FinancialsSettingsList.
                txtFinancialsDataList.Text = FinancialsDataViewList
                FinancialsSettings.LoadFile() 'Load the settings list in FinancialsSettingsList.
                DisplayFinancialsList()
            End If

            'Restore View Data - Calculations settings
            If Settings.<FormSettings>.<CalculationsDbPath>.Value <> Nothing Then CalculationsDbPath = Settings.<FormSettings>.<CalculationsDbPath>.Value
            If Settings.<FormSettings>.<CalculationsDataViewList>.Value <> Nothing Then
                CalculationsDataViewList = Settings.<FormSettings>.<CalculationsDataViewList>.Value
                CalculationsSettings.ListFileName = CalculationsDataViewList 'Set the file name in CalculationsSettingsList.
                txtCalcsDataList.Text = CalculationsDataViewList
                CalculationsSettings.LoadFile() 'Load the settings list in CalculationsSettingsList.
                DisplayCalculationsList()
            End If

            If Settings.<FormSettings>.<NewsDbPath>.Value <> Nothing Then NewsDbPath = Settings.<FormSettings>.<NewsDbPath>.Value
            If Settings.<FormSettings>.<OtherDbPath>.Value <> Nothing Then OtherDbPath = Settings.<FormSettings>.<OtherDbPath>.Value

            'Calculations Settings:
            'Restore Calculations Sub Tab selection (Copy Data - Select Data - Simple Calculations - Curve Fitting - Filters)
            If Settings.<FormSettings>.<SelectedCalculationsTabIndex>.Value <> Nothing Then TabControl2.SelectedIndex = Settings.<FormSettings>.<SelectedCalculationsTabIndex>.Value

            'Restore saved calculations settings file names:
            If Settings.<FormSettings>.<CopyDataSettingsFile>.Value <> Nothing Then CopyDataSettingsFile = Settings.<FormSettings>.<CopyDataSettingsFile>.Value
            If Settings.<FormSettings>.<SelectDataSettingsFile>.Value <> Nothing Then SelectDataSettingsFile = Settings.<FormSettings>.<SelectDataSettingsFile>.Value
            If Settings.<FormSettings>.<SimpleCalcsSettingsFile>.Value <> Nothing Then SimpleCalcsSettingsFile = Settings.<FormSettings>.<SimpleCalcsSettingsFile>.Value
            If Settings.<FormSettings>.<DateCalcsSettingsFile>.Value <> Nothing Then DateCalcSettingsFile = Settings.<FormSettings>.<DateCalcsSettingsFile>.Value
            If Settings.<FormSettings>.<DateSelectionSettingsFile>.Value <> Nothing Then DateSelectSettingsFile = Settings.<FormSettings>.<DateSelectionSettingsFile>.Value

            'Restore Copy Data settings: (The CopyDataSettingsFile may update these)
            If Settings.<FormSettings>.<CopyDataInputDb>.Value <> Nothing Then cmbCopyDataInputDb.SelectedIndex = cmbCopyDataInputDb.FindStringExact(Settings.<FormSettings>.<CopyDataInputDb>.Value)
            If Settings.<FormSettings>.<CopyDataInputData>.Value <> Nothing Then cmbCopyDataInputData.SelectedIndex = cmbCopyDataInputData.FindStringExact(Settings.<FormSettings>.<CopyDataInputData>.Value)
            If Settings.<FormSettings>.<CopyDataOutputDb>.Value <> Nothing Then cmbCopyDataOutputDb.SelectedIndex = cmbCopyDataOutputDb.FindStringExact(Settings.<FormSettings>.<CopyDataOutputDb>.Value)
            If Settings.<FormSettings>.<CopyDataOutputTable>.Value <> Nothing Then cmbCopyDataOutputData.SelectedIndex = cmbCopyDataOutputData.FindStringExact(Settings.<FormSettings>.<CopyDataOutputTable>.Value)
            SetUpCopyDataTab()
            LoadCopyDataSettingsFile()

            'Restore Select Data settings: (The SelectDataSettingsFile may update these)
            If Settings.<FormSettings>.<SelectDataInputDb>.Value <> Nothing Then cmbSelectDataInputDb.SelectedIndex = cmbSelectDataInputDb.FindStringExact(Settings.<FormSettings>.<SelectDataInputDb>.Value)
            If Settings.<FormSettings>.<SelectDataInputData>.Value <> Nothing Then cmbSelectDataInputData.SelectedIndex = cmbSelectDataInputData.FindStringExact(Settings.<FormSettings>.<SelectDataInputData>.Value)
            If Settings.<FormSettings>.<SelectDataOutputDb>.Value <> Nothing Then cmbSelectDataOutputDb.SelectedIndex = cmbSelectDataOutputDb.FindStringExact(Settings.<FormSettings>.<SelectDataOutputDb>.Value)
            If Settings.<FormSettings>.<SelectDataOutputTable>.Value <> Nothing Then cmbSelectDataOutputData.SelectedIndex = cmbSelectDataOutputData.FindStringExact(Settings.<FormSettings>.<SelectDataOutputTable>.Value)
            SetUpSelectDataTab()
            LoadSelectDataSettingsFile()

            'Restore Simple Calculations settings: (The SimpleCalcsSettingsFile may update these)
            If Settings.<FormSettings>.<SimpleCalcsVerticalSplitterDist>.Value <> Nothing Then SplitContainer2.SplitterDistance = Settings.<FormSettings>.<SimpleCalcsVerticalSplitterDist>.Value
            If Settings.<FormSettings>.<SimpleCalcsLhsHorSplitterDist>.Value <> Nothing Then SplitContainer3.SplitterDistance = Settings.<FormSettings>.<SimpleCalcsLhsHorSplitterDist>.Value
            If Settings.<FormSettings>.<SimpleCalcsRhsHorSplitterDist>.Value <> Nothing Then SplitContainer4.SplitterDistance = Settings.<FormSettings>.<SimpleCalcsRhsHorSplitterDist>.Value
            SetUpSimpleCalculationsTab()
            LoadSimpleCalcsSettingsFile()

            'Restore Date Calculations Settings:
            If Settings.<FormSettings>.<DateCalcsFixedDate>.Value <> Nothing Then txtFixedDate.Text = Settings.<FormSettings>.<DateCalcsFixedDate>.Value
            If Settings.<FormSettings>.<DateCalcsDateFormatString>.Value <> Nothing Then txtDateFormatString.Text = Settings.<FormSettings>.<DateCalcsDateFormatString>.Value
            LoadDateCalcsSettingsFile()

            'Restore Date Selection Settings:
            LoadDateSelectSettingsFile()

            'Restore Daily Prices Tab Settings:
            SetUpDailyPricesTab()

            'Restore Daily Prices Calculations Settings:
            If Settings.<FormSettings>.<DailyPricesInputDb>.Value <> Nothing Then cmbDailyPriceInputDb.SelectedIndex = cmbDailyPriceInputDb.FindStringExact(Settings.<FormSettings>.<DailyPricesInputDb>.Value)
            If Settings.<FormSettings>.<DailyPricesInputTable>.Value <> Nothing Then cmbDailyPriceInputTable.SelectedIndex = cmbDailyPriceInputTable.FindStringExact(Settings.<FormSettings>.<DailyPricesInputTable>.Value)
            If Settings.<FormSettings>.<DailyPricesOutoutDb>.Value <> Nothing Then cmbDailyPriceOutputDb.SelectedIndex = cmbDailyPriceOutputDb.FindStringExact(Settings.<FormSettings>.<DailyPricesOutoutDb>.Value)
            If Settings.<FormSettings>.<DailyPricesOutputTable>.Value <> Nothing Then cmbDailyPriceOutputTable.SelectedIndex = cmbDailyPriceOutputTable.FindStringExact(Settings.<FormSettings>.<DailyPricesOutputTable>.Value)
            If Settings.<FormSettings>.<DailyPricesCalculationType>.Value <> Nothing Then
                cmbDailyPriceCalcType.SelectedIndex = cmbDailyPriceCalcType.FindStringExact(Settings.<FormSettings>.<DailyPricesCalculationType>.Value)
                If cmbDailyPriceCalcType.SelectedIndex = -1 Then
                    cmbDailyPriceCalcType.SelectedIndex = 0
                End If
            Else
                    cmbDailyPriceCalcType.SelectedIndex = 0
            End If



            'NOTE: THIS DATA IS NOW STORED IN SharePricesSettingsList
            ''Read the list of ViewSharePricesItems:
            'Dim ViewSharePricesItems = From item In Settings.<FormSettings>.<ViewSharePricesItems>.<Description>
            'lstSharePrices.Items.Clear()
            'For Each item In ViewSharePricesItems
            '    lstSharePrices.Items.Add(item.Value)
            'Next

            'NOTE: THIS DATA IS NOW STROED IN FinancialsSettingsList
            ''Read the list of ViewFinancialsItems:
            'Dim ViewFinancialsItems = From item In Settings.<FormSettings>.<ViewFinancialsItems>.<Description>

            'lstFinancials.Items.Clear()
            'For Each item In ViewFinancialsItems
            '    lstFinancials.Items.Add(item.Value)
            'Next
        Else
            Debug.Print("Settings file not found.")

        End If
    End Sub

    Private Sub ReadApplicationInfo()
        'Read the Application Information.

        If ApplicationInfo.FileExists Then
            ApplicationInfo.ReadFile()
        Else
            'There is no Application_Info.xml file.
            DefaultAppProperties() 'Create a new Application Info file with default application properties:
        End If
    End Sub

    Private Sub DefaultAppProperties()
        'These properties will be saved in the Application_Info.xml file in the application directory.
        'If this file is deleted, it will be re-created using these default application properties.

        'Change this to show your application Name, Description and Creation Date.
        ApplicationInfo.Name = "ADVL_Shares_1"

        'ApplicationInfo.ApplicationDir is set when the application is started.
        ApplicationInfo.ExecutablePath = Application.ExecutablePath

        ApplicationInfo.Description = "The ADVL_Shares_1 application stores historical data for publicly traded shares and performs processing and analysis techniques aimed at optimizing share trading returns."
        ApplicationInfo.CreationDate = "29-Nov-2016 12:00:00"

        'Author -----------------------------------------------------------------------------------------------------------
        'Change this to show your Name, Description and Contact information.
        ApplicationInfo.Author.Name = "Signalworks Pty Ltd"
        ApplicationInfo.Author.Description = "Signalworks Pty Ltd" & vbCrLf &
            "Australian Proprietary Company" & vbCrLf &
            "ABN 26 066 681 598" & vbCrLf &
            "Registration Date 05/10/1994"

        ApplicationInfo.Author.Contact = "http://www.andorville.com.au/"

        'File Associations: -----------------------------------------------------------------------------------------------
        'Add any file associations here.
        'The file extension and a description of files that can be opened by this application are specified.
        'The example below specifies a coordinate system parameter file type with the file extension .ADVLCoord.
        'Dim Assn1 As New ADVL_System_Utilities.FileAssociation
        'Assn1.Extension = "ADVLCoord"
        'Assn1.Description = "Andorville (TM) software coordinate system parameter file"
        'ApplicationInfo.FileAssociations.Add(Assn1)

        'Version ----------------------------------------------------------------------------------------------------------
        ApplicationInfo.Version.Major = My.Application.Info.Version.Major
        ApplicationInfo.Version.Minor = My.Application.Info.Version.Minor
        ApplicationInfo.Version.Build = My.Application.Info.Version.Build
        ApplicationInfo.Version.Revision = My.Application.Info.Version.Revision

        'Copyright --------------------------------------------------------------------------------------------------------
        'Add your copyright information here.
        ApplicationInfo.Copyright.OwnerName = "Signalworks Pty Ltd, ABN 26 066 681 598"
        ApplicationInfo.Copyright.PublicationYear = "2016"

        'Trademarks -------------------------------------------------------------------------------------------------------
        'Add your trademark information here.
        Dim Trademark1 As New ADVL_Utilities_Library_1.Trademark
        Trademark1.OwnerName = "Signalworks Pty Ltd, ABN 26 066 681 598"
        Trademark1.Text = "Andorville"
        Trademark1.Registered = False
        Trademark1.GenericTerm = "software"
        ApplicationInfo.Trademarks.Add(Trademark1)
        Dim Trademark2 As New ADVL_Utilities_Library_1.Trademark
        Trademark2.OwnerName = "Signalworks Pty Ltd, ABN 26 066 681 598"
        Trademark2.Text = "AL-H7"
        Trademark2.Registered = False
        Trademark2.GenericTerm = "software"
        ApplicationInfo.Trademarks.Add(Trademark2)

        'License -------------------------------------------------------------------------------------------------------
        'Add your license information here.
        ApplicationInfo.License.CopyrightOwnerName = "Signalworks Pty Ltd, ABN 26 066 681 598"
        ApplicationInfo.License.PublicationYear = "2016"

        'License Links:
        'http://choosealicense.com/
        'http://www.apache.org/licenses/
        'http://opensource.org/

        'Apache License 2.0 ---------------------------------------------
        ApplicationInfo.License.Code = ADVL_Utilities_Library_1.License.Codes.Apache_License_2_0
        ApplicationInfo.License.Notice = ApplicationInfo.License.ApacheLicenseNotice 'Get the pre-defined Aapche license notice.
        ApplicationInfo.License.Text = ApplicationInfo.License.ApacheLicenseText     'Get the pre-defined Apache license text.

        'Code to use other pre-defined license types is shown below:

        'GNU General Public License, version 3 --------------------------
        'ApplicationInfo.License.Type = ADVL_Utilities_Library_1.License.Types.GNU_GPL_V3_0
        'ApplicationInfo.License.Notice = 'Add the License Notice to ADVL_Utilities_Library_1 License class.
        'ApplicationInfo.License.Text = 'Add the License Text to ADVL_Utilities_Library_1 License class.

        'The MIT License ------------------------------------------------
        'ApplicationInfo.License.Type = ADVL_Utilities_Library_1.License.Types.MIT_License
        'ApplicationInfo.License.Notice = ApplicationInfo.License.MITLicenseNotice
        'ApplicationInfo.License.Text = ApplicationInfo.License.MITLicenseText

        'No License Specified -------------------------------------------
        'ApplicationInfo.License.Type = ADVL_Utilities_Library_1.License.Types.None
        'ApplicationInfo.License.Notice = ""
        'ApplicationInfo.License.Text = ""

        'The Unlicense --------------------------------------------------
        'ApplicationInfo.License.Type = ADVL_Utilities_Library_1.License.Types.The_Unlicense
        'ApplicationInfo.License.Notice = ApplicationInfo.License.UnLicenseNotice
        'ApplicationInfo.License.Text = ApplicationInfo.License.UnLicenseText

        'Unknown License ------------------------------------------------
        'ApplicationInfo.License.Type = ADVL_Utilities_Library_1.License.Types.Unknown
        'ApplicationInfo.License.Notice = ""
        'ApplicationInfo.License.Text = ""

        'Source Code: --------------------------------------------------------------------------------------------------
        'Add your source code information here if required.
        'THIS SECTION WILL BE UPDATED TO ALLOW A GITHUB LINK.
        ApplicationInfo.SourceCode.Language = "Visual Basic 2015"
        ApplicationInfo.SourceCode.FileName = ""
        ApplicationInfo.SourceCode.FileSize = 0
        ApplicationInfo.SourceCode.FileHash = ""
        ApplicationInfo.SourceCode.WebLink = ""
        ApplicationInfo.SourceCode.Contact = ""
        ApplicationInfo.SourceCode.Comments = ""

        'ModificationSummary: -----------------------------------------------------------------------------------------
        'Add any source code modification here is required.
        ApplicationInfo.ModificationSummary.BaseCodeName = ""
        ApplicationInfo.ModificationSummary.BaseCodeDescription = ""
        ApplicationInfo.ModificationSummary.BaseCodeVersion.Major = 0
        ApplicationInfo.ModificationSummary.BaseCodeVersion.Minor = 0
        ApplicationInfo.ModificationSummary.BaseCodeVersion.Build = 0
        ApplicationInfo.ModificationSummary.BaseCodeVersion.Revision = 0
        ApplicationInfo.ModificationSummary.Description = "This is the first released version of the application. No earlier base code used."

        'Library List: ------------------------------------------------------------------------------------------------
        'Add the ADVL_Utilties_Library_1 library:
        Dim NewLib As New ADVL_Utilities_Library_1.LibrarySummary
        NewLib.Name = "ADVL_System_Utilities"
        NewLib.Description = "System Utility classes used in Andorville (TM) software development system applications"
        NewLib.CreationDate = "7-Jan-2016 12:00:00"
        NewLib.LicenseNotice = "Copyright 2016 Signalworks Pty Ltd, ABN 26 066 681 598" & vbCrLf &
                               vbCrLf &
                               "Licensed under the Apache License, Version 2.0 (the ""License"");" & vbCrLf &
                               "you may not use this file except in compliance with the License." & vbCrLf &
                               "You may obtain a copy of the License at" & vbCrLf &
                               vbCrLf &
                               "http://www.apache.org/licenses/LICENSE-2.0" & vbCrLf &
                               vbCrLf &
                               "Unless required by applicable law or agreed to in writing, software" & vbCrLf &
                               "distributed under the License is distributed on an ""AS IS"" BASIS," & vbCrLf &
                               "WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied." & vbCrLf &
                               "See the License for the specific language governing permissions and" & vbCrLf &
                               "limitations under the License." & vbCrLf

        NewLib.CopyrightNotice = "Copyright 2016 Signalworks Pty Ltd, ABN 26 066 681 598"

        NewLib.Version.Major = 1
        NewLib.Version.Minor = 0
        NewLib.Version.Build = 1
        NewLib.Version.Revision = 0

        NewLib.Author.Name = "Signalworks Pty Ltd"
        NewLib.Author.Description = "Signalworks Pty Ltd" & vbCrLf &
            "Australian Proprietary Company" & vbCrLf &
            "ABN 26 066 681 598" & vbCrLf &
            "Registration Date 05/10/1994"

        NewLib.Author.Contact = "http://www.andorville.com.au/"

        Dim NewClass1 As New ADVL_Utilities_Library_1.ClassSummary
        NewClass1.Name = "ZipComp"
        NewClass1.Description = "The ZipComp class is used to compress files into and extract files from a zip file."
        NewLib.Classes.Add(NewClass1)
        Dim NewClass2 As New ADVL_Utilities_Library_1.ClassSummary
        NewClass2.Name = "XSequence"
        NewClass2.Description = "The XSequence class is used to run an XML property sequence (XSequence) file. XSequence files are used to record and replay processing sequences in Andorville (TM) software applications."
        NewLib.Classes.Add(NewClass2)
        Dim NewClass3 As New ADVL_Utilities_Library_1.ClassSummary
        NewClass3.Name = "XMessage"
        NewClass3.Description = "The XMessage class is used to read an XML Message (XMessage). An XMessage is a simplified XSequence used to exchange information between Andorville (TM) software applications."
        NewLib.Classes.Add(NewClass3)
        Dim NewClass4 As New ADVL_Utilities_Library_1.ClassSummary
        NewClass4.Name = "Location"
        NewClass4.Description = "The Location class consists of properties and methods to store data in a location, which is either a directory or archive file."
        NewLib.Classes.Add(NewClass4)
        Dim NewClass5 As New ADVL_Utilities_Library_1.ClassSummary
        NewClass5.Name = "Project"
        NewClass5.Description = "An Andorville (TM) software application can store data within one or more projects. Each project stores a set of related data files. The Project class contains properties and methods used to manage a project."
        NewLib.Classes.Add(NewClass5)
        Dim NewClass6 As New ADVL_Utilities_Library_1.ClassSummary
        NewClass6.Name = "ProjectSummary"
        NewClass6.Description = "ProjectSummary stores a summary of a project."
        NewLib.Classes.Add(NewClass6)
        Dim NewClass7 As New ADVL_Utilities_Library_1.ClassSummary
        NewClass7.Name = "DataFileInfo"
        NewClass7.Description = "The DataFileInfo class stores information about a data file."
        NewLib.Classes.Add(NewClass7)
        Dim NewClass8 As New ADVL_Utilities_Library_1.ClassSummary
        NewClass8.Name = "Message"
        NewClass8.Description = "The Message class contains text properties and methods used to display messages in an Andorville (TM) software application."
        NewLib.Classes.Add(NewClass8)
        Dim NewClass9 As New ADVL_Utilities_Library_1.ClassSummary
        NewClass9.Name = "ApplicationSummary"
        NewClass9.Description = "The ApplicationSummary class stores a summary of an Andorville (TM) software application."
        NewLib.Classes.Add(NewClass9)
        Dim NewClass10 As New ADVL_Utilities_Library_1.ClassSummary
        NewClass10.Name = "LibrarySummary"
        NewClass10.Description = "The LibrarySummary class stores a summary of a software library used by an application."
        NewLib.Classes.Add(NewClass10)
        Dim NewClass11 As New ADVL_Utilities_Library_1.ClassSummary
        NewClass11.Name = "ClassSummary"
        NewClass11.Description = "The ClassSummary class stores a summary of a class contained in a software library."
        NewLib.Classes.Add(NewClass11)
        Dim NewClass12 As New ADVL_Utilities_Library_1.ClassSummary
        NewClass12.Name = "ModificationSummary"
        NewClass12.Description = "The ModificationSummary class stores a summary of any modifications made to an application or library."
        NewLib.Classes.Add(NewClass12)
        Dim NewClass13 As New ADVL_Utilities_Library_1.ClassSummary
        NewClass13.Name = "ApplicationInfo"
        NewClass13.Description = "The ApplicationInfo class stores information about an Andorville (TM) software application."
        NewLib.Classes.Add(NewClass13)
        Dim NewClass14 As New ADVL_Utilities_Library_1.ClassSummary
        NewClass14.Name = "Version"
        NewClass14.Description = "The Version class stores application, library or project version information."
        NewLib.Classes.Add(NewClass14)
        Dim NewClass15 As New ADVL_Utilities_Library_1.ClassSummary
        NewClass15.Name = "Author"
        NewClass15.Description = "The Author class stores information about an Author."
        NewLib.Classes.Add(NewClass15)
        Dim NewClass16 As New ADVL_Utilities_Library_1.ClassSummary
        NewClass16.Name = "FileAssociation"
        NewClass16.Description = "The FileAssociation class stores the file association extension and description. An application can open files on its file association list."
        NewLib.Classes.Add(NewClass16)
        Dim NewClass17 As New ADVL_Utilities_Library_1.ClassSummary
        NewClass17.Name = "Copyright"
        NewClass17.Description = "The Copyright class stores copyright information."
        NewLib.Classes.Add(NewClass17)
        Dim NewClass18 As New ADVL_Utilities_Library_1.ClassSummary
        NewClass18.Name = "License"
        NewClass18.Description = "The License class stores license information."
        NewLib.Classes.Add(NewClass18)
        Dim NewClass19 As New ADVL_Utilities_Library_1.ClassSummary
        NewClass19.Name = "SourceCode"
        NewClass19.Description = "The SourceCode class stores information about the source code for the application."
        NewLib.Classes.Add(NewClass19)
        Dim NewClass20 As New ADVL_Utilities_Library_1.ClassSummary
        NewClass20.Name = "Usage"
        NewClass20.Description = "The Usage class stores information about application or project usage."
        NewLib.Classes.Add(NewClass20)
        Dim NewClass21 As New ADVL_Utilities_Library_1.ClassSummary
        NewClass21.Name = "Trademark"
        NewClass21.Description = "The Trademark class stored information about a trademark used by the author of an application or data."
        NewLib.Classes.Add(NewClass21)

        ApplicationInfo.Libraries.Add(NewLib)

        'Add other library information here: --------------------------------------------------------------------------

    End Sub

    'Save the form settings if the form is being minimised:
    Protected Overrides Sub WndProc(ByRef m As Message)
        If m.Msg = &H112 Then 'SysCommand
            If m.WParam.ToInt32 = &HF020 Then 'Form is being minimised
                SaveFormSettings()
            End If
        End If
        MyBase.WndProc(m)
    End Sub

    Private Sub SaveProjectSettings()
        'Save the project settings in an XML file.
        'Add any Project Settings to be saved into the settingsData XDocument.
        Dim settingsData = <?xml version="1.0" encoding="utf-8"?>
                           <!---->
                           <!--Project settings for ADVL_Coordinates_1 application.-->
                           <ProjectSettings>
                           </ProjectSettings>

        Dim SettingsFileName As String = "ProjectSettings_" & ApplicationInfo.Name & "_" & Me.Text & ".xml"
        Project.SaveXmlSettings(SettingsFileName, settingsData)

    End Sub

    Private Sub RestoreProjectSettings()
        'Restore the project settings from an XML document.

        Dim SettingsFileName As String = "ProjectSettings_" & ApplicationInfo.Name & "_" & Me.Text & ".xml"

        If Project.SettingsFileExists(SettingsFileName) Then
            Dim Settings As System.Xml.Linq.XDocument
            Project.ReadXmlSettings(SettingsFileName, Settings)

            If IsNothing(Settings) Then 'There is no Settings XML data.
                Exit Sub
            End If

            'Restore a Project Setting example:
            If Settings.<ProjectSettings>.<Setting1>.Value = Nothing Then
                'Project setting not saved.
                'Setting1 = ""
            Else
                'Setting1 = Settings.<ProjectSettings>.<Setting1>.Value
            End If

            'Continue restoring saved settings.

        End If

    End Sub

#End Region 'Process XML Files ----------------------------------------------------------------------------------------------------------------------------------------------------------------

#Region " Form Display Methods - Code used to display this form." '----------------------------------------------------------------------------------------------------------------------------

    Private Sub Main_Load(sender As Object, e As EventArgs) Handles Me.Load

        'Write the startup messages in a stringbuilder object.
        'Messages cannot be written using Message.Add until this is set up later in the startup sequence.
        Dim sb As New System.Text.StringBuilder
        sb.Append("------------------- Starting Application: ADVL Shares Application ---------------------------------------------------------------- " & vbCrLf)

        'Set the Application Directory path: ------------------------------------------------
        Project.ApplicationDir = My.Application.Info.DirectoryPath.ToString

        'Read the Application Information file: ---------------------------------------------
        ApplicationInfo.ApplicationDir = My.Application.Info.DirectoryPath.ToString 'Set the Application Directory property

        If ApplicationInfo.ApplicationLocked Then
            MessageBox.Show("The application is locked. If the application is not already in use, remove the 'Application_Info.lock file from the application directory: " & ApplicationInfo.ApplicationDir, "Notice", MessageBoxButtons.OK)
            Dim dr As System.Windows.Forms.DialogResult
            dr = MessageBox.Show("Press 'Yes' to unlock the application", "Notice", MessageBoxButtons.YesNo)
            If dr = System.Windows.Forms.DialogResult.Yes Then
                ApplicationInfo.UnlockApplication()
            Else
                Application.Exit()
                'System.Windows.Forms.Application.Exit()
            End If
        End If

        ReadApplicationInfo()
        ApplicationInfo.LockApplication()

        'Read the Application Usage information: --------------------------------------------
        ApplicationUsage.StartTime = Now
        ApplicationUsage.SaveLocn.Type = ADVL_Utilities_Library_1.FileLocation.Types.Directory
        ApplicationUsage.SaveLocn.Path = Project.ApplicationDir
        ApplicationUsage.RestoreUsageInfo()
        sb.Append("Application usage: Total duration = " & Format(ApplicationUsage.TotalDuration.TotalHours, "#0.##") & " hours" & vbCrLf)

        'Restore Project information: -------------------------------------------------------
        Project.ApplicationName = ApplicationInfo.Name
        Project.ReadLastProjectInfo()
        Project.ReadProjectInfoFile()
        Project.Usage.StartTime = Now

        'Project.ReadProjectInfoFile()

        ApplicationInfo.SettingsLocn = Project.SettingsLocn

        'Set up the Message object:
        Message.ApplicationName = ApplicationInfo.Name
        Message.SettingsLocn = Project.SettingsLocn

        'Set up Settings Lists:
        FinancialsSettings.FileLocation = Project.DataLocn
        SharePricesSettings.FileLocation = Project.DataLocn
        CalculationsSettings.FileLocation = Project.DataLocn

        'Set up the simple calculations data grid view:
        'Dim TextBoxCol0 As New DataGridViewTextBoxColumn
        'dgvCopyData.Columns.Add(TextBoxCol0)
        'dgvCopyData.Columns(0).HeaderText = "Parameter Name"
        'dgvCopyData.Columns(0).Width = 160

        'NOTE: Separate tabs now used for fifferent calculation types.
        'cmbCalcType.Items.Add("Copy columns")
        'cmbCalcType.Items.Add("Select data")
        'cmbCalcType.Items.Add("Calculations")

        'Initialise Calculations - Copy Data Tab
        cmbCopyDataInputDb.Items.Add("Share Prices")
        cmbCopyDataInputDb.Items.Add("Financials")
        cmbCopyDataInputDb.Items.Add("Calculations")

        cmbCopyDataOutputDb.Items.Add("Share Prices")
        cmbCopyDataOutputDb.Items.Add("Financials")
        cmbCopyDataOutputDb.Items.Add("Calculations")

        'Initialise Calculations - Select Data Tab
        cmbSelectDataInputDb.Items.Add("Share Prices")
        cmbSelectDataInputDb.Items.Add("Financials")
        cmbSelectDataInputDb.Items.Add("Calculations")

        cmbSelectDataOutputDb.Items.Add("Share Prices")
        cmbSelectDataOutputDb.Items.Add("Financials")
        cmbSelectDataOutputDb.Items.Add("Calculations")

        'Initialise Calculations - Simple Calculations Tab
        cmbSimpleCalcDb.Items.Add("Share Prices")
        cmbSimpleCalcDb.Items.Add("Financials")
        cmbSimpleCalcDb.Items.Add("Calculations")

        'Initialise Calculations - Date Calculations Tab
        cmbDateCalcDb.Items.Add("Share Prices")
        cmbDateCalcDb.Items.Add("Financials")
        cmbDateCalcDb.Items.Add("Calculations")

        cmbDateCalcType.Items.Add("Date at start of month")
        cmbDateCalcType.Items.Add("Date at end of month")
        cmbDateCalcType.Items.Add("Date of Start Date add N Days")
        cmbDateCalcType.Items.Add("Date of Start Date minus N Days")
        cmbDateCalcType.Items.Add("Fixed Date")

        'Initialise Calculations - Date Selections Tab
        cmbDateSelectInputDb.Items.Add("Share Prices")
        cmbDateSelectInputDb.Items.Add("Financials")
        cmbDateSelectInputDb.Items.Add("Calculations")

        cmbDateSelectOutputDb.Items.Add("Share Prices")
        cmbDateSelectOutputDb.Items.Add("Financials")
        cmbDateSelectOutputDb.Items.Add("Calculations")

        cmbDateSelectionType.Items.Add("Select Input data with Input date = Output date")
        cmbDateSelectionType.Items.Add("Select first Input data after Output date")
        cmbDateSelectionType.Items.Add("Select last Input data before Output date")

        'Initialise Calculations - Daily Prices Tab
        cmbDailyPriceInputDb.Items.Add("Share Prices")
        cmbDailyPriceInputDb.Items.Add("Financials")
        cmbDailyPriceInputDb.Items.Add("Calculations")

        cmbDailyPriceOutputDb.Items.Add("Share Prices")
        cmbDailyPriceOutputDb.Items.Add("Financials")
        cmbDailyPriceOutputDb.Items.Add("Calculations")

        cmbDailyPriceCalcType.Items.Add("Count the daily number of companies and total value traded") 'Daily trading statistics.
        cmbDailyPriceCalcType.Items.Add("Find trading gaps") 'To detect trading halts, new listings or delistings.
        cmbDailyPriceCalcType.Items.Add("Find first trade and last trade dates") 'To detect new listings and delistings.
        cmbDailyPriceCalcType.Items.Add("Find price level changes") 'To detect share splits and consolidations.
        cmbDailyPriceCalcType.Items.Add("Find price spikes") 'To detect over-bought or over-sold shares.




        'Initialise Utilities - Date Calculations Tab:
        txtYear.Text = "2000"
        txtMonth.Text = "1"
        cmbYearMonthDateCalc.Items.Add("Date of end of month")
        cmbYearMonthDateCalc.Items.Add("Date of start of month")
        cmbYearMonthDateCalc.SelectedIndex = 0

        txtStartDate.Text = "01 Jan 2000"
        txtNDays.Text = "1"
        cmbStartDateNDaysCalc.Items.Add("Date of Start Date + N Days")
        cmbStartDateNDaysCalc.Items.Add("Date of Start Date - N Days")
        cmbStartDateNDaysCalc.SelectedIndex = 0



        'Set up the Simple Calculations tab: -----------------------------------------------
        'dgvSimpleCalcsParameterList.
        'Use these default settings: (The last form setting are sometimes not restored correctly!)
        SplitContainer2.SplitterDistance = 640 'Vertical spliiter distance
        SplitContainer3.SplitterDistance = 320 'LHS horizontal splitter distance
        SplitContainer4.SplitterDistance = 320 'RHS horizontal splitter distance

        'Set up context menus:
        txtCopyDataSettings.ContextMenuStrip = ContextMenuStrip1


        RestoreFormSettings() 'Restore the form settings
        RestoreProjectSettings() 'Restore the Project settings

        'Show the project information: ------------------------------------------------------
        txtProjectName.Text = Project.Name
        txtProjectDescription.Text = Project.Description
        Select Case Project.Type
            Case ADVL_Utilities_Library_1.Project.Types.Directory
                txtProjectType.Text = "Directory"
            Case ADVL_Utilities_Library_1.Project.Types.Archive
                txtProjectType.Text = "Archive"
            Case ADVL_Utilities_Library_1.Project.Types.Hybrid
                txtProjectType.Text = "Hybrid"
            Case ADVL_Utilities_Library_1.Project.Types.None
                txtProjectType.Text = "None"
        End Select
        txtCreationDate.Text = Format(Project.Usage.FirstUsed, "d-MMM-yyyy H:mm:ss")
        txtLastUsed.Text = Format(Project.Usage.LastUsed, "d-MMM-yyyy H:mm:ss")
        Select Case Project.SettingsLocn.Type
            Case ADVL_Utilities_Library_1.FileLocation.Types.Directory
                txtSettingsLocationType.Text = "Directory"
            Case ADVL_Utilities_Library_1.FileLocation.Types.Archive
                txtSettingsLocationType.Text = "Archive"
        End Select
        txtSettingsLocationPath.Text = Project.SettingsLocn.Path
        Select Case Project.DataLocn.Type
            Case ADVL_Utilities_Library_1.FileLocation.Types.Directory
                txtDataLocationType.Text = "Directory"
            Case ADVL_Utilities_Library_1.FileLocation.Types.Archive
                txtDataLocationType.Text = "Archive"
        End Select
        txtDataLocationPath.Text = Project.DataLocn.Path



        sb.Append("------------------- Started OK ------------------------------------------------------------------------------------------------------------------------ " & vbCrLf & vbCrLf)
        Me.Show() 'Show this form before showing the Message form
        Message.Add(sb.ToString)

    End Sub

    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        'Exit the Application

        DisconnectFromAppNet() 'Disconnect from the Application Network.

        'SaveFormSettings() 'Save the settings of this form. 'THESE ARE SAVED WHEN THE FORM_CLOSING EVENT TRIGGERS.
        SaveProjectSettings() 'Save project settings.

        'Save the settings file used to for data views:
        SharePricesSettings.SaveFile()
        FinancialsSettings.SaveFile()
        CalculationsSettings.SaveFile()

        ApplicationInfo.WriteFile() 'Update the Application Information file.
        ApplicationInfo.UnlockApplication()

        Project.SaveLastProjectInfo() 'Save information about the last project used.

        'Project.SaveProjectInfoFile() 'Update the Project Information file. This is not required unless there is a change made to the project.

        Project.Usage.SaveUsageInfo() 'Save Project usage information.

        ApplicationUsage.SaveUsageInfo() 'Save Application usage information.

        Application.Exit()

    End Sub

    Private Sub Main_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        'Save the form settings if the form state is normal. (A minimised form will have the incorrect size and location.)
        If WindowState = FormWindowState.Normal Then
            SaveFormSettings()
        End If
    End Sub

#End Region 'Form Display Methods -------------------------------------------------------------------------------------------------------------------------------------------------------------

#Region " Open and Close Forms - Code used to open and close other forms." '-------------------------------------------------------------------------------------------------------------------

    Private Sub btnMessages_Click(sender As Object, e As EventArgs) Handles btnMessages.Click
        'Show the Messages form.
        Message.ApplicationName = ApplicationInfo.Name
        Message.SettingsLocn = Project.SettingsLocn
        Message.Show()
        Message.MessageForm.BringToFront()
    End Sub

    'THIS CODE IS USED IF A SINGLE SHARE PRICES FORM IS TO BE SHOWN:
    'Private Sub btnViewSharePrices_Click(sender As Object, e As EventArgs) Handles btnViewSharePrices.Click
    '    'Open the Share Prices form:
    '    If IsNothing(SharePrices) Then
    '        SharePrices = New frmSharePrices
    '        SharePrices.Show()
    '    Else
    '        SharePrices.Show()
    '    End If
    '    SharePrices.DatabasePath = SharePriceDbPath
    '    SharePrices.FillCmbSelectTable()
    'End Sub

    'THIS CODE IS USED IF A SINGLE SHARE PRICES FORM IS TO BE SHOWN:
    'Private Sub SharePrices_FormClosed(sender As Object, e As FormClosedEventArgs) Handles SharePrices.FormClosed
    '    SharePrices = Nothing
    'End Sub

    'THIS CODE IS USED IF MULTIPLE SHARE PRICES FORMA ARE TO BE SHOWN:
    Private Sub btnViewSharePrices_Click(sender As Object, e As EventArgs) Handles btnViewSharePrices.Click
        'Open a form to view share price data.

        'Check if one or more forms are selected on lstSharePrices:
        If lstSharePrices.SelectedIndices.Count > 0 Then 'Open each selected form
            For Each item In lstSharePrices.SelectedIndices
                OpenSharePricesFormNo(item)
            Next
        Else  'Open a new share prices form:
            OpenNewSharePricesForm()
        End If

    End Sub

    Private Sub OpenSharePricesFormNo(ByVal Index As Integer)
        'Open the Share Prices form with specified Index number.

        If SharePricesFormList.Count < Index + 1 Then
            'Insert null entries into SharePricesList then add a new form at the specified index position:
            Dim I As Integer
            For I = SharePricesFormList.Count To Index
                SharePricesFormList.Add(Nothing)
            Next
            SharePrices = New frmSharePrices
            SharePricesFormList(Index) = SharePrices
            SharePricesFormList(Index).FormNo = Index
            SharePricesFormList(Index).Show
            'ElseIf SharePricesList(Index) = Nothing Then
        ElseIf IsNothing(SharePricesFormList(Index)) Then
            'Add the new form at specified index position:
            SharePrices = New frmSharePrices
            SharePricesFormList(Index) = SharePrices
            SharePricesFormList(Index).FormNo = Index
            SharePricesFormList(Index).Show()
        Else
            'The form at the specified index poistion is already displayed.
            SharePricesFormList(Index).BringToFront()
        End If
    End Sub

    Private Sub OpenNewSharePricesForm()
        'Code to show multiple instances if the form:
        SharePrices = New frmSharePrices
        If SharePricesFormList.Count = 0 Then
            SharePricesFormList.Add(SharePrices)
            SharePricesFormList(0).FormNo = 0
            SharePricesFormList(0).Show()
        Else
            Dim I As Integer
            Dim FormAdded As Boolean = False
            For I = 0 To SharePricesFormList.Count - 1 'Check if there are closed forms in SharePricesList. They can be re-used.
                If IsNothing(SharePricesFormList(I)) Then
                    SharePricesFormList(I) = SharePrices
                    SharePricesFormList(I).FormNo = I
                    SharePricesFormList(I).Show()
                    FormAdded = True
                    Exit For
                End If
            Next
            If FormAdded = False Then 'Add a new form to SharePricesList.
                Dim FormNo As Integer
                SharePricesFormList.Add(SharePrices)
                FormNo = SharePricesFormList.Count - 1
                SharePricesFormList(FormNo).FormNo = FormNo
                SharePricesFormList(FormNo).Show()
            End If
        End If
    End Sub

    'THIS CODE IS USED IF MULTIPLE SHARE PRICES FORMA ARE TO BE SHOWN:
    Public Sub SharePricesFormClosed()
        'This subroutine is called when the SharePrices form has been closed.
        'The subroutine is usually called from the FormClosed event of the SharePrices form.
        'The SharePrices form may have multiple instances.
        'The ClosedFormNumber property should contains the number of the instance of the SharePrices form.
        'This property should be updated by the SharePrices form when it is being closed.
        'The ClosedFormNumber property value is used to determine which element in SharePricesList should be set to Nothing.

        'ERROR: When application is closed with SharePricesList forms open: !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
        'An unhandled exception of type 'System.ArgumentOutOfRangeException' occurred in mscorlib.dll
        'Additional Information: Index was out of range. Must be non-negative And less than the size of the collection.
        If SharePricesFormList.Count < ClosedFormNo + 1 Then
            'ClosedFormNo is too large to exist in SharePricesList
            Exit Sub
        End If

        If IsNothing(SharePricesFormList(ClosedFormNo)) Then
            'The form is already set to nothing
        Else
            SharePricesFormList(ClosedFormNo) = Nothing
        End If
    End Sub

    Private Sub btnViewFinancials_Click(sender As Object, e As EventArgs) Handles btnViewFinancials.Click
        'Open a form to view Company Financial data.

        'Check if one or more forms are selected on lstSharePrices:
        If lstFinancials.SelectedIndices.Count > 0 Then 'Open each selected form
            For Each item In lstFinancials.SelectedIndices
                OpenFinancialsFormNo(item)
            Next
        Else  'Open a new share prices form:
            OpenNewFinancialsForm()
        End If
    End Sub

    Private Sub OpenFinancialsFormNo(ByVal Index As Integer)
        'Open the Financials form with specified Index number.

        If FinancialsFormList.Count < Index + 1 Then
            'Insert null entries into FinancialsList then add a new form at the specified index position:
            Dim I As Integer
            For I = FinancialsFormList.Count To Index
                FinancialsFormList.Add(Nothing)
            Next
            'Financials = New frmFinancials
            'FinancialsList(Index) = Financials
            FinancialsFormList(Index) = New frmFinancials
            FinancialsFormList(Index).FormNo = Index
            FinancialsFormList(Index).Show
            'ElseIf FinancialsList(Index) = Nothing Then
        ElseIf IsNothing(FinancialsFormList(Index)) Then
            'Add the new form at specified index position:
            'Financials = New frmFinancials
            'FinancialsList(Index) = Financials
            FinancialsFormList(Index) = New frmFinancials
            FinancialsFormList(Index).FormNo = Index
            FinancialsFormList(Index).Show()
        Else
            'The form at the specified index poistion is already displayed.
            FinancialsFormList(Index).BringToFront()
        End If
    End Sub

    Private Sub OpenNewFinancialsForm()
        'Code to show multiple instances if the Financials form:
        Financials = New frmFinancials
        If FinancialsFormList.Count = 0 Then
            FinancialsFormList.Add(Financials)
            FinancialsFormList(0).FormNo = 0
            FinancialsFormList(0).Show()
        Else
            Dim I As Integer
            Dim FormAdded As Boolean = False
            For I = 0 To FinancialsFormList.Count - 1 'Check if there are closed forms in FinancialsFormList. They can be re-used.
                If IsNothing(FinancialsFormList(I)) Then
                    FinancialsFormList(I) = Financials
                    FinancialsFormList(I).FormNo = I
                    FinancialsFormList(I).Show()
                    FormAdded = True
                    Exit For
                End If
            Next
            If FormAdded = False Then 'Add a new form to FinancialsFormList.
                Dim FormNo As Integer
                FinancialsFormList.Add(Financials)
                FormNo = FinancialsFormList.Count - 1
                FinancialsFormList(FormNo).FormNo = FormNo
                FinancialsFormList(FormNo).Show()
            End If
        End If
    End Sub

    Public Sub FinancialsFormClosed()
        'This subroutine is called when the Financials form has been closed.
        'The subroutine is usually called from the FormClosed event of the Financials form.
        'The Financials form may have multiple instances.
        'The ClosedFormNumber property should contains the number of the instance of the Financials form.
        'This property should be updated by the Financials form when it is being closed.
        'The ClosedFormNumber property value is used to determine which element in FinancialsFormList should be set to Nothing.

        If FinancialsFormList.Count < ClosedFormNo + 1 Then
            'ClosedFormNo is too large to exist in FinancialsFormList
            Exit Sub
        End If

        If IsNothing(FinancialsFormList(ClosedFormNo)) Then
            'The form is already set to nothing
        Else
            FinancialsFormList(ClosedFormNo) = Nothing
        End If
    End Sub

    Private Sub btnViewCalcs_Click(sender As Object, e As EventArgs) Handles btnViewCalcs.Click
        'Open a form to view Calculations data.

        'Check if one or more forms are selected on lstCalculations:
        If lstCalculations.SelectedIndices.Count > 0 Then 'Open each selected form
            For Each item In lstCalculations.SelectedIndices
                OpenCalculationsFormNo(item)
            Next
        Else  'Open a new Calculations form:
            OpenNewCalculationsForm()
        End If
    End Sub

    Private Sub OpenCalculationsFormNo(ByVal Index As Integer)
        'Open the Calculations form with specified Index number.

        If CalculationsFormList.Count < Index + 1 Then
            'Insert null entries into CalculationsList then add a new form at the specified index position:
            Dim I As Integer
            For I = CalculationsFormList.Count To Index
                CalculationsFormList.Add(Nothing)
            Next
            CalculationsFormList(Index) = New frmCalculations
            CalculationsFormList(Index).FormNo = Index
            CalculationsFormList(Index).Show
        ElseIf IsNothing(CalculationsFormList(Index)) Then
            'Add the new form at specified index position:
            CalculationsFormList(Index) = New frmCalculations
            CalculationsFormList(Index).FormNo = Index
            CalculationsFormList(Index).Show()
        Else
            'The form at the specified index poistion is already displayed.
            CalculationsFormList(Index).BringToFront()
        End If
    End Sub

    Private Sub OpenNewCalculationsForm()
        'Code to show multiple instances if the Calculations form:
        Calculations = New frmCalculations
        If CalculationsFormList.Count = 0 Then
            CalculationsFormList.Add(Calculations)
            CalculationsFormList(0).FormNo = 0
            CalculationsFormList(0).Show()
        Else
            Dim I As Integer
            Dim FormAdded As Boolean = False
            For I = 0 To CalculationsFormList.Count - 1 'Check if there are closed forms in CalculationsFormList. They can be re-used.
                If IsNothing(CalculationsFormList(I)) Then
                    CalculationsFormList(I) = Calculations
                    CalculationsFormList(I).FormNo = I
                    CalculationsFormList(I).Show()
                    FormAdded = True
                    Exit For
                End If
            Next
            If FormAdded = False Then 'Add a new form to CalculationsFormList.
                Dim FormNo As Integer
                CalculationsFormList.Add(Calculations)
                FormNo = CalculationsFormList.Count - 1
                CalculationsFormList(FormNo).FormNo = FormNo
                CalculationsFormList(FormNo).Show()
            End If
        End If
    End Sub

    Public Sub CalculationsFormClosed()
        'This subroutine is called when the Calculations form has been closed.
        'The subroutine is usually called from the FormClosed event of the Calculations form.
        'The Calculations form may have multiple instances.
        'The ClosedFormNumber property should contains the number of the instance of the Calculations form.
        'This property should be updated by the Calculations form when it is being closed.
        'The ClosedFormNumber property value is used to determine which element in CalculationsFormList should be set to Nothing.

        If CalculationsFormList.Count < ClosedFormNo + 1 Then
            'ClosedFormNo is too large to exist in CalculationsFormList
            Exit Sub
        End If

        If IsNothing(CalculationsFormList(ClosedFormNo)) Then
            'The form is already set to nothing
        Else
            CalculationsFormList(ClosedFormNo) = Nothing
        End If
    End Sub

    Private Sub btnView_Click(sender As Object, e As EventArgs) Handles btnView.Click
        'Show the Processing Sequence form:
        If IsNothing(Sequence) Then
            Sequence = New frmSequence
            Sequence.Show()
        Else
            Sequence.Show()
        End If
    End Sub

    Private Sub Sequence_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Sequence.FormClosed
        Sequence = Nothing
    End Sub


#End Region 'Open and Close Forms -------------------------------------------------------------------------------------------------------------------------------------------------------------

#Region " Form Methods - The main actions performed by this form." '---------------------------------------------------------------------------------------------------------------------------

    Private Sub btnProject_Click(sender As Object, e As EventArgs) Handles btnProject.Click
        Project.SelectProject()
    End Sub

    Private Sub btnAppInfo_Click(sender As Object, e As EventArgs) Handles btnAppInfo.Click
        ApplicationInfo.ShowInfo()
    End Sub

    Private Sub btnAndorville_Click(sender As Object, e As EventArgs) Handles btnAndorville.Click
        ApplicationInfo.ShowInfo()
    End Sub

#Region " Project Events Code"

    Private Sub Project_Message(Msg As String) Handles Project.Message
        'Display the Project message:
        Message.Add(Msg & vbCrLf)
    End Sub

    Private Sub Project_ErrorMessage(Msg As String) Handles Project.ErrorMessage
        'Display the Project error message:
        'Message.SetWarningStyle()
        'Message.Add(Msg & vbCrLf)
        'Message.SetNormalStyle()
        Message.AddWarning(Msg & vbCrLf)
    End Sub

    Private Sub Project_Closing() Handles Project.Closing
        'The current project is closing.

        SaveFormSettings() 'Save the form settings - they are saved in the Project before is closes.
        SaveProjectSettings() 'Update this subroutine if project settings need to be saved.

        'Save the current project usage information:
        Project.Usage.SaveUsageInfo()
    End Sub

    Private Sub Project_Selected() Handles Project.Selected
        'A new project has been selected.

        RestoreFormSettings()
        Project.ReadProjectInfoFile()
        Project.Usage.StartTime = Now

        ApplicationInfo.SettingsLocn = Project.SettingsLocn
        Message.SettingsLocn = Project.SettingsLocn

        'Restore the new project settings:
        RestoreProjectSettings() 'Update this subroutine if project settings need to be restored.

        'Show the project information:
        txtProjectName.Text = Project.Name
        txtProjectDescription.Text = Project.Description
        Select Case Project.Type
            Case ADVL_Utilities_Library_1.Project.Types.Directory
                txtProjectType.Text = "Directory"
            Case ADVL_Utilities_Library_1.Project.Types.Archive
                txtProjectType.Text = "Archive"
            Case ADVL_Utilities_Library_1.Project.Types.Hybrid
                txtProjectType.Text = "Hybrid"
            Case ADVL_Utilities_Library_1.Project.Types.None
                txtProjectType.Text = "None"
        End Select

        txtCreationDate.Text = Format(Project.CreationDate, "d-MMM-yyyy H:mm:ss")
        txtLastUsed.Text = Format(Project.Usage.LastUsed, "d-MMM-yyyy H:mm:ss")
        Select Case Project.SettingsLocn.Type
            Case ADVL_Utilities_Library_1.FileLocation.Types.Directory
                txtSettingsLocationType.Text = "Directory"
            Case ADVL_Utilities_Library_1.FileLocation.Types.Archive
                txtSettingsLocationType.Text = "Archive"
        End Select
        txtSettingsLocationPath.Text = Project.SettingsLocn.Path
        Select Case Project.DataLocn.Type
            Case ADVL_Utilities_Library_1.FileLocation.Types.Directory
                txtDataLocationType.Text = "Directory"
            Case ADVL_Utilities_Library_1.FileLocation.Types.Archive
                txtDataLocationType.Text = "Archive"
        End Select
        txtDataLocationPath.Text = Project.DataLocn.Path

    End Sub

#End Region 'Project Events Code

#Region " Online/Offline Code"

    Private Sub btnOnline_Click(sender As Object, e As EventArgs) Handles btnOnline.Click
        'Connect to or disconnect from the Application Network.
        If ConnectedToAppnet = False Then
            ConnectToAppNet()
        Else
            DisconnectFromAppNet()
        End If
    End Sub

    Private Sub ConnectToAppNet()
        'Connect to the Application Network. (Message Exchange)

        Dim Result As Boolean

        If IsNothing(client) Then
            client = New ServiceReference1.MsgServiceClient(New System.ServiceModel.InstanceContext(New MsgServiceCallback))
        End If

        If client.State = ServiceModel.CommunicationState.Faulted Then
            Message.SetWarningStyle()
            Message.Add("client state is faulted. Connection not made!" & vbCrLf)
        Else
            Try
                client.Endpoint.Binding.SendTimeout = New System.TimeSpan(0, 0, 8) 'Temporarily set the send timeaout to 8 seconds

                Result = client.Connect(ApplicationInfo.Name, ServiceReference1.clsConnectionAppTypes.Application, False, False) 'Application Name is "Application_Template"
                'appName, appType, getAllWarnings, getAllMessages

                If Result = True Then
                    Message.Add("Connected to the Application Network as " & ApplicationInfo.Name & vbCrLf)
                    client.Endpoint.Binding.SendTimeout = New System.TimeSpan(1, 0, 0) 'Restore the send timeaout to 1 hour
                    btnOnline.Text = "Online"
                    btnOnline.ForeColor = Color.ForestGreen
                    ConnectedToAppnet = True
                    SendApplicationInfo()
                Else
                    Message.Add("Connection to the Application Network failed!" & vbCrLf)
                    client.Endpoint.Binding.SendTimeout = New System.TimeSpan(1, 0, 0) 'Restore the send timeaout to 1 hour
                End If
            Catch ex As System.TimeoutException
                Message.Add("Timeout error. Check if the Application Network is running." & vbCrLf)
            Catch ex As Exception
                Message.Add("Error message: " & ex.Message & vbCrLf)
                client.Endpoint.Binding.SendTimeout = New System.TimeSpan(1, 0, 0) 'Restore the send timeaout to 1 hour
            End Try
        End If

    End Sub

    Private Sub DisconnectFromAppNet()
        'Disconnect from the Application Network.

        Dim Result As Boolean

        If IsNothing(client) Then
            Message.Add("Already disconnected from the Application Network." & vbCrLf)
            btnOnline.Text = "Offline"
            btnOnline.ForeColor = Color.Red
            ConnectedToAppnet = False
        Else
            If client.State = ServiceModel.CommunicationState.Faulted Then
                Message.Add("client state is faulted." & vbCrLf)
            Else
                Try
                    Message.Add("Running client.Disconnect(ApplicationName)   ApplicationName = " & ApplicationInfo.Name & vbCrLf)
                    client.Disconnect(ApplicationInfo.Name) 'NOTE: If Application Network has closed, this application freezes at this line! Try Catch EndTry added to fix this.
                    btnOnline.Text = "Offline"
                    btnOnline.ForeColor = Color.Red
                    ConnectedToAppnet = False
                Catch ex As Exception
                    Message.SetWarningStyle()
                    Message.Add("Error disconnecting from Application Network: " & ex.Message & vbCrLf)
                End Try
            End If
        End If
    End Sub

    Private Sub SendApplicationInfo()
        'Send the application information to the Administrator connections.

        If IsNothing(client) Then
            Message.Add("No client connection available!" & vbCrLf)
        Else
            If client.State = ServiceModel.CommunicationState.Faulted Then
                Message.Add("client state is faulted. Message not sent!" & vbCrLf)
            Else
                'Create the XML instructions to send application information.
                Dim decl As New XDeclaration("1.0", "utf-8", "yes")
                Dim doc As New XDocument(decl, Nothing) 'Create an XDocument to store the instructions.
                Dim xmessage As New XElement("XMsg") 'This indicates the start of the message in the XMessage class
                Dim applicationInfo As New XElement("ApplicationInfo")
                Dim name As New XElement("Name", Me.ApplicationInfo.Name)
                applicationInfo.Add(name)

                Dim exePath As New XElement("ExecutablePath", Me.ApplicationInfo.ExecutablePath)
                applicationInfo.Add(exePath)

                Dim directory As New XElement("Directory", Me.ApplicationInfo.ApplicationDir)
                applicationInfo.Add(directory)
                Dim description As New XElement("Description", Me.ApplicationInfo.Description)
                applicationInfo.Add(description)
                xmessage.Add(applicationInfo)
                doc.Add(xmessage)
                client.SendMessage("ApplicationNetwork", doc.ToString)
            End If
        End If

    End Sub

#End Region 'Online/Offline code

#Region " Process XMessages" '=========================================================================================================================================================

    Private Sub XMsg_Instruction(Info As String, Locn As String) Handles XMsg.Instruction
        'Process an XMessage instruction.
        'An XMessage is a simplified XSequence. It is used to exchange information between Andorville (TM) applications.
        '
        'An XSequence file is an AL-H7 (TM) Information Vector Sequence stored in an XML format.
        'AL-H7(TM) is the name of a programming system that uses sequences of information and location value pairs to store data items or processing steps.
        'A single information and location value pair is called a knowledge element (or noxel).
        'Any program, mathematical expression or data set can be expressed as an Information Vector Sequence.

        'Add code here to process the XMessage instructions.
        'See other Andorville(TM) applciations for examples.

        Select Case Locn

            Case "EndOfSequence"
                'End of Information Vector Sequence reached.
            Case Else
                Message.SetWarningStyle()
                Message.Add("Unknown location: " & Locn & vbCrLf)
                Message.SetNormalStyle()
        End Select

    End Sub

    Private Sub SendMessage()
        'Code used to send a message after a timer delay.
        'The message destination is stored in MessageDest
        'The message text is stored in MessageText
        Timer1.Interval = 100 '100ms delay
        Timer1.Enabled = True 'Start the timer.
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick

        If IsNothing(client) Then
            Message.AddWarning("No client connection available!" & vbCrLf)
        Else
            If client.State = ServiceModel.CommunicationState.Faulted Then
                Message.AddWarning("client state is faulted. Message not sent!" & vbCrLf)
            Else
                Try
                    Message.Add("Sending a message. Number of characters: " & MessageText.Length & vbCrLf)
                    client.SendMessage(ClientAppName, MessageText)
                    MessageText = "" 'Clear the message after it has been sent.
                    ClientAppName = "" 'Clear the Client Application Name after the message has been sent.
                    ClientAppLocn = "" 'Clear the Client Application Location after the message has been sent.
                Catch ex As Exception
                    Message.AddWarning("Error sending message: " & ex.Message & vbCrLf)
                End Try
            End If
        End If

        'Stop timer:
        Timer1.Enabled = False
    End Sub

#End Region 'Process XMessages --------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Settings Tab" '==============================================================================================================================================================

    Private Sub btnFindSharePriceDatabase_Click(sender As Object, e As EventArgs) Handles btnFindSharePriceDatabase.Click
        'Find a Share Price database:

        If SharePriceDbPath = "" Then
            OpenFileDialog1.InitialDirectory = System.Environment.SpecialFolder.MyDocuments
            OpenFileDialog1.Filter = "Database |*.accdb; *.mdb"
            OpenFileDialog1.FileName = ""
        Else
            Dim fInfo As New System.IO.FileInfo(SharePriceDbPath)
            OpenFileDialog1.InitialDirectory = fInfo.DirectoryName
            OpenFileDialog1.Filter = "Database |*.accdb; *.mdb"
            OpenFileDialog1.FileName = fInfo.Name
        End If

        If OpenFileDialog1.ShowDialog() = vbOK Then
            SharePriceDbPath = OpenFileDialog1.FileName
        End If
    End Sub

    Private Sub btnFindFinancialsDatabase_Click(sender As Object, e As EventArgs) Handles btnFindFinancialsDatabase.Click
        'Find a Historical Financials database:

        If FinancialsDbPath = "" Then
            OpenFileDialog1.InitialDirectory = System.Environment.SpecialFolder.MyDocuments
            OpenFileDialog1.Filter = "Database |*.accdb; *.mdb"
            OpenFileDialog1.FileName = ""
        Else
            Dim fInfo As New System.IO.FileInfo(FinancialsDbPath)
            OpenFileDialog1.InitialDirectory = fInfo.DirectoryName
            OpenFileDialog1.Filter = "Database |*.accdb; *.mdb"
            OpenFileDialog1.FileName = fInfo.Name
        End If

        If OpenFileDialog1.ShowDialog() = vbOK Then
            FinancialsDbPath = OpenFileDialog1.FileName
        End If
    End Sub

    Private Sub btnFindCalcsDatabase_Click(sender As Object, e As EventArgs) Handles btnFindCalcsDatabase.Click
        'Find a Calculations database:

        If CalculationsDbPath = "" Then
            OpenFileDialog1.InitialDirectory = System.Environment.SpecialFolder.MyDocuments
            OpenFileDialog1.Filter = "Database |*.accdb; *.mdb"
            OpenFileDialog1.FileName = ""
        Else
            Dim fInfo As New System.IO.FileInfo(CalculationsDbPath)
            OpenFileDialog1.InitialDirectory = fInfo.DirectoryName
            OpenFileDialog1.Filter = "Database |*.accdb; *.mdb"
            OpenFileDialog1.FileName = fInfo.Name
        End If

        If OpenFileDialog1.ShowDialog() = vbOK Then
            CalculationsDbPath = OpenFileDialog1.FileName
        End If
    End Sub

    Private Sub btnFindNewsDirectory_Click(sender As Object, e As EventArgs) Handles btnFindNewsDirectory.Click
        'Find a News directory:

    End Sub

#End Region 'Settings Tab -------------------------------------------------------------------------------------------------------------------------------------------------------------

#Region " View Data Tab" '=============================================================================================================================================================

#Region " View Share Prices Sub Tab" '=================================================================================================================================================

    Public Sub UpdateSharePricesDataDescr(ByVal IndexNo As Integer, ByVal Description As String)
        'Set the Share Prices data description in lstSharePrices list box.
        '  IndexNo is the index number of the item in the list.
        '  Description is the data description to be entered at that index number'

        Dim ListCount As Integer = lstSharePrices.Items.Count

        If IndexNo >= ListCount Then
            'Pad out entries in lstSharePrices list box:
            Dim I As Integer
            For I = ListCount To IndexNo
                lstSharePrices.Items.Add("")
            Next
        End If
        lstSharePrices.Items(IndexNo) = Description
    End Sub

    Private Sub btnInsertViewSPBefore_Click(sender As Object, e As EventArgs) Handles btnInsertViewSPBefore.Click
        'Insert a new Share Prices view before the item selected in the share prices list.
        'If no item is selected, insert the new view at the start of the list.

        Dim SelectedIndex As Integer = lstSharePrices.SelectedIndex 'The index of the selected view.
        Dim NViews As Integer = lstSharePrices.Items.Count 'The number of views in the Share Prices list.
        If NViews = 0 Then
            lstSharePrices.Items.Add("")
            Dim NewSettings As New DataViewSettings
            SharePricesSettings.List.Add(NewSettings)
            OpenNewSharePricesForm()
        ElseIf NViews = 1 Then
            lstSharePrices.Items.Insert(0, "")
            'Insert a new Settings entry in FinancialSettings a position 0:
            Dim NewSettings As New DataViewSettings
            SharePricesSettings.List.Insert(0, NewSettings)
            OpenSharePricesFormNo(0) 'Open the new blank view in the first position.
        Else
            If SelectedIndex >= 0 Then
                lstSharePrices.Items.Insert(SelectedIndex, "")
                'Insert a new Settings entry in SharePricesSettings:
                Dim NewSettings As New DataViewSettings
                SharePricesSettings.List.Insert(SelectedIndex, NewSettings)
                OpenSharePricesFormNo(SelectedIndex)
            Else
                'No item selected
                lstSharePrices.Items.Insert(0, "")
                Dim NewSettings As New DataViewSettings
                SharePricesSettings.List.Insert(0, NewSettings)
                OpenSharePricesFormNo(0)
            End If
        End If
    End Sub

    Private Sub btnInsertViewSPAfter_Click(sender As Object, e As EventArgs) Handles btnInsertViewSPAfter.Click
        'Insert a new Share Prices view after the item selected in the share prices list.
        'If no item is selected, insert the new view at the end of the list.

        Dim SelectedIndex As Integer = lstSharePrices.SelectedIndex 'The index of the selected view.
        Dim NViews As Integer = lstSharePrices.Items.Count 'The number of views in the Financials list.
        If NViews = 0 Then
            lstSharePrices.Items.Add("")
            Dim NewSettings As New DataViewSettings
            SharePricesSettings.List.Add(NewSettings)
            OpenNewSharePricesForm()
        ElseIf NViews = 1 Then
            'Add a new View at the end of the list.
            lstSharePrices.Items.Add("")
            Dim NewSettings As New DataViewSettings
            SharePricesSettings.List.Add(NewSettings)
            OpenSharePricesFormNo(1) 'Open the new blank view in the new second position.
        Else
            If SelectedIndex >= 0 Then
                If SelectedIndex = NViews - 1 Then 'Last item selected
                    'Add the new View to the end of the list.
                    lstSharePrices.Items.Add("")
                    Dim NewSettings As New DataViewSettings
                    SharePricesSettings.List.Add(NewSettings)
                    'OpenFinancialsFormNo(NViews + 1)
                    OpenSharePricesFormNo(NViews)
                Else
                    lstSharePrices.Items.Insert(SelectedIndex + 1, "")
                    'Insert a new Settings entry in FinancialSettings:
                    Dim NewSettings As New DataViewSettings
                    SharePricesSettings.List.Insert(SelectedIndex + 1, NewSettings)
                    OpenSharePricesFormNo(SelectedIndex + 1)
                End If
            Else
                'No item selected
                'Add the new View to the end of the list.
                lstSharePrices.Items.Add("")
                Dim NewSettings As New DataViewSettings
                SharePricesSettings.List.Add(NewSettings)
                OpenSharePricesFormNo(NViews + 1)
            End If
        End If
    End Sub

    Private Sub btnDeleteViewSP_Click(sender As Object, e As EventArgs) Handles btnDeleteViewSP.Click
        'Delete selected view

        Dim SelectedIndex As Integer = lstSharePrices.SelectedIndex 'The index of the selected view.
        Dim NViews As Integer = lstSharePrices.Items.Count 'The number of views in the Share Prices list.

        If NViews = 0 Then
            'No Views to delete.
        Else
            lstSharePrices.Items.RemoveAt(SelectedIndex) 'Remove the selected View in lstFinancials
            'Close the form if it is open:
            If IsNothing(SharePricesFormList(SelectedIndex)) Then
            Else
                SharePricesFormList(SelectedIndex).CloseForm
            End If
            'Delete the entry in SharePricesSettings
            SharePricesSettings.List.RemoveAt(SelectedIndex)
        End If
    End Sub

    Private Sub btnSaveSharePricesDataList_Click(sender As Object, e As EventArgs) Handles btnSaveSharePricesDataList.Click
        'Save the Share Prices data view list.
        SaveSharePricesDataList()
    End Sub

    Private Sub SaveSharePricesDataList()
        If Trim(txtSharePricesDataList.Text) = "" Then
            Message.AddWarning("No file name has been specified to save the list of Share Prices data views!" & vbCrLf)
        Else
            If txtSharePricesDataList.Text.EndsWith(".SPDataList") Then
                txtSharePricesDataList.Text = Trim(txtSharePricesDataList.Text)
            Else
                txtSharePricesDataList.Text = Trim(txtSharePricesDataList.Text) & ".SPDataList"
            End If
            SharePricesSettings.ListFileName = txtSharePricesDataList.Text
            SharePriceDataViewList = txtSharePricesDataList.Text
            SharePricesSettings.SaveFile()
            SharePriceSettingsChanged = False
        End If
    End Sub

    Private Sub btnFindSharePricesDataList_Click(sender As Object, e As EventArgs) Handles btnFindSharePricesDataList.Click
        'Find a Share Prices data view list:

        Select Case Project.DataLocn.Type
            Case ADVL_Utilities_Library_1.FileLocation.Types.Directory
                'Select a Share Prices data view list from the project directory:
                OpenFileDialog1.InitialDirectory = Project.DataLocn.Path
                OpenFileDialog1.Filter = "Share Prices Data View List | *.SPDataList"
                If OpenFileDialog1.ShowDialog() = DialogResult.OK Then
                    Dim DataFileName As String = System.IO.Path.GetFileName(OpenFileDialog1.FileName)
                    SharePricesSettings.ListFileName = DataFileName
                    SharePriceDataViewList = DataFileName
                    txtSharePricesDataList.Text = DataFileName
                    SharePricesSettings.LoadFile()
                    DisplaySharePricesList()
                End If
            Case ADVL_Utilities_Library_1.FileLocation.Types.Archive
                'Select a Financials Data View list file from the project archive:
                'Show the zip archive file selection form:
                Zip = New ADVL_Utilities_Library_1.ZipComp
                Zip.ArchivePath = Project.DataLocn.Path
                Zip.SelectFile() 'Show the SelectFile form.
                Zip.SelectFileForm.ApplicationName = Project.ApplicationName
                Zip.SelectFileForm.SettingsLocn = Project.SettingsLocn
                Zip.SelectFileForm.Show()
                Zip.SelectFileForm.RestoreFormSettings()
                Zip.SelectFileForm.FileExtension = ".SPDataList"
                Zip.SelectFileForm.GetFileList()
                If Zip.SelectedFile <> "" Then
                    'A file has been selected
                    SharePricesSettings.ListFileName = Zip.SelectedFile
                    SharePriceDataViewList = Zip.SelectedFile
                    txtSharePricesDataList.Text = Zip.SelectedFile
                    SharePricesSettings.LoadFile()
                    DisplaySharePricesList()
                End If
        End Select
    End Sub

    Private Sub DisplaySharePricesList()
        'Display the Share Prices data view list descriptions in lstSharePrices:
        lstSharePrices.Items.Clear()
        Dim I As Integer
        For I = 0 To SharePricesSettings.NRecords - 1
            lstSharePrices.Items.Add(SharePricesSettings.List(I).Description)
        Next
    End Sub

#End Region 'View Share Prices Sub Tab ------------------------------------------------------------------------------------------------------------------------------------------------

#Region " View Financials Sub Tab" '===================================================================================================================================================

    Public Sub UpdateFinancialsDataDescr(ByVal IndexNo As Integer, ByVal Description As String)
        'Set the Financials data description in lstFinancials list box.
        '  IndexNo is the index number of the item in the list.
        '  Description is the data description to be entered at that index number'

        Dim ListCount As Integer = lstFinancials.Items.Count

        If IndexNo >= ListCount Then
            'Pad out entries in lstFinancials list box:
            Dim I As Integer
            For I = ListCount To IndexNo
                lstFinancials.Items.Add("")
            Next
        End If
        lstFinancials.Items(IndexNo) = Description
    End Sub

    Private Sub btnInsertViewFinBefore_Click(sender As Object, e As EventArgs) Handles btnInsertViewFinBefore.Click
        'Insert a new Financials view before the item selected in the financials list.
        'If no item is selected, insert the new view at the start of the list.

        Dim SelectedIndex As Integer = lstFinancials.SelectedIndex 'The index of the selected view.
        Dim NViews As Integer = lstFinancials.Items.Count 'The number of views in the Financials list.
        If NViews = 0 Then
            lstFinancials.Items.Add("")
            Dim NewSettings As New DataViewSettings
            FinancialsSettings.List.Add(NewSettings)
            OpenNewFinancialsForm()
        ElseIf NViews = 1 Then
            ''Move the existing View down one position and insert a blank View in the first position.
            'lstFinancials.Items.Add(lstFinancials.Items(0).ToString)
            'lstFinancials.Items(0) = ""

            'Dim OldSettingsFilename As String = "FormSettings_" & ApplicationInfo.Name & "_Financials_0.xml" '"FormSettings_" & Main.ApplicationInfo.Name & "_" & Me.Text & "_" & FormNo & ".xml"
            'Dim NewSettingsFileName As String = "FormSettings_" & ApplicationInfo.Name & "_Financials_1.xml"
            'Project.RenameSettingsFile(OldSettingsFilename, NewSettingsFileName)

            lstFinancials.Items.Insert(0, "")
            'Insert a new Settings entry in FinancialSettings a position 0:
            Dim NewSettings As New DataViewSettings
            FinancialsSettings.List.Insert(0, NewSettings)

            OpenFinancialsFormNo(0) 'Open the new blank view in the first position.
        Else
            If SelectedIndex >= 0 Then
                'Move the last View down one position:
                'lstFinancials.Items.Add(lstFinancials.Items(NViews - 1).ToString)

                lstFinancials.Items.Insert(SelectedIndex, "")

                'Dim OldSettingsFilename As String = "FormSettings_" & ApplicationInfo.Name & "_Financials_" & NViews - 1 & ".xml" '"FormSettings_" & Main.ApplicationInfo.Name & "_" & Me.Text & "_" & FormNo & ".xml"
                'Dim NewSettingsFileName As String = "FormSettings_" & ApplicationInfo.Name & "_Financials_" & NViews & ".xml"
                'Project.RenameSettingsFile(OldSettingsFilename, NewSettingsFileName)
                'Dim I As Integer
                'For I = NViews - 1 To SelectedIndex Step -1
                '    lstFinancials.Items(I) = lstFinancials.Items(I - 1)
                '    OldSettingsFilename = "FormSettings_" & ApplicationInfo.Name & "_Financials_" & I - 1 & ".xml"
                '    NewSettingsFileName = "FormSettings_" & ApplicationInfo.Name & "_Financials_" & I & ".xml"
                '    Project.RenameSettingsFile(OldSettingsFilename, NewSettingsFileName)
                'Next

                'Insert a new Settings entry in FinancialSettings:
                Dim NewSettings As New DataViewSettings
                FinancialsSettings.List.Insert(SelectedIndex, NewSettings)


                'lstFinancials.Items(SelectedIndex) = ""
                OpenFinancialsFormNo(SelectedIndex)
            Else
                'No item selected
                ''Move the last View down one position:
                'lstFinancials.Items.Add(lstFinancials.Items(NViews - 1).ToString)
                'Dim OldSettingsFilename As String = "FormSettings_" & ApplicationInfo.Name & "_Financials_" & NViews - 1 & ".xml" '"FormSettings_" & Main.ApplicationInfo.Name & "_" & Me.Text & "_" & FormNo & ".xml"
                'Dim NewSettingsFileName As String = "FormSettings_" & ApplicationInfo.Name & "_Financials_" & NViews & ".xml"
                'Project.RenameSettingsFile(OldSettingsFilename, NewSettingsFileName)
                'Dim I As Integer
                'For I = NViews - 1 To 1 Step -1
                '    lstFinancials.Items(I) = lstFinancials.Items(I - 1)
                '    OldSettingsFilename = "FormSettings_" & ApplicationInfo.Name & "_Financials_" & I - 1 & ".xml"
                '    NewSettingsFileName = "FormSettings_" & ApplicationInfo.Name & "_Financials_" & I & ".xml"
                '    Project.RenameSettingsFile(OldSettingsFilename, NewSettingsFileName)
                'Next
                'lstFinancials.Items(0) = ""

                lstFinancials.Items.Insert(0, "")
                Dim NewSettings As New DataViewSettings
                FinancialsSettings.List.Insert(0, NewSettings)

                OpenFinancialsFormNo(0)
            End If
        End If
    End Sub

    Private Sub btnInsertViewFinAfter_Click(sender As Object, e As EventArgs) Handles btnInsertViewFinAfter.Click
        'Insert a new Financials view after the item selected in the financials list.
        'If no item is selected, insert the new view at the end of the list.

        Dim SelectedIndex As Integer = lstFinancials.SelectedIndex 'The index of the selected view.
        Dim NViews As Integer = lstFinancials.Items.Count 'The number of views in the Financials list.
        If NViews = 0 Then
            lstFinancials.Items.Add("")
            Dim NewSettings As New DataViewSettings
            FinancialsSettings.List.Add(NewSettings)
            OpenNewFinancialsForm()
        ElseIf NViews = 1 Then
            'Add a new View at the end of the list.
            lstFinancials.Items.Add("")
            Dim NewSettings As New DataViewSettings
            FinancialsSettings.List.Add(NewSettings)
            OpenFinancialsFormNo(1) 'Open the new blank view in the new second position.
        Else
            If SelectedIndex >= 0 Then
                If SelectedIndex = NViews - 1 Then 'Last item selected
                    'Add the new View to the end of the list.
                    lstFinancials.Items.Add("")
                    Dim NewSettings As New DataViewSettings
                    FinancialsSettings.List.Add(NewSettings)
                    'OpenFinancialsFormNo(NViews + 1)
                    OpenFinancialsFormNo(NViews)
                Else
                    'Move the last View down one position:
                    'lstFinancials.Items.Add(lstFinancials.Items(NViews - 1).ToString)

                    'Dim OldSettingsFilename As String = "FormSettings_" & ApplicationInfo.Name & "_Financials_" & NViews - 1 & ".xml"
                    'Dim NewSettingsFileName As String = "FormSettings_" & ApplicationInfo.Name & "_Financials_" & NViews & ".xml"
                    'Project.RenameSettingsFile(OldSettingsFilename, NewSettingsFileName)
                    'Dim I As Integer
                    ''For I = NViews - 1 To SelectedIndex + 1 Step -1
                    'For I = NViews - 1 To SelectedIndex + 2 Step -1
                    '    lstFinancials.Items(I) = lstFinancials.Items(I - 1)
                    '    OldSettingsFilename = "FormSettings_" & ApplicationInfo.Name & "_Financials_" & I - 1 & ".xml"
                    '    NewSettingsFileName = "FormSettings_" & ApplicationInfo.Name & "_Financials_" & I & ".xml"
                    '    Project.RenameSettingsFile(OldSettingsFilename, NewSettingsFileName)
                    'Next

                    lstFinancials.Items.Insert(SelectedIndex + 1, "")
                    'Insert a new Settings entry in FinancialSettings:
                    Dim NewSettings As New DataViewSettings
                    FinancialsSettings.List.Insert(SelectedIndex + 1, NewSettings)

                    ''Insert a new entry in lstFinancials:
                    'lstFinancials.Items(SelectedIndex + 1) = ""
                    OpenFinancialsFormNo(SelectedIndex + 1)
                End If
            Else
                'No item selected
                'Add the new View to the end of the list.
                lstFinancials.Items.Add("")
                Dim NewSettings As New DataViewSettings
                FinancialsSettings.List.Add(NewSettings)
                OpenFinancialsFormNo(NViews + 1)
            End If
        End If

    End Sub

    Private Sub btnDeleteViewFin_Click(sender As Object, e As EventArgs) Handles btnDeleteViewFin.Click
        'Delete selected view

        Dim SelectedIndex As Integer = lstFinancials.SelectedIndex 'The index of the selected view.
        Dim NViews As Integer = lstFinancials.Items.Count 'The number of views in the Financials list.

        If NViews = 0 Then
            'No Views to delete.
        Else
            lstFinancials.Items.RemoveAt(SelectedIndex) 'Remove the selected View in lstFinancials
            'Close the form if it is open:
            If IsNothing(FinancialsFormList(SelectedIndex)) Then
            Else
                FinancialsFormList(SelectedIndex).CloseForm
            End If

            ''Rename the settings files
            'Dim I As Integer
            'Dim OldSettingsFilename As String
            'Dim NewSettingsFileName As String
            'For I = SelectedIndex To NViews - 2
            '    OldSettingsFilename = "FormSettings_" & ApplicationInfo.Name & "_Financials_" & I + 1 & ".xml"
            '    NewSettingsFileName = "FormSettings_" & ApplicationInfo.Name & "_Financials_" & I & ".xml"
            '    Project.RenameSettingsFile(OldSettingsFilename, NewSettingsFileName)
            'Next

            'Delete the entry in FinancialsSettings
            FinancialsSettings.List.RemoveAt(SelectedIndex)

        End If



    End Sub

    Private Sub btnSaveFinDataList_Click(sender As Object, e As EventArgs) Handles btnSaveFinDataList.Click
        'Save the Financials data view list.
        SaveFinDataList()
    End Sub

    Private Sub SaveFinDataList()
        If Trim(txtFinancialsDataList.Text) = "" Then
            Message.AddWarning("No file name has been specified to save the list of Financial data views!" & vbCrLf)
        Else
            If txtFinancialsDataList.Text.EndsWith(".FinDataList") Then
                txtFinancialsDataList.Text = Trim(txtFinancialsDataList.Text)
            Else
                txtFinancialsDataList.Text = Trim(txtFinancialsDataList.Text) & ".FinDataList"
            End If
            FinancialsSettings.ListFileName = txtFinancialsDataList.Text
            FinancialsDataViewList = txtFinancialsDataList.Text
            FinancialsSettings.SaveFile()
            FinancialsSettingsChanged = False
        End If
    End Sub

    Private Sub btnFindFinDataList_Click(sender As Object, e As EventArgs) Handles btnFindFinDataList.Click
        'Find a Financials data view list:

        Select Case Project.DataLocn.Type
            Case ADVL_Utilities_Library_1.FileLocation.Types.Directory
                'Select a Financials data view list from the project directory:
                OpenFileDialog1.InitialDirectory = Project.DataLocn.Path
                OpenFileDialog1.Filter = "Financials Data View List | *.FinDataList"
                If OpenFileDialog1.ShowDialog() = DialogResult.OK Then
                    Dim DataFileName As String = System.IO.Path.GetFileName(OpenFileDialog1.FileName)
                    FinancialsSettings.ListFileName = DataFileName
                    FinancialsDataViewList = DataFileName
                    txtFinancialsDataList.Text = DataFileName
                    FinancialsSettings.LoadFile()
                    DisplayFinancialsList()
                End If
            Case ADVL_Utilities_Library_1.FileLocation.Types.Archive
                'Select a Financials Data View list file from the project archive:
                'Show the zip archive file selection form:
                Zip = New ADVL_Utilities_Library_1.ZipComp
                Zip.ArchivePath = Project.DataLocn.Path
                Zip.SelectFile() 'Show the SelectFile form.
                Zip.SelectFileForm.ApplicationName = Project.ApplicationName
                Zip.SelectFileForm.SettingsLocn = Project.SettingsLocn
                Zip.SelectFileForm.Show()
                Zip.SelectFileForm.RestoreFormSettings()
                Zip.SelectFileForm.FileExtension = ".FinDataList"
                Zip.SelectFileForm.GetFileList()
                If Zip.SelectedFile <> "" Then
                    'A file has been selected
                    FinancialsSettings.ListFileName = Zip.SelectedFile
                    FinancialsDataViewList = Zip.SelectedFile
                    txtFinancialsDataList.Text = Zip.SelectedFile
                    FinancialsSettings.LoadFile()
                    DisplayFinancialsList()
                End If
        End Select
    End Sub

    Private Sub DisplayFinancialsList()
        'Display the Financials data view list descriptions in lstFinancials:
        lstFinancials.Items.Clear()
        Dim I As Integer
        For I = 0 To FinancialsSettings.NRecords - 1
            lstFinancials.Items.Add(FinancialsSettings.List(I).Description)
        Next
    End Sub


#End Region 'View Financials Sub Tab --------------------------------------------------------------------------------------------------------------------------------------------------

#Region " View Calculations Sub Tab" '=================================================================================================================================================

    Public Sub UpdateCalculationsDataDescr(ByVal IndexNo As Integer, ByVal Description As String)
        'Set the Calculations data description in lstCalculations list box.
        '  IndexNo is the index number of the item in the list.
        '  Description is the data description to be entered at that index number'

        Dim ListCount As Integer = lstCalculations.Items.Count

        If IndexNo >= ListCount Then
            'Pad out entries in lstCalculations list box:
            Dim I As Integer
            For I = ListCount To IndexNo
                lstCalculations.Items.Add("")
            Next
        End If
        lstCalculations.Items(IndexNo) = Description
    End Sub

    Private Sub btnDeleteViewCalcs_Click(sender As Object, e As EventArgs) Handles btnDeleteViewCalcs.Click
        'Delete selected view

        Dim SelectedIndex As Integer = lstCalculations.SelectedIndex 'The index of the selected view.
        Dim NViews As Integer = lstCalculations.Items.Count 'The number of views in the Financials list.

        If NViews = 0 Then
            'No Views to delete.
        Else
            lstCalculations.Items.RemoveAt(SelectedIndex) 'Remove the selected View in lstFinancials
            'Close the form if it is open:
            If SelectedIndex > CalculationsFormList.Count - 1 Then
                'No entry exists in the CalculationsFormList at SelectedIndex
                Exit Sub
            End If
            If IsNothing(CalculationsFormList(SelectedIndex)) Then
            Else
                CalculationsFormList(SelectedIndex).CloseForm
            End If

            'Delete the entry in CalculationsSettings
            CalculationsSettings.List.RemoveAt(SelectedIndex)

        End If
    End Sub

    Private Sub btnInsertViewCalcsBefore_Click(sender As Object, e As EventArgs) Handles btnInsertViewCalcsBefore.Click
        'Insert a new Calculations view before the item selected in the calculations list.
        'If no item is selected, insert the new view at the start of the list.

        Dim SelectedIndex As Integer = lstCalculations.SelectedIndex 'The index of the selected view.
        Dim NViews As Integer = lstCalculations.Items.Count 'The number of views in the Calculations list.
        If NViews = 0 Then
            lstCalculations.Items.Add("")
            Dim NewSettings As New DataViewSettings
            CalculationsSettings.List.Add(NewSettings)
            OpenNewCalculationsForm()
        ElseIf NViews = 1 Then
            'Insert a new Settings entry in CalculationsSettings a position 0:
            lstCalculations.Items.Insert(0, "")
            Dim NewSettings As New DataViewSettings
            CalculationsSettings.List.Insert(0, NewSettings)

            OpenCalculationsFormNo(0) 'Open the new blank view in the first position.
        Else
            If SelectedIndex >= 0 Then
                lstCalculations.Items.Insert(SelectedIndex, "")
                'Insert a new Settings entry in FinancialSettings:
                Dim NewSettings As New DataViewSettings
                CalculationsSettings.List.Insert(SelectedIndex, NewSettings)
                OpenCalculationsFormNo(SelectedIndex)
            Else
                'No item selected
                lstCalculations.Items.Insert(0, "")
                Dim NewSettings As New DataViewSettings
                CalculationsSettings.List.Insert(0, NewSettings)
                OpenCalculationsFormNo(0)
            End If
        End If
    End Sub

    Private Sub btnInsertViewCalcsAfter_Click(sender As Object, e As EventArgs) Handles btnInsertViewCalcsAfter.Click
        'Insert a new Calculations view after the item selected in the calculations list.
        'If no item is selected, insert the new view at the end of the list.

        Dim SelectedIndex As Integer = lstCalculations.SelectedIndex 'The index of the selected view.
        Dim NViews As Integer = lstCalculations.Items.Count 'The number of views in the Financials list.
        If NViews = 0 Then
            lstCalculations.Items.Add("")
            Dim NewSettings As New DataViewSettings
            CalculationsSettings.List.Add(NewSettings)
            OpenNewCalculationsForm()
        ElseIf NViews = 1 Then
            'Add a new View at the end of the list.
            lstCalculations.Items.Add("")
            Dim NewSettings As New DataViewSettings
            CalculationsSettings.List.Add(NewSettings)
            OpenCalculationsFormNo(1) 'Open the new blank view in the new second position.
        Else
            If SelectedIndex >= 0 Then
                If SelectedIndex = NViews - 1 Then 'Last item selected
                    'Add the new View to the end of the list.
                    lstCalculations.Items.Add("")
                    Dim NewSettings As New DataViewSettings
                    CalculationsSettings.List.Add(NewSettings)
                    OpenCalculationsFormNo(NViews)
                Else
                    'Insert a new Settings entry in FinancialSettings:
                    lstCalculations.Items.Insert(SelectedIndex + 1, "")
                    Dim NewSettings As New DataViewSettings
                    CalculationsSettings.List.Insert(SelectedIndex + 1, NewSettings)
                    OpenCalculationsFormNo(SelectedIndex + 1)
                End If
            Else
                'No item selected
                'Add the new View to the end of the list.
                lstCalculations.Items.Add("")
                Dim NewSettings As New DataViewSettings
                CalculationsSettings.List.Add(NewSettings)
                OpenCalculationsFormNo(NViews + 1)
            End If
        End If
    End Sub

    Private Sub btnSaveCalcsDataList_Click(sender As Object, e As EventArgs) Handles btnSaveCalcsDataList.Click
        'Save the Calculations data view list.
        SaveCalcDataList()
    End Sub

    Private Sub SaveCalcDataList()
        If Trim(txtCalcsDataList.Text) = "" Then
            Message.AddWarning("No file name has been specified to save the list of Calculations data views!" & vbCrLf)
        Else
            If txtCalcsDataList.Text.EndsWith(".CalcDataList") Then
                txtCalcsDataList.Text = Trim(txtCalcsDataList.Text)
            Else
                txtCalcsDataList.Text = Trim(txtCalcsDataList.Text) & ".CalcDataList"
            End If
            CalculationsSettings.ListFileName = txtCalcsDataList.Text
            CalculationsDataViewList = txtCalcsDataList.Text
            CalculationsSettings.SaveFile()
            CalculationsSettingsChanged = False
        End If
    End Sub

    Private Sub btnFindCalcsDataList_Click(sender As Object, e As EventArgs) Handles btnFindCalcsDataList.Click
        'Find a Calculations data view list:

        Select Case Project.DataLocn.Type
            Case ADVL_Utilities_Library_1.FileLocation.Types.Directory
                'Select a Calculations data view list from the project directory:
                OpenFileDialog1.InitialDirectory = Project.DataLocn.Path
                OpenFileDialog1.Filter = "Calculations Data View List | *.CalcDataList"
                If OpenFileDialog1.ShowDialog() = DialogResult.OK Then
                    Dim DataFileName As String = System.IO.Path.GetFileName(OpenFileDialog1.FileName)
                    CalculationsSettings.ListFileName = DataFileName
                    CalculationsDataViewList = DataFileName
                    txtCalcsDataList.Text = DataFileName
                    CalculationsSettings.LoadFile()
                    DisplayCalculationsList()
                End If
            Case ADVL_Utilities_Library_1.FileLocation.Types.Archive
                'Select a Calculations Data View list file from the project archive:
                'Show the zip archive file selection form:
                Zip = New ADVL_Utilities_Library_1.ZipComp
                Zip.ArchivePath = Project.DataLocn.Path
                Zip.SelectFile() 'Show the SelectFile form.
                Zip.SelectFileForm.ApplicationName = Project.ApplicationName
                Zip.SelectFileForm.SettingsLocn = Project.SettingsLocn
                Zip.SelectFileForm.Show()
                Zip.SelectFileForm.RestoreFormSettings()
                Zip.SelectFileForm.FileExtension = ".CalcDataList"
                Zip.SelectFileForm.GetFileList()
                If Zip.SelectedFile <> "" Then
                    'A file has been selected
                    CalculationsSettings.ListFileName = Zip.SelectedFile
                    CalculationsDataViewList = Zip.SelectedFile
                    txtCalcsDataList.Text = Zip.SelectedFile
                    CalculationsSettings.LoadFile()
                    DisplayCalculationsList()
                End If
        End Select
    End Sub

    Private Sub DisplayCalculationsList()
        'Display the Calculations data view list descriptions in lstCalculations:
        lstCalculations.Items.Clear()
        Dim I As Integer
        For I = 0 To CalculationsSettings.NRecords - 1
            lstCalculations.Items.Add(CalculationsSettings.List(I).Description)
        Next
    End Sub

#End Region 'View Calculations Sub Tab ------------------------------------------------------------------------------------------------------------------------------------------------

#Region " View News Sub Tab" '=========================================================================================================================================================

#End Region 'View News Sub Tab --------------------------------------------------------------------------------------------------------------------------------------------------------

    Public Sub ViewTableData(ByVal IndexNo As Integer, ByVal Description As String, ByVal Query As String)
        'Set the description and Query of the Table View.
    End Sub



#End Region 'View Data Tab ------------------------------------------------------------------------------------------------------------------------------------------------------------

#Region " Calculations Tab" '==================================================================================================================================================================

    Private Sub TabControl2_GotFocus(sender As Object, e As EventArgs) Handles TabControl2.GotFocus
        'The Calculations Tab has got the focus

        If SharePriceSettingsChanged = True Then
            SaveSharePricesDataList()
        End If

        If FinancialsSettingsChanged = True Then
            SaveFinDataList()
        End If

        If CalculationsSettingsChanged = True Then
            SaveCalcDataList()
        End If

    End Sub

#Region " Copy Data Sub Tab" '=================================================================================================================================================================

    Private Sub SetUpCopyDataTab()

        dgvCopyData.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
        dgvCopyData.Columns.Clear()
        'dgvCopyData.Rows.Clear()
        txtCopyDataSettings.Text = CopyDataSettingsFile

        Dim ComboBoxCol0 As New DataGridViewComboBoxColumn
        dgvCopyData.Columns.Add(ComboBoxCol0)
        dgvCopyData.Columns(0).HeaderText = "Input Column"
        dgvCopyData.Columns(0).Width = 160

        Select Case cmbCopyDataInputDb.SelectedItem.ToString
            Case "Share Prices"
                If cmbCopyDataInputData.SelectedIndex > -1 Then
                    For Each item In SharePricesSettings.List(cmbCopyDataInputData.SelectedIndex).TableCols
                        ComboBoxCol0.Items.Add(item)
                    Next
                    txtCopyDataInputQuery.Text = SharePricesSettings.List(cmbCopyDataInputData.SelectedIndex).Query
                End If
            Case "Financials"
                If cmbCopyDataInputData.SelectedIndex > -1 Then
                    For Each item In FinancialsSettings.List(cmbCopyDataInputData.SelectedIndex).TableCols
                        ComboBoxCol0.Items.Add(item)
                    Next
                    txtCopyDataInputQuery.Text = FinancialsSettings.List(cmbCopyDataInputData.SelectedIndex).Query
                End If
            Case "Calculations"
                If cmbCopyDataInputData.SelectedIndex > -1 Then
                    For Each item In CalculationsSettings.List(cmbCopyDataInputData.SelectedIndex).TableCols
                        ComboBoxCol0.Items.Add(item)
                    Next
                    txtCopyDataInputQuery.Text = CalculationsSettings.List(cmbCopyDataInputData.SelectedIndex).Query
                End If
        End Select

        Dim ComboBoxCol1 As New DataGridViewComboBoxColumn
        dgvCopyData.Columns.Add(ComboBoxCol1)
        dgvCopyData.Columns(1).HeaderText = "Output Column"
        dgvCopyData.Columns(1).Width = 160

        If cmbCopyDataOutputDb.Items.Count = 0 Then
            Exit Sub
        End If
        If cmbCopyDataOutputData.Items.Count = 0 Then
            Exit Sub
        End If

        If cmbCopyDataOutputDb.SelectedIndex = -1 Then
            cmbCopyDataOutputDb.SelectedIndex = 0
        End If
        If cmbCopyDataOutputData.SelectedIndex = -1 Then
            cmbCopyDataOutputData.SelectedIndex = 0
        End If

        Select Case cmbCopyDataOutputDb.SelectedItem.ToString
            Case "Share Prices"
                For Each item In SharePricesSettings.List(cmbCopyDataOutputData.SelectedIndex).TableCols
                    ComboBoxCol1.Items.Add(item)
                Next
                txtCopyDataOutputQuery.Text = SharePricesSettings.List(cmbCopyDataOutputData.SelectedIndex).Query
            Case "Financials"
                For Each item In FinancialsSettings.List(cmbCopyDataOutputData.SelectedIndex).TableCols
                    ComboBoxCol1.Items.Add(item)
                Next
                txtCopyDataOutputQuery.Text = FinancialsSettings.List(cmbCopyDataOutputData.SelectedIndex).Query
            Case "Calculations"
                For Each item In CalculationsSettings.List(cmbCopyDataOutputData.SelectedIndex).TableCols
                    ComboBoxCol1.Items.Add(item)
                Next
                txtCopyDataOutputQuery.Text = CalculationsSettings.List(cmbCopyDataOutputData.SelectedIndex).Query
        End Select

        'Restore Selections in dgvCopyData
        RestoreCopyDataSelections()

    End Sub

    Private Sub RestoreCopyDataSelections()
        'Load the Calculations Settings file.
        'CopyColumnsSettingsFile contains the file name.

        If CopyDataSettingsFile = "" Then
            'Message.AddWarning("No Copy Columns settings file name has been specified!" & vbCrLf)
        Else
            Dim XDoc As System.Xml.Linq.XDocument
            Project.ReadXmlData(CopyDataSettingsFile, XDoc)

            'Don't update these selections: (Only dgvCopyData is being updated.)
            'If XDoc.<CopyColumnsSettings>.<InputDatabase>.Value <> Nothing Then cmbCopyDataInputDb.SelectedIndex = cmbCopyDataInputDb.FindStringExact(XDoc.<CopyColumnsSettings>.<InputDatabase>.Value)
            'If XDoc.<CopyColumnsSettings>.<InputData>.Value <> Nothing Then cmbCopyDataInputData.SelectedIndex = cmbCopyDataInputData.FindStringExact(XDoc.<CopyColumnsSettings>.<InputData>.Value)
            'If XDoc.<CopyColumnsSettings>.<OutputDatabase>.Value <> Nothing Then cmbCopyDataOutputDb.SelectedIndex = cmbCopyDataOutputDb.FindStringExact(XDoc.<CopyColumnsSettings>.<OutputDatabase>.Value)
            'If XDoc.<CopyColumnsSettings>.<OutputTable>.Value <> Nothing Then cmbCopyDataOutputData.SelectedIndex = cmbCopyDataOutputData.FindStringExact(XDoc.<CopyColumnsSettings>.<OutputTable>.Value)

            dgvCopyData.Rows.Clear()

            Dim settings = From item In XDoc.<CopyColumnsSettings>.<CopyList>.<CopyColumn>

            For Each item In settings
                dgvCopyData.Rows.Add(item.<From>.Value, item.<To>.Value)
            Next
        End If
    End Sub

    Private Sub dgvCopyData_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles dgvCopyData.DataError
        'This code stops an unnecessary error dialog appearing.
        If e.Context = DataGridViewDataErrorContexts.Formatting Or e.Context = DataGridViewDataErrorContexts.PreferredSize Then
            e.ThrowException = False
        End If
    End Sub

    Private Sub GetCopyDataInputDataList()
        'Fill the cmbCopyDataInputData list of data views.

        cmbCopyDataInputData.Items.Clear()
        Dim I As Integer

        Select Case cmbCopyDataInputDb.SelectedItem.ToString
            Case "Share Prices"
                'SharePriceDataViewList
                For I = 0 To SharePricesSettings.NRecords - 1
                    cmbCopyDataInputData.Items.Add(SharePricesSettings.List(I).Description)
                Next

            Case "Financials"
                'FinancialsDataViewList
                For I = 0 To FinancialsSettings.NRecords - 1
                    cmbCopyDataInputData.Items.Add(FinancialsSettings.List(I).Description)
                Next

            Case "Calculations"
                'CalculationsDataViewList
                For I = 0 To CalculationsSettings.NRecords - 1
                    cmbCopyDataInputData.Items.Add(CalculationsSettings.List(I).Description)
                Next
        End Select
    End Sub

    Private Sub cmbCopyDataInputDb_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbCopyDataInputDb.SelectedIndexChanged
        GetCopyDataInputDataList()
    End Sub

    Private Sub cmbCopyDataInputData_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbCopyDataInputData.SelectedIndexChanged
        'Open selected input data.
        SetUpCopyDataTab()
    End Sub

    Private Sub cmbCopyDataOutputDb_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbCopyDataOutputDb.SelectedIndexChanged
        GetCopyDataOutputDataList()
    End Sub

    Private Sub GetCopyDataOutputDataList()
        'Fill the cmbCopyDataOutputData list of data views.

        cmbCopyDataOutputData.Items.Clear()
        Dim I As Integer

        Select Case cmbCopyDataOutputDb.SelectedItem.ToString
            Case "Share Prices"
                'SharePriceDataViewList
                For I = 0 To SharePricesSettings.NRecords - 1
                    cmbCopyDataOutputData.Items.Add(SharePricesSettings.List(I).Description)
                Next
            Case "Financials"
                'FinancialsDataViewList
                For I = 0 To FinancialsSettings.NRecords - 1
                    cmbCopyDataOutputData.Items.Add(FinancialsSettings.List(I).Description)
                Next
            Case "Calculations"
                'CalculationsDataViewList
                For I = 0 To CalculationsSettings.NRecords - 1
                    cmbCopyDataOutputData.Items.Add(CalculationsSettings.List(I).Description)
                Next
        End Select
    End Sub

    Private Sub cmbCopyDataOutputData_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbCopyDataOutputData.SelectedIndexChanged
        'The Copy Data output table has been selected.
        SetUpCopyDataTab()
    End Sub

    Private Sub btnSaveCopyDataSettings_Click(sender As Object, e As EventArgs) Handles btnSaveCopyDataSettings.Click
        'Save the Copy Data Settings.

        If Trim(txtCopyDataSettings.Text) = "" Then
            Message.AddWarning("No file name has been sepcified!" & vbCrLf)
        Else
            If txtCopyDataSettings.Text.EndsWith(".CopyColumns") Then
                txtCopyDataSettings.Text = Trim(txtCopyDataSettings.Text)
            Else
                txtCopyDataSettings.Text = Trim(txtCopyDataSettings.Text) & ".CopyColumns"
            End If

            dgvCopyData.AllowUserToAddRows = False 'This removes the last edit row from the DataGridView.

            Dim XDoc = <?xml version="1.0" encoding="utf-8"?>
                       <CopyColumnsSettings>
                           <!---->
                           <!--Copy Columns Settings-->
                           <!---->
                           <HostApplication><%= ApplicationInfo.Name %></HostApplication>
                           <InputDatabase><%= cmbCopyDataInputDb.SelectedItem.ToString %></InputDatabase>
                           <InputData><%= cmbCopyDataInputData.SelectedItem.ToString %></InputData>
                           <OutputDatabase><%= cmbCopyDataOutputDb.SelectedItem.ToString %></OutputDatabase>
                           <OutputTable><%= cmbCopyDataOutputData.SelectedItem.ToString %></OutputTable>
                           <CopyList>
                               <%= From item In dgvCopyData.Rows
                                   Select
                                               <CopyColumn>
                                                   <From><%= item.Cells(0).Value %></From>
                                                   <To><%= item.Cells(1).Value %></To>
                                               </CopyColumn>
                               %>
                           </CopyList>
                       </CopyColumnsSettings>
            Project.SaveXmlData(txtCopyDataSettings.Text, XDoc)
            dgvCopyData.AllowUserToAddRows = True 'Allow user to add rows again.
            CopyDataSettingsFile = txtCopyDataSettings.Text
        End If

    End Sub

    Private Sub btnFindCopyDataSettings_Click(sender As Object, e As EventArgs) Handles btnFindCopyDataSettings.Click
        'Find a Copt Data Settings file.

        Select Case Project.DataLocn.Type
            Case ADVL_Utilities_Library_1.FileLocation.Types.Directory
                'Select a Copy Columns settings file from the project directory:
                OpenFileDialog1.InitialDirectory = Project.DataLocn.Path
                OpenFileDialog1.Filter = "Copy Columns settings file | *.CopyColumns"
                If OpenFileDialog1.ShowDialog() = DialogResult.OK Then
                    Dim DataFileName As String = System.IO.Path.GetFileName(OpenFileDialog1.FileName)
                    CopyDataSettingsFile = DataFileName
                    txtCopyDataSettings.Text = DataFileName
                    'LoadCopyColumnsSettingsFile()
                    LoadCopyDataSettingsFile()
                End If
            Case ADVL_Utilities_Library_1.FileLocation.Types.Archive
                'Select a Copy Columns settings file from the project archive:
                'Show the zip archive file selection form:
                Zip = New ADVL_Utilities_Library_1.ZipComp
                Zip.ArchivePath = Project.DataLocn.Path
                Zip.SelectFile() 'Show the SelectFile form.
                Zip.SelectFileForm.ApplicationName = Project.ApplicationName
                Zip.SelectFileForm.SettingsLocn = Project.SettingsLocn
                Zip.SelectFileForm.Show()
                Zip.SelectFileForm.RestoreFormSettings()
                Zip.SelectFileForm.FileExtension = ".CopyColumns"
                Zip.SelectFileForm.GetFileList()
                If Zip.SelectedFile <> "" Then
                    'A file has been selected
                    CopyDataSettingsFile = Zip.SelectedFile
                    txtCopyDataSettings.Text = Zip.SelectedFile
                    'LoadCopyColumnsSettingsFile()
                    LoadCopyDataSettingsFile()
                End If
        End Select
    End Sub

    Private Sub btnNewCopyDataSettings_Click(sender As Object, e As EventArgs) Handles btnNewCopyDataSettings.Click
        'New Copy Data Settings.

        CopyDataSettingsFile = ""
        txtCopyDataSettings.Text = ""
        'cmbCopyDataInputDb.SelectedIndex = -1
        'cmbCopyDataInputData.SelectedIndex = -1
        cmbCopyDataInputData.SelectedIndex = 0
        'cmbCopyDataOutputDb.SelectedIndex = -1
        'cmbCopyDataOutputData.SelectedIndex = -1
        cmbCopyDataOutputData.SelectedIndex = 0
        dgvCopyData.Rows.Clear()

    End Sub

    Private Sub LoadCopyDataSettingsFile()
        'Load the Calculations Settings file.
        'CopyColumnsSettingsFile contains the file name.

        If CopyDataSettingsFile = "" Then
            Message.AddWarning("No Copy Columns settings file name has been specified!" & vbCrLf)
        Else
            Dim XDoc As System.Xml.Linq.XDocument
            Project.ReadXmlData(CopyDataSettingsFile, XDoc)

            If XDoc.<CopyColumnsSettings>.<InputDatabase>.Value <> Nothing Then cmbCopyDataInputDb.SelectedIndex = cmbCopyDataInputDb.FindStringExact(XDoc.<CopyColumnsSettings>.<InputDatabase>.Value)
            If XDoc.<CopyColumnsSettings>.<InputData>.Value <> Nothing Then cmbCopyDataInputData.SelectedIndex = cmbCopyDataInputData.FindStringExact(XDoc.<CopyColumnsSettings>.<InputData>.Value)
            If XDoc.<CopyColumnsSettings>.<OutputDatabase>.Value <> Nothing Then cmbCopyDataOutputDb.SelectedIndex = cmbCopyDataOutputDb.FindStringExact(XDoc.<CopyColumnsSettings>.<OutputDatabase>.Value)
            If XDoc.<CopyColumnsSettings>.<OutputTable>.Value <> Nothing Then cmbCopyDataOutputData.SelectedIndex = cmbCopyDataOutputData.FindStringExact(XDoc.<CopyColumnsSettings>.<OutputTable>.Value)

            dgvCopyData.Rows.Clear()

            Dim settings = From item In XDoc.<CopyColumnsSettings>.<CopyList>.<CopyColumn>

            For Each item In settings
                dgvCopyData.Rows.Add(item.<From>.Value, item.<To>.Value)
            Next
        End If
    End Sub

    Private Sub btnApplyCopyDataSettings_Click(sender As Object, e As EventArgs) Handles btnApplyCopyDataSettings.Click
        'Apply the Copy Columns Simple Calculations settings.

        ApplyCopyData()

    End Sub

    Private Sub ApplyCopyData()
        'Apply the Copy Columns Simple Calculations settings.

        LoadCopyDataDsInputData()
        LoadCopyDataDsOutputData()

        'dsInput contains the input table
        'dsOutput contains the output table

        dgvCopyData.AllowUserToAddRows = False

        Dim NCols As Integer = dgvCopyData.RowCount

        Dim InCols(0 To NCols) As String 'Array to contain the Input table Column names.
        Dim OutCols(0 To NCols) As String 'Array to contain the corresponding Output table Column names.

        Dim I As Integer

        'Message for debugging:
        Message.Add("Starting Copy Data --------------------" & vbCrLf)
        For I = 1 To NCols
            InCols(I) = dgvCopyData.Rows(I - 1).Cells(0).Value
            OutCols(I) = dgvCopyData.Rows(I - 1).Cells(1).Value
            Message.Add("Copy Data item no: " & I & "  from column name: " & InCols(I) & "  to column name: " & OutCols(I) & vbCrLf)
        Next
        Message.Add("Copying Data ----------------------- " & vbCrLf)
        Dim Count As Integer = 0

        Try
            For Each item In dsInput.Tables("myData").Rows
                Count += 1
                If Count Mod 100 = 0 Then Message.Add("Copying record number: " & Count & vbCrLf)
                Dim newRow As DataRow = dsOutput.Tables("myData").NewRow
                For I = 1 To NCols
                    newRow(OutCols(I)) = item(InCols(I))
                Next
                'Message.Add("New ASX_Code: " & newRow("ASX_Code") & "  New Report_Date: " & newRow("Report_Date") & vbCrLf)
                dsOutput.Tables("myData").Rows.Add(newRow)
            Next

        Catch ex As Exception
            Message.AddWarning("Error copying data: " & ex.Message & vbCrLf)
        End Try

        Try
            Message.Add("Updating Database ------------------ " & vbCrLf)
            outputDa.Update(dsOutput.Tables("myData"))
            Message.Add("Copy Data Complete ----------------- " & vbCrLf)
        Catch ex As Exception
            Message.AddWarning("Error updating database table: " & ex.Message & vbCrLf)
        End Try

        'Clear the data list when finished:
        dsInput.Clear()
        dsInput.Tables.Clear()
        dsOutput.Clear()
        dsOutput.Tables.Clear()

        dgvCopyData.AllowUserToAddRows = False
        dgvCopyData.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
    End Sub

    Private Sub LoadCopyDataDsInputData()
        'Open selected input data in dsInput.

        dsInput.Clear()
        dsInput.Tables.Clear()
        Dim Query As String
        Dim connString As String
        Dim myConnection As OleDb.OleDbConnection = New OleDb.OleDbConnection
        Dim da As OleDb.OleDbDataAdapter

        Select Case cmbCopyDataInputDb.SelectedItem.ToString
            Case "Share Prices"
                'Query = SharePricesSettings.List(cmbCopyDataInputData.SelectedIndex).Query
                Query = txtCopyDataInputQuery.Text
                connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source =" & SharePriceDbPath
                myConnection.ConnectionString = connString
                myConnection.Open()
                da = New OleDb.OleDbDataAdapter(Query, myConnection)
                da.MissingSchemaAction = MissingSchemaAction.AddWithKey
                da.Fill(dsInput, "myData")
                myConnection.Close()
            Case "Financials"
                'Query = FinancialsSettings.List(cmbCopyDataInputData.SelectedIndex).Query
                Query = txtCopyDataInputQuery.Text
                connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source =" & FinancialsDbPath
                myConnection.ConnectionString = connString
                myConnection.Open()
                da = New OleDb.OleDbDataAdapter(Query, myConnection)
                da.MissingSchemaAction = MissingSchemaAction.AddWithKey
                da.Fill(dsInput, "myData")
                myConnection.Close()
            Case "Calculations"
                'Query = FinancialsSettings.List(cmbCopyDataInputData.SelectedIndex).Query
                Query = txtCopyDataInputQuery.Text
                connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source =" & CalculationsDbPath
                myConnection.ConnectionString = connString
                myConnection.Open()
                da = New OleDb.OleDbDataAdapter(Query, myConnection)
                da.MissingSchemaAction = MissingSchemaAction.AddWithKey
                da.Fill(dsInput, "myData")
                myConnection.Close()
        End Select
    End Sub

    Private Sub LoadCopyDataDsOutputData()
        'Open selected output data in dsOutput.

        dsOutput.Clear()
        dsOutput.Tables.Clear()

        Select Case cmbCopyDataOutputDb.SelectedItem.ToString
            Case "Share Prices"
                'outputQuery = SharePricesSettings.List(cmbCopyDataOutputData.SelectedIndex).Query
                outputQuery = txtCopyDataOutputQuery.Text
                outputConnString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source =" & SharePriceDbPath
                outputConnection.ConnectionString = outputConnString
                outputConnection.Open()
                outputDa = New OleDb.OleDbDataAdapter(outputQuery, outputConnection)
                outputDa.MissingSchemaAction = MissingSchemaAction.AddWithKey
                Dim outputCmdBuilder As New OleDb.OleDbCommandBuilder(outputDa)
                outputDa.Fill(dsOutput, "myData")
                outputConnection.Close()
            Case "Financials"
                'outputQuery = FinancialsSettings.List(cmbCopyDataOutputData.SelectedIndex).Query
                outputQuery = txtCopyDataOutputQuery.Text
                outputConnString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source =" & FinancialsDbPath
                outputConnection.ConnectionString = outputConnString
                outputConnection.Open()
                outputDa = New OleDb.OleDbDataAdapter(outputQuery, outputConnection)
                outputDa.MissingSchemaAction = MissingSchemaAction.AddWithKey
                Dim outputCmdBuilder As New OleDb.OleDbCommandBuilder(outputDa)
                outputDa.Fill(dsOutput, "myData")
                outputConnection.Close()
            Case "Calculations"
                'outputQuery = CalculationsSettings.List(cmbCopyDataOutputData.SelectedIndex).Query
                outputQuery = txtCopyDataOutputQuery.Text
                outputConnString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source =" & CalculationsDbPath
                outputConnection.ConnectionString = outputConnString
                outputConnection.Open()
                outputDa = New OleDb.OleDbDataAdapter(outputQuery, outputConnection)
                outputDa.MissingSchemaAction = MissingSchemaAction.AddWithKey
                Dim outputCmdBuilder As New OleDb.OleDbCommandBuilder(outputDa)
                outputDa.Fill(dsOutput, "myData")
                outputConnection.Close()
        End Select
    End Sub

    Private Sub btnAddCopyDataToSequence_Click(sender As Object, e As EventArgs) Handles btnAddCopyDataToSequence.Click
        'Add the Copy Data sequence to the Processing Sequence.

        If IsNothing(Sequence) Then
            Message.AddWarning("The Processing Sequence form is not open." & vbCrLf)
        Else
            'Write new instructions to the Sequence text.
            Dim Posn As Integer
            Posn = Sequence.rtbSequence.SelectionStart

            Sequence.rtbSequence.SelectedText = vbCrLf & "<!--Copy Data: Settings used to copy data from an Input table to an Output tabe :-->" & vbCrLf
            Sequence.rtbSequence.SelectedText = "  <CopyData>" & vbCrLf

            'Input data parameters:
            Sequence.rtbSequence.SelectedText = "    <InputDatabase>" & cmbCopyDataInputDb.SelectedItem.ToString & "</InputDatabase>" & vbCrLf
            Select Case cmbCopyDataInputDb.SelectedItem.ToString
                Case "Share Prices"
                    Sequence.rtbSequence.SelectedText = "    <InputDatabasePath>" & SharePriceDbPath & "</InputDatabasePath>" & vbCrLf
                Case "Financials"
                    Sequence.rtbSequence.SelectedText = "    <InputDatabasePath>" & FinancialsDbPath & "</InputDatabasePath>" & vbCrLf
                Case "Calculations"
                    Sequence.rtbSequence.SelectedText = "    <InputDatabasePath>" & CalculationsDbPath & "</InputDatabasePath>" & vbCrLf
            End Select
            Sequence.rtbSequence.SelectedText = "    <InputQuery>" & txtCopyDataInputQuery.Text & "</InputQuery>" & vbCrLf

            'Output data parameters:
            Sequence.rtbSequence.SelectedText = "    <OutputDatabase>" & cmbCopyDataOutputDb.SelectedItem.ToString & "</OutputDatabase>" & vbCrLf
            Select Case cmbCopyDataOutputDb.SelectedItem.ToString
                Case "Share Prices"
                    Sequence.rtbSequence.SelectedText = "    <OutputDatabasePath>" & SharePriceDbPath & "</OutputDatabasePath>" & vbCrLf
                Case "Financials"
                    Sequence.rtbSequence.SelectedText = "    <OutputDatabasePath>" & FinancialsDbPath & "</OutputDatabasePath>" & vbCrLf
                Case "Calculations"
                    Sequence.rtbSequence.SelectedText = "    <OutputDatabasePath>" & CalculationsDbPath & "</OutputDatabasePath>" & vbCrLf
            End Select
            Sequence.rtbSequence.SelectedText = "    <OutputQuery>" & txtCopyDataOutputQuery.Text & "</OutputQuery>" & vbCrLf

            'List of columns to copy:
            Sequence.rtbSequence.SelectedText = "    <CopyList>" & vbCrLf
            dgvCopyData.AllowUserToAddRows = False 'This removes the last edit row from the DataGridView.
            Dim NRows As Integer = dgvCopyData.Rows.Count
            Dim RowNo As Integer
            For RowNo = 0 To NRows - 1
                Sequence.rtbSequence.SelectedText = "      <CopyColumn>" & vbCrLf
                Sequence.rtbSequence.SelectedText = "        <From>" & dgvCopyData.Rows(RowNo).Cells(0).Value & "</From>" & vbCrLf
                Sequence.rtbSequence.SelectedText = "        <To>" & dgvCopyData.Rows(RowNo).Cells(1).Value & "</To>" & vbCrLf
                Sequence.rtbSequence.SelectedText = "      </CopyColumn>" & vbCrLf
            Next
            dgvCopyData.AllowUserToAddRows = True 'Allow user to add rows again.
            Sequence.rtbSequence.SelectedText = "    </CopyList>" & vbCrLf

            Sequence.rtbSequence.SelectedText = "    <Command>Apply</Command>" & vbCrLf

            Sequence.rtbSequence.SelectedText = "  </CopyData>" & vbCrLf
            Sequence.rtbSequence.SelectedText = "<!---->" & vbCrLf

            Sequence.FormatXmlText()

        End If

    End Sub

#End Region 'Copy Data Sub Tab ----------------------------------------------------------------------------------------------------------------------------------------------------------------

#Region " Select Data Sub Tab" '===============================================================================================================================================================

    Private Sub dgvSelectData_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles dgvSelectData.DataError
        'This code stops an unnecessary error dialog appearing.
        If e.Context = DataGridViewDataErrorContexts.Formatting Or e.Context = DataGridViewDataErrorContexts.PreferredSize Then
            e.ThrowException = False
        End If
    End Sub

    Private Sub dgvSelectConstraints_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles dgvSelectConstraints.DataError
        'This code stops an unnecessary error dialog appearing.
        If e.Context = DataGridViewDataErrorContexts.Formatting Or e.Context = DataGridViewDataErrorContexts.PreferredSize Then
            e.ThrowException = False
        End If
    End Sub

    Private Sub SetUpSelectDataTab()
        'Set up the Select Data tab.

        'Set up dgvSelectData ----------------------------------------------------
        dgvSelectData.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
        dgvSelectData.Columns.Clear()

        txtSelectDataSettings.Text = SelectDataSettingsFile

        Dim ComboBoxCol0 As New DataGridViewComboBoxColumn
        dgvSelectData.Columns.Add(ComboBoxCol0)
        dgvSelectData.Columns(0).HeaderText = "Input Column"

        Select Case cmbSelectDataInputDb.SelectedItem.ToString
            Case "Share Prices"
                For Each item In SharePricesSettings.List(cmbSelectDataInputData.SelectedIndex).TableCols
                    ComboBoxCol0.Items.Add(item)
                Next
                txtSelectDataInputQuery.Text = SharePricesSettings.List(cmbSelectDataInputData.SelectedIndex).Query
            Case "Financials"
                For Each item In FinancialsSettings.List(cmbSelectDataInputData.SelectedIndex).TableCols
                    ComboBoxCol0.Items.Add(item)
                Next
                txtSelectDataInputQuery.Text = FinancialsSettings.List(cmbSelectDataInputData.SelectedIndex).Query
            Case "Calculations"
                For Each item In CalculationsSettings.List(cmbSelectDataInputData.SelectedIndex).TableCols
                    ComboBoxCol0.Items.Add(item)
                Next
                txtSelectDataInputQuery.Text = CalculationsSettings.List(cmbSelectDataInputData.SelectedIndex).Query
        End Select

        Dim ComboBoxCol1 As New DataGridViewComboBoxColumn
        dgvSelectData.Columns.Add(ComboBoxCol1)
        dgvSelectData.Columns(1).HeaderText = "Output Column"
        'dgvSelectData.Columns(1).Width = 160

        If cmbSelectDataOutputDb.SelectedIndex = -1 Then
            cmbSelectDataOutputDb.SelectedIndex = 0
        End If
        If cmbSelectDataOutputData.SelectedIndex = -1 Then
            cmbSelectDataOutputData.SelectedIndex = 0
        End If

        Select Case cmbSelectDataOutputDb.SelectedItem.ToString
            Case "Share Prices"
                For Each item In SharePricesSettings.List(cmbSelectDataOutputData.SelectedIndex).TableCols
                    ComboBoxCol1.Items.Add(item)
                Next
                txtSelectDataOutputQuery.Text = SharePricesSettings.List(cmbSelectDataOutputData.SelectedIndex).Query
            Case "Financials"
                For Each item In FinancialsSettings.List(cmbSelectDataOutputData.SelectedIndex).TableCols
                    ComboBoxCol1.Items.Add(item)
                Next
                txtSelectDataOutputQuery.Text = FinancialsSettings.List(cmbSelectDataOutputData.SelectedIndex).Query
            Case "Calculations"
                For Each item In CalculationsSettings.List(cmbSelectDataOutputData.SelectedIndex).TableCols
                    ComboBoxCol1.Items.Add(item)
                Next
                txtSelectDataOutputQuery.Text = CalculationsSettings.List(cmbSelectDataOutputData.SelectedIndex).Query
        End Select

        'Set up dgvSelectConstraints ---------------------------------------------
        dgvSelectConstraints.Columns.Clear()

        Dim ComboBoxCol20 As New DataGridViewComboBoxColumn
        dgvSelectConstraints.Columns.Add(ComboBoxCol20)
        dgvSelectConstraints.Columns(0).HeaderText = "WHERE Input Column"
        'dgvSelectConstraints.Columns(0).Width = 160

        Select Case cmbSelectDataInputDb.SelectedItem.ToString
            Case "Share Prices"
                For Each item In SharePricesSettings.List(cmbSelectDataInputData.SelectedIndex).TableCols
                    ComboBoxCol20.Items.Add(item)
                Next
            Case "Financials"
                For Each item In FinancialsSettings.List(cmbSelectDataInputData.SelectedIndex).TableCols
                    ComboBoxCol20.Items.Add(item)
                Next
            Case "Calculations"
                For Each item In CalculationsSettings.List(cmbSelectDataInputData.SelectedIndex).TableCols
                    ComboBoxCol20.Items.Add(item)
                Next
        End Select


        Dim ComboBoxCol21 As New DataGridViewComboBoxColumn
        dgvSelectConstraints.Columns.Add(ComboBoxCol21)
        dgvSelectConstraints.Columns(1).HeaderText = "= Output Column"

        Select Case cmbSelectDataOutputDb.SelectedItem.ToString
            Case "Share Prices"
                For Each item In SharePricesSettings.List(cmbSelectDataOutputData.SelectedIndex).TableCols
                    ComboBoxCol21.Items.Add(item)
                Next
            Case "Financials"
                For Each item In FinancialsSettings.List(cmbSelectDataOutputData.SelectedIndex).TableCols
                    ComboBoxCol21.Items.Add(item)
                Next
            Case "Calculations"
                For Each item In CalculationsSettings.List(cmbSelectDataOutputData.SelectedIndex).TableCols
                    ComboBoxCol21.Items.Add(item)
                Next
        End Select

        'NOTE: Dont load the saved settings because different settings may be required!!!
        'LoadCopyColumnsSettingsFile()
        'LoadSelectDataSettingsFile() 

        RestoreSelectDataSelections
    End Sub

    Private Sub RestoreSelectDataSelections()
        'Restore the Select Data settings. (Leave the input and output data selections unchanged.)

        If SelectDataSettingsFile = "" Then
            'Message.AddWarning("No Select Data settings file name has been specified!" & vbCrLf)
        Else
            Dim XDoc As System.Xml.Linq.XDocument
            Project.ReadXmlData(SelectDataSettingsFile, XDoc)

            'Don't update these selections: (Only dgvCopyData is being updated.)
            'If XDoc.<CopyColumnsSettings>.<InputDatabase>.Value <> Nothing Then cmbCopyDataInputDb.SelectedIndex = cmbCopyDataInputDb.FindStringExact(XDoc.<CopyColumnsSettings>.<InputDatabase>.Value)
            'If XDoc.<CopyColumnsSettings>.<InputData>.Value <> Nothing Then cmbCopyDataInputData.SelectedIndex = cmbCopyDataInputData.FindStringExact(XDoc.<CopyColumnsSettings>.<InputData>.Value)
            'If XDoc.<CopyColumnsSettings>.<OutputDatabase>.Value <> Nothing Then cmbCopyDataOutputDb.SelectedIndex = cmbCopyDataOutputDb.FindStringExact(XDoc.<CopyColumnsSettings>.<OutputDatabase>.Value)
            'If XDoc.<CopyColumnsSettings>.<OutputTable>.Value <> Nothing Then cmbCopyDataOutputData.SelectedIndex = cmbCopyDataOutputData.FindStringExact(XDoc.<CopyColumnsSettings>.<OutputTable>.Value)

            dgvSelectData.Rows.Clear()

            Dim settings = From item In XDoc.<SelectDataSettings>.<SelectDataList>.<CopyColumn>

            For Each item In settings
                dgvSelectData.Rows.Add(item.<From>.Value, item.<To>.Value)
            Next

            dgvSelectData.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            dgvSelectData.AutoResizeColumns()

            dgvSelectConstraints.Rows.Clear()
            'dgvSelectConstraints.AutoResizeColumns()
            Dim settings2 = From item In XDoc.<SelectDataSettings>.<SelectConstraintsList>.<Constraint>
            For Each item In settings2
                dgvSelectConstraints.Rows.Add(item.<WhereInputColumn>.Value, item.<EqualsOutputColumn>.Value)
            Next
            dgvSelectConstraints.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            dgvSelectConstraints.AutoResizeColumns()
            dgvSelectData.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            dgvSelectConstraints.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

        End If
    End Sub

    Private Sub GetSelectDataInputDataList()
        'Fill the cmbSelectDataInputData list of data views.

        cmbSelectDataInputData.Items.Clear()
        Dim I As Integer

        Select Case cmbSelectDataInputDb.SelectedItem.ToString
            Case "Share Prices"
                'SharePriceDataViewList
                For I = 0 To SharePricesSettings.NRecords - 1
                    cmbSelectDataInputData.Items.Add(SharePricesSettings.List(I).Description)
                Next

            Case "Financials"
                'FinancialsDataViewList
                For I = 0 To FinancialsSettings.NRecords - 1
                    cmbSelectDataInputData.Items.Add(FinancialsSettings.List(I).Description)
                Next

            Case "Calculations"
                'CalculationsDataViewList
                For I = 0 To CalculationsSettings.NRecords - 1
                    cmbSelectDataInputData.Items.Add(CalculationsSettings.List(I).Description)
                Next
        End Select
    End Sub

    Private Sub cmbSelectDataInputDb_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbSelectDataInputDb.SelectedIndexChanged
        GetSelectDataInputDataList()
    End Sub

    Private Sub cmbSelectDataInputData_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbSelectDataInputData.SelectedIndexChanged
        'Open selected input data.
        SetUpSelectDataTab()
    End Sub

    Private Sub cmbSelectDataOutputDb_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbSelectDataOutputDb.SelectedIndexChanged
        'GetSelectDataOutputTableList()
        GetSelectDataOutputDataList()
    End Sub

    Private Sub GetSelectDataOutputDataList()
        'Fill the cmbSelectDataOutputData list.

        cmbSelectDataOutputData.Items.Clear()
        Dim I As Integer

        Select Case cmbSelectDataOutputDb.SelectedItem.ToString
            Case "Share Prices"
                'SharePriceDataViewList
                For I = 0 To SharePricesSettings.NRecords - 1
                    cmbSelectDataOutputData.Items.Add(SharePricesSettings.List(I).Description)
                Next
            Case "Financials"
                'FinancialsDataViewList
                For I = 0 To FinancialsSettings.NRecords - 1
                    cmbSelectDataOutputData.Items.Add(FinancialsSettings.List(I).Description)
                Next
            Case "Calculations"
                'CalculationsDataViewList
                For I = 0 To CalculationsSettings.NRecords - 1
                    cmbSelectDataOutputData.Items.Add(CalculationsSettings.List(I).Description)
                Next
        End Select

    End Sub

    Private Sub cmbSelectDataOutputData_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbSelectDataOutputData.SelectedIndexChanged
        'The Select Data output table has been selected.
        SetUpSelectDataTab()
    End Sub

    Private Sub btnSaveSelectDataSettings_Click(sender As Object, e As EventArgs) Handles btnSaveSelectDataSettings.Click
        'Save the Select Data Settings.

        If Trim(txtSelectDataSettings.Text) = "" Then
            Message.AddWarning("No file name has been specified!" & vbCrLf)
        Else
            If txtSelectDataSettings.Text.EndsWith(".SelectData") Then
                txtSelectDataSettings.Text = Trim(txtSelectDataSettings.Text)
            Else
                txtSelectDataSettings.Text = Trim(txtSelectDataSettings.Text) & ".SelectData"
            End If

            dgvSelectData.AllowUserToAddRows = False 'This removes the last edit row from the DataGridView.
            dgvSelectConstraints.AllowUserToAddRows = False

            Dim XDoc = <?xml version="1.0" encoding="utf-8"?>
                       <SelectDataSettings>
                           <!---->
                           <!--Select Data Settings-->
                           <!---->
                           <HostApplication><%= ApplicationInfo.Name %></HostApplication>
                           <InputDatabase><%= cmbSelectDataInputDb.SelectedItem.ToString %></InputDatabase>
                           <InputData><%= cmbSelectDataInputData.SelectedItem.ToString %></InputData>
                           <OutputDatabase><%= cmbSelectDataOutputDb.SelectedItem.ToString %></OutputDatabase>
                           <OutputTable><%= cmbSelectDataOutputData.SelectedItem.ToString %></OutputTable>
                           <SelectDataList>
                               <%= From item In dgvSelectData.Rows
                                   Select
                                               <CopyColumn>
                                                   <From><%= item.Cells(0).Value %></From>
                                                   <To><%= item.Cells(1).Value %></To>
                                               </CopyColumn>
                               %>
                           </SelectDataList>
                           <SelectConstraintsList>
                               <%= From item In dgvSelectConstraints.Rows
                                   Select
                                               <Constraint>
                                                   <WhereInputColumn><%= item.Cells(0).Value %></WhereInputColumn>
                                                   <EqualsOutputColumn><%= item.Cells(1).Value %></EqualsOutputColumn>
                                               </Constraint>
                               %>
                           </SelectConstraintsList>
                       </SelectDataSettings>
            Project.SaveXmlData(txtSelectDataSettings.Text, XDoc)
            dgvSelectData.AllowUserToAddRows = True 'Allow user to add rows again.
            dgvSelectConstraints.AllowUserToAddRows = True
            SelectDataSettingsFile = txtSelectDataSettings.Text
        End If

    End Sub

    Private Sub btnFindSelectDataSettings_Click(sender As Object, e As EventArgs) Handles btnFindSelectDataSettings.Click
        'Find a Select Data Settings file.

        Select Case Project.DataLocn.Type
            Case ADVL_Utilities_Library_1.FileLocation.Types.Directory
                'Select a Select Data settings file from the project directory:
                OpenFileDialog1.InitialDirectory = Project.DataLocn.Path
                OpenFileDialog1.Filter = "Select Data settings file | *.SelectData"
                If OpenFileDialog1.ShowDialog() = DialogResult.OK Then
                    Dim DataFileName As String = System.IO.Path.GetFileName(OpenFileDialog1.FileName)
                    SelectDataSettingsFile = DataFileName
                    txtSelectDataSettings.Text = DataFileName
                    LoadSelectDataSettingsFile()
                End If
            Case ADVL_Utilities_Library_1.FileLocation.Types.Archive
                'Select a Copy Columns settings file from the project archive:
                'Show the zip archive file selection form:
                Zip = New ADVL_Utilities_Library_1.ZipComp
                Zip.ArchivePath = Project.DataLocn.Path
                Zip.SelectFile() 'Show the SelectFile form.
                Zip.SelectFileForm.ApplicationName = Project.ApplicationName
                Zip.SelectFileForm.SettingsLocn = Project.SettingsLocn
                Zip.SelectFileForm.Show()
                Zip.SelectFileForm.RestoreFormSettings()
                Zip.SelectFileForm.FileExtension = ".SelectData"
                Zip.SelectFileForm.GetFileList()
                If Zip.SelectedFile <> "" Then
                    'A file has been selected
                    SelectDataSettingsFile = Zip.SelectedFile
                    txtSelectDataSettings.Text = Zip.SelectedFile
                    LoadSelectDataSettingsFile()
                End If
        End Select
    End Sub

    Private Sub LoadSelectDataSettingsFile()
        'Load the Select Data Settings file.
        'SelectDataSettingsFile contains the file name.

        If SelectDataSettingsFile = "" Then
            Message.AddWarning("No Select Data settings file name has been specified!" & vbCrLf)
        Else
            Dim XDoc As System.Xml.Linq.XDocument
            Project.ReadXmlData(SelectDataSettingsFile, XDoc)

            If XDoc.<SelectDataSettings>.<InputDatabase>.Value <> Nothing Then cmbSelectDataInputDb.SelectedIndex = cmbSelectDataInputDb.FindStringExact(XDoc.<SelectDataSettings>.<InputDatabase>.Value)
            If XDoc.<SelectDataSettings>.<InputData>.Value <> Nothing Then cmbSelectDataInputData.SelectedIndex = cmbSelectDataInputData.FindStringExact(XDoc.<SelectDataSettings>.<InputData>.Value)
            If XDoc.<SelectDataSettings>.<OutputDatabase>.Value <> Nothing Then cmbSelectDataOutputDb.SelectedIndex = cmbSelectDataOutputDb.FindStringExact(XDoc.<SelectDataSettings>.<OutputDatabase>.Value)
            If XDoc.<SelectDataSettings>.<OutputTable>.Value <> Nothing Then cmbSelectDataOutputData.SelectedIndex = cmbSelectDataOutputData.FindStringExact(XDoc.<SelectDataSettings>.<OutputTable>.Value)

            dgvSelectData.Rows.Clear()
            'dgvSelectData.AutoResizeColumns()
            Dim settings = From item In XDoc.<SelectDataSettings>.<SelectDataList>.<CopyColumn>
            For Each item In settings
                dgvSelectData.Rows.Add(item.<From>.Value, item.<To>.Value)
            Next
            dgvSelectData.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            dgvSelectData.AutoResizeColumns()


            dgvSelectConstraints.Rows.Clear()
            'dgvSelectConstraints.AutoResizeColumns()
            Dim settings2 = From item In XDoc.<SelectDataSettings>.<SelectConstraintsList>.<Constraint>
            For Each item In settings2
                dgvSelectConstraints.Rows.Add(item.<WhereInputColumn>.Value, item.<EqualsOutputColumn>.Value)
            Next
            dgvSelectConstraints.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            dgvSelectConstraints.AutoResizeColumns()
            dgvSelectData.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            dgvSelectConstraints.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

        End If
    End Sub

    Private Sub btnApplySelectDataSettings_Click(sender As Object, e As EventArgs) Handles btnApplySelectDataSettings.Click
        'Apply the Select Data Calculations settings.
        ApplySelectData()
    End Sub

    Private Sub ApplySelectData()
        'Apply the Select Data Calculations settings.

        LoadSelectDataDsInputData()
        LoadSelectDataDsOutputData()

        'dsInput contains the input table
        'dsOutput contains the output table

        dgvSelectData.AllowUserToAddRows = False
        dgvSelectConstraints.AllowUserToAddRows = False

        Dim NCols As Integer = dgvSelectData.RowCount

        Dim InCols(0 To NCols) As String 'Array to contain the Input table Column names.
        Dim OutCols(0 To NCols) As String 'Array to contain the corresponding Output table Column names.

        Dim I As Integer

        For I = 1 To NCols
            InCols(I) = dgvSelectData.Rows(I - 1).Cells(0).Value 'The Input data column names to copy.
            OutCols(I) = dgvSelectData.Rows(I - 1).Cells(1).Value 'The Output table column names to paste.
        Next

        Dim NCons As Integer = dgvSelectConstraints.RowCount
        Dim InCons(0 To NCons) As String 'The Input data contraint column names.
        Dim OutCons(0 To NCons) As String 'The Output table constraint column names.

        For I = 1 To NCons
            InCons(I) = dgvSelectConstraints.Rows(I - 1).Cells(0).Value
            OutCons(I) = dgvSelectConstraints.Rows(I - 1).Cells(1).Value
        Next

        Dim myQuery As String
        Message.Add("Selecting Data ----------------------- " & vbCrLf)
        Dim Count As Integer = 0
        For Each item In dsOutput.Tables("myData").Rows
            Count += 1
            If Count Mod 100 = 0 Then Message.Add("Selecting record number: " & Count & vbCrLf)
            myQuery = InCons(1) & " = '" & item(OutCons(1)) & "'"
            For I = 2 To NCons
                myQuery = myQuery & " And " & InCons(I) & " = '" & item(OutCons(I)) & "'"
            Next

            Dim myRecords = dsInput.Tables("myData").Select(myQuery)
            If myRecords.Count = 0 Then
                Message.AddWarning("No records found with this constraint: " & myQuery & vbCrLf)
            ElseIf myRecords.Count = 1 Then
                'Message.Add("One record found with this constraint: " & myQuery & vbCrLf)
                For I = 1 To NCols
                    item(OutCols(I)) = myRecords(0).Item(InCols(I))
                Next
            Else
                Message.AddWarning("More than one record found with this constraint: " & myQuery & vbCrLf)
            End If
        Next

        Try
            Message.Add("Updating Database ------------------ " & vbCrLf)
            outputDa.Update(dsOutput.Tables("myData"))
            Message.Add("Select Data Complete ----------------- " & vbCrLf)
        Catch ex As Exception
            Message.AddWarning("Error updating database table: " & ex.Message & vbCrLf)
        End Try

        'Clear the data list when finished:
        dsInput.Clear()
        dsInput.Tables.Clear()
        dsOutput.Clear()
        dsOutput.Tables.Clear()

        dgvSelectData.AllowUserToAddRows = True
        dgvSelectConstraints.AllowUserToAddRows = True
        dgvSelectData.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
        dgvSelectConstraints.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
    End Sub

    Private Sub LoadSelectDataDsInputData()
        'Open selected input data in dsInput.

        dsInput.Clear()
        dsInput.Tables.Clear()
        Dim Query As String
        Dim connString As String
        Dim myConnection As OleDb.OleDbConnection = New OleDb.OleDbConnection
        Dim da As OleDb.OleDbDataAdapter

        Select Case cmbSelectDataInputDb.SelectedItem.ToString
            Case "Share Prices"
                'Query = SharePricesSettings.List(cmbSelectDataInputData.SelectedIndex).Query
                Query = txtSelectDataInputQuery.Text
                connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source =" & SharePriceDbPath
                myConnection.ConnectionString = connString
                myConnection.Open()
                da = New OleDb.OleDbDataAdapter(Query, myConnection)
                da.MissingSchemaAction = MissingSchemaAction.AddWithKey
                da.Fill(dsInput, "myData")
                myConnection.Close()
            Case "Financials"
                'Query = FinancialsSettings.List(cmbSelectDataInputData.SelectedIndex).Query
                Query = txtSelectDataInputQuery.Text
                connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source =" & FinancialsDbPath
                myConnection.ConnectionString = connString
                myConnection.Open()
                da = New OleDb.OleDbDataAdapter(Query, myConnection)
                da.MissingSchemaAction = MissingSchemaAction.AddWithKey
                da.Fill(dsInput, "myData")
                myConnection.Close()
            Case "Calculations"
                'Query = CalculationsSettings.List(cmbSelectDataInputData.SelectedIndex).Query
                Query = txtSelectDataInputQuery.Text
                connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source =" & CalculationsDbPath
                myConnection.ConnectionString = connString
                myConnection.Open()
                da = New OleDb.OleDbDataAdapter(Query, myConnection)
                da.MissingSchemaAction = MissingSchemaAction.AddWithKey
                da.Fill(dsInput, "myData")
                myConnection.Close()
        End Select
    End Sub

    Private Sub LoadSelectDataDsOutputData()
        'Open selected output data in dsOutput.

        dsOutput.Clear()
        dsOutput.Tables.Clear()

        Select Case cmbSelectDataOutputDb.SelectedItem.ToString
            Case "Share Prices"
                'outputQuery = SharePricesSettings.List(cmbSelectDataOutputData.SelectedIndex).Query
                outputQuery = txtSelectDataOutputQuery.Text
                outputConnString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source =" & SharePriceDbPath
                outputConnection.ConnectionString = outputConnString
                outputConnection.Open()
                outputDa = New OleDb.OleDbDataAdapter(outputQuery, outputConnection)
                outputDa.MissingSchemaAction = MissingSchemaAction.AddWithKey
                Dim outputCmdBuilder As New OleDb.OleDbCommandBuilder(outputDa)
                outputDa.Fill(dsOutput, "myData")
                outputConnection.Close()
            Case "Financials"
                'outputQuery = FinancialsSettings.List(cmbSelectDataOutputData.SelectedIndex).Query
                outputQuery = txtSelectDataOutputQuery.Text
                outputConnString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source =" & FinancialsDbPath
                outputConnection.ConnectionString = outputConnString
                outputConnection.Open()
                outputDa = New OleDb.OleDbDataAdapter(outputQuery, outputConnection)
                outputDa.MissingSchemaAction = MissingSchemaAction.AddWithKey
                Dim outputCmdBuilder As New OleDb.OleDbCommandBuilder(outputDa)
                outputDa.Fill(dsOutput, "myData")
                outputConnection.Close()
            Case "Calculations"
                'outputQuery = CalculationsSettings.List(cmbSelectDataOutputData.SelectedIndex).Query
                outputQuery = txtSelectDataOutputQuery.Text
                outputConnString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source =" & CalculationsDbPath
                outputConnection.ConnectionString = outputConnString
                outputConnection.Open()
                outputDa = New OleDb.OleDbDataAdapter(outputQuery, outputConnection)
                outputDa.MissingSchemaAction = MissingSchemaAction.AddWithKey
                Dim outputCmdBuilder As New OleDb.OleDbCommandBuilder(outputDa)
                outputDa.Fill(dsOutput, "myData")
                outputConnection.Close()
        End Select
    End Sub

    Private Sub btnShowInputDataSet_Click(sender As Object, e As EventArgs) Handles btnShowInputDataSet.Click
        dgvDataSet.DataSource = dsInput.Tables("myData")
    End Sub

    Private Sub btnShowOutputDataSet_Click(sender As Object, e As EventArgs) Handles btnShowOutputDataSet.Click
        dgvDataSet.DataSource = dsOutput.Tables("myData")
    End Sub

    Private Sub btnShowSelectedInput_Click(sender As Object, e As EventArgs) Handles btnShowSelectedInput.Click
        'Select data from the Input DataSet and show it.

        Try
            Dim SelectedInput = dsInput.Tables("myData").Select(txtSelectInput.Text)
            'dgvDataSet.DataSource = SelectedInput(0)(0)

            txtCount.Text = SelectedInput.Count
        Catch ex As Exception
            Message.AddWarning(ex.Message & vbCrLf)
        End Try

    End Sub

    Private Sub btnAddSelectDataToSequence_Click(sender As Object, e As EventArgs) Handles btnAddSelectDataToSequence.Click
        'Add the Select Data sequence to the Processing Sequence.

        If IsNothing(Sequence) Then
            Message.AddWarning("The Processing Sequence form is not open." & vbCrLf)
        Else
            'Write new instructions to the Sequence text.
            Dim Posn As Integer
            Posn = Sequence.rtbSequence.SelectionStart

            Sequence.rtbSequence.SelectedText = vbCrLf & "<!--Select Data: Settings used to select data from an Input table and copy it to an Output table :-->" & vbCrLf
            Sequence.rtbSequence.SelectedText = "  <SelectData>" & vbCrLf

            'Input data parameters:
            Sequence.rtbSequence.SelectedText = "    <InputDatabase>" & cmbSelectDataInputDb.SelectedItem.ToString & "</InputDatabase>" & vbCrLf
            Select Case cmbSelectDataInputDb.SelectedItem.ToString
                Case "Share Prices"
                    Sequence.rtbSequence.SelectedText = "    <InputDatabasePath>" & SharePriceDbPath & "</InputDatabasePath>" & vbCrLf
                Case "Financials"
                    Sequence.rtbSequence.SelectedText = "    <InputDatabasePath>" & FinancialsDbPath & "</InputDatabasePath>" & vbCrLf
                Case "Calculations"
                    Sequence.rtbSequence.SelectedText = "    <InputDatabasePath>" & CalculationsDbPath & "</InputDatabasePath>" & vbCrLf
            End Select
            Sequence.rtbSequence.SelectedText = "    <InputQuery>" & txtSelectDataInputQuery.Text & "</InputQuery>" & vbCrLf

            'Output data parameters:
            Sequence.rtbSequence.SelectedText = "    <OutputDatabase>" & cmbSelectDataOutputDb.SelectedItem.ToString & "</OutputDatabase>" & vbCrLf
            Select Case cmbSelectDataOutputDb.SelectedItem.ToString
                Case "Share Prices"
                    Sequence.rtbSequence.SelectedText = "    <OutputDatabasePath>" & SharePriceDbPath & "</OutputDatabasePath>" & vbCrLf
                Case "Financials"
                    Sequence.rtbSequence.SelectedText = "    <OutputDatabasePath>" & FinancialsDbPath & "</OutputDatabasePath>" & vbCrLf
                Case "Calculations"
                    Sequence.rtbSequence.SelectedText = "    <OutputDatabasePath>" & CalculationsDbPath & "</OutputDatabasePath>" & vbCrLf
            End Select
            Sequence.rtbSequence.SelectedText = "    <OutputQuery>" & txtSelectDataOutputQuery.Text & "</OutputQuery>" & vbCrLf

            'Select constraints list:
            Sequence.rtbSequence.SelectedText = "    <SelectConstraintList>" & vbCrLf
            dgvSelectConstraints.AllowUserToAddRows = False 'This removes the last edit row from the DataGridView.
            Dim NRows As Integer = dgvSelectConstraints.Rows.Count
            Dim RowNo As Integer
            For RowNo = 0 To NRows - 1
                Sequence.rtbSequence.SelectedText = "      <Constraint>" & vbCrLf
                Sequence.rtbSequence.SelectedText = "        <WhereInputColumn>" & dgvSelectConstraints.Rows(RowNo).Cells(0).Value & "</WhereInputColumn>" & vbCrLf
                Sequence.rtbSequence.SelectedText = "        <EqualsOutputColumn>" & dgvSelectConstraints.Rows(RowNo).Cells(1).Value & "</EqualsOutputColumn>" & vbCrLf
                Sequence.rtbSequence.SelectedText = "      </Constraint>" & vbCrLf
            Next
            dgvSelectConstraints.AllowUserToAddRows = True 'Allow user to add rows again.
            Sequence.rtbSequence.SelectedText = "    </SelectConstraintList>" & vbCrLf

            'List of columns to copy:
            Sequence.rtbSequence.SelectedText = "    <SelectDataList>" & vbCrLf
            dgvSelectData.AllowUserToAddRows = False 'This removes the last edit row from the DataGridView.
            'Dim NRows As Integer = dgvSelectData.Rows.Count
            NRows = dgvSelectData.Rows.Count
            'Dim RowNo As Integer
            For RowNo = 0 To NRows - 1
                Sequence.rtbSequence.SelectedText = "      <CopyColumn>" & vbCrLf
                Sequence.rtbSequence.SelectedText = "        <From>" & dgvSelectData.Rows(RowNo).Cells(0).Value & "</From>" & vbCrLf
                Sequence.rtbSequence.SelectedText = "        <To>" & dgvSelectData.Rows(RowNo).Cells(1).Value & "</To>" & vbCrLf
                Sequence.rtbSequence.SelectedText = "      </CopyColumn>" & vbCrLf
            Next
            dgvSelectData.AllowUserToAddRows = True 'Allow user to add rows again.
            Sequence.rtbSequence.SelectedText = "    </SelectDataList>" & vbCrLf

            Sequence.rtbSequence.SelectedText = "    <Command>Apply</Command>" & vbCrLf
            Sequence.rtbSequence.SelectedText = "  </SelectData>" & vbCrLf
            Sequence.rtbSequence.SelectedText = "<!---->" & vbCrLf

            Sequence.FormatXmlText()

        End If

    End Sub

#End Region 'Select Data Sub Tab --------------------------------------------------------------------------------------------------------------------------------------------------------------

#Region " Simple Calculations Sub Tab" '=======================================================================================================================================================

    Private Sub dgvSimpleCalcsParameterList_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles dgvSimpleCalcsParameterList.DataError
        'This code stops an unnecessary error dialog appearing.
        If e.Context = DataGridViewDataErrorContexts.Formatting Or e.Context = DataGridViewDataErrorContexts.PreferredSize Then
            e.ThrowException = False
        End If
    End Sub

    Private Sub dgvSimpleCalcsInputData_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles dgvSimpleCalcsCalculations.DataError
        'This code stops an unnecessary error dialog appearing.
        If e.Context = DataGridViewDataErrorContexts.Formatting Or e.Context = DataGridViewDataErrorContexts.PreferredSize Then
            e.ThrowException = False
        End If
    End Sub

    Private Sub dgvSimpleCalcsCalculations_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles dgvSimpleCalcsInputData.DataError
        'This code stops an unnecessary error dialog appearing.
        If e.Context = DataGridViewDataErrorContexts.Formatting Or e.Context = DataGridViewDataErrorContexts.PreferredSize Then
            e.ThrowException = False
        End If
    End Sub

    Private Sub dgvSimpleCalcsOutputData_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles dgvSimpleCalcsOutputData.DataError
        'This code stops an unnecessary error dialog appearing.
        If e.Context = DataGridViewDataErrorContexts.Formatting Or e.Context = DataGridViewDataErrorContexts.PreferredSize Then
            e.ThrowException = False
        End If
    End Sub

    Private Sub SetUpSimpleCalculationsTab()

        Debug.Print("Running SetUpSimpleCalculationsTab()")
        Debug.Print("dgvSimpleCalcsParameterList.Rows.Count = " & dgvSimpleCalcsParameterList.Rows.Count)
        txtSimpleCalcSettings.Text = SimpleCalcsSettingsFile

        dgvSimpleCalcsParameterList.AllowUserToAddRows = False 'This removes the last edit row from the DataGridView.
        Debug.Print("dgvSimpleCalcsParameterList.AllowUserToAddRows = False")
        Debug.Print("dgvSimpleCalcsParameterList.Rows.Count = " & dgvSimpleCalcsParameterList.Rows.Count)
        dgvSimpleCalcsCalculations.AllowUserToAddRows = False
        dgvSimpleCalcsInputData.AllowUserToAddRows = False
        dgvSimpleCalcsOutputData.AllowUserToAddRows = False


        'Set up dgvSimpleCalcsParameterList ----------------------------------------------------
        dgvSimpleCalcsParameterList.Columns.Clear()

        Dim TextBoxCol0 As New DataGridViewTextBoxColumn
        dgvSimpleCalcsParameterList.Columns.Add(TextBoxCol0)
        dgvSimpleCalcsParameterList.Columns(0).HeaderText = "Name"
        dgvSimpleCalcsParameterList.Columns(0).Width = 140

        Dim ComboBoxCol1 As New DataGridViewComboBoxColumn
        dgvSimpleCalcsParameterList.Columns.Add(ComboBoxCol1)
        dgvSimpleCalcsParameterList.Columns(1).HeaderText = "Type"
        dgvSimpleCalcsParameterList.Columns(1).Width = 70
        ComboBoxCol1.Items.Add("Variable")
        ComboBoxCol1.Items.Add("Constant")
        'ComboBoxCol1.Items.Add("Date") 'NOTE: This modification was tried but won't work becauser the Param dictionary can store only single values and noty dates!!!

        Dim TextBoxCol2 As New DataGridViewTextBoxColumn
        dgvSimpleCalcsParameterList.Columns.Add(TextBoxCol2)
        dgvSimpleCalcsParameterList.Columns(2).HeaderText = "Value"
        dgvSimpleCalcsParameterList.Columns(2).Width = 60

        Dim TextBoxCol3 As New DataGridViewTextBoxColumn
        dgvSimpleCalcsParameterList.Columns.Add(TextBoxCol3)
        dgvSimpleCalcsParameterList.Columns(3).HeaderText = "Description"
        dgvSimpleCalcsParameterList.Columns(3).Width = 260

        'Set up dgvSimpleCalcsInputData ---------------------------------------------
        dgvSimpleCalcsInputData.Columns.Clear()

        Dim ComboBoxCol20 As New DataGridViewComboBoxColumn
        dgvSimpleCalcsInputData.Columns.Add(ComboBoxCol20)
        dgvSimpleCalcsInputData.Columns(0).HeaderText = "Input Parameter"
        dgvSimpleCalcsInputData.Columns(0).Width = 140

        'Add the list of Parameters
        For Each item In dgvSimpleCalcsParameterList.Rows
            ComboBoxCol20.Items.Add(item.Cells(0).Value)
            Debug.Print(item.Cells(0).Value)
        Next

        Dim ComboBoxCol21 As New DataGridViewComboBoxColumn
        dgvSimpleCalcsInputData.Columns.Add(ComboBoxCol21)
        dgvSimpleCalcsInputData.Columns(1).HeaderText = "Input Column"
        dgvSimpleCalcsInputData.Columns(1).Width = 140

        'If dsOutput.Tables.Count > 0 Then
        '    'For Each item In dsInput.Tables("myData").Columns
        '    For Each item In dsOutput.Tables("myData").Columns
        '        ComboBoxCol21.Items.Add(item.Columnname)
        '    Next
        'End If

        If cmbSimpleCalcDb.SelectedIndex = -1 Then
            cmbSimpleCalcDb.SelectedIndex = 0
        End If
        If cmbSimpleCalcData.SelectedIndex = -1 Then
            cmbSimpleCalcData.SelectedIndex = 0
        End If

        Select Case cmbSimpleCalcDb.SelectedItem.ToString
            Case "Share Prices"
                For Each item In SharePricesSettings.List(cmbSimpleCalcData.SelectedIndex).TableCols
                    ComboBoxCol21.Items.Add(item)
                Next
                txtSimpleCalcsQuery.Text = SharePricesSettings.List(cmbSimpleCalcData.SelectedIndex).Query
            Case "Financials"
                For Each item In FinancialsSettings.List(cmbSimpleCalcData.SelectedIndex).TableCols
                    ComboBoxCol21.Items.Add(item)
                Next
                txtSimpleCalcsQuery.Text = FinancialsSettings.List(cmbSimpleCalcData.SelectedIndex).Query
            Case "Calculations"
                For Each item In CalculationsSettings.List(cmbSimpleCalcData.SelectedIndex).TableCols
                    ComboBoxCol21.Items.Add(item)
                Next
                txtSimpleCalcsQuery.Text = CalculationsSettings.List(cmbSimpleCalcData.SelectedIndex).Query
        End Select


        'Set up dgvSimpleCalcsCalculations ---------------------------------------------
        dgvSimpleCalcsCalculations.Columns.Clear()

        Dim ComboBoxCol30 As New DataGridViewComboBoxColumn
        dgvSimpleCalcsCalculations.Columns.Add(ComboBoxCol30)
        dgvSimpleCalcsCalculations.Columns(0).HeaderText = "Input 1"
        dgvSimpleCalcsCalculations.Columns(0).Width = 140

        'Add the list of Parameters
        For Each item In dgvSimpleCalcsParameterList.Rows
            If item.Cells(0).Value <> Nothing Then ComboBoxCol30.Items.Add(item.Cells(0).Value)
            'Debug.Print(item.Cells(0).Value)
        Next

        Dim ComboBoxCol31 As New DataGridViewComboBoxColumn
        dgvSimpleCalcsCalculations.Columns.Add(ComboBoxCol31)
        dgvSimpleCalcsCalculations.Columns(1).HeaderText = "Input 2"
        dgvSimpleCalcsCalculations.Columns(1).Width = 140

        'Add the list of Parameters
        For Each item In dgvSimpleCalcsParameterList.Rows
            If item.Cells(0).Value <> Nothing Then ComboBoxCol31.Items.Add(item.Cells(0).Value)
        Next

        Dim ComboBoxCol32 As New DataGridViewComboBoxColumn
        dgvSimpleCalcsCalculations.Columns.Add(ComboBoxCol32)
        dgvSimpleCalcsCalculations.Columns(2).HeaderText = "Operation"
        dgvSimpleCalcsCalculations.Columns(2).Width = 110

        ComboBoxCol32.Items.Add("Input 1 + Input 2")
        ComboBoxCol32.Items.Add("Input 1 - Input 2")
        ComboBoxCol32.Items.Add("Input 1 x Input 2")
        ComboBoxCol32.Items.Add("Input 1 / Input 2")

        Dim ComboBoxCol33 As New DataGridViewComboBoxColumn
        dgvSimpleCalcsCalculations.Columns.Add(ComboBoxCol33)
        dgvSimpleCalcsCalculations.Columns(3).HeaderText = "Output"
        dgvSimpleCalcsCalculations.Columns(3).Width = 140

        'Add the list of Parameters
        For Each item In dgvSimpleCalcsParameterList.Rows
            If item.Cells(0).Value <> Nothing Then ComboBoxCol33.Items.Add(item.Cells(0).Value)
        Next

        'Set up dgvSimpleCalcsOutputData ---------------------------------------------
        dgvSimpleCalcsOutputData.Columns.Clear()

        Dim ComboBoxCol40 As New DataGridViewComboBoxColumn
        dgvSimpleCalcsOutputData.Columns.Add(ComboBoxCol40)
        dgvSimpleCalcsOutputData.Columns(0).HeaderText = "Output Parameter"
        dgvSimpleCalcsOutputData.Columns(0).Width = 140

        'Add the list of Parameters
        For Each item In dgvSimpleCalcsParameterList.Rows
            If item.Cells(0).Value <> Nothing Then ComboBoxCol40.Items.Add(item.Cells(0).Value)
        Next

        Dim ComboBoxCol41 As New DataGridViewComboBoxColumn
        dgvSimpleCalcsOutputData.Columns.Add(ComboBoxCol41)
        dgvSimpleCalcsOutputData.Columns(1).HeaderText = "Output Column"
        dgvSimpleCalcsOutputData.Columns(1).Width = 140

        'If dsOutput.Tables.Count > 0 Then
        '    'For Each item In dsInput.Tables("myData").Columns
        '    For Each item In dsOutput.Tables("myData").Columns
        '        ComboBoxCol41.Items.Add(item.Columnname)
        '    Next
        'End If

        Select Case cmbSimpleCalcDb.SelectedItem.ToString
            Case "Share Prices"
                For Each item In SharePricesSettings.List(cmbSimpleCalcData.SelectedIndex).TableCols
                    ComboBoxCol41.Items.Add(item)
                Next
            Case "Financials"
                For Each item In FinancialsSettings.List(cmbSimpleCalcData.SelectedIndex).TableCols
                    ComboBoxCol41.Items.Add(item)
                Next
            Case "Calculations"
                For Each item In CalculationsSettings.List(cmbSimpleCalcData.SelectedIndex).TableCols
                    ComboBoxCol41.Items.Add(item)
                Next
        End Select

        dgvSimpleCalcsParameterList.AllowUserToAddRows = True 'Allow user to add rows again.
        dgvSimpleCalcsCalculations.AllowUserToAddRows = True
        dgvSimpleCalcsInputData.AllowUserToAddRows = True
        dgvSimpleCalcsOutputData.AllowUserToAddRows = True

        'dgvSimpleCalcsParameterList.AllowUserToResizeColumns = True
        'dgvSimpleCalcsCalculations.AllowUserToResizeColumns = True
        'dgvSimpleCalcsInputData.AllowUserToResizeColumns = True
        'dgvSimpleCalcsOutputData.AllowUserToResizeColumns = True

        dgvSimpleCalcsParameterList.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
        dgvSimpleCalcsCalculations.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
        dgvSimpleCalcsInputData.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
        dgvSimpleCalcsOutputData.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

        'LoadCopyColumnsSettingsFile()
        'LoadSelectDataSettingsFile()
        'LoadSimpleCalcsSettingsFile()

        RestoreSimpleCalcsSelections()
    End Sub

    Private Sub RestoreSimpleCalcsSelections()
        'Restore the Simple Calculations settings from the settings file. (Leave the input and output data selections unchanged.)
        'SimpleCalcsSettingsFile contains the file name.

        If SimpleCalcsSettingsFile = "" Then
            Message.AddWarning("No Simple Calculations settings file name has been specified!" & vbCrLf)
        Else
            Dim XDoc As System.Xml.Linq.XDocument
            Project.ReadXmlData(SimpleCalcsSettingsFile, XDoc)

            'Don't update these selections: (Only dgvSimpleCalcsParameterList, dgvSimpleCalcsInputData, dgvSimpleCalcsCalculations and dgvSimpleCalcsOutputData are being updated.)
            'If XDoc.<SimpleCalculationsSettings>.<SelectedDatabase>.Value <> Nothing Then cmbSimpleCalcDb.SelectedIndex = cmbSimpleCalcDb.FindStringExact(XDoc.<SimpleCalculationsSettings>.<SelectedDatabase>.Value)
            'If XDoc.<SimpleCalculationsSettings>.<SelectedTable>.Value <> Nothing Then cmbSimpleCalcData.SelectedIndex = cmbSimpleCalcData.FindStringExact(XDoc.<SimpleCalculationsSettings>.<SelectedTable>.Value)

            dgvSimpleCalcsParameterList.Rows.Clear()
            Dim settings = From item In XDoc.<SimpleCalculationsSettings>.<ParameterList>.<Parameter>
            For Each item In settings
                dgvSimpleCalcsParameterList.Rows.Add(item.<Name>.Value, item.<Type>.Value, item.<Value>.Value, item.<Description>.Value)
            Next
            dgvSimpleCalcsParameterList.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            dgvSimpleCalcsParameterList.AutoResizeColumns()

            UpdateSimpleCalcsParams()

            dgvSimpleCalcsInputData.Rows.Clear()
            Dim settings2 = From item In XDoc.<SimpleCalculationsSettings>.<InputDataList>.<InputData>
            For Each item In settings2
                dgvSimpleCalcsInputData.Rows.Add(item.<Parameter>.Value, item.<Column>.Value)
            Next
            dgvSimpleCalcsInputData.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            dgvSimpleCalcsInputData.AutoResizeColumns()

            dgvSimpleCalcsCalculations.Rows.Clear()
            Dim settings3 = From item In XDoc.<SimpleCalculationsSettings>.<CalculationList>.<Calculation>
            For Each item In settings3
                dgvSimpleCalcsCalculations.Rows.Add(item.<Input1>.Value, item.<Input2>.Value, item.<Operation>.Value, item.<Output>.Value)
            Next
            dgvSimpleCalcsCalculations.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            dgvSimpleCalcsCalculations.AutoResizeColumns()

            dgvSimpleCalcsOutputData.Rows.Clear()
            Dim settings4 = From item In XDoc.<SimpleCalculationsSettings>.<OutputDataList>.<OutputData>
            For Each item In settings4
                dgvSimpleCalcsOutputData.Rows.Add(item.<Parameter>.Value, item.<Column>.Value)
            Next
            dgvSimpleCalcsOutputData.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            dgvSimpleCalcsOutputData.AutoResizeColumns()

            'dgvSimpleCalcsParameterList.AllowUserToResizeColumns = True
            'dgvSimpleCalcsCalculations.AllowUserToResizeColumns = True
            'dgvSimpleCalcsInputData.AllowUserToResizeColumns = True
            'dgvSimpleCalcsOutputData.AllowUserToResizeColumns = True

            dgvSimpleCalcsParameterList.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            dgvSimpleCalcsCalculations.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            dgvSimpleCalcsInputData.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            dgvSimpleCalcsOutputData.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

        End If
    End Sub

    Private Sub cmbSimpleCalcDb_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbSimpleCalcDb.SelectedIndexChanged
        'GetSimpleSelectedTableList()
        GetSimpleCalcsOutputDataList()
    End Sub

    Private Sub cmbSimpleCalcData_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbSimpleCalcData.SelectedIndexChanged
        SetUpSimpleCalculationsTab()
    End Sub

    Private Sub GetSimpleCalcsOutputDataList()
        'Fill the cmbSimpleCalcData list.

        cmbSimpleCalcData.Items.Clear()
        Dim I As Integer

        Select Case cmbSimpleCalcDb.SelectedItem.ToString
            Case "Share Prices"
                'SharePriceDataViewList
                For I = 0 To SharePricesSettings.NRecords - 1
                    cmbSimpleCalcData.Items.Add(SharePricesSettings.List(I).Description)
                Next
            Case "Financials"
                'FinancialsDataViewList
                For I = 0 To FinancialsSettings.NRecords - 1
                    cmbSimpleCalcData.Items.Add(FinancialsSettings.List(I).Description)
                Next
            Case "Calculations"
                'CalculationsDataViewList
                For I = 0 To CalculationsSettings.NRecords - 1
                    cmbSimpleCalcData.Items.Add(CalculationsSettings.List(I).Description)
                Next
        End Select

    End Sub

    'Private Sub GetSimpleSelectedTableList()
    '    'Fill the cmbSimpleCalcTable list of data views.

    '    'Database access for MS Access:
    '    Dim connectionString As String 'Declare a connection string for MS Access - defines the database or server to be used.
    '    Dim conn As System.Data.OleDb.OleDbConnection 'Declare a connection for MS Access - used by the Data Adapter to connect to and disconnect from the database.
    '    Dim dt As DataTable

    '    cmbSimpleCalcData.Items.Clear()
    '    Dim I As Integer

    '    Select Case cmbSimpleCalcDb.SelectedItem.ToString
    '        Case "Share Prices"
    '            ''SharePriceDataViewList
    '            'For I = 0 To SharePricesSettings.NRecords - 1
    '            '    'cmbSelectDataInputData.Items.Add(SharePricesSettings.List(I).Description)
    '            '    cmbSelectDataOutputTable.Items.Add(SharePricesSettings.List(I).Description)
    '            'Next
    '            If SharePriceDbPath = "" Then
    '                Message.AddWarning("No Share Prices database has been selected!" & vbCrLf)
    '                Exit Sub
    '            End If
    '            'Access 2007+:
    '            connectionString = "provider=Microsoft.ACE.OLEDB.12.0;" +
    '             "data source = " + SharePriceDbPath
    '        Case "Financials"
    '            ''FinancialsDataViewList
    '            'For I = 0 To FinancialsSettings.NRecords - 1
    '            '    'cmbCopyDataInputData.Items.Add(FinancialsSettings.List(I).Description)
    '            '    cmbSelectDataOutputTable.Items.Add(FinancialsSettings.List(I).Description)
    '            'Next
    '            If FinancialsDbPath = "" Then
    '                Message.AddWarning("No Financials database has been selected!" & vbCrLf)
    '                Exit Sub
    '            End If
    '            'Access 2007+:
    '            connectionString = "provider=Microsoft.ACE.OLEDB.12.0;" +
    '             "data source = " + FinancialsDbPath
    '        Case "Calculations"
    '            If CalculationsDbPath = "" Then
    '                Message.AddWarning("No Calculations database has been selected!" & vbCrLf)
    '                Exit Sub
    '            End If
    '            'Access 2007+:
    '            connectionString = "provider=Microsoft.ACE.OLEDB.12.0;" +
    '             "data source = " + CalculationsDbPath
    '    End Select

    '    'Connect to the Access database:
    '    conn = New System.Data.OleDb.OleDbConnection(connectionString)

    '    Try
    '        conn.Open()
    '        Dim restrictions As String() = New String() {Nothing, Nothing, Nothing, "TABLE"} 'This restriction removes system tables
    '        dt = conn.GetSchema("Tables", restrictions)

    '        Dim dr As DataRow
    '        'Dim I As Integer 'Loop index
    '        Dim MaxI As Integer
    '        MaxI = dt.Rows.Count
    '        For I = 0 To MaxI - 1
    '            dr = dt.Rows(0)
    '            cmbSimpleCalcData.Items.Add(dt.Rows(I).Item(2).ToString)
    '        Next I

    '        conn.Close()

    '    Catch ex As Exception
    '        Message.AddWarning("Error reading list of tables in the selected database. " & ex.Message & vbCrLf)
    '    End Try
    'End Sub

    Private Sub btnSaveSimpleCalcSettings_Click(sender As Object, e As EventArgs) Handles btnSaveSimpleCalcSettings.Click
        'Save the Simple Calculations settings:

        If Trim(txtSimpleCalcSettings.Text) = "" Then
            Message.AddWarning("No file name has been specified!" & vbCrLf)
        Else
            If txtSimpleCalcSettings.Text.EndsWith(".SimpleCalcs") Then
                txtSimpleCalcSettings.Text = Trim(txtSimpleCalcSettings.Text)
            Else
                txtSimpleCalcSettings.Text = Trim(txtSimpleCalcSettings.Text) & ".SimpleCalcs"
            End If

            dgvSimpleCalcsParameterList.AllowUserToAddRows = False 'This removes the last edit row from the DataGridView.
            dgvSimpleCalcsCalculations.AllowUserToAddRows = False
            dgvSimpleCalcsInputData.AllowUserToAddRows = False
            dgvSimpleCalcsOutputData.AllowUserToAddRows = False

            Dim XDoc = <?xml version="1.0" encoding="utf-8"?>
                       <SimpleCalculationsSettings>
                           <!---->
                           <!--Simple Calculations Settings-->
                           <HostApplication><%= ApplicationInfo.Name %></HostApplication>
                           <SelectedDatabase><%= cmbSimpleCalcDb.SelectedItem.ToString %></SelectedDatabase>
                           <SelectedData><%= cmbSimpleCalcData.SelectedItem.ToString %></SelectedData>
                           <ParameterList>
                               <%= From item In dgvSimpleCalcsParameterList.Rows
                                   Select
                                   <Parameter>
                                       <Name><%= item.Cells(0).Value %></Name>
                                       <Type><%= item.Cells(1).Value %></Type>
                                       <Value><%= item.Cells(2).Value %></Value>
                                       <Description><%= item.Cells(3).Value %></Description>
                                   </Parameter> %>
                           </ParameterList>
                           <InputDataList>
                               <%= From item In dgvSimpleCalcsInputData.Rows
                                   Select
                                   <InputData>
                                       <Parameter><%= item.Cells(0).Value %></Parameter>
                                       <Column><%= item.Cells(1).Value %></Column>
                                   </InputData> %>
                           </InputDataList>
                           <CalculationList>
                               <%= From item In dgvSimpleCalcsCalculations.Rows
                                   Select
                                   <Calculation>
                                       <Input1><%= item.Cells(0).Value %></Input1>
                                       <Input2><%= item.Cells(1).Value %></Input2>
                                       <Operation><%= item.Cells(2).Value %></Operation>
                                       <Output><%= item.Cells(3).Value %></Output>
                                   </Calculation> %>
                           </CalculationList>
                           <OutputDataList>
                               <%= From item In dgvSimpleCalcsOutputData.Rows
                                   Select
                                   <OutputData>
                                       <Parameter><%= item.Cells(0).Value %></Parameter>
                                       <Column><%= item.Cells(1).Value %></Column>
                                   </OutputData> %>
                           </OutputDataList>
                       </SimpleCalculationsSettings>

            Project.SaveXmlData(txtSimpleCalcSettings.Text, XDoc)
            dgvSimpleCalcsParameterList.AllowUserToAddRows = True 'Allow user to add rows again.
            dgvSimpleCalcsCalculations.AllowUserToAddRows = True
            dgvSimpleCalcsInputData.AllowUserToAddRows = True
            dgvSimpleCalcsOutputData.AllowUserToAddRows = True
            SimpleCalcsSettingsFile = txtSimpleCalcSettings.Text

        End If
    End Sub

    Private Sub LoadSimpleCalcsSettingsFile()
        'Load the Simple Calculations Settings file.
        'SimpleCalcsSettingsFile contains the file name.

        If SimpleCalcsSettingsFile = "" Then
            Message.AddWarning("No Simple Calculations settings file name has been specified!" & vbCrLf)
        Else
            Dim XDoc As System.Xml.Linq.XDocument
            Project.ReadXmlData(SimpleCalcsSettingsFile, XDoc)

            If XDoc.<SimpleCalculationsSettings>.<SelectedDatabase>.Value <> Nothing Then cmbSimpleCalcDb.SelectedIndex = cmbSimpleCalcDb.FindStringExact(XDoc.<SimpleCalculationsSettings>.<SelectedDatabase>.Value)
            If XDoc.<SimpleCalculationsSettings>.<SelectedData>.Value <> Nothing Then cmbSimpleCalcData.SelectedIndex = cmbSimpleCalcData.FindStringExact(XDoc.<SimpleCalculationsSettings>.<SelectedData>.Value)

            dgvSimpleCalcsParameterList.Rows.Clear()
            Dim settings = From item In XDoc.<SimpleCalculationsSettings>.<ParameterList>.<Parameter>
            For Each item In settings
                dgvSimpleCalcsParameterList.Rows.Add(item.<Name>.Value, item.<Type>.Value, item.<Value>.Value, item.<Description>.Value)
            Next
            dgvSimpleCalcsParameterList.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            dgvSimpleCalcsParameterList.AutoResizeColumns()

            UpdateSimpleCalcsParams()

            dgvSimpleCalcsInputData.Rows.Clear()
            Dim settings2 = From item In XDoc.<SimpleCalculationsSettings>.<InputDataList>.<InputData>
            For Each item In settings2
                dgvSimpleCalcsInputData.Rows.Add(item.<Parameter>.Value, item.<Column>.Value)
            Next
            dgvSimpleCalcsInputData.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            dgvSimpleCalcsInputData.AutoResizeColumns()

            dgvSimpleCalcsCalculations.Rows.Clear()
            Dim settings3 = From item In XDoc.<SimpleCalculationsSettings>.<CalculationList>.<Calculation>
            For Each item In settings3
                dgvSimpleCalcsCalculations.Rows.Add(item.<Input1>.Value, item.<Input2>.Value, item.<Operation>.Value, item.<Output>.Value)
            Next
            dgvSimpleCalcsCalculations.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            dgvSimpleCalcsCalculations.AutoResizeColumns()

            dgvSimpleCalcsOutputData.Rows.Clear()
            Dim settings4 = From item In XDoc.<SimpleCalculationsSettings>.<OutputDataList>.<OutputData>
            For Each item In settings4
                dgvSimpleCalcsOutputData.Rows.Add(item.<Parameter>.Value, item.<Column>.Value)
            Next
            dgvSimpleCalcsOutputData.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            dgvSimpleCalcsOutputData.AutoResizeColumns()

            'dgvSimpleCalcsParameterList.AllowUserToResizeColumns = True
            'dgvSimpleCalcsCalculations.AllowUserToResizeColumns = True
            'dgvSimpleCalcsInputData.AllowUserToResizeColumns = True
            'dgvSimpleCalcsOutputData.AllowUserToResizeColumns = True

            dgvSimpleCalcsParameterList.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            dgvSimpleCalcsCalculations.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            dgvSimpleCalcsInputData.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            dgvSimpleCalcsOutputData.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

        End If
    End Sub

    Private Sub btnFindSimpleCalcSettings_Click(sender As Object, e As EventArgs) Handles btnFindSimpleCalcSettings.Click
        'Find a Simple Calculations Settings file.

        Select Case Project.DataLocn.Type
            Case ADVL_Utilities_Library_1.FileLocation.Types.Directory
                'Select a Simple Calculations settings file from the project directory:
                OpenFileDialog1.InitialDirectory = Project.DataLocn.Path
                OpenFileDialog1.Filter = "Simple Calculations settings file | *.SimpleCalcs"
                If OpenFileDialog1.ShowDialog() = DialogResult.OK Then
                    Dim DataFileName As String = System.IO.Path.GetFileName(OpenFileDialog1.FileName)
                    SimpleCalcsSettingsFile = DataFileName
                    txtSimpleCalcSettings.Text = DataFileName
                    LoadSimpleCalcsSettingsFile()
                End If
            Case ADVL_Utilities_Library_1.FileLocation.Types.Archive
                'Select a Copy Columns settings file from the project archive:
                'Show the zip archive file selection form:
                Zip = New ADVL_Utilities_Library_1.ZipComp
                Zip.ArchivePath = Project.DataLocn.Path
                Zip.SelectFile() 'Show the SelectFile form.
                Zip.SelectFileForm.ApplicationName = Project.ApplicationName
                Zip.SelectFileForm.SettingsLocn = Project.SettingsLocn
                Zip.SelectFileForm.Show()
                Zip.SelectFileForm.RestoreFormSettings()
                Zip.SelectFileForm.FileExtension = ".SimpleCalcs"
                Zip.SelectFileForm.GetFileList()
                If Zip.SelectedFile <> "" Then
                    'A file has been selected
                    SimpleCalcsSettingsFile = Zip.SelectedFile
                    txtSimpleCalcSettings.Text = Zip.SelectedFile
                    LoadSimpleCalcsSettingsFile()
                End If
        End Select

    End Sub

    Private Sub UpdateSimpleCalcsParams()
        'Update the Parameter selection options in dgvCalculations, dgvInputData and dgvOutputData

        'Save the current settings in dgvSimpleCalcsCalculations
        dgvSimpleCalcsCalculations.AllowUserToAddRows = False 'This removes the last edit row from the DataGridView.
        Dim NCalcsRows As Integer = dgvSimpleCalcsCalculations.RowCount 'The number of rows in the Calculation grid.
        Dim SavedCalcsData(0 To NCalcsRows, 0 To 4) As String 'Array used to temporarily saved the contents of the Calculation grid.
        Dim I As Integer

        'Save the contents of the Calculation grid:
        For I = 0 To NCalcsRows - 1
            SavedCalcsData(I, 0) = dgvSimpleCalcsCalculations.Rows(I).Cells(0).Value
            SavedCalcsData(I, 1) = dgvSimpleCalcsCalculations.Rows(I).Cells(1).Value
            SavedCalcsData(I, 2) = dgvSimpleCalcsCalculations.Rows(I).Cells(2).Value
            SavedCalcsData(I, 3) = dgvSimpleCalcsCalculations.Rows(I).Cells(3).Value
        Next

        'Save the current settings in dgvSimpleCalcsInputData
        dgvSimpleCalcsInputData.AllowUserToAddRows = False 'This removes the last edit row from the DataGridView.
        Dim NInputRows As Integer = dgvSimpleCalcsInputData.RowCount 'The number of rows in the Input grid.
        Dim SavedInputData(0 To NInputRows, 0 To 2) As String 'Array used to temporarily saved the contents of the Input grid.

        'Save the contents of the Input grid:
        For I = 0 To NInputRows - 1
            SavedInputData(I, 0) = dgvSimpleCalcsInputData.Rows(I).Cells(0).Value
            SavedInputData(I, 1) = dgvSimpleCalcsInputData.Rows(I).Cells(1).Value
        Next

        'Save the current settings in dgvSimpleCalcsOutputData
        dgvSimpleCalcsOutputData.AllowUserToAddRows = False 'This removes the last edit row from the DataGridView.
        Dim NOutputRows As Integer = dgvSimpleCalcsOutputData.RowCount 'The number of rows in the Input grid.
        Dim SavedOutputData(0 To NOutputRows, 0 To 2) As String 'Array used to temporarily saved the contents of the Output grid.

        'Save the contents of the Input grid:
        For I = 0 To NOutputRows - 1
            SavedOutputData(I, 0) = dgvSimpleCalcsOutputData.Rows(I).Cells(0).Value
            SavedOutputData(I, 1) = dgvSimpleCalcsOutputData.Rows(I).Cells(1).Value
        Next

        'Set up dgvSimpleCalcsCalculations ---------------------------------------------
        dgvSimpleCalcsCalculations.Columns.Clear()

        Dim ComboBoxCol30 As New DataGridViewComboBoxColumn
        dgvSimpleCalcsCalculations.Columns.Add(ComboBoxCol30)
        dgvSimpleCalcsCalculations.Columns(0).HeaderText = "Input 1"
        dgvSimpleCalcsCalculations.Columns(0).Width = 140

        'Add the list of Parameters
        If dgvSimpleCalcsParameterList.RowCount > 0 Then
            For Each item In dgvSimpleCalcsParameterList.Rows
                If item.Cells(0).Value <> Nothing Then ComboBoxCol30.Items.Add(item.Cells(0).Value)
            Next
        End If

        Dim ComboBoxCol31 As New DataGridViewComboBoxColumn
        dgvSimpleCalcsCalculations.Columns.Add(ComboBoxCol31)
        dgvSimpleCalcsCalculations.Columns(1).HeaderText = "Input 2"
        dgvSimpleCalcsCalculations.Columns(1).Width = 140

        'Add the list of Parameters
        If dgvSimpleCalcsParameterList.RowCount > 0 Then
            For Each item In dgvSimpleCalcsParameterList.Rows
                If item.Cells(0).Value <> Nothing Then ComboBoxCol31.Items.Add(item.Cells(0).Value)
            Next
        End If

        Dim ComboBoxCol32 As New DataGridViewComboBoxColumn
        dgvSimpleCalcsCalculations.Columns.Add(ComboBoxCol32)
        dgvSimpleCalcsCalculations.Columns(2).HeaderText = "Operation"
        dgvSimpleCalcsCalculations.Columns(2).Width = 110

        ComboBoxCol32.Items.Add("Input 1 + Input 2")
        ComboBoxCol32.Items.Add("Input 1 - Input 2")
        ComboBoxCol32.Items.Add("Input 1 x Input 2")
        ComboBoxCol32.Items.Add("Input 1 / Input 2")

        Dim ComboBoxCol33 As New DataGridViewComboBoxColumn
        dgvSimpleCalcsCalculations.Columns.Add(ComboBoxCol33)
        dgvSimpleCalcsCalculations.Columns(3).HeaderText = "Output"
        dgvSimpleCalcsCalculations.Columns(3).Width = 140

        'Add the list of Parameters
        If dgvSimpleCalcsParameterList.RowCount > 0 Then
            For Each item In dgvSimpleCalcsParameterList.Rows
                If item.Cells(0).Value <> Nothing Then ComboBoxCol33.Items.Add(item.Cells(0).Value)
            Next
        End If

        'Set up dgvSimpleCalcsInputData ---------------------------------------------
        dgvSimpleCalcsInputData.Columns.Clear()

        Dim ComboBoxCol20 As New DataGridViewComboBoxColumn
        dgvSimpleCalcsInputData.Columns.Add(ComboBoxCol20)
        dgvSimpleCalcsInputData.Columns(0).HeaderText = "Input Parameter"
        dgvSimpleCalcsInputData.Columns(0).Width = 140

        'Add the list of Parameters
        If dgvSimpleCalcsParameterList.RowCount > 0 Then
            For Each item In dgvSimpleCalcsParameterList.Rows
                If item.Cells(0).Value <> Nothing Then ComboBoxCol20.Items.Add(item.Cells(0).Value)
            Next
        End If

        Dim ComboBoxCol21 As New DataGridViewComboBoxColumn
        dgvSimpleCalcsInputData.Columns.Add(ComboBoxCol21)
        dgvSimpleCalcsInputData.Columns(1).HeaderText = "Input Column"
        dgvSimpleCalcsInputData.Columns(1).Width = 140

        'If dsOutput.Tables.Count > 0 Then
        '    For Each item In dsOutput.Tables("myData").Columns
        '        ComboBoxCol21.Items.Add(item.Columnname)
        '    Next
        'End If

        If cmbSimpleCalcDb.SelectedIndex = -1 Then
            cmbSimpleCalcDb.SelectedIndex = 0
        End If
        If cmbSimpleCalcData.SelectedIndex = -1 Then
            cmbSimpleCalcData.SelectedIndex = 0
        End If

        Select Case cmbSimpleCalcDb.SelectedItem.ToString
            Case "Share Prices"
                For Each item In SharePricesSettings.List(cmbSimpleCalcData.SelectedIndex).TableCols
                    ComboBoxCol21.Items.Add(item)
                Next
            Case "Financials"
                For Each item In FinancialsSettings.List(cmbSimpleCalcData.SelectedIndex).TableCols
                    ComboBoxCol21.Items.Add(item)
                Next
            Case "Calculations"
                For Each item In CalculationsSettings.List(cmbSimpleCalcData.SelectedIndex).TableCols
                    ComboBoxCol21.Items.Add(item)
                Next
        End Select

        'Set up dgvSimpleCalcsOutputData ---------------------------------------------
        dgvSimpleCalcsOutputData.Columns.Clear()

        Dim ComboBoxCol40 As New DataGridViewComboBoxColumn
        dgvSimpleCalcsOutputData.Columns.Add(ComboBoxCol40)
        dgvSimpleCalcsOutputData.Columns(0).HeaderText = "Output Parameter"
        dgvSimpleCalcsOutputData.Columns(0).Width = 140

        'Add the list of Parameters
        If dgvSimpleCalcsParameterList.RowCount > 0 Then
            For Each item In dgvSimpleCalcsParameterList.Rows
                'ComboBoxCol40.Items.Add(item.Cells(0).Value)
                If item.Cells(0).Value <> Nothing Then ComboBoxCol40.Items.Add(item.Cells(0).Value)
            Next
        End If

        Dim ComboBoxCol41 As New DataGridViewComboBoxColumn
        dgvSimpleCalcsOutputData.Columns.Add(ComboBoxCol41)
        dgvSimpleCalcsOutputData.Columns(1).HeaderText = "Output Column"
        dgvSimpleCalcsOutputData.Columns(1).Width = 140

        'If dsOutput.Tables.Count > 0 Then
        '    For Each item In dsOutput.Tables("myData").Columns
        '        ComboBoxCol41.Items.Add(item.Columnname)
        '    Next
        'End If

        Select Case cmbSimpleCalcDb.SelectedItem.ToString
            Case "Share Prices"
                For Each item In SharePricesSettings.List(cmbSimpleCalcData.SelectedIndex).TableCols
                    ComboBoxCol41.Items.Add(item)
                Next
            Case "Financials"
                For Each item In FinancialsSettings.List(cmbSimpleCalcData.SelectedIndex).TableCols
                    ComboBoxCol41.Items.Add(item)
                Next
            Case "Calculations"
                For Each item In CalculationsSettings.List(cmbSimpleCalcData.SelectedIndex).TableCols
                    ComboBoxCol41.Items.Add(item)
                Next
        End Select

        'Restore the contents of the Calculation grid:
        For I = 0 To NCalcsRows - 1
            dgvSimpleCalcsCalculations.Rows.Add(SavedCalcsData(I, 0), SavedCalcsData(I, 1), SavedCalcsData(I, 2), SavedCalcsData(I, 3))
        Next

        'Restore the contents of the Input grid:
        For I = 0 To NInputRows - 1
            dgvSimpleCalcsInputData.Rows.Add(SavedInputData(I, 0), SavedInputData(I, 1))
        Next

        'Restore the contents of the Output grid:
        For I = 0 To NOutputRows - 1
            dgvSimpleCalcsOutputData.Rows.Add(SavedOutputData(I, 0), SavedOutputData(I, 1))
        Next

        dgvSimpleCalcsCalculations.AllowUserToAddRows = True 'Allow user to add rows again.
        dgvSimpleCalcsInputData.AllowUserToAddRows = True
        dgvSimpleCalcsOutputData.AllowUserToAddRows = True

        'dgvSimpleCalcsParameterList.AllowUserToResizeColumns = True
        'dgvSimpleCalcsCalculations.AllowUserToResizeColumns = True
        'dgvSimpleCalcsInputData.AllowUserToResizeColumns = True
        'dgvSimpleCalcsOutputData.AllowUserToResizeColumns = True

        dgvSimpleCalcsParameterList.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
        dgvSimpleCalcsCalculations.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
        dgvSimpleCalcsInputData.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
        dgvSimpleCalcsOutputData.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

    End Sub

    Private Sub dgvSimpleCalcsParameterList_LostFocus(sender As Object, e As EventArgs) Handles dgvSimpleCalcsParameterList.LostFocus
        UpdateSimpleCalcsParams()
    End Sub

    'NOTE: The selected data will be loaded into dsOutput only when it is required for the calculations.
    'Private Sub cmbSimpleCalcData_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbSimpleCalcData.SelectedIndexChanged
    '    'The Simple Calculations output table has been selected.

    '    'dsOutput will be used for the Input and Output data columns
    '    dsOutput.Clear()
    '    dsOutput.Tables.Clear()

    '    Select Case cmbSimpleCalcDb.SelectedItem.ToString
    '        Case "Share Prices"
    '            outputQuery = "SELECT * FROM " & cmbSimpleCalcData.SelectedItem.ToString
    '            outputConnString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source =" & SharePriceDbPath
    '            outputConnection.ConnectionString = outputConnString
    '            outputConnection.Open()
    '            outputDa = New OleDb.OleDbDataAdapter(outputQuery, outputConnection)
    '            outputDa.MissingSchemaAction = MissingSchemaAction.AddWithKey
    '            Dim outputCmdBuilder As New OleDb.OleDbCommandBuilder(outputDa)
    '            outputDa.Fill(dsOutput, "myData")
    '            outputConnection.Close()
    '            SetUpSimpleCalculationsTab()
    '        Case "Financials"
    '            'Query = FinancialsSettings.List(cmbInputData.SelectedIndex).Query
    '            outputQuery = "SELECT * FROM " & cmbSimpleCalcData.SelectedItem.ToString
    '            outputConnString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source =" & FinancialsDbPath
    '            outputConnection.ConnectionString = outputConnString
    '            outputConnection.Open()
    '            outputDa = New OleDb.OleDbDataAdapter(outputQuery, outputConnection)
    '            outputDa.MissingSchemaAction = MissingSchemaAction.AddWithKey
    '            Dim outputCmdBuilder As New OleDb.OleDbCommandBuilder(outputDa)
    '            outputDa.Fill(dsOutput, "myData")
    '            outputConnection.Close()
    '            SetUpSimpleCalculationsTab()
    '        Case "Calculations"
    '            outputQuery = "SELECT * FROM " & cmbSimpleCalcData.SelectedItem.ToString
    '            outputConnString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source =" & CalculationsDbPath
    '            outputConnection.ConnectionString = outputConnString
    '            outputConnection.Open()
    '            outputDa = New OleDb.OleDbDataAdapter(outputQuery, outputConnection)
    '            outputDa.MissingSchemaAction = MissingSchemaAction.AddWithKey
    '            Dim outputCmdBuilder As New OleDb.OleDbCommandBuilder(outputDa)
    '            outputDa.Fill(dsOutput, "myData")
    '            outputConnection.Close()
    '            SetUpSimpleCalculationsTab()
    '    End Select

    'End Sub

    Private Sub btnApplySimpleCalcSettings_Click(sender As Object, e As EventArgs) Handles btnApplySimpleCalcSettings.Click
        'Apply the Simple Calculations.
        ApplySimpleCalcs()
    End Sub

    Private Sub ApplySimpleCalcs()
        'Apply the Simple Calculations.

        LoadSimpleCalcsDsOutputData()

        'dsOutput contains the selected data

        dgvSimpleCalcsParameterList.AllowUserToAddRows = False 'This removes the last edit row from the DataGridView.
        dgvSimpleCalcsCalculations.AllowUserToAddRows = False
        dgvSimpleCalcsInputData.AllowUserToAddRows = False
        dgvSimpleCalcsOutputData.AllowUserToAddRows = False

        'Set up the Parameter List:
        Dim Param As New Dictionary(Of String, DbSingle)

        'Get parameters from dgvSimpleCalcsParameterList
        For Each item In dgvSimpleCalcsParameterList.Rows
            If item.Cells(1).Value = "Constant" Then
                Param.Add(item.Cells(0).Value, GetDbSingle(item.Cells(2).Value, False)) 'For a Constant, Param("Name").Value = constant value, .NullValue = False. This value remains a constant.
                'ElseIf item.Cells(1).Value = "Date" Then
                '    Param.Add()
                'NOTE: This code won't work because the Param disctionary can store only single values, not dates!!!
            Else
                Param.Add(item.Cells(0).Value, GetDbSingle(0, True)) 'Initially Param("Name").Value = 0, .NullValue = True. The value will be read later from the database.
            End If
        Next

        'Get Input settings:
        Dim InputData As New List(Of ParameterLocation)

        'Get Input Data settings from dgvSimpleCalcsInputData
        For Each item In dgvSimpleCalcsInputData.Rows
            InputData.Add(GetParamLocn(item.Cells(0).Value, item.Cells(1).Value))
        Next

        'Get Calculations list:
        Dim Calc As New List(Of Calculation)

        For Each item In dgvSimpleCalcsCalculations.Rows
            Calc.Add(GetCalc(item.Cells(0).Value, item.Cells(1).Value, item.Cells(2).Value, item.Cells(3).Value))
        Next

        'Get Output settings:
        Dim OutputData As New List(Of ParameterLocation)
        'Get Output Data settings from dgvSimpleCalcsOutputData
        For Each item In dgvSimpleCalcsOutputData.Rows
            OutputData.Add(GetParamLocn(item.Cells(0).Value, item.Cells(1).Value))
        Next

        Message.Add("Starting Simple Calculations ----------------------- " & vbCrLf)
        Dim Count As Integer = 0

        'Process each row in dsOutput
        For Each item In dsOutput.Tables("myData").Rows
            Count += 1
            If Count Mod 100 = 0 Then Message.Add("Selecting record number: " & Count & vbCrLf)

            'Get Input Data values:
            For Each inputItem In InputData
                If IsDBNull(item(inputItem.ColName)) Then
                    Param(inputItem.ParamName) = GetDbSingle(0, True)
                Else
                    Param(inputItem.ParamName) = GetDbSingle(item(inputItem.ColName), False)
                End If
            Next

            'Perform Calculations:
            For Each calcItem In Calc
                'Check for null values
                If Param(calcItem.Input1).NullValue Then
                    'Input1 is null: cannot perform calculation, set Output to null.
                    Param(calcItem.Output).NullValue = True
                Else
                    If Param(calcItem.Input2).NullValue Then
                        'Input2 is null: cannot perform calculation, set Output to null.
                        Param(calcItem.Output).NullValue = True
                    Else
                        'Input1 and Input2 have non null values. Perform calculation
                        Select Case calcItem.Operation
                            Case "Input 1 + Input 2"
                                Param(calcItem.Output).Value = Param(calcItem.Input1).Value + Param(calcItem.Input2).Value
                                Param(calcItem.Output).NullValue = False
                            Case "Input 1 - Input 2"
                                Param(calcItem.Output).Value = Param(calcItem.Input1).Value - Param(calcItem.Input2).Value
                                Param(calcItem.Output).NullValue = False
                            Case "Input 1 x Input 2"
                                Param(calcItem.Output).Value = Param(calcItem.Input1).Value * Param(calcItem.Input2).Value
                                Param(calcItem.Output).NullValue = False
                            Case "Input 1 / Input 2"
                                If Param(calcItem.Input2).Value = 0 Then
                                    'Attempting to divide by zero! Set output value to DbNull.
                                    Param(calcItem.Output).NullValue = True
                                Else
                                    Param(calcItem.Output).Value = Param(calcItem.Input1).Value / Param(calcItem.Input2).Value
                                    Param(calcItem.Output).NullValue = False
                                End If
                            Case Else
                                Message.AddWarning("Unrecognised calculation operation: " & calcItem.Operation & vbCrLf)
                        End Select
                    End If
                End If
            Next

            'Write Output Data values:
            For Each outputItem In OutputData
                If Param(outputItem.ParamName).NullValue Then
                    'Write a dbNull to the DataTable
                    item(outputItem.ColName) = DBNull.Value
                Else
                    'Write the output value to the DataTable
                    item(outputItem.ColName) = Param(outputItem.ParamName).Value
                End If
            Next
        Next

        Try
            Message.Add("Updating Database ------------------ " & vbCrLf)
            outputDa.Update(dsOutput.Tables("myData"))
            Message.Add("Simple Calculations Complete ----------------- " & vbCrLf)
        Catch ex As Exception
            Message.AddWarning("Error updating database table: " & ex.Message & vbCrLf)
        End Try

        'Clear the data when finished:
        dsOutput.Clear()
        dsOutput.Tables.Clear()

        dgvSimpleCalcsParameterList.AllowUserToAddRows = True
        dgvSimpleCalcsInputData.AllowUserToAddRows = True
        dgvSimpleCalcsCalculations.AllowUserToAddRows = True
        dgvSimpleCalcsOutputData.AllowUserToAddRows = True
        dgvSimpleCalcsParameterList.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
        dgvSimpleCalcsInputData.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
        dgvSimpleCalcsCalculations.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
        dgvSimpleCalcsOutputData.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

    End Sub

    Private Sub LoadSimpleCalcsDsOutputData()
        'Open selected data in dsOutput.

        dsOutput.Clear()
        dsOutput.Tables.Clear()

        Select Case cmbSimpleCalcDb.SelectedItem.ToString
            Case "Share Prices"
                'outputQuery = SharePricesSettings.List(cmbSimpleCalcData.SelectedIndex).Query
                outputQuery = txtSimpleCalcsQuery.Text
                outputConnString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source =" & SharePriceDbPath
                outputConnection.ConnectionString = outputConnString
                outputConnection.Open()
                outputDa = New OleDb.OleDbDataAdapter(outputQuery, outputConnection)
                outputDa.MissingSchemaAction = MissingSchemaAction.AddWithKey
                Dim outputCmdBuilder As New OleDb.OleDbCommandBuilder(outputDa)
                outputDa.Fill(dsOutput, "myData")
                outputConnection.Close()
            Case "Financials"
                'outputQuery = FinancialsSettings.List(cmbSimpleCalcData.SelectedIndex).Query
                outputQuery = txtSimpleCalcsQuery.Text
                outputConnString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source =" & FinancialsDbPath
                outputConnection.ConnectionString = outputConnString
                outputConnection.Open()
                outputDa = New OleDb.OleDbDataAdapter(outputQuery, outputConnection)
                outputDa.MissingSchemaAction = MissingSchemaAction.AddWithKey
                Dim outputCmdBuilder As New OleDb.OleDbCommandBuilder(outputDa)
                outputDa.Fill(dsOutput, "myData")
                outputConnection.Close()
            Case "Calculations"
                'outputQuery = CalculationsSettings.List(cmbSimpleCalcData.SelectedIndex).Query
                outputQuery = txtSimpleCalcsQuery.Text
                outputConnString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source =" & CalculationsDbPath
                outputConnection.ConnectionString = outputConnString
                outputConnection.Open()
                outputDa = New OleDb.OleDbDataAdapter(outputQuery, outputConnection)
                outputDa.MissingSchemaAction = MissingSchemaAction.AddWithKey
                Dim outputCmdBuilder As New OleDb.OleDbCommandBuilder(outputDa)
                outputDa.Fill(dsOutput, "myData")
                outputConnection.Close()
        End Select
    End Sub

    Private Function GetDbSingle(ByVal Value As Single, ByVal NullValue As Boolean) As DbSingle
        Dim NewDbSingle As New DbSingle

        NewDbSingle.Value = Value
        NewDbSingle.NullValue = NullValue
        Return NewDbSingle
        'GetDbSingle.Value = Value
        'GetDbSingle.NullValue = NullValue
    End Function

    Private Function GetCalc(ByVal Input1 As String, ByVal Input2 As String, ByVal Operation As String, ByVal Output As String) As Calculation
        Dim NewCalculation As New Calculation

        NewCalculation.Input1 = Input1
        NewCalculation.Input2 = Input2
        NewCalculation.Operation = Operation
        NewCalculation.Output = Output
        Return NewCalculation
        'GetCalc.Input1 = Input1
        'GetCalc.Input2 = Input2
        'GetCalc.Operation = Operation
        'GetCalc.Output = Output
    End Function

    Private Function GetParamLocn(ByVal ParamName As String, ByVal ColName As String) As ParameterLocation
        Dim NewParamLocn As New ParameterLocation

        NewParamLocn.ParamName = ParamName
        NewParamLocn.ColName = ColName
        Return NewParamLocn

        'GetParamLocn.ParamName = ParamName
        'GetParamLocn.ColName = ColName
    End Function

    Private Sub btnAddSimpleCalcsToSequence_Click(sender As Object, e As EventArgs) Handles btnAddSimpleCalcsToSequence.Click
        'Add the Simple Calculations sequence to the Processing Sequence.

        If IsNothing(Sequence) Then
            Message.AddWarning("The Processing Sequence form is not open." & vbCrLf)
        Else
            'Write new instructions to the Sequence text.
            Dim Posn As Integer
            Posn = Sequence.rtbSequence.SelectionStart

            Sequence.rtbSequence.SelectedText = vbCrLf & "<!--Simple Calculations: Settings used to perform simple calculations on data in the selected table :-->" & vbCrLf
            Sequence.rtbSequence.SelectedText = "  <SimpleCalculations>" & vbCrLf

            'Selected data parameters:
            Sequence.rtbSequence.SelectedText = "    <SelectedDatabase>" & cmbSimpleCalcDb.SelectedItem.ToString & "</SelectedDatabase>" & vbCrLf
            Select Case cmbSimpleCalcDb.SelectedItem.ToString
                Case "Share Prices"
                    Sequence.rtbSequence.SelectedText = "    <SelectedDatabasePath>" & SharePriceDbPath & "</SelectedDatabasePath>" & vbCrLf
                Case "Financials"
                    Sequence.rtbSequence.SelectedText = "    <SelectedDatabasePath>" & FinancialsDbPath & "</SelectedDatabasePath>" & vbCrLf
                Case "Calculations"
                    Sequence.rtbSequence.SelectedText = "    <SelectedDatabasePath>" & CalculationsDbPath & "</SelectedDatabasePath>" & vbCrLf
            End Select
            Sequence.rtbSequence.SelectedText = "    <DataQuery>" & txtSimpleCalcsQuery.Text & "</DataQuery>" & vbCrLf

            'Parameter List:
            Sequence.rtbSequence.SelectedText = "    <ParameterList>" & vbCrLf
            dgvSimpleCalcsParameterList.AllowUserToAddRows = False 'This removes the last edit row from the DataGridView.
            Dim NRows As Integer = dgvSimpleCalcsParameterList.Rows.Count
            Dim RowNo As Integer
            For RowNo = 0 To NRows - 1
                Sequence.rtbSequence.SelectedText = "      <Parameter>" & vbCrLf
                Sequence.rtbSequence.SelectedText = "        <Name>" & dgvSimpleCalcsParameterList.Rows(RowNo).Cells(0).Value & "</Name>" & vbCrLf
                Sequence.rtbSequence.SelectedText = "        <Type>" & dgvSimpleCalcsParameterList.Rows(RowNo).Cells(1).Value & "</Type>" & vbCrLf
                Sequence.rtbSequence.SelectedText = "        <Value>" & dgvSimpleCalcsParameterList.Rows(RowNo).Cells(2).Value & "</Value>" & vbCrLf
                Sequence.rtbSequence.SelectedText = "        <Description>" & dgvSimpleCalcsParameterList.Rows(RowNo).Cells(3).Value & "</Description>" & vbCrLf
                Sequence.rtbSequence.SelectedText = "      </Parameter>" & vbCrLf
            Next
            Sequence.rtbSequence.SelectedText = "    </ParameterList>" & vbCrLf
            dgvSimpleCalcsParameterList.AllowUserToAddRows = True 'Allow user to add rows again.

            'Input Data List:
            Sequence.rtbSequence.SelectedText = "    <InputDataList>" & vbCrLf
            dgvSimpleCalcsInputData.AllowUserToAddRows = False 'This removes the last edit row from the DataGridView.
            NRows = dgvSimpleCalcsInputData.Rows.Count
            For RowNo = 0 To NRows - 1
                Sequence.rtbSequence.SelectedText = "      <InputData>" & vbCrLf
                Sequence.rtbSequence.SelectedText = "        <Parameter>" & dgvSimpleCalcsInputData.Rows(RowNo).Cells(0).Value & "</Parameter>" & vbCrLf
                Sequence.rtbSequence.SelectedText = "        <Column>" & dgvSimpleCalcsInputData.Rows(RowNo).Cells(1).Value & "</Column>" & vbCrLf
                Sequence.rtbSequence.SelectedText = "      </InputData>" & vbCrLf
            Next
            Sequence.rtbSequence.SelectedText = "    </InputDataList>" & vbCrLf
            dgvSimpleCalcsInputData.AllowUserToAddRows = True 'Allow user to add rows again.

            'Calculation List:
            Sequence.rtbSequence.SelectedText = "    <CalculationList>" & vbCrLf
            dgvSimpleCalcsCalculations.AllowUserToAddRows = False 'This removes the last edit row from the DataGridView.
            NRows = dgvSimpleCalcsCalculations.Rows.Count
            For RowNo = 0 To NRows - 1
                Sequence.rtbSequence.SelectedText = "      <Calculation>" & vbCrLf
                Sequence.rtbSequence.SelectedText = "        <Input1>" & dgvSimpleCalcsCalculations.Rows(RowNo).Cells(0).Value & "</Input1>" & vbCrLf
                Sequence.rtbSequence.SelectedText = "        <Input2>" & dgvSimpleCalcsCalculations.Rows(RowNo).Cells(1).Value & "</Input2>" & vbCrLf
                Sequence.rtbSequence.SelectedText = "        <Operation>" & dgvSimpleCalcsCalculations.Rows(RowNo).Cells(2).Value & "</Operation>" & vbCrLf
                Sequence.rtbSequence.SelectedText = "        <Output>" & dgvSimpleCalcsCalculations.Rows(RowNo).Cells(3).Value & "</Output>" & vbCrLf
                Sequence.rtbSequence.SelectedText = "      </Calculation>" & vbCrLf
            Next
            Sequence.rtbSequence.SelectedText = "    </CalculationList>" & vbCrLf
            dgvSimpleCalcsCalculations.AllowUserToAddRows = True 'Allow user to add rows again.

            'Output Data List:
            Sequence.rtbSequence.SelectedText = "    <OutputDataList>" & vbCrLf
            dgvSimpleCalcsOutputData.AllowUserToAddRows = False 'This removes the last edit row from the DataGridView.
            NRows = dgvSimpleCalcsOutputData.Rows.Count
            For RowNo = 0 To NRows - 1
                Sequence.rtbSequence.SelectedText = "      <OutputData>" & vbCrLf
                Sequence.rtbSequence.SelectedText = "        <Parameter>" & dgvSimpleCalcsOutputData.Rows(RowNo).Cells(0).Value & "</Parameter>" & vbCrLf
                Sequence.rtbSequence.SelectedText = "        <Column>" & dgvSimpleCalcsOutputData.Rows(RowNo).Cells(1).Value & "</Column>" & vbCrLf
                Sequence.rtbSequence.SelectedText = "      </OutputData>" & vbCrLf
            Next
            Sequence.rtbSequence.SelectedText = "    </OutputDataList>" & vbCrLf
            dgvSimpleCalcsOutputData.AllowUserToAddRows = True 'Allow user to add rows again.

            Sequence.rtbSequence.SelectedText = "    <Command>Apply</Command>" & vbCrLf
            Sequence.rtbSequence.SelectedText = "  </SimpleCalculations>" & vbCrLf
            Sequence.rtbSequence.SelectedText = "<!---->" & vbCrLf

            Sequence.FormatXmlText()

        End If

    End Sub

    Private Sub btnNewSimpleCalcSettings_Click(sender As Object, e As EventArgs) Handles btnNewSimpleCalcSettings.Click
        'New Simple Calculations Settings.

        SimpleCalcsSettingsFile = ""
        txtSimpleCalcSettings.Text = ""
        cmbSimpleCalcData.SelectedIndex = 0

        dgvSimpleCalcsParameterList.Rows.Clear()
        dgvSimpleCalcsCalculations.Rows.Clear()
        dgvSimpleCalcsInputData.Rows.Clear()
        dgvSimpleCalcsOutputData.Rows.Clear()

    End Sub


#End Region 'Simple Calculations Sub Tab ------------------------------------------------------------------------------------------------------------------------------------------------------

#Region " Date Calculations Sub Tab" '=========================================================================================================================================================

    Private Sub cmbDateCalcDb_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbDateCalcDb.SelectedIndexChanged
        GetDateCalcsOutputDataList()
    End Sub

    Private Sub GetDateCalcsOutputDataList()
        'Fill the cmbDateCalcData list.

        cmbDateCalcData.Items.Clear()
        Dim I As Integer

        Select Case cmbDateCalcDb.SelectedItem.ToString
            Case "Share Prices"
                'SharePriceDataViewList
                For I = 0 To SharePricesSettings.NRecords - 1
                    cmbDateCalcData.Items.Add(SharePricesSettings.List(I).Description)
                Next
            Case "Financials"
                'FinancialsDataViewList
                For I = 0 To FinancialsSettings.NRecords - 1
                    cmbDateCalcData.Items.Add(FinancialsSettings.List(I).Description)
                Next
            Case "Calculations"
                'CalculationsDataViewList
                For I = 0 To CalculationsSettings.NRecords - 1
                    cmbDateCalcData.Items.Add(CalculationsSettings.List(I).Description)
                Next
        End Select

    End Sub

    Private Sub SetUpDateCalcsTab()

        txtDateCalcSettings.Text = DateCalcSettingsFile

        cmbDateCalcParam1.Items.Clear()
        Select Case cmbDateCalcDb.SelectedItem.ToString
            Case "Share Prices"
                For Each item In SharePricesSettings.List(cmbDateCalcData.SelectedIndex).TableCols
                    cmbDateCalcParam1.Items.Add(item)
                Next
                txtDateCalcsQuery.Text = SharePricesSettings.List(cmbDateCalcData.SelectedIndex).Query
            Case "Financials"
                For Each item In FinancialsSettings.List(cmbDateCalcData.SelectedIndex).TableCols
                    cmbDateCalcParam1.Items.Add(item)
                Next
                txtDateCalcsQuery.Text = FinancialsSettings.List(cmbDateCalcData.SelectedIndex).Query
            Case "Calculations"
                For Each item In CalculationsSettings.List(cmbDateCalcData.SelectedIndex).TableCols
                    cmbDateCalcParam1.Items.Add(item)
                Next
                txtDateCalcsQuery.Text = CalculationsSettings.List(cmbDateCalcData.SelectedIndex).Query
        End Select

        cmbDateCalcParam2.Items.Clear()
        Select Case cmbDateCalcDb.SelectedItem.ToString
            Case "Share Prices"
                For Each item In SharePricesSettings.List(cmbDateCalcData.SelectedIndex).TableCols
                    cmbDateCalcParam2.Items.Add(item)
                Next
            Case "Financials"
                For Each item In FinancialsSettings.List(cmbDateCalcData.SelectedIndex).TableCols
                    cmbDateCalcParam2.Items.Add(item)
                Next
            Case "Calculations"
                For Each item In CalculationsSettings.List(cmbDateCalcData.SelectedIndex).TableCols
                    cmbDateCalcParam2.Items.Add(item)
                Next
        End Select

        cmbDateCalcOutput.Items.Clear()
        Select Case cmbDateCalcDb.SelectedItem.ToString
            Case "Share Prices"
                For Each item In SharePricesSettings.List(cmbDateCalcData.SelectedIndex).TableCols
                    cmbDateCalcOutput.Items.Add(item)
                Next
            Case "Financials"
                For Each item In FinancialsSettings.List(cmbDateCalcData.SelectedIndex).TableCols
                    cmbDateCalcOutput.Items.Add(item)
                Next
            Case "Calculations"
                For Each item In CalculationsSettings.List(cmbDateCalcData.SelectedIndex).TableCols
                    cmbDateCalcOutput.Items.Add(item)
                Next
        End Select

    End Sub

    Private Sub btnSaveDateCalcSettings_Click(sender As Object, e As EventArgs) Handles btnSaveDateCalcSettings.Click
        'Save the Date Calculations Settings.

        If Trim(txtDateCalcSettings.Text) = "" Then
            Message.AddWarning("No file name has been sepcified!" & vbCrLf)
        Else
            If txtDateCalcSettings.Text.EndsWith(".DateCalcs") Then
                txtDateCalcSettings.Text = Trim(txtDateCalcSettings.Text)
            Else
                txtDateCalcSettings.Text = Trim(txtDateCalcSettings.Text) & ".DateCalcs"
            End If

            If cmbDateCalcDb.SelectedIndex = -1 Then
                Message.AddWarning("A database has not been selected" & vbCrLf)
                Exit Sub
            End If
            If cmbDateCalcData.SelectedIndex = -1 Then
                Message.AddWarning("A dataset has not been selected" & vbCrLf)
                Exit Sub
            End If
            If cmbDateCalcType.SelectedIndex = -1 Then
                Message.AddWarning("A date calculation type has not been selected" & vbCrLf)
                Exit Sub
            End If
            If cmbDateCalcParam1.SelectedIndex = -1 Then
                Message.AddWarning("Parameter 1 has not been selected" & vbCrLf)
                Exit Sub
            End If
            If cmbDateCalcParam2.SelectedIndex = -1 Then
                Message.AddWarning("Parameter 2 has not been selected" & vbCrLf)
                Exit Sub
            End If
            If cmbDateCalcOutput.SelectedIndex = -1 Then
                Message.AddWarning("The output column has not been selected" & vbCrLf)
                Exit Sub
            End If

            Dim XDoc = <?xml version="1.0" encoding="utf-8"?>
                       <DateCalculationsSettings>
                           <!---->
                           <!--Date Calculations Settings-->
                           <!---->
                           <HostApplication><%= ApplicationInfo.Name %></HostApplication>
                           <SelectedDatabase><%= cmbDateCalcDb.SelectedItem.ToString %></SelectedDatabase>
                           <SelectedData><%= cmbDateCalcData.SelectedItem.ToString %></SelectedData>
                           <CalculationType><%= cmbDateCalcType.SelectedItem.ToString %></CalculationType>
                           <InputColumn1><%= cmbDateCalcParam1.SelectedItem.ToString %></InputColumn1>
                           <InputColumn2><%= cmbDateCalcParam2.SelectedItem.ToString %></InputColumn2>
                           <OutputColumn><%= cmbDateCalcOutput.SelectedItem.ToString %></OutputColumn>
                       </DateCalculationsSettings>
            Project.SaveXmlData(txtDateCalcSettings.Text, XDoc)
            DateCalcSettingsFile = txtDateCalcSettings.Text
        End If
    End Sub

    Private Sub btnFindDateCalcSettings_Click(sender As Object, e As EventArgs) Handles btnFindDateCalcSettings.Click
        'Find a Date Calculations Settings file.

        Select Case Project.DataLocn.Type
            Case ADVL_Utilities_Library_1.FileLocation.Types.Directory
                'Select a Date Calculations settings file from the project directory:
                OpenFileDialog1.InitialDirectory = Project.DataLocn.Path
                OpenFileDialog1.Filter = "Date Calculations settings file | *.DateCalcs"
                If OpenFileDialog1.ShowDialog() = DialogResult.OK Then
                    Dim DataFileName As String = System.IO.Path.GetFileName(OpenFileDialog1.FileName)
                    DateCalcSettingsFile = DataFileName
                    txtDateCalcSettings.Text = DataFileName
                    LoadDateCalcsSettingsFile()
                End If
            Case ADVL_Utilities_Library_1.FileLocation.Types.Archive
                'Select a Date Calculations settings file from the project archive:
                'Show the zip archive file selection form:
                Zip = New ADVL_Utilities_Library_1.ZipComp
                Zip.ArchivePath = Project.DataLocn.Path
                Zip.SelectFile() 'Show the SelectFile form.
                Zip.SelectFileForm.ApplicationName = Project.ApplicationName
                Zip.SelectFileForm.SettingsLocn = Project.SettingsLocn
                Zip.SelectFileForm.Show()
                Zip.SelectFileForm.RestoreFormSettings()
                Zip.SelectFileForm.FileExtension = ".DateCalcs"
                Zip.SelectFileForm.GetFileList()
                If Zip.SelectedFile <> "" Then
                    'A file has been selected
                    DateCalcSettingsFile = Zip.SelectedFile
                    txtDateCalcSettings.Text = Zip.SelectedFile
                    LoadDateCalcsSettingsFile()
                End If
        End Select
    End Sub

    Private Sub LoadDateCalcsSettingsFile()
        'Load the Date Calculations Settings file.
        'DateCalcSettingsFile contains the file name.

        If DateCalcSettingsFile = "" Then
            Message.AddWarning("No Date Calculations settings file name has been specified!" & vbCrLf)
        Else
            Dim XDoc As System.Xml.Linq.XDocument
            Project.ReadXmlData(DateCalcSettingsFile, XDoc)

            If XDoc.<DateCalculationsSettings>.<SelectedDatabase>.Value <> Nothing Then cmbDateCalcDb.SelectedIndex = cmbDateCalcDb.FindStringExact(XDoc.<DateCalculationsSettings>.<SelectedDatabase>.Value)
            If XDoc.<DateCalculationsSettings>.<SelectedData>.Value <> Nothing Then cmbDateCalcData.SelectedIndex = cmbDateCalcData.FindStringExact(XDoc.<DateCalculationsSettings>.<SelectedData>.Value)
            If XDoc.<DateCalculationsSettings>.<CalculationType>.Value <> Nothing Then cmbDateCalcType.SelectedIndex = cmbDateCalcType.FindStringExact(XDoc.<DateCalculationsSettings>.<CalculationType>.Value)
            If XDoc.<DateCalculationsSettings>.<InputColumn1>.Value <> Nothing Then cmbDateCalcParam1.SelectedIndex = cmbDateCalcParam1.FindStringExact(XDoc.<DateCalculationsSettings>.<InputColumn1>.Value)
            If XDoc.<DateCalculationsSettings>.<InputColumn2>.Value <> Nothing Then cmbDateCalcParam2.SelectedIndex = cmbDateCalcParam2.FindStringExact(XDoc.<DateCalculationsSettings>.<InputColumn2>.Value)
            If XDoc.<DateCalculationsSettings>.<OutputColumn>.Value <> Nothing Then cmbDateCalcOutput.SelectedIndex = cmbDateCalcOutput.FindStringExact(XDoc.<DateCalculationsSettings>.<OutputColumn>.Value)

        End If
    End Sub

    Private Sub btnApplyDateCalcSettings_Click(sender As Object, e As EventArgs) Handles btnApplyDateCalcSettings.Click
        'Apply the date calculations
        ApplyDateCalcs()
    End Sub

    Private Sub ApplyDateCalcs()
        'Apply the date calculations

        LoadDateCalcsDsOutputData()
        'dsOutput contains the selected data

        Select Case cmbDateCalcType.SelectedItem.ToString
            Case "Date at start of month"
                'Dim MonthCol As String = cmbDateCalcParam1.SelectedItem.ToString
                Dim MonthCol As String = cmbDateCalcParam1.Text
                'Dim YearCol As String = cmbDateCalcParam2.SelectedItem.ToString
                Dim YearCol As String = cmbDateCalcParam2.Text
                'Dim DateCol As String = cmbDateCalcOutput.SelectedItem.ToString
                Dim DateCol As String = cmbDateCalcOutput.Text

                Dim Month As Integer
                Dim Year As Integer
                Dim CalcDate As DateTime

                Message.Add("Starting Date Calculations ----------------------- " & vbCrLf)
                Dim Count As Integer = 0

                'Process each row in dsOutput
                For Each item In dsOutput.Tables("myData").Rows
                    Count += 1
                    If Count Mod 100 = 0 Then Message.Add("Processing record number: " & Count & vbCrLf)

                    If IsDBNull(item(MonthCol)) Then
                        'Write null value in output date column:
                        item(DateCol) = DBNull.Value
                    Else
                        If IsDBNull(item(YearCol)) Then
                            'Write null value in output date column:
                            item(DateCol) = DBNull.Value
                        Else
                            'Calculate date at start of month:
                            Month = item(MonthCol)
                            Year = item(YearCol)
                            CalcDate = DateSerial(Year, Month, 1)
                            item(DateCol) = CalcDate
                        End If
                    End If
                Next

            Case "Date at end of month"
                'Dim MonthCol As String = cmbDateCalcParam1.SelectedItem.ToString
                Dim MonthCol As String = cmbDateCalcParam1.Text
                'Dim YearCol As String = cmbDateCalcParam2.SelectedItem.ToString
                Dim YearCol As String = cmbDateCalcParam2.Text
                'Dim DateCol As String = cmbDateCalcOutput.SelectedItem.ToString
                Dim DateCol As String = cmbDateCalcOutput.Text

                Dim Month As Integer
                Dim Year As Integer
                Dim CalcDate As DateTime

                Message.Add("Starting Date Calculations ----------------------- " & vbCrLf)
                Dim Count As Integer = 0

                'Process each row in dsOutput
                For Each item In dsOutput.Tables("myData").Rows
                    Count += 1
                    If Count Mod 100 = 0 Then Message.Add("Processing record number: " & Count & vbCrLf)

                    If IsDBNull(item(MonthCol)) Then
                        'Write null value in output date column:
                        item(DateCol) = DBNull.Value
                    Else
                        If IsDBNull(item(YearCol)) Then
                            'Write null value in output date column:
                            item(DateCol) = DBNull.Value
                        Else
                            'Calculate date at end of month:
                            Month = item(MonthCol)
                            Year = item(YearCol)
                            CalcDate = DateSerial(Year, Month + 1, 0)
                            item(DateCol) = CalcDate
                        End If
                    End If
                Next

            Case "Date of Start Date add N Days"
                'Dim StartDateCol As String = cmbDateCalcParam1.SelectedItem.ToString
                Dim StartDateCol As String = cmbDateCalcParam1.Text
                'Dim NDaysCol As String = cmbDateCalcParam2.SelectedItem.ToString
                Dim NDaysCol As String = cmbDateCalcParam2.Text
                'Dim DateCol As String = cmbDateCalcOutput.SelectedItem.ToString
                Dim DateCol As String = cmbDateCalcOutput.Text

                Dim StartDate As DateTime
                Dim NDays As Double
                Dim CalcDate As DateTime

                Message.Add("Starting Date Calculations ----------------------- " & vbCrLf)
                Dim Count As Integer = 0

                'Process each row in dsOutput
                For Each item In dsOutput.Tables("myData").Rows
                    Count += 1
                    If Count Mod 100 = 0 Then Message.Add("Processing record number: " & Count & vbCrLf)

                    If IsDBNull(item(StartDateCol)) Then
                        'Write null value in output date column:
                        item(DateCol) = DBNull.Value
                    Else
                        If IsDBNull(item(NDaysCol)) Then
                            'Write null value in output date column:
                            item(DateCol) = DBNull.Value
                        Else
                            'Calculate date add NDays:
                            StartDate = item(StartDateCol)
                            NDays = item(NDaysCol)
                            CalcDate = StartDate.AddDays(NDays)
                            item(DateCol) = CalcDate
                        End If
                    End If
                Next

            Case "Date of Start Date minus N Days"
                'Dim StartDateCol As String = cmbDateCalcParam1.SelectedItem.ToString
                Dim StartDateCol As String = cmbDateCalcParam1.Text
                'Dim NDaysCol As String = cmbDateCalcParam2.SelectedItem.ToString
                Dim NDaysCol As String = cmbDateCalcParam2.Text
                'Dim DateCol As String = cmbDateCalcOutput.SelectedItem.ToString
                Dim DateCol As String = cmbDateCalcOutput.Text

                Dim StartDate As DateTime
                Dim NDays As Double
                Dim CalcDate As DateTime

                Message.Add("Starting Date Calculations ----------------------- " & vbCrLf)
                Dim Count As Integer = 0

                'Process each row in dsOutput
                For Each item In dsOutput.Tables("myData").Rows
                    Count += 1
                    If Count Mod 100 = 0 Then Message.Add("Processing record number: " & Count & vbCrLf)

                    If IsDBNull(item(StartDateCol)) Then
                        'Write null value in output date column:
                        item(DateCol) = DBNull.Value
                    Else
                        If IsDBNull(item(NDaysCol)) Then
                            'Write null value in output date column:
                            item(DateCol) = DBNull.Value
                        Else
                            'Calculate date add NDays:
                            StartDate = item(StartDateCol)
                            NDays = item(NDaysCol)
                            CalcDate = StartDate.AddDays(-NDays)
                            item(DateCol) = CalcDate
                        End If
                    End If
                Next

            Case "Fixed Date"

                Dim DateFormatString As String = txtDateFormatString.Text
                Dim FixedDate As DateTime = DateTime.ParseExact(txtFixedDate.Text, DateFormatString, Nothing)
                'Dim DateCol As String = cmbDateCalcOutput.SelectedItem.ToString
                Dim DateCol As String = cmbDateCalcOutput.Text

                Message.Add("Starting Date Calculations ----------------------- " & vbCrLf)
                Dim Count As Integer = 0

                'Process each row in dsOutput
                For Each item In dsOutput.Tables("myData").Rows
                    Count += 1
                    If Count Mod 100 = 0 Then Message.Add("Processing record number: " & Count & vbCrLf)

                    item(DateCol) = FixedDate

                    'If IsDBNull(item(StartDateCol)) Then
                    '    'Write null value in output date column:
                    '    item(DateCol) = DBNull.Value
                    'Else
                    '    If IsDBNull(item(NDaysCol)) Then
                    '        'Write null value in output date column:
                    '        item(DateCol) = DBNull.Value
                    '    Else
                    '        'Calculate date add NDays:
                    '        StartDate = item(StartDateCol)
                    '        NDays = item(NDaysCol)
                    '        CalcDate = StartDate.AddDays(-NDays)
                    '        item(DateCol) = CalcDate
                    '    End If
                    'End If
                Next

            Case Else
                Message.AddWarning("Unknown Date Calculation type: " & cmbDateCalcType.SelectedItem.ToString & vbCrLf)
                Exit Sub
        End Select

        Try
            Message.Add("Updating Database ------------------ " & vbCrLf)
            outputDa.Update(dsOutput.Tables("myData"))
            Message.Add("Date Calculations Complete ----------------- " & vbCrLf)
        Catch ex As Exception
            Message.AddWarning("Error updating database table: " & ex.Message & vbCrLf)
        End Try

        'Clear the data when finished:
        dsOutput.Clear()
        dsOutput.Tables.Clear()
    End Sub

    Private Sub LoadDateCalcsDsOutputData()
        'Open selected data in dsOutput.

        dsOutput.Clear()
        dsOutput.Tables.Clear()

        Select Case cmbDateCalcDb.SelectedItem.ToString
            Case "Share Prices"
                'outputQuery = SharePricesSettings.List(cmbDateCalcData.SelectedIndex).Query
                outputQuery = txtDateCalcsQuery.Text
                outputConnString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source =" & SharePriceDbPath
                outputConnection.ConnectionString = outputConnString
                outputConnection.Open()
                outputDa = New OleDb.OleDbDataAdapter(outputQuery, outputConnection)
                outputDa.MissingSchemaAction = MissingSchemaAction.AddWithKey
                Dim outputCmdBuilder As New OleDb.OleDbCommandBuilder(outputDa)
                outputDa.Fill(dsOutput, "myData")
                outputConnection.Close()
            Case "Financials"
                'outputQuery = FinancialsSettings.List(cmbDateCalcData.SelectedIndex).Query
                outputQuery = txtDateCalcsQuery.Text
                outputConnString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source =" & FinancialsDbPath
                outputConnection.ConnectionString = outputConnString
                outputConnection.Open()
                outputDa = New OleDb.OleDbDataAdapter(outputQuery, outputConnection)
                outputDa.MissingSchemaAction = MissingSchemaAction.AddWithKey
                Dim outputCmdBuilder As New OleDb.OleDbCommandBuilder(outputDa)
                outputDa.Fill(dsOutput, "myData")
                outputConnection.Close()
            Case "Calculations"
                'outputQuery = CalculationsSettings.List(cmbDateCalcData.SelectedIndex).Query
                outputQuery = txtDateCalcsQuery.Text
                outputConnString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source =" & CalculationsDbPath
                outputConnection.ConnectionString = outputConnString
                outputConnection.Open()
                outputDa = New OleDb.OleDbDataAdapter(outputQuery, outputConnection)
                outputDa.MissingSchemaAction = MissingSchemaAction.AddWithKey
                Dim outputCmdBuilder As New OleDb.OleDbCommandBuilder(outputDa)
                outputDa.Fill(dsOutput, "myData")
                outputConnection.Close()
        End Select
    End Sub

    Private Sub cmbDateCalcType_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbDateCalcType.SelectedIndexChanged

        Select Case cmbDateCalcType.SelectedItem.ToString
            Case "Date at start of month"
                'Enable Input Column 1 and Input Column 2:
                Label36.Enabled = True
                Label40.Enabled = True
                cmbDateCalcParam1.Enabled = True
                Label37.Enabled = True
                Label41.Enabled = True
                cmbDateCalcParam2.Enabled = True

                Label40.Text = "Month"
                Label41.Text = "Year"
                Label42.Text = "Date"

                'Disable Fixed Date inputs:
                Label65.Enabled = False
                txtFixedDate.Enabled = False
                Label66.Enabled = False
                txtDateFormatString.Enabled = False

            Case "Date at end of month"
                'Enable Input Column 1 and Input Column 2:
                Label36.Enabled = True
                Label40.Enabled = True
                cmbDateCalcParam1.Enabled = True
                Label37.Enabled = True
                Label41.Enabled = True
                cmbDateCalcParam2.Enabled = True

                Label40.Text = "Month"
                Label41.Text = "Year"
                Label42.Text = "Date"

                'Disable Fixed Date inputs:
                Label65.Enabled = False
                txtFixedDate.Enabled = False
                Label66.Enabled = False
                txtDateFormatString.Enabled = False

            Case "Date of Start Date add N Days"
                'Enable Input Column 1 and Input Column 2:
                Label36.Enabled = True
                Label40.Enabled = True
                cmbDateCalcParam1.Enabled = True
                Label37.Enabled = True
                Label41.Enabled = True
                cmbDateCalcParam2.Enabled = True

                Label40.Text = "Start Date"
                Label41.Text = "N Days"
                Label42.Text = "Date"

                'Disable Fixed Date inputs:
                Label65.Enabled = False
                txtFixedDate.Enabled = False
                Label66.Enabled = False
                txtDateFormatString.Enabled = False

            Case "Date of Start Date minus N Days"
                'Enable Input Column 1 and Input Column 2:
                Label36.Enabled = True
                Label40.Enabled = True
                cmbDateCalcParam1.Enabled = True
                Label37.Enabled = True
                Label41.Enabled = True
                cmbDateCalcParam2.Enabled = True

                Label40.Text = "Start Date"
                Label41.Text = "N Days"
                Label42.Text = "Date"

                'Disable Fixed Date inputs:
                Label65.Enabled = False
                txtFixedDate.Enabled = False
                Label66.Enabled = False
                txtDateFormatString.Enabled = False

            Case "Fixed Date"

                'Enable Fixed Date inputs:
                Label65.Enabled = True
                txtFixedDate.Enabled = True
                Label66.Enabled = True
                txtDateFormatString.Enabled = True

                'Disable Input Column 1 and Input Column 2:
                Label36.Enabled = False
                Label40.Enabled = False
                cmbDateCalcParam1.Enabled = False
                Label37.Enabled = False
                Label41.Enabled = False
                cmbDateCalcParam2.Enabled = False

        End Select
    End Sub

    Private Sub cmbDateCalcData_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbDateCalcData.SelectedIndexChanged
        SetUpDateCalcsTab()
    End Sub

    Private Sub btnAddDateCalcsToSequence_Click(sender As Object, e As EventArgs) Handles btnAddDateCalcsToSequence.Click
        'Add the Date Calculations sequence to the Processing Sequence.

        If IsNothing(Sequence) Then
            Message.AddWarning("The Processing Sequence form is not open." & vbCrLf)
        Else
            'Write new instructions to the Sequence text.
            Dim Posn As Integer
            Posn = Sequence.rtbSequence.SelectionStart

            Sequence.rtbSequence.SelectedText = vbCrLf & "<!--Date Calculations: Settings used to perform date calculations on data in the selected table :-->" & vbCrLf
            Sequence.rtbSequence.SelectedText = "  <DateCalculations>" & vbCrLf

            'Selected data parameters:
            Sequence.rtbSequence.SelectedText = "    <SelectedDatabase>" & cmbDateCalcDb.SelectedItem.ToString & "</SelectedDatabase>" & vbCrLf
            Select Case cmbSimpleCalcDb.SelectedItem.ToString
                Case "Share Prices"
                    Sequence.rtbSequence.SelectedText = "    <SelectedDatabasePath>" & SharePriceDbPath & "</SelectedDatabasePath>" & vbCrLf
                Case "Financials"
                    Sequence.rtbSequence.SelectedText = "    <SelectedDatabasePath>" & FinancialsDbPath & "</SelectedDatabasePath>" & vbCrLf
                Case "Calculations"
                    Sequence.rtbSequence.SelectedText = "    <SelectedDatabasePath>" & CalculationsDbPath & "</SelectedDatabasePath>" & vbCrLf
            End Select
            Sequence.rtbSequence.SelectedText = "    <DataQuery>" & txtDateCalcsQuery.Text & "</DataQuery>" & vbCrLf

            Sequence.rtbSequence.SelectedText = "    <CalculationType>" & cmbDateCalcType.SelectedItem.ToString & "</CalculationType>" & vbCrLf

            Select Case cmbDateCalcType.SelectedItem.ToString
                Case "Date at start of month"
                    Sequence.rtbSequence.SelectedText = "    <MonthColumn>" & cmbDateCalcParam1.SelectedItem.ToString & "</MonthColumn>" & vbCrLf
                    Sequence.rtbSequence.SelectedText = "    <YearColumn>" & cmbDateCalcParam2.SelectedItem.ToString & "</YearColumn>" & vbCrLf
                    Sequence.rtbSequence.SelectedText = "    <OutputDateColumn>" & cmbDateCalcOutput.SelectedItem.ToString & "</OutputDateColumn>" & vbCrLf

                Case "Date at end of month"
                    Sequence.rtbSequence.SelectedText = "    <MonthColumn>" & cmbDateCalcParam1.SelectedItem.ToString & "</MonthColumn>" & vbCrLf
                    Sequence.rtbSequence.SelectedText = "    <YearColumn>" & cmbDateCalcParam2.SelectedItem.ToString & "</YearColumn>" & vbCrLf
                    Sequence.rtbSequence.SelectedText = "    <OutputDateColumn>" & cmbDateCalcOutput.SelectedItem.ToString & "</OutputDateColumn>" & vbCrLf

                Case "Date of Start Date add N Days"
                    Sequence.rtbSequence.SelectedText = "    <StartDateColumn>" & cmbDateCalcParam1.SelectedItem.ToString & "</StartDateColumn>" & vbCrLf
                    Sequence.rtbSequence.SelectedText = "    <NDaysColumn>" & cmbDateCalcParam2.SelectedItem.ToString & "</NDaysColumn>" & vbCrLf
                    Sequence.rtbSequence.SelectedText = "    <OutputDateColumn>" & cmbDateCalcOutput.SelectedItem.ToString & "</OutputDateColumn>" & vbCrLf

                Case "Date of Start Date minus N Days"
                    Sequence.rtbSequence.SelectedText = "    <StartDateColumn>" & cmbDateCalcParam1.SelectedItem.ToString & "</StartDateColumn>" & vbCrLf
                    Sequence.rtbSequence.SelectedText = "    <NDaysColumn>" & cmbDateCalcParam2.SelectedItem.ToString & "</NDaysColumn>" & vbCrLf
                    Sequence.rtbSequence.SelectedText = "    <OutputDateColumn>" & cmbDateCalcOutput.SelectedItem.ToString & "</OutputDateColumn>" & vbCrLf

                Case "Fixed Date"
                    Sequence.rtbSequence.SelectedText = "    <FixedDate>" & txtFixedDate.Text & "</FixedDate>" & vbCrLf
                    Sequence.rtbSequence.SelectedText = "    <DateFormatString>" & txtDateFormatString.Text & "</DateFormatString>" & vbCrLf
                    Sequence.rtbSequence.SelectedText = "    <OutputDateColumn>" & cmbDateCalcOutput.SelectedItem.ToString & "</OutputDateColumn>" & vbCrLf

            End Select
            Sequence.rtbSequence.SelectedText = "    <Command>" & "Apply" & "</Command>" & vbCrLf
            Sequence.rtbSequence.SelectedText = "  </DateCalculations>" & vbCrLf
            Sequence.rtbSequence.SelectedText = "<!---->" & vbCrLf
            Sequence.FormatXmlText()
        End If

    End Sub

#End Region 'Date Calculations Sub Tab --------------------------------------------------------------------------------------------------------------------------------------------------------

#Region " Date Selections Sub Tab" '===========================================================================================================================================================

    Private Sub dgvDateSelectData_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles dgvDateSelectData.DataError
        'This code stops an unnecessary error dialog appearing.
        If e.Context = DataGridViewDataErrorContexts.Formatting Or e.Context = DataGridViewDataErrorContexts.PreferredSize Then
            e.ThrowException = False
        End If
    End Sub

    Private Sub dgvDateSelectConstraints_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles dgvDateSelectConstraints.DataError
        'This code stops an unnecessary error dialog appearing.
        If e.Context = DataGridViewDataErrorContexts.Formatting Or e.Context = DataGridViewDataErrorContexts.PreferredSize Then
            e.ThrowException = False
        End If
    End Sub

    Private Sub SetUpDateSelectionsTab()
        'Set up the Date Selections tab.

        'Set up dgvDateSelectData ----------------------------------------------------
        dgvDateSelectData.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
        dgvDateSelectData.Columns.Clear()

        cmbDateSelInputDateCol.Items.Clear()
        cmbDateSelOutputDateCol.Items.Clear()

        txtDateSelSettings.Text = DateSelectSettingsFile

        Dim ComboBoxCol0 As New DataGridViewComboBoxColumn
        dgvDateSelectData.Columns.Add(ComboBoxCol0)
        dgvDateSelectData.Columns(0).HeaderText = "Input Column"

        Select Case cmbDateSelectInputDb.SelectedItem.ToString
            Case "Share Prices"
                For Each item In SharePricesSettings.List(cmbDateSelectInputData.SelectedIndex).TableCols
                    ComboBoxCol0.Items.Add(item)
                    cmbDateSelInputDateCol.Items.Add(item)
                Next
                txtDateSelInputQuery.Text = SharePricesSettings.List(cmbDateSelectInputData.SelectedIndex).Query
            Case "Financials"
                For Each item In FinancialsSettings.List(cmbDateSelectInputData.SelectedIndex).TableCols
                    ComboBoxCol0.Items.Add(item)
                    cmbDateSelInputDateCol.Items.Add(item)
                Next
                txtDateSelInputQuery.Text = FinancialsSettings.List(cmbDateSelectInputData.SelectedIndex).Query
            Case "Calculations"
                For Each item In CalculationsSettings.List(cmbDateSelectInputData.SelectedIndex).TableCols
                    ComboBoxCol0.Items.Add(item)
                    cmbDateSelInputDateCol.Items.Add(item)
                Next
                txtDateSelInputQuery.Text = CalculationsSettings.List(cmbDateSelectInputData.SelectedIndex).Query
        End Select

        Dim ComboBoxCol1 As New DataGridViewComboBoxColumn
        dgvDateSelectData.Columns.Add(ComboBoxCol1)
        dgvDateSelectData.Columns(1).HeaderText = "Output Column"
        'dgvSelectData.Columns(1).Width = 160

        If cmbDateSelectOutputDb.SelectedIndex = -1 Then
            cmbDateSelectOutputDb.SelectedIndex = 0
        End If
        If cmbDateSelectOutputData.SelectedIndex = -1 Then
            cmbDateSelectOutputData.SelectedIndex = 0
        End If

        Select Case cmbDateSelectOutputDb.SelectedItem.ToString
            Case "Share Prices"
                For Each item In SharePricesSettings.List(cmbDateSelectOutputData.SelectedIndex).TableCols
                    ComboBoxCol1.Items.Add(item)
                    cmbDateSelOutputDateCol.Items.Add(item)
                Next
                txtDateSelOutputQuery.Text = SharePricesSettings.List(cmbDateSelectOutputData.SelectedIndex).Query
            Case "Financials"
                For Each item In FinancialsSettings.List(cmbDateSelectOutputData.SelectedIndex).TableCols
                    ComboBoxCol1.Items.Add(item)
                    cmbDateSelOutputDateCol.Items.Add(item)
                Next
                txtDateSelOutputQuery.Text = FinancialsSettings.List(cmbDateSelectOutputData.SelectedIndex).Query
            Case "Calculations"
                For Each item In CalculationsSettings.List(cmbDateSelectOutputData.SelectedIndex).TableCols
                    ComboBoxCol1.Items.Add(item)
                    cmbDateSelOutputDateCol.Items.Add(item)
                Next
                txtDateSelOutputQuery.Text = CalculationsSettings.List(cmbDateSelectOutputData.SelectedIndex).Query
        End Select

        'Set up dgvDateSelectConstraints ---------------------------------------------
        dgvDateSelectConstraints.Columns.Clear()

        Dim ComboBoxCol20 As New DataGridViewComboBoxColumn
        dgvDateSelectConstraints.Columns.Add(ComboBoxCol20)
        dgvDateSelectConstraints.Columns(0).HeaderText = "WHERE Input Column"

        Select Case cmbDateSelectInputDb.SelectedItem.ToString
            Case "Share Prices"
                For Each item In SharePricesSettings.List(cmbDateSelectInputData.SelectedIndex).TableCols
                    ComboBoxCol20.Items.Add(item)
                Next
            Case "Financials"
                For Each item In FinancialsSettings.List(cmbDateSelectInputData.SelectedIndex).TableCols
                    ComboBoxCol20.Items.Add(item)
                Next
            Case "Calculations"
                For Each item In CalculationsSettings.List(cmbDateSelectInputData.SelectedIndex).TableCols
                    ComboBoxCol20.Items.Add(item)
                Next
        End Select


        Dim ComboBoxCol21 As New DataGridViewComboBoxColumn
        dgvDateSelectConstraints.Columns.Add(ComboBoxCol21)
        dgvDateSelectConstraints.Columns(1).HeaderText = "= Output Column"

        Select Case cmbDateSelectOutputDb.SelectedItem.ToString
            Case "Share Prices"
                For Each item In SharePricesSettings.List(cmbDateSelectOutputData.SelectedIndex).TableCols
                    ComboBoxCol21.Items.Add(item)
                Next
            Case "Financials"
                For Each item In FinancialsSettings.List(cmbDateSelectOutputData.SelectedIndex).TableCols
                    ComboBoxCol21.Items.Add(item)
                Next
            Case "Calculations"
                For Each item In CalculationsSettings.List(cmbDateSelectOutputData.SelectedIndex).TableCols
                    ComboBoxCol21.Items.Add(item)
                Next
        End Select




        'NOTE: Dont load the saved settings because different settings may be required!!!
        'LoadCopyColumnsSettingsFile()
        'LoadSelectDataSettingsFile() 

        'RestoreSelectDataSelections()
        RestoreDateSelectSelections()

        'Resize the columns to fit the data:
        dgvDateSelectData.AutoResizeColumns()
        dgvDateSelectConstraints.AutoResizeColumns()

        'This code allows user to resize column widths:
        'dgvDateSelectData.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
        'dgvDateSelectConstraints.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

    End Sub

    Private Sub RestoreDateSelectSelections()
        'Restore the Date Selections settings. (Leave the input and output data selections unchanged.)

        If DateSelectSettingsFile = "" Then
            'Message.AddWarning("No Date Selections settings file name has been specified!" & vbCrLf)
        Else
            Dim XDoc As System.Xml.Linq.XDocument
            Project.ReadXmlData(DateSelectSettingsFile, XDoc)

            'Don't update these selections: (Only dgvCopyData is being updated.)
            'If XDoc.<CopyColumnsSettings>.<InputDatabase>.Value <> Nothing Then cmbCopyDataInputDb.SelectedIndex = cmbCopyDataInputDb.FindStringExact(XDoc.<CopyColumnsSettings>.<InputDatabase>.Value)
            'If XDoc.<CopyColumnsSettings>.<InputData>.Value <> Nothing Then cmbCopyDataInputData.SelectedIndex = cmbCopyDataInputData.FindStringExact(XDoc.<CopyColumnsSettings>.<InputData>.Value)
            'If XDoc.<CopyColumnsSettings>.<OutputDatabase>.Value <> Nothing Then cmbCopyDataOutputDb.SelectedIndex = cmbCopyDataOutputDb.FindStringExact(XDoc.<CopyColumnsSettings>.<OutputDatabase>.Value)
            'If XDoc.<CopyColumnsSettings>.<OutputTable>.Value <> Nothing Then cmbCopyDataOutputData.SelectedIndex = cmbCopyDataOutputData.FindStringExact(XDoc.<CopyColumnsSettings>.<OutputTable>.Value)

            dgvDateSelectData.Rows.Clear()

            Dim settings = From item In XDoc.<DateSelectSettings>.<SelectDataList>.<CopyColumn>

            For Each item In settings
                dgvDateSelectData.Rows.Add(item.<From>.Value, item.<To>.Value)
            Next

            dgvDateSelectData.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            dgvDateSelectData.AutoResizeColumns()

            dgvDateSelectConstraints.Rows.Clear()
            Dim settings2 = From item In XDoc.<DateSelectSettings>.<SelectConstraintsList>.<Constraint>
            For Each item In settings2
                dgvDateSelectConstraints.Rows.Add(item.<WhereInputColumn>.Value, item.<EqualsOutputColumn>.Value)
            Next
            dgvDateSelectConstraints.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            dgvDateSelectConstraints.AutoResizeColumns()
            dgvDateSelectData.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            dgvDateSelectConstraints.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

        End If
    End Sub

    Private Sub GetDateSelectInputDataList()
        'Fill the cmbDateSelectInputData list of data views.

        cmbDateSelectInputData.Items.Clear()
        Dim I As Integer

        Select Case cmbDateSelectInputDb.SelectedItem.ToString
            Case "Share Prices"
                'SharePriceDataViewList
                For I = 0 To SharePricesSettings.NRecords - 1
                    cmbDateSelectInputData.Items.Add(SharePricesSettings.List(I).Description)
                Next

            Case "Financials"
                'FinancialsDataViewList
                For I = 0 To FinancialsSettings.NRecords - 1
                    cmbDateSelectInputData.Items.Add(FinancialsSettings.List(I).Description)
                Next

            Case "Calculations"
                'CalculationsDataViewList
                For I = 0 To CalculationsSettings.NRecords - 1
                    cmbDateSelectInputData.Items.Add(CalculationsSettings.List(I).Description)
                Next
        End Select
    End Sub

    Private Sub cmbDateSelectInputDb_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbDateSelectInputDb.SelectedIndexChanged
        GetDateSelectInputDataList()
    End Sub

    Private Sub cmbDateSelectInputData_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbDateSelectInputData.SelectedIndexChanged
        SetUpDateSelectionsTab()
    End Sub

    Private Sub cmbDateSelectOutputDb_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbDateSelectOutputDb.SelectedIndexChanged
        GetDateSelectOutputDataList
    End Sub

    Private Sub GetDateSelectOutputDataList()
        'Fill the cmbSelectDataOutputData list.

        cmbDateSelectOutputData.Items.Clear()
        Dim I As Integer

        Select Case cmbDateSelectOutputDb.SelectedItem.ToString
            Case "Share Prices"
                'SharePriceDataViewList
                For I = 0 To SharePricesSettings.NRecords - 1
                    cmbDateSelectOutputData.Items.Add(SharePricesSettings.List(I).Description)
                Next
            Case "Financials"
                'FinancialsDataViewList
                For I = 0 To FinancialsSettings.NRecords - 1
                    cmbDateSelectOutputData.Items.Add(FinancialsSettings.List(I).Description)
                Next
            Case "Calculations"
                'CalculationsDataViewList
                For I = 0 To CalculationsSettings.NRecords - 1
                    cmbDateSelectOutputData.Items.Add(CalculationsSettings.List(I).Description)
                Next
        End Select

    End Sub

    Private Sub cmbDateSelectOutputData_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbDateSelectOutputData.SelectedIndexChanged
        SetUpDateSelectionsTab()
    End Sub



    Private Sub cmbDateSelInputDateCol_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbDateSelInputDateCol.SelectedIndexChanged

    End Sub

    Private Sub cmbDateSelOutputDateCol_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbDateSelOutputDateCol.SelectedIndexChanged

    End Sub

    Private Sub btnAddDateSelectToSequence_Click(sender As Object, e As EventArgs) Handles btnAddDateSelectToSequence.Click
        'Add the Date Select sequence to the Processing Sequence.

        If IsNothing(Sequence) Then
            Message.AddWarning("The Processing Sequence form is not open." & vbCrLf)
        Else
            'Write new instructions to the Sequence text.
            Dim Posn As Integer
            Posn = Sequence.rtbSequence.SelectionStart

            Sequence.rtbSequence.SelectedText = vbCrLf & "<!--Date Select: Settings used to select data based on dates from an Input table and copy it to an Output table :-->" & vbCrLf
            Sequence.rtbSequence.SelectedText = "  <DateSelect>" & vbCrLf


            'Input data parameters:
            Sequence.rtbSequence.SelectedText = "    <InputDatabase>" & cmbDateSelectInputDb.SelectedItem.ToString & "</InputDatabase>" & vbCrLf
            Select Case cmbDateSelectInputDb.SelectedItem.ToString
                Case "Share Prices"
                    Sequence.rtbSequence.SelectedText = "    <InputDatabasePath>" & SharePriceDbPath & "</InputDatabasePath>" & vbCrLf
                Case "Financials"
                    Sequence.rtbSequence.SelectedText = "    <InputDatabasePath>" & FinancialsDbPath & "</InputDatabasePath>" & vbCrLf
                Case "Calculations"
                    Sequence.rtbSequence.SelectedText = "    <InputDatabasePath>" & CalculationsDbPath & "</InputDatabasePath>" & vbCrLf
            End Select
            Sequence.rtbSequence.SelectedText = "    <InputQuery>" & txtDateSelInputQuery.Text & "</InputQuery>" & vbCrLf

            'Output data parameters:
            Sequence.rtbSequence.SelectedText = "    <OutputDatabase>" & cmbDateSelectOutputDb.SelectedItem.ToString & "</OutputDatabase>" & vbCrLf
            Select Case cmbDateSelectOutputDb.SelectedItem.ToString
                Case "Share Prices"
                    Sequence.rtbSequence.SelectedText = "    <OutputDatabasePath>" & SharePriceDbPath & "</OutputDatabasePath>" & vbCrLf
                Case "Financials"
                    Sequence.rtbSequence.SelectedText = "    <OutputDatabasePath>" & FinancialsDbPath & "</OutputDatabasePath>" & vbCrLf
                Case "Calculations"
                    Sequence.rtbSequence.SelectedText = "    <OutputDatabasePath>" & CalculationsDbPath & "</OutputDatabasePath>" & vbCrLf
            End Select
            Sequence.rtbSequence.SelectedText = "    <OutputQuery>" & txtDateSelOutputQuery.Text & "</OutputQuery>" & vbCrLf

            'Date Selection Type:
            'Sequence.rtbSequence.SelectedText = "    <DateSelectionType>" & cmbDateSelectionType.SelectedItem.ToString & "</DateSelectionType>" & vbCrLf
            Sequence.rtbSequence.SelectedText = "    <DateSelectionType>" & cmbDateSelectionType.Text & "</DateSelectionType>" & vbCrLf
            'Sequence.rtbSequence.SelectedText = "    <InputDateColumn>" & cmbDateSelInputDateCol.SelectedItem.ToString & "</InputDateColumn>" & vbCrLf
            Sequence.rtbSequence.SelectedText = "    <InputDateColumn>" & cmbDateSelInputDateCol.Text & "</InputDateColumn>" & vbCrLf
            ' Sequence.rtbSequence.SelectedText = "    <OutputDateColumn>" & cmbDateSelOutputDateCol.SelectedItem.ToString & "</OutputDateColumn>" & vbCrLf
            Sequence.rtbSequence.SelectedText = "    <OutputDateColumn>" & cmbDateSelOutputDateCol.Text & "</OutputDateColumn>" & vbCrLf

            'Select constraints list:
            Sequence.rtbSequence.SelectedText = "    <SelectConstraintList>" & vbCrLf
            dgvDateSelectConstraints.AllowUserToAddRows = False 'This removes the last edit row from the DataGridView.
            Dim NRows As Integer = dgvDateSelectConstraints.Rows.Count
            Dim RowNo As Integer
            For RowNo = 0 To NRows - 1
                Sequence.rtbSequence.SelectedText = "      <Constraint>" & vbCrLf
                Sequence.rtbSequence.SelectedText = "        <WhereInputColumn>" & dgvDateSelectConstraints.Rows(RowNo).Cells(0).Value & "</WhereInputColumn>" & vbCrLf
                Sequence.rtbSequence.SelectedText = "        <EqualsOutputColumn>" & dgvDateSelectConstraints.Rows(RowNo).Cells(1).Value & "</EqualsOutputColumn>" & vbCrLf
                Sequence.rtbSequence.SelectedText = "      </Constraint>" & vbCrLf
            Next
            dgvDateSelectConstraints.AllowUserToAddRows = True 'Allow user to add rows again.
            Sequence.rtbSequence.SelectedText = "    </SelectConstraintList>" & vbCrLf

            'List of columns to copy:
            Sequence.rtbSequence.SelectedText = "    <SelectDataList>" & vbCrLf
            dgvDateSelectData.AllowUserToAddRows = False 'This removes the last edit row from the DataGridView.
            NRows = dgvDateSelectData.Rows.Count
            For RowNo = 0 To NRows - 1
                Sequence.rtbSequence.SelectedText = "      <CopyColumn>" & vbCrLf
                Sequence.rtbSequence.SelectedText = "        <From>" & dgvDateSelectData.Rows(RowNo).Cells(0).Value & "</From>" & vbCrLf
                Sequence.rtbSequence.SelectedText = "        <To>" & dgvDateSelectData.Rows(RowNo).Cells(1).Value & "</To>" & vbCrLf
                Sequence.rtbSequence.SelectedText = "      </CopyColumn>" & vbCrLf
            Next
            dgvDateSelectData.AllowUserToAddRows = True 'Allow user to add rows again.
            Sequence.rtbSequence.SelectedText = "    </SelectDataList>" & vbCrLf

            Sequence.rtbSequence.SelectedText = "    <Command>Apply</Command>" & vbCrLf
            Sequence.rtbSequence.SelectedText = "  </DateSelect>" & vbCrLf
            Sequence.rtbSequence.SelectedText = "<!---->" & vbCrLf

            Sequence.FormatXmlText()
        End If

    End Sub

    Private Sub btnSaveDateSelSettings_Click(sender As Object, e As EventArgs) Handles btnSaveDateSelSettings.Click
        'Save the Date Select Settings.

        If Trim(txtDateSelSettings.Text) = "" Then
            Message.AddWarning("No file name has been specified!" & vbCrLf)
        Else
            If txtDateSelSettings.Text.EndsWith(".DateSelect") Then
                txtDateSelSettings.Text = Trim(txtDateSelSettings.Text)
            Else
                txtDateSelSettings.Text = Trim(txtDateSelSettings.Text) & ".DateSelect"
            End If

            dgvDateSelectData.AllowUserToAddRows = False 'This removes the last edit row from the DataGridView.
            dgvDateSelectConstraints.AllowUserToAddRows = False

            Try
                Dim myXDoc = <?xml version="1.0" encoding="utf-8"?>
                             <DateSelectSettings>
                                 <!---->
                                 <!--Date Select Settings-->
                                 <!---->
                                 <HostApplication><%= ApplicationInfo.Name %></HostApplication>
                                 <InputDatabase><%= cmbDateSelectInputDb.SelectedItem.ToString %></InputDatabase>
                                 <InputData><%= cmbDateSelectInputData.SelectedItem.ToString %></InputData>
                                 <OutputDatabase><%= cmbDateSelectOutputDb.SelectedItem.ToString %></OutputDatabase>
                                 <OutputTable><%= cmbDateSelectOutputData.SelectedItem.ToString %></OutputTable>
                                 <DateSelectionType><%= cmbDateSelectionType.SelectedItem.ToString %></DateSelectionType>
                                 <InputDateColumn><%= cmbDateSelInputDateCol.SelectedItem.ToString %></InputDateColumn>
                                 <OutputDateColumn><%= cmbDateSelOutputDateCol.SelectedItem.ToString %></OutputDateColumn>
                                 <SelectDataList>
                                     <%= From item In dgvDateSelectData.Rows
                                         Select
                                                   <CopyColumn>
                                                       <From><%= item.Cells(0).Value %></From>
                                                       <To><%= item.Cells(1).Value %></To>
                                                   </CopyColumn>
                                     %>
                                 </SelectDataList>
                                 <SelectConstraintsList>
                                     <%= From item In dgvDateSelectConstraints.Rows
                                         Select
                                                   <Constraint>
                                                       <WhereInputColumn><%= item.Cells(0).Value %></WhereInputColumn>
                                                       <EqualsOutputColumn><%= item.Cells(1).Value %></EqualsOutputColumn>
                                                   </Constraint>
                                     %>
                                 </SelectConstraintsList>
                             </DateSelectSettings>

                Project.SaveXmlData(txtDateSelSettings.Text, myXDoc)
            Catch ex As Exception
                Message.AddWarning(ex.Message & vbCrLf)
            End Try



            dgvDateSelectData.AllowUserToAddRows = True 'Allow user to add rows again.
            dgvDateSelectConstraints.AllowUserToAddRows = True
            DateSelectSettingsFile = txtDateSelSettings.Text
        End If
    End Sub

    Private Sub btnFindDateSelSettings_Click(sender As Object, e As EventArgs) Handles btnFindDateSelSettings.Click
        'Find a Date Select Settings file.

        Select Case Project.DataLocn.Type
            Case ADVL_Utilities_Library_1.FileLocation.Types.Directory
                'Select a Date Select settings file from the project directory:
                OpenFileDialog1.InitialDirectory = Project.DataLocn.Path
                OpenFileDialog1.Filter = "Date Select settings file | *.DateSelect"
                If OpenFileDialog1.ShowDialog() = DialogResult.OK Then
                    Dim DataFileName As String = System.IO.Path.GetFileName(OpenFileDialog1.FileName)
                    DateSelectSettingsFile = DataFileName
                    txtDateSelSettings.Text = DataFileName
                    LoadDateSelectSettingsFile()
                End If
            Case ADVL_Utilities_Library_1.FileLocation.Types.Archive
                'Select a Date Select settings file from the project archive:
                'Show the zip archive file selection form:
                Zip = New ADVL_Utilities_Library_1.ZipComp
                Zip.ArchivePath = Project.DataLocn.Path
                Zip.SelectFile() 'Show the SelectFile form.
                Zip.SelectFileForm.ApplicationName = Project.ApplicationName
                Zip.SelectFileForm.SettingsLocn = Project.SettingsLocn
                Zip.SelectFileForm.Show()
                Zip.SelectFileForm.RestoreFormSettings()
                Zip.SelectFileForm.FileExtension = ".DateSelect"
                Zip.SelectFileForm.GetFileList()
                If Zip.SelectedFile <> "" Then
                    'A file has been selected
                    DateSelectSettingsFile = Zip.SelectedFile
                    txtDateSelSettings.Text = Zip.SelectedFile
                    LoadDateSelectSettingsFile()
                End If
        End Select
    End Sub

    Private Sub LoadDateSelectSettingsFile()
        'Load the Date Select Settings file.
        'DateSelectSettingsFile contains the file name.

        If DateSelectSettingsFile = "" Then
            Message.AddWarning("No Date Select settings file name has been specified!" & vbCrLf)
        Else
            Dim XDoc As System.Xml.Linq.XDocument
            Project.ReadXmlData(DateSelectSettingsFile, XDoc)

            If XDoc.<DateSelectSettings>.<InputDatabase>.Value <> Nothing Then cmbDateSelectInputDb.SelectedIndex = cmbDateSelectInputDb.FindStringExact(XDoc.<DateSelectSettings>.<InputDatabase>.Value)
            If XDoc.<DateSelectSettings>.<InputData>.Value <> Nothing Then cmbDateSelectInputData.SelectedIndex = cmbDateSelectInputData.FindStringExact(XDoc.<DateSelectSettings>.<InputData>.Value)
            If XDoc.<DateSelectSettings>.<OutputDatabase>.Value <> Nothing Then cmbDateSelectOutputDb.SelectedIndex = cmbDateSelectOutputDb.FindStringExact(XDoc.<DateSelectSettings>.<OutputDatabase>.Value)
            If XDoc.<DateSelectSettings>.<OutputTable>.Value <> Nothing Then cmbDateSelectOutputData.SelectedIndex = cmbDateSelectOutputData.FindStringExact(XDoc.<DateSelectSettings>.<OutputTable>.Value)
            If XDoc.<DateSelectSettings>.<DateSelectionType>.Value <> Nothing Then cmbDateSelectionType.SelectedIndex = cmbDateSelectionType.FindStringExact(XDoc.<DateSelectSettings>.<DateSelectionType>.Value)
            If XDoc.<DateSelectSettings>.<InputDateColumn>.Value <> Nothing Then cmbDateSelInputDateCol.SelectedIndex = cmbDateSelInputDateCol.FindStringExact(XDoc.<DateSelectSettings>.<InputDateColumn>.Value)
            If XDoc.<DateSelectSettings>.<OutputDateColumn>.Value <> Nothing Then cmbDateSelOutputDateCol.SelectedIndex = cmbDateSelOutputDateCol.FindStringExact(XDoc.<DateSelectSettings>.<OutputDateColumn>.Value)

            dgvDateSelectData.Rows.Clear()
            Dim settings = From item In XDoc.<DateSelectSettings>.<SelectDataList>.<CopyColumn>
            For Each item In settings
                dgvDateSelectData.Rows.Add(item.<From>.Value, item.<To>.Value)
            Next
            dgvDateSelectData.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            dgvDateSelectData.AutoResizeColumns()


            dgvDateSelectConstraints.Rows.Clear()
            Dim settings2 = From item In XDoc.<DateSelectSettings>.<SelectConstraintsList>.<Constraint>
            For Each item In settings2
                dgvDateSelectConstraints.Rows.Add(item.<WhereInputColumn>.Value, item.<EqualsOutputColumn>.Value)
            Next
            dgvDateSelectConstraints.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            dgvDateSelectConstraints.AutoResizeColumns()
            dgvDateSelectData.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            dgvDateSelectConstraints.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

        End If
    End Sub

    Private Sub btnApplyDateSelSettings_Click(sender As Object, e As EventArgs) Handles btnApplyDateSelSettings.Click
        'Apply the Date Select settings.

        ApplyDateSelections()

    End Sub

    Private Sub ApplyDateSelections()
        'Apply the Date Select settings.

        LoadDateSelectDsInputData()
        LoadDateSelectDsOutputData()

        'dsInput contains the input table
        'dsOutput contains the output table

        'Message.Add("dsInput records: " & dsInput.Tables("myData").Rows.Count & vbCrLf)
        'Message.Add("dsOutput records: " & dsOutput.Tables("myData").Rows.Count & vbCrLf)

        dgvDateSelectData.AllowUserToAddRows = False
        dgvDateSelectConstraints.AllowUserToAddRows = False

        Dim NCols As Integer = dgvDateSelectData.RowCount

        Dim InCols(0 To NCols) As String 'Array to contain the Input table Column names.
        Dim OutCols(0 To NCols) As String 'Array to contain the corresponding Output table Column names.

        Dim I As Integer

        For I = 1 To NCols
            InCols(I) = dgvDateSelectData.Rows(I - 1).Cells(0).Value 'The Input data column names to copy.
            OutCols(I) = dgvDateSelectData.Rows(I - 1).Cells(1).Value 'The Output table column names to paste.
        Next

        Dim NCons As Integer = dgvDateSelectConstraints.RowCount
        Dim InCons(0 To NCons) As String 'The Input data contraint column names.
        Dim OutCons(0 To NCons) As String 'The Output table constraint column names.

        For I = 1 To NCons
            InCons(I) = dgvDateSelectConstraints.Rows(I - 1).Cells(0).Value
            OutCons(I) = dgvDateSelectConstraints.Rows(I - 1).Cells(1).Value
        Next

        Dim InputDateCol As String = cmbDateSelInputDateCol.SelectedItem.ToString
        Dim OutputDateCol As String = cmbDateSelOutputDateCol.SelectedItem.ToString

        Dim myQuery As String
        Dim mySort As String

        Dim Count As Integer = 0

        Select Case cmbDateSelectionType.SelectedItem.ToString
            Case "Select Input data with Input date = Output date"
                Message.Add("Selecting Data ----------------------- " & vbCrLf)
                For Each item In dsOutput.Tables("myData").Rows
                    Count += 1
                    If Count Mod 100 = 0 Then Message.Add("Selecting record number: " & Count & vbCrLf)
                    myQuery = InCons(1) & " = '" & item(OutCons(1)) & "'"
                    For I = 2 To NCons
                        myQuery = myQuery & " And " & InCons(I) & " = '" & item(OutCons(I)) & "'"
                    Next
                    'myQuery = myQuery & " And " & InputDateCol & " = '" & item(OutputDateCol) & "'"
                    'myQuery = myQuery & " And " & InputDateCol & " = #" & item(OutputDateCol) & "#"
                    myQuery = myQuery & " And " & InputDateCol & " = #" & Format(item(OutputDateCol), "MM-dd-yyyy") & "#" 'Dates in a query must have this format!

                    Dim myRecords = dsInput.Tables("myData").Select(myQuery)
                    If myRecords.Count = 0 Then
                        Message.AddWarning("No records found with this constraint: " & myQuery & vbCrLf)
                    ElseIf myRecords.Count = 1 Then
                        'Message.Add("One record found with this constraint: " & myQuery & vbCrLf)
                        For I = 1 To NCols
                            item(OutCols(I)) = myRecords(0).Item(InCols(I))
                        Next
                    Else
                        Message.AddWarning("More than one record found with this constraint: " & myQuery & vbCrLf)
                    End If
                Next

            Case "Select first Input data after Output date"
                Message.Add("Selecting Data ----------------------- " & vbCrLf)
                For Each item In dsOutput.Tables("myData").Rows
                    Count += 1
                    If Count Mod 100 = 0 Then Message.Add("Selecting record number: " & Count & vbCrLf)
                    myQuery = InCons(1) & " = '" & item(OutCons(1)) & "'"
                    For I = 2 To NCons
                        myQuery = myQuery & " And " & InCons(I) & " = '" & item(OutCons(I)) & "'"
                    Next
                    'myQuery = myQuery & " And " & InputDateCol & " > '" & item(OutputDateCol) & "'" & " Order By " & InputDateCol
                    'myQuery = myQuery & " And " & InputDateCol & " > #" & item(OutputDateCol) & "#" & " Order By " & InputDateCol
                    'myQuery = myQuery & " And " & InputDateCol & " > #" & item(OutputDateCol) & "#"
                    myQuery = myQuery & " And " & InputDateCol & " > #" & Format(item(OutputDateCol), "MM-dd-yyyy") & "#" 'Dates in a query must have this format!
                    'mySort = InputDateCol & " DESC"
                    mySort = InputDateCol & " ASC"

                    'Dim myRecords = dsInput.Tables("myData").Select(myQuery)
                    Dim myRecords = dsInput.Tables("myData").Select(myQuery, mySort)
                    If myRecords.Count = 0 Then
                        Message.AddWarning("No records found with this constraint: " & myQuery & vbCrLf)
                    ElseIf myRecords.Count = 1 Then
                        'Message.Add("One record found with this constraint: " & myQuery & vbCrLf)
                        For I = 1 To NCols
                            item(OutCols(I)) = myRecords(0).Item(InCols(I))
                        Next
                    Else
                        'Message.AddWarning("More than one record found with this constraint: " & myQuery & vbCrLf)
                        'Use the data from the first match
                        For I = 1 To NCols
                            item(OutCols(I)) = myRecords(0).Item(InCols(I))
                        Next
                    End If
                Next

            Case "Select last Input data before Output date"
                Message.Add("Selecting Data ----------------------- " & vbCrLf)
                For Each item In dsOutput.Tables("myData").Rows
                    Count += 1
                    If Count Mod 100 = 0 Then Message.Add("Selecting record number: " & Count & vbCrLf)
                    myQuery = InCons(1) & " = '" & item(OutCons(1)) & "'"
                    For I = 2 To NCons
                        myQuery = myQuery & " And " & InCons(I) & " = '" & item(OutCons(I)) & "'"
                    Next
                    'myQuery = myQuery & " And " & InputDateCol & " < '" & item(OutputDateCol) & "'" & " Order By '" & InputDateCol & "' Desc"
                    'myQuery = myQuery & " And " & InputDateCol & " < #" & item(OutputDateCol) & "#" & " Order By " & InputDateCol & " Desc"
                    'myQuery = myQuery & " And " & InputDateCol & " < #" & item(OutputDateCol) & "#"
                    myQuery = myQuery & " And " & InputDateCol & " < #" & Format(item(OutputDateCol), "MM-dd-yyyy") & "#" 'Dates in a query must have this format!
                    'mySort = InputDateCol & " ASC"
                    mySort = InputDateCol & " DESC"

                    'Dim myRecords = dsInput.Tables("myData").Select(myQuery)
                    Dim myRecords = dsInput.Tables("myData").Select(myQuery, mySort)
                    If myRecords.Count = 0 Then
                        Message.AddWarning("No records found with this constraint: " & myQuery & vbCrLf)
                    ElseIf myRecords.Count = 1 Then
                        'Message.Add("One record found with this constraint: " & myQuery & vbCrLf)
                        For I = 1 To NCols
                            item(OutCols(I)) = myRecords(0).Item(InCols(I))
                        Next
                    Else
                        'Message.AddWarning("More than one record found with this constraint: " & myQuery & vbCrLf)
                        'Use the data from the first match
                        For I = 1 To NCols
                            item(OutCols(I)) = myRecords(0).Item(InCols(I))
                        Next
                    End If
                Next


        End Select

        Try
            Message.Add("Updating Database ------------------ " & vbCrLf)
            outputDa.Update(dsOutput.Tables("myData"))
            Message.Add("Select Data Complete ----------------- " & vbCrLf)
        Catch ex As Exception
            Message.AddWarning("Error updating database table: " & ex.Message & vbCrLf)
        End Try



        'Clear the data list when finished:
        dsInput.Clear()
        dsInput.Tables.Clear()
        dsOutput.Clear()
        dsOutput.Tables.Clear()

        dgvDateSelectData.AllowUserToAddRows = True
        dgvDateSelectConstraints.AllowUserToAddRows = True
        dgvDateSelectData.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
        dgvDateSelectConstraints.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
    End Sub

    Private Sub LoadDateSelectDsInputData()
        'Open selected input data in dsInput.

        dsInput.Clear()
        dsInput.Tables.Clear()
        Dim Query As String
        Dim connString As String
        Dim myConnection As OleDb.OleDbConnection = New OleDb.OleDbConnection
        Dim da As OleDb.OleDbDataAdapter

        Select Case cmbDateSelectInputDb.SelectedItem.ToString
            Case "Share Prices"
                'Query = SharePricesSettings.List(cmbDateSelectInputData.SelectedIndex).Query
                Query = txtDateSelInputQuery.Text
                connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source =" & SharePriceDbPath
                myConnection.ConnectionString = connString
                myConnection.Open()
                da = New OleDb.OleDbDataAdapter(Query, myConnection)
                da.MissingSchemaAction = MissingSchemaAction.AddWithKey
                da.Fill(dsInput, "myData")
                myConnection.Close()
            Case "Financials"
                'Query = FinancialsSettings.List(cmbDateSelectInputData.SelectedIndex).Query
                Query = txtDateSelInputQuery.Text
                connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source =" & FinancialsDbPath
                myConnection.ConnectionString = connString
                myConnection.Open()
                da = New OleDb.OleDbDataAdapter(Query, myConnection)
                da.MissingSchemaAction = MissingSchemaAction.AddWithKey
                da.Fill(dsInput, "myData")
                myConnection.Close()
            Case "Calculations"
                'Query = CalculationsSettings.List(cmbDateSelectInputData.SelectedIndex).Query
                Query = txtDateSelInputQuery.Text
                connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source =" & CalculationsDbPath
                myConnection.ConnectionString = connString
                myConnection.Open()
                da = New OleDb.OleDbDataAdapter(Query, myConnection)
                da.MissingSchemaAction = MissingSchemaAction.AddWithKey
                da.Fill(dsInput, "myData")
                myConnection.Close()
        End Select
    End Sub

    Private Sub LoadDateSelectDsOutputData()
        'Open selected output data in dsOutput.

        dsOutput.Clear()
        dsOutput.Tables.Clear()

        Select Case cmbDateSelectOutputDb.SelectedItem.ToString
            Case "Share Prices"
                'outputQuery = SharePricesSettings.List(cmbDateSelectOutputData.SelectedIndex).Query
                outputQuery = txtDateSelOutputQuery.Text
                outputConnString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source =" & SharePriceDbPath
                outputConnection.ConnectionString = outputConnString
                outputConnection.Open()
                outputDa = New OleDb.OleDbDataAdapter(outputQuery, outputConnection)
                outputDa.MissingSchemaAction = MissingSchemaAction.AddWithKey
                Dim outputCmdBuilder As New OleDb.OleDbCommandBuilder(outputDa)
                outputDa.Fill(dsOutput, "myData")
                outputConnection.Close()
            Case "Financials"
                'outputQuery = FinancialsSettings.List(cmbDateSelectOutputData.SelectedIndex).Query
                outputQuery = txtDateSelOutputQuery.Text
                outputConnString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source =" & FinancialsDbPath
                outputConnection.ConnectionString = outputConnString
                outputConnection.Open()
                outputDa = New OleDb.OleDbDataAdapter(outputQuery, outputConnection)
                outputDa.MissingSchemaAction = MissingSchemaAction.AddWithKey
                Dim outputCmdBuilder As New OleDb.OleDbCommandBuilder(outputDa)
                outputDa.Fill(dsOutput, "myData")
                outputConnection.Close()
            Case "Calculations"
                'outputQuery = CalculationsSettings.List(cmbDateSelectOutputData.SelectedIndex).Query
                outputQuery = txtDateSelOutputQuery.Text
                outputConnString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source =" & CalculationsDbPath
                outputConnection.ConnectionString = outputConnString
                outputConnection.Open()
                outputDa = New OleDb.OleDbDataAdapter(outputQuery, outputConnection)
                outputDa.MissingSchemaAction = MissingSchemaAction.AddWithKey
                Dim outputCmdBuilder As New OleDb.OleDbCommandBuilder(outputDa)
                outputDa.Fill(dsOutput, "myData")
                outputConnection.Close()
        End Select
    End Sub


#End Region 'Date Selections Sub Tab ----------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Daily Prices Sub Tab" '==============================================================================================================================================================

    Private Sub dgvDailyPricesInput_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles dgvDailyPricesInput.DataError
        'This code stops an unnecessary error dialog appearing.
        If e.Context = DataGridViewDataErrorContexts.Formatting Or e.Context = DataGridViewDataErrorContexts.PreferredSize Then
            e.ThrowException = False
        End If
    End Sub

    Private Sub dgvDailyPricesOutput_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles dgvDailyPricesOutput.DataError
        'This code stops an unnecessary error dialog appearing.
        If e.Context = DataGridViewDataErrorContexts.Formatting Or e.Context = DataGridViewDataErrorContexts.PreferredSize Then
            e.ThrowException = False
        End If
    End Sub

    Private Sub SetUpDailyPricesTab()
        'Set up the Daily Prices tab.

        Dim connectionString As String 'Declare a connection string for MS Access - defines the database or server to be used.
        Dim conn As System.Data.OleDb.OleDbConnection 'Declare a connection for MS Access - used by the Data Adapter to connect to and disconnect from the database.
        Dim commandString As String 'Declare a command string - contains the query to be passed to the database.
        Dim ds As DataSet 'Declate a Dataset.
        Dim dt As DataTable
        Dim dataAdapter As System.Data.OleDb.OleDbDataAdapter
        Dim NFields As Integer 'The number of fields in dt
        Dim I As Integer 'Loop index

        'Set up dgvDailyPricesInput ----------------------------------------------------
        dgvDailyPricesInput.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
        dgvDailyPricesInput.Columns.Clear()

        If cmbDailyPriceInputDb.SelectedIndex > -1 Then
            Select Case cmbDailyPriceInputDb.SelectedItem.ToString
                Case "Share Prices"
                    connectionString = "provider=Microsoft.ACE.OLEDB.12.0;data source = " + SharePriceDbPath
                Case "Financials"
                    connectionString = "provider=Microsoft.ACE.OLEDB.12.0;data source = " + FinancialsDbPath
                Case "Calculations"
                    connectionString = "provider=Microsoft.ACE.OLEDB.12.0;data source = " + CalculationsDbPath
            End Select

            conn = New System.Data.OleDb.OleDbConnection(connectionString) 'Connect to the Access database:
            conn.Open()

            If cmbDailyPriceInputTable.SelectedIndex > -1 Then
                If cmbDailyPriceInputTable.SelectedItem.ToString = "" Then
                    commandString = ""
                Else
                    commandString = "SELECT Top 500 * FROM " + cmbDailyPriceInputTable.SelectedItem.ToString
                    dataAdapter = New System.Data.OleDb.OleDbDataAdapter(commandString, conn)
                    ds = New DataSet
                    dataAdapter.Fill(ds, "SelTable") 'ds was defined earlier as a DataSet
                    dt = ds.Tables("SelTable")
                    NFields = dt.Columns.Count
                End If
            Else
                commandString = ""
            End If
        Else
            commandString = ""
        End If

        Dim TextBoxCol0 As New DataGridViewTextBoxColumn
        dgvDailyPricesInput.Columns.Add(TextBoxCol0)
        dgvDailyPricesInput.Columns(0).HeaderText = "Input Parameter"

        Dim ComboBoxCol1 As New DataGridViewComboBoxColumn
        dgvDailyPricesInput.Columns.Add(ComboBoxCol1)
        dgvDailyPricesInput.Columns(1).HeaderText = "Input Table Column"
        dgvDailyPricesInput.Columns(1).Width = 240
        If commandString = "" Then
            'No column names available.
        Else
            'Add the available column names to ComboBox Col1:
            For I = 0 To NFields - 1
                ComboBoxCol1.Items.Add(dt.Columns(I).ColumnName.ToString)
            Next
        End If


        'Set up dgvDailyPricesOutput ----------------------------------------------------
        dgvDailyPricesOutput.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
        dgvDailyPricesOutput.Columns.Clear()


        If cmbDailyPriceOutputDb.SelectedIndex > -1 Then
            Select Case cmbDailyPriceOutputDb.SelectedItem.ToString
                Case "Share Prices"
                    connectionString = "provider=Microsoft.ACE.OLEDB.12.0;data source = " + SharePriceDbPath
                Case "Financials"
                    connectionString = "provider=Microsoft.ACE.OLEDB.12.0;data source = " + FinancialsDbPath
                Case "Calculations"
                    connectionString = "provider=Microsoft.ACE.OLEDB.12.0;data source = " + CalculationsDbPath
            End Select

            conn = New System.Data.OleDb.OleDbConnection(connectionString) 'Connect to the Access database:
            conn.Open()

            If cmbDailyPriceOutputTable.SelectedIndex > -1 Then
                If cmbDailyPriceOutputTable.SelectedItem.ToString = "" Then
                    commandString = ""
                Else
                    commandString = "SELECT Top 500 * FROM " + cmbDailyPriceOutputTable.SelectedItem.ToString
                    dataAdapter = New System.Data.OleDb.OleDbDataAdapter(commandString, conn)
                    ds = New DataSet
                    dataAdapter.Fill(ds, "SelTable") 'ds was defined earlier as a DataSet
                    dt = ds.Tables("SelTable")
                    NFields = dt.Columns.Count
                End If
            Else
                commandString = ""
            End If
        Else
            commandString = ""
        End If



        Dim TextBoxCol10 As New DataGridViewTextBoxColumn
        dgvDailyPricesOutput.Columns.Add(TextBoxCol10)
        dgvDailyPricesOutput.Columns(0).HeaderText = "Output Parameter"

        Dim ComboBoxCol11 As New DataGridViewComboBoxColumn
        dgvDailyPricesOutput.Columns.Add(ComboBoxCol11)
        dgvDailyPricesOutput.Columns(1).HeaderText = "Output Table Column"
        dgvDailyPricesOutput.Columns(1).Width = 240

        If commandString = "" Then
            'No column names available.
        Else
            'Add the available column names to ComboBox Col1:
            For I = 0 To NFields - 1
                ComboBoxCol11.Items.Add(dt.Columns(I).ColumnName.ToString)
            Next
        End If




        If cmbDailyPriceCalcType.SelectedIndex > -1 Then
            dgvDailyPricesInput.AllowUserToAddRows = False
            dgvDailyPricesInput.Rows.Clear()

            dgvDailyPricesOutput.AllowUserToAddRows = False
            dgvDailyPricesOutput.Rows.Clear()

            Select Case cmbDailyPriceCalcType.SelectedItem.ToString
                Case "Count the daily number of companies and total value traded"
                    dgvDailyPricesInput.Rows.Add()
                    dgvDailyPricesInput.Rows(dgvDailyPricesInput.RowCount - 1).Cells(0).Value = "Trade_Date"
                    dgvDailyPricesInput.Rows.Add()
                    dgvDailyPricesInput.Rows(dgvDailyPricesInput.RowCount - 1).Cells(0).Value = "ASX_Code"
                    dgvDailyPricesInput.Rows.Add()
                    dgvDailyPricesInput.Rows(dgvDailyPricesInput.RowCount - 1).Cells(0).Value = "Close_Price"
                    dgvDailyPricesInput.Rows.Add()
                    dgvDailyPricesInput.Rows(dgvDailyPricesInput.RowCount - 1).Cells(0).Value = "Volume"

                    dgvDailyPricesOutput.Rows.Add()
                    dgvDailyPricesOutput.Rows(dgvDailyPricesOutput.RowCount - 1).Cells(0).Value = "Trade_Date"
                    dgvDailyPricesOutput.Rows.Add()
                    dgvDailyPricesOutput.Rows(dgvDailyPricesOutput.RowCount - 1).Cells(0).Value = "NCompanies_Traded"
                    dgvDailyPricesOutput.Rows.Add()
                    dgvDailyPricesOutput.Rows(dgvDailyPricesOutput.RowCount - 1).Cells(0).Value = "Value_Traded"


                Case "Find trading gaps"

                Case "Find first trade and last trade dates"

                Case "Find price level changes"

                Case "Find price spikes"

            End Select

            dgvDailyPricesInput.AutoResizeColumn(0)
            dgvDailyPricesOutput.AutoResizeColumn(0)

            dgvDailyPricesInput.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            dgvDailyPricesOutput.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

        End If

    End Sub

    Private Sub cmbDailyPriceInputDb_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbDailyPriceInputDb.SelectedIndexChanged
        'The Input database has been selected.
        'Fill the list of input tables.

        Dim connectionString As String 'Declare a connection string for MS Access - defines the database or server to be used.

        Select Case cmbDailyPriceInputDb.SelectedItem.ToString
            Case "Share Prices"
                If SharePriceDbPath = "" Then
                    Message.AddWarning("No database selected!" & vbCrLf)
                    Exit Sub
                Else
                    connectionString = "provider=Microsoft.ACE.OLEDB.12.0;data source = " + SharePriceDbPath
                End If
            Case "Financials"
                If FinancialsDbPath = "" Then
                    Message.AddWarning("No database selected!" & vbCrLf)
                    Exit Sub
                Else
                    connectionString = "provider=Microsoft.ACE.OLEDB.12.0;data source = " + FinancialsDbPath
                End If
            Case "Calculations"
                If CalculationsDbPath = "" Then
                    Message.AddWarning("No database selected!" & vbCrLf)
                    Exit Sub
                Else
                    connectionString = "provider=Microsoft.ACE.OLEDB.12.0;data source = " + CalculationsDbPath
                End If
        End Select

        Dim conn As System.Data.OleDb.OleDbConnection 'Declare a connection for MS Access - used by the Data Adapter to connect to and disconnect from the database.
        Dim dt As DataTable

        cmbDailyPriceInputTable.Text = ""
        cmbDailyPriceInputTable.Items.Clear()

        'Dim ds As DataSet = New DataSet

        conn = New System.Data.OleDb.OleDbConnection(connectionString)
        conn.Open()

        Dim restrictions As String() = New String() {Nothing, Nothing, Nothing, "TABLE"} 'This restriction removes system tables
        dt = conn.GetSchema("Tables", restrictions)

        'Fill cmbDailyPriceInputTable:
        Dim dr As DataRow
        Dim I As Integer 'Loop index
        Dim MaxI As Integer = dt.Rows.Count

        For I = 0 To MaxI - 1
            dr = dt.Rows(0)
            cmbDailyPriceInputTable.Items.Add(dt.Rows(I).Item(2).ToString)
        Next

        conn.Close()

    End Sub

    Private Sub cmbDailyPriceOutputDb_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbDailyPriceOutputDb.SelectedIndexChanged
        'The Output database has been selected.
        'Fill the list of output tables.

        Dim connectionString As String 'Declare a connection string for MS Access - defines the database or server to be used.

        Select Case cmbDailyPriceOutputDb.SelectedItem.ToString
            Case "Share Prices"
                If SharePriceDbPath = "" Then
                    Message.AddWarning("No database selected!" & vbCrLf)
                    Exit Sub
                Else
                    connectionString = "provider=Microsoft.ACE.OLEDB.12.0;data source = " + SharePriceDbPath
                End If
            Case "Financials"
                If FinancialsDbPath = "" Then
                    Message.AddWarning("No database selected!" & vbCrLf)
                    Exit Sub
                Else
                    connectionString = "provider=Microsoft.ACE.OLEDB.12.0;data source = " + FinancialsDbPath
                End If
            Case "Calculations"
                If CalculationsDbPath = "" Then
                    Message.AddWarning("No database selected!" & vbCrLf)
                    Exit Sub
                Else
                    connectionString = "provider=Microsoft.ACE.OLEDB.12.0;data source = " + CalculationsDbPath
                End If
        End Select

        Dim conn As System.Data.OleDb.OleDbConnection 'Declare a connection for MS Access - used by the Data Adapter to connect to and disconnect from the database.
        Dim dt As DataTable

        cmbDailyPriceOutputTable.Text = ""
        cmbDailyPriceOutputTable.Items.Clear()

        'Dim ds As DataSet = New DataSet

        conn = New System.Data.OleDb.OleDbConnection(connectionString)
        conn.Open()

        Dim restrictions As String() = New String() {Nothing, Nothing, Nothing, "TABLE"} 'This restriction removes system tables
        dt = conn.GetSchema("Tables", restrictions)

        'Fill cmbDailyPriceInputTable:
        Dim dr As DataRow
        Dim I As Integer 'Loop index
        Dim MaxI As Integer = dt.Rows.Count

        For I = 0 To MaxI - 1
            dr = dt.Rows(0)
            cmbDailyPriceOutputTable.Items.Add(dt.Rows(I).Item(2).ToString)
        Next

        conn.Close()
    End Sub

    Private Sub cmbDailyPriceCalcType_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbDailyPriceCalcType.SelectedIndexChanged
        'Daily prices calculation type changed.
        SetUpDailyPricesTab()
    End Sub

    Private Sub cmbDailyPriceInputTable_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbDailyPriceInputTable.SelectedIndexChanged
        'New Input table selected.
        'Update the list of available columns in dgvDailyPricesInput.
        SetUpDailyPricesTab()
    End Sub

    Private Sub cmbDailyPriceOutputTable_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbDailyPriceOutputTable.SelectedIndexChanged
        'New Output table selected.
        'Update the list of available columns in dgvDailyPricesOutput.
        SetUpDailyPricesTab()
    End Sub


#End Region 'Daily Prices Sub Tab -------------------------------------------------------------------------------------------------------------------------------------------------------------

#Region " Curve Fitting Sub Tab" '=============================================================================================================================================================

#End Region 'Curve Fitting Sub Tab ------------------------------------------------------------------------------------------------------------------------------------------------------------

#Region " Filters Sub Tab" '===================================================================================================================================================================

#End Region 'Filters Sub Tab ------------------------------------------------------------------------------------------------------------------------------------------------------------------

#End Region 'Calculations Tab -----------------------------------------------------------------------------------------------------------------------------------------------------------------

#Region " Processing Sequence Tab" '===========================================================================================================================================================

    Private Sub btnOpenSequence_Click(sender As Object, e As EventArgs) Handles btnOpenSequence.Click
        'Open a processing sequence file.

        Dim SelectedFileName As String = ""

        SelectedFileName = Project.SelectDataFile("Sequence", "Sequence")
        Message.Add("Selected Processing Sequence: " & SelectedFileName & vbCrLf)

        If SelectedFileName = "" Then

        Else
            txtName.Text = SelectedFileName

            Dim xmlSeq As System.Xml.Linq.XDocument

            Project.ReadXmlData(SelectedFileName, xmlSeq)

            If xmlSeq Is Nothing Then
                Exit Sub
            End If

            txtDescription.Text = xmlSeq.<ProcessingSequence>.<Description>.Value

            XSeq.SequenceName = SelectedFileName
            XSeq.SequenceDescription = txtDescription.Text

            'rtbSequence.Text = xmlSeq.ToString
            'FormatXmlText()

            'Import.ImportSequenceName = SelectedFileName
            'Import.ImportSequenceDescription = xmlSeq.<ProcessingSequence>.<Description>.Value
            'txtDescription.Text = Import.ImportSequenceDescription

        End If
    End Sub



#End Region 'Processing Sequence Tab ----------------------------------------------------------------------------------------------------------------------------------------------------------

    Private Sub DeleteSettingsFileToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DeleteSettingsFileToolStripMenuItem.Click
        'Delete the Settings File corresponding to the text box clicked.
        Message.Add("ContextMenuStrip1.SourceControl.Name = " & ContextMenuStrip1.SourceControl.Name & vbCrLf)

        Select Case ContextMenuStrip1.SourceControl.Name
            Case "txtCopyDataSettings" 'Delete the selected Copy Data settings file.
                'Dim SettingsFile As String = txtCopyDataSettings.Text
                Project.DeleteData(txtCopyDataSettings.Text)
                CopyDataSettingsFile = ""
                txtCopyDataSettings.Text = ""
                cmbCopyDataInputData.SelectedIndex = 0
                cmbCopyDataOutputData.SelectedIndex = 0
                dgvCopyData.Rows.Clear()

            Case "txtSimpleCalcSettings" 'Delete the selected Simple Calculations settings file.
                Project.DeleteData(txtSimpleCalcSettings.Text)
                SimpleCalcsSettingsFile = ""
                txtSimpleCalcSettings.Text = ""
                cmbSimpleCalcData.SelectedIndex = 0
                dgvSimpleCalcsParameterList.Rows.Clear()
                dgvSimpleCalcsCalculations.Rows.Clear()
                dgvSimpleCalcsInputData.Rows.Clear()
                dgvSimpleCalcsOutputData.Rows.Clear()


        End Select

    End Sub

#Region " Utilities Tab" '=====================================================================================================================================================================

#Region "Date Calculations Sub Tab" '==========================================================================================================================================================

    Private Sub cmbYearMonthDateCalc_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbYearMonthDateCalc.SelectedIndexChanged

        UpdateYearMonthCalcDate()
    End Sub

    Private Sub UpdateYearMonthCalcDate()
        'Update the Year, Month date calculation.

        If cmbYearMonthDateCalc.SelectedIndex = -1 Then Exit Sub

        Select Case cmbYearMonthDateCalc.SelectedItem.ToString
            Case "Date of end of month"
                Dim Month As Integer = Val(txtMonth.Text)
                Dim Year As Integer = Val(txtYear.Text)
                Dim calcDate As DateTime = DateSerial(Year, Month + 1, 0)
                txtYearMonthCalcDate.Text = calcDate.ToLongDateString

            Case "Date of start of month"
                Dim Month As Integer = Val(txtMonth.Text)
                Dim Year As Integer = Val(txtYear.Text)
                Dim calcDate As DateTime = DateSerial(Year, Month, 1)
                txtYearMonthCalcDate.Text = calcDate.ToLongDateString
        End Select
    End Sub

    Private Sub cmbStartDateNDaysCalc_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbStartDateNDaysCalc.SelectedIndexChanged

        UpdateStartDateNDaysCalcDate()
    End Sub

    Private Sub UpdateStartDateNDaysCalcDate()
        'Update the StartDate, NDays date calculation.

        If cmbStartDateNDaysCalc.SelectedIndex = -1 Then Exit Sub

        Select Case cmbStartDateNDaysCalc.SelectedItem.ToString
            Case "Date of Start Date + N Days"
                Dim StartDate As DateTime = DateTime.ParseExact(txtStartDate.Text, "dd MMM yyyy", Nothing)
                'Dim StartDate As Date = Date.Parse(txtStartDate.Text)
                Dim NDays As Double = Val(txtNDays.Text)
                Dim calcDate As DateTime = StartDate.AddDays(NDays)
                'txtStartDateNDaysCalc.Text = Format(calcDate, "dd MMM yyyy")
                txtStartDateNDaysCalc.Text = calcDate.ToLongDateString

            Case "Date of Start Date - N Days"
                Dim StartDate As DateTime = DateTime.ParseExact(txtStartDate.Text, "dd MMM yyyy", Nothing)
                'Dim StartDate As Date = Date.Parse(txtStartDate.Text)
                Dim NDays As Double = Val(txtNDays.Text)
                Dim calcDate As DateTime = StartDate.AddDays(-NDays)
                'txtStartDateNDaysCalc.Text = Format(calcDate, "dd MMM yyyy")
                txtStartDateNDaysCalc.Text = calcDate.ToLongDateString

        End Select
    End Sub

    Private Sub txtYear_TextChanged(sender As Object, e As EventArgs) Handles txtYear.TextChanged
        UpdateYearMonthCalcDate()
    End Sub

    Private Sub txtMonth_TextChanged(sender As Object, e As EventArgs) Handles txtMonth.TextChanged
        UpdateYearMonthCalcDate()
    End Sub

    Private Sub txtStartDate_TextChanged(sender As Object, e As EventArgs) Handles txtStartDate.TextChanged
        UpdateStartDateNDaysCalcDate()
    End Sub

    Private Sub txtNDays_TextChanged(sender As Object, e As EventArgs) Handles txtNDays.TextChanged
        UpdateStartDateNDaysCalcDate()
    End Sub




#End Region 'Date Calculations Sub Tab --------------------------------------------------------------------------------------------------------------------------------------------------------


#End Region 'Utilities Tab --------------------------------------------------------------------------------------------------------------------------------------------------------------------

#Region " Run XSequence Code" '================================================================================================================================================================

    Private Sub XSeq_ErrorMsg(ErrMsg As String) Handles XSeq.ErrorMsg
        Message.Add(ErrMsg & vbCrLf)
    End Sub

    Private Sub XSeq_Instruction(Info As String, Locn As String) Handles XSeq.Instruction
        'Execute each instruction that is produced by running the XSequence file.

        Select Case Locn
            'Parameter Code: -----------------------------------------------------------------------------------------------------------------------------
            Case "Parameter:Name"
                XSeq.NewParameter.Name = Info
            Case "Parameter:Description"
                XSeq.NewParameter.Description = Info
            Case "Parameter:Value"
                XSeq.NewParameter.Value = Info
            Case "Parameter:Command"
                Select Case Info
                    Case "Add"
                        XSeq.AddParameter()
                    Case Else
                        Message.AddWarning("Unknown Parameter:Command Information Value: " & Info & vbCrLf)
                End Select
            'Copy Data Code: -----------------------------------------------------------------------------------------------------------------------------
            Case "CopyData:InputDatabase"
                cmbCopyDataInputDb.SelectedIndex = cmbCopyDataInputDb.FindStringExact(Info)
                'Set up dgvCopyData:
                dgvCopyData.AllowUserToAddRows = False 'This removes the last edit row that allows the user to add new rows.
                dgvCopyData.Rows.Clear()
            Case "CopyData:InputDatabasePath"
                'This information is not used.
                'Consider writing a warning message if the current database path is different from the one in this Processing Sequence.
            Case "CopyData:InputQuery"
                txtCopyDataInputQuery.Text = Info
            Case "CopyData:InputQuery:ReadParameter"
                If XSeq.Parameter.ContainsKey(Info) Then
                    txtCopyDataInputQuery.Text = XSeq.Parameter(Info).Value
                Else
                    Message.AddWarning("Instruction error - CopyData:InputQuery:ReadParameter - The following parameter was not found: " & Info & vbCrLf)
                End If
            Case "CopyData:OutputDatabase"
                cmbCopyDataOutputDb.SelectedIndex = cmbCopyDataOutputDb.FindStringExact(Info)
            Case "CopyData:OutputDatabasePath"
                'This information is not used.
                'Consider writing a warning message if the current database path is different from the one in this Processing Sequence.
            Case "CopyData:OutputQuery"
                txtCopyDataOutputQuery.Text = Info
            Case "CopyData:OutputQuery:ReadParameter"
                'If XSeq.Parameter(Info).
                If XSeq.Parameter.ContainsKey(Info) Then
                    txtCopyDataOutputQuery.Text = XSeq.Parameter(Info).Value
                Else
                    Message.AddWarning("Instruction error - CopyData:OutputQuery:ReadParameter - The following parameter was not found: " & Info & vbCrLf)
                End If
            Case "CopyData:CopyList:CopyColumn:From"
                dgvCopyData.Rows.Add() 'Add a new blank row.
                dgvCopyData.Rows(dgvCopyData.Rows.Count - 1).Cells(0).Value = Info 'Add the From column name to the last row.
            Case "CopyData:CopyList:CopyColumn:To"
                dgvCopyData.Rows(dgvCopyData.Rows.Count - 1).Cells(1).Value = Info 'Add the To column name to the last row.
            Case "CopyData:Command"
                Select Case Info
                    Case "Apply"
                        ApplyCopyData()
                    Case Else
                        Message.AddWarning("Unknown CopyData:Command Information Value: " & Info & vbCrLf)
                End Select


            'Select Data Code: ---------------------------------------------------------------------------------------------------------------------------
            Case "SelectData:InputDatabase"
                cmbSelectDataInputDb.SelectedIndex = cmbSelectDataInputDb.FindStringExact(Info)
                'Set up dgvCopyData:
                dgvSelectData.AllowUserToAddRows = False 'This removes the last edit row that allows the user to add new rows.
                dgvSelectData.Rows.Clear()
                dgvSelectConstraints.AllowUserToAddRows = False 'This removes the last edit row that allows the user to add new rows.
                dgvSelectConstraints.Rows.Clear()
            Case "SelectData:InputDatabasePath"
                'This information is not used.
                'Consider writing a warning message if the current database path is different from the one in this Processing Sequence.
            Case "SelectData:InputQuery"
                txtSelectDataInputQuery.Text = Info
            Case "SelectData:InputQuery:ReadParameter"
                If XSeq.Parameter.ContainsKey(Info) Then
                    txtSelectDataInputQuery.Text = XSeq.Parameter(Info).Value
                Else
                    Message.AddWarning("Instruction error - SelectData:InputQuery:ReadParameter - The following parameter was not found: " & Info & vbCrLf)
                End If
            Case "SelectData:OutputDatabase"
                cmbSelectDataOutputDb.SelectedIndex = cmbSelectDataOutputDb.FindStringExact(Info)
            Case "SelectData:OutputDatabasePath"
                'This information is not used.
                'Consider writing a warning message if the current database path is different from the one in this Processing Sequence.
            Case "SelectData:OutputQuery"
                txtSelectDataOutputQuery.Text = Info
            Case "SelectData:OutputQuery:ReadParameter"
                If XSeq.Parameter.ContainsKey(Info) Then
                    txtSelectDataOutputQuery.Text = XSeq.Parameter(Info).Value
                Else
                    Message.AddWarning("Instruction error - SelectData:OutputQuery:ReadParameter - The following parameter was not found: " & Info & vbCrLf)
                End If
            Case "SelectData:SelectConstraintList:Constraint:WhereInputColumn"
                dgvSelectConstraints.Rows.Add() 'Add a new blank row.
                dgvSelectConstraints.Rows(dgvSelectConstraints.Rows.Count - 1).Cells(0).Value = Info 'Add the WhereInputColumn name to the last row.
            Case "SelectData:SelectConstraintList:Constraint:EqualsOutputColumn"
                dgvSelectConstraints.Rows(dgvSelectConstraints.Rows.Count - 1).Cells(1).Value = Info 'Add the EqualsOutputColumn name to the last row.
            Case "SelectData:SelectDataList:CopyColumn:From"
                dgvSelectData.Rows.Add() 'Add a new blank row.
                dgvSelectData.Rows(dgvSelectData.Rows.Count - 1).Cells(0).Value = Info 'Add the From column name to the last row.
            Case "SelectData:SelectDataList:CopyColumn:To"
                dgvSelectData.Rows(dgvSelectData.Rows.Count - 1).Cells(1).Value = Info 'Add the To column name to the last row.
            Case "SelectData:Command"
                Select Case Info
                    Case "Apply"
                        'ApplyDateSelections()
                        ApplySelectData()
                    Case Else
                        Message.AddWarning("Unknown SelectData:Command Information Value: " & Info & vbCrLf)
                End Select


            'Simple Calculations Code: -------------------------------------------------------------------------------------------------------------------
            Case "SimpleCalculations:SelectedDatabase"
                cmbSimpleCalcDb.SelectedIndex = cmbSimpleCalcDb.FindStringExact(Info)
                'Set up DataGridViews:
                dgvSimpleCalcsParameterList.AllowUserToAddRows = False 'This removes the last edit row that allows the user to add new rows.
                dgvSimpleCalcsParameterList.Rows.Clear()
                dgvSimpleCalcsInputData.AllowUserToAddRows = False 'This removes the last edit row that allows the user to add new rows.
                dgvSimpleCalcsInputData.Rows.Clear()
                dgvSimpleCalcsCalculations.AllowUserToAddRows = False 'This removes the last edit row that allows the user to add new rows.
                dgvSimpleCalcsCalculations.Rows.Clear()
                dgvSimpleCalcsOutputData.AllowUserToAddRows = False 'This removes the last edit row that allows the user to add new rows.
                dgvSimpleCalcsOutputData.Rows.Clear()

            Case "SimpleCalculations:SelectedDatabasePath"
                  'This information is not used.
                'Consider writing a warning message if the current database path is different from the one in this Processing Sequence.
            Case "SimpleCalculations:DataQuery"
                txtSimpleCalcsQuery.Text = Info
            Case "SimpleCalculations:DataQuery:ReadParameter"
                If XSeq.Parameter.ContainsKey(Info) Then
                    txtSimpleCalcsQuery.Text = XSeq.Parameter(Info).Value
                Else
                    Message.AddWarning("Instruction error - SimpleCalculations:DataQuery:ReadParameter - The following parameter was not found: " & Info & vbCrLf)
                End If
            Case "SimpleCalculations:ParameterList:Parameter:Name"
                dgvSimpleCalcsParameterList.Rows.Add() 'Add a new blank row.
                dgvSimpleCalcsParameterList.Rows(dgvSimpleCalcsParameterList.Rows.Count - 1).Cells(0).Value = Info 'Add the Parameter Name to the last row.
            Case "SimpleCalculations:ParameterList:Parameter:Type"
                dgvSimpleCalcsParameterList.Rows(dgvSimpleCalcsParameterList.Rows.Count - 1).Cells(1).Value = Info 'Add the Parameter Type to the last row.
            Case "SimpleCalculations:ParameterList:Parameter:Value"
                dgvSimpleCalcsParameterList.Rows(dgvSimpleCalcsParameterList.Rows.Count - 1).Cells(2).Value = Info 'Add the Parameter Value to the last row.
            Case "SimpleCalculations:ParameterList:Parameter:Description"
                dgvSimpleCalcsParameterList.Rows(dgvSimpleCalcsParameterList.Rows.Count - 1).Cells(3).Value = Info 'Add the Parameter Description to the last row.
            Case "SimpleCalculations:InputDataList:InputData:Parameter"
                dgvSimpleCalcsInputData.Rows.Add() 'Add a new blank row.
                dgvSimpleCalcsInputData.Rows(dgvSimpleCalcsInputData.Rows.Count - 1).Cells(0).Value = Info 'Add the Input Parameter Name to the last row.
            Case "SimpleCalculations:InputDataList:InputData:Column"
                dgvSimpleCalcsInputData.Rows(dgvSimpleCalcsInputData.Rows.Count - 1).Cells(1).Value = Info 'Add the Input Column Name to the last row.
            Case "SimpleCalculations:CalculationList:Calculation:Input1"
                dgvSimpleCalcsCalculations.Rows.Add() 'Add a new blank row.
                dgvSimpleCalcsCalculations.Rows(dgvSimpleCalcsCalculations.Rows.Count - 1).Cells(0).Value = Info 'Add the Input1 Parameter Name to the last row.
            Case "SimpleCalculations:CalculationList:Calculation:Input2"
                dgvSimpleCalcsCalculations.Rows(dgvSimpleCalcsCalculations.Rows.Count - 1).Cells(1).Value = Info 'Add the Input2 Parameter Name to the last row.
            Case "SimpleCalculations:CalculationList:Calculation:Operation"
                dgvSimpleCalcsCalculations.Rows(dgvSimpleCalcsCalculations.Rows.Count - 1).Cells(2).Value = Info 'Add the Operation to the last row.
            Case "SimpleCalculations:CalculationList:Calculation:Output"
                dgvSimpleCalcsCalculations.Rows(dgvSimpleCalcsCalculations.Rows.Count - 1).Cells(3).Value = Info 'Add the Output Parameter Name to the last row.
            Case "SimpleCalculations:OutputDataList:OutputData:Parameter"
                dgvSimpleCalcsOutputData.Rows.Add() 'Add a new blank row.
                dgvSimpleCalcsOutputData.Rows(dgvSimpleCalcsOutputData.Rows.Count - 1).Cells(0).Value = Info 'Add the Output Parameter Name to the last row.
            Case "SimpleCalculations:OutputDataList:OutputData:Column"
                dgvSimpleCalcsOutputData.Rows(dgvSimpleCalcsOutputData.Rows.Count - 1).Cells(1).Value = Info 'Add the Output Column Name to the last row.
            Case "SimpleCalculations:Command"
                Select Case Info
                    Case "Apply"
                        ApplySimpleCalcs()
                    Case Else
                        Message.AddWarning("Unknown SimpleCalculations:Command Information Value: " & Info & vbCrLf)
                End Select





            'Date Calculations Code: ---------------------------------------------------------------------------------------------------------------------
            Case "DateCalculations:SelectedDatabase"
                cmbDateCalcDb.SelectedIndex = cmbDateCalcDb.FindStringExact(Info)
            Case "DateCalculations:SelectedDatabasePath"
                  'This information is not used.
                'Consider writing a warning message if the current database path is different from the one in this Processing Sequence.
            Case "DateCalculations:DataQuery"
                txtDateCalcsQuery.Text = Info
            Case "DateCalculations:DataQuery:ReadParameter"
                If XSeq.Parameter.ContainsKey(Info) Then
                    txtDateCalcsQuery.Text = XSeq.Parameter(Info).Value
                Else
                    Message.AddWarning("Instruction error - DataCalculations:DataQuery:ReadParameter - The following parameter was not found: " & Info & vbCrLf)
                End If
            Case "DateCalculations:CalculationType"
                cmbDateCalcType.SelectedIndex = cmbDateCalcType.FindStringExact(Info)
            Case "DateCalculations:FixedDate"
                txtFixedDate.Text = Info
            Case "DateCalculations:FixedDate:ReadParameter"
                If XSeq.Parameter.ContainsKey(Info) Then
                    txtFixedDate.Text = XSeq.Parameter(Info).Value
                Else
                    Message.AddWarning("Instruction error - DateCalculations:FixedDate:ReadParameter - The following parameter was not found: " & Info & vbCrLf)
                End If
            Case "DateCalculations:DateFormatString"
                txtDateFormatString.Text = Info
            Case "DateCalculations:MonthColumn"
                cmbDateCalcParam1.Text = Info
            Case "DateCalculations:YearColumn"
                cmbDateCalcParam2.Text = Info
            Case "DateCalculations:StartDateColumn"
                cmbDateCalcParam1.Text = Info
            Case "DateCalculations:NDaysColumn"
                cmbDateCalcParam2.Text = Info
            Case "DateCalculations:OutputDateColumn"
                cmbDateCalcOutput.Text = Info
            Case "DateCalculations:Command"
                Select Case Info
                    Case "Apply"
                        ApplyDateCalcs()
                    Case Else
                        Message.AddWarning("Unknown CopyData:Command Information Value: " & Info & vbCrLf)
                End Select


            'Date Select Code: ---------------------------------------------------------------------------------------------------------------------------
            Case "DateSelect:InputDatabase"
                cmbDateSelectInputDb.SelectedIndex = cmbDateSelectInputDb.FindStringExact(Info)
                'Set up DataGridViews:
                dgvDateSelectData.AllowUserToAddRows = False 'This removes the last edit row that allows the user to add new rows.
                dgvDateSelectData.Rows.Clear()
                dgvDateSelectConstraints.AllowUserToAddRows = False 'This removes the last edit row that allows the user to add new rows.
                dgvDateSelectConstraints.Rows.Clear()
            Case "DateSelect:InputDatabasePath"
                'This information is not used.
                'Consider writing a warning message if the current database path is different from the one in this Processing Sequence.
            Case "DateSelect:InputQuery"
                txtDateSelInputQuery.Text = Info
            Case "DateSelect:InputQuery:ReadParameter"
                If XSeq.Parameter.ContainsKey(Info) Then
                    txtDateSelInputQuery.Text = XSeq.Parameter(Info).Value
                Else
                    Message.AddWarning("Instruction error - DateSelect:InputQuery:ReadParameter - The following parameter was not found: " & Info & vbCrLf)
                End If
            Case "DateSelect:OutputDatabase"
                cmbDateSelectOutputDb.SelectedIndex = cmbDateSelectOutputDb.FindStringExact(Info)
            Case "DateSelect:OutputDatabasePath"
                'This information is not used.
                'Consider writing a warning message if the current database path is different from the one in this Processing Sequence.
            Case "DateSelect:OutputQuery"
                txtDateSelOutputQuery.Text = Info
            Case "DateSelect:OutputQuery:ReadParameter"
                If XSeq.Parameter.ContainsKey(Info) Then
                    txtDateSelOutputQuery.Text = XSeq.Parameter(Info).Value
                Else
                    Message.AddWarning("Instruction error - DateSelect:OutputQuery:ReadParameter - The following parameter was not found: " & Info & vbCrLf)
                End If
            Case "DateSelect:DateSelectionType"
                cmbDateSelectionType.Text = Info
            Case "DateSelect:InputDateColumn"
                cmbDateSelInputDateCol.Text = Info
            Case "DateSelect:OutputDateColumn"
                cmbDateSelOutputDateCol.Text = Info
            Case "DateSelect:SelectConstraintList:Constraint:WhereInputColumn"
                dgvDateSelectConstraints.Rows.Add() 'Add a new blank row.
                dgvDateSelectConstraints.Rows(dgvDateSelectConstraints.Rows.Count - 1).Cells(0).Value = Info 'Add the WhereInputColumn name to the last row.
            Case "DateSelect:SelectConstraintList:Constraint:EqualsOutputColumn"
                dgvDateSelectConstraints.Rows(dgvDateSelectConstraints.Rows.Count - 1).Cells(1).Value = Info 'Add the EqualsOutputColumn name to the last row.
            Case "DateSelect:SelectDataList:CopyColumn:From"
                dgvDateSelectData.Rows.Add() 'Add a new blank row.
                dgvDateSelectData.Rows(dgvDateSelectData.Rows.Count - 1).Cells(0).Value = Info 'Add the From column name to the last row.
            Case "DateSelect:SelectDataList:CopyColumn:To"
                dgvDateSelectData.Rows(dgvDateSelectData.Rows.Count - 1).Cells(1).Value = Info 'Add the To column name to the last row.
            Case "DateSelect:Command"
                Select Case Info
                    Case "Apply"
                        ApplyDateSelections()
                    Case Else
                        Message.AddWarning("Unknown DateSelect:Command Information Value: " & Info & vbCrLf)
                End Select



            'End of Sequence Code: -----------------------------------------------------------------------------------------------------------------------
            Case "EndOfSequence"
                'ApplyCopyData()
                XSeq.Parameter.Clear() 'Clear the Parameter dictionary.
                Message.Add("Processing sequence has completed." & vbCrLf)

            Case Else
                Message.AddWarning("Unknown Information Location: " & Locn & vbCrLf)
        End Select


    End Sub





#End Region 'Run XSequence Code ---------------------------------------------------------------------------------------------------------------------------------------------------------------


#End Region 'Form Methods ---------------------------------------------------------------------------------------------------------------------------------------------------------------------

#Region " Form Events - Events that can be triggered by this form." '==========================================================================================================================
#End Region 'Form Events ----------------------------------------------------------------------------------------------------------------------------------------------------------------------

#Region " Classes - Other classes used in this form." '========================================================================================================================================

    Private Class DbSingle
        'DbSingle is used to represent single precision values obtained brom a database. It includes a NullValue boolean property.

        Private _value As Single = 0 'Single precision value.
        Property Value As Single
            Get
                Return _value
            End Get
            Set(value As Single)
                _value = value
            End Set
        End Property

        Private _nullValue As Boolean = True 'If True, the Value is DbNull.
        Property NullValue As Boolean
            Get
                Return _nullValue
            End Get
            Set(value As Boolean)
                _nullValue = value
            End Set
        End Property

    End Class

    Private Class Calculation
        'Used to represent  a calculation operation

        Private _input1 As String = "" 'Stores the Input1 parameter name.
        Property Input1 As String
            Get
                Return _input1
            End Get
            Set(value As String)
                _input1 = value
            End Set
        End Property

        Private _input2 As String = "" 'Stores the Input2 parameter name.
        Property Input2 As String
            Get
                Return _input2
            End Get
            Set(value As String)
                _input2 = value
            End Set
        End Property

        Private _operation As String = "" 'Stores the nae of the calculation operation (Input 1 + Input 2, Input 1 - Input 2, Input 1 x Input 2, Input 1 / Input 2)
        Property Operation As String
            Get
                Return _operation
            End Get
            Set(value As String)
                _operation = value
            End Set
        End Property

        Private _output As String = "" 'Stores the Output parameter name.
        Property Output As String
            Get
                Return _output
            End Get
            Set(value As String)
                _output = value
            End Set
        End Property

    End Class

    Private Class ParameterLocation
        'Used to stores the Table ColumnName corresponding to a ParameterName.

        Private _paramName As String = "" 'The name of a prameter.
        Property ParamName As String
            Get
                Return _paramName
            End Get
            Set(value As String)
                _paramName = value
            End Set
        End Property

        Private _colName As String = "" 'The name of the corresponding DataTable Column.
        Property ColName As String
            Get
                Return _colName
            End Get
            Set(value As String)
                _colName = value
            End Set
        End Property

    End Class










#End Region 'Classes --------------------------------------------------------------------------------------------------------------------------------------------------------------------------



End Class


