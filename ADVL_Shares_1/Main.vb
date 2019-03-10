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

Imports System.Security.Permissions
<PermissionSet(SecurityAction.Demand, Name:="FullTrust")>
<System.Runtime.InteropServices.ComVisibleAttribute(True)>
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
    '
    'Calling JavaScript from VB.NET:
    'The following Imports statement and permissions are required for the Main form:
    'Imports System.Security.Permissions
    '<PermissionSet(SecurityAction.Demand, Name:="FullTrust")> _
    '<System.Runtime.InteropServices.ComVisibleAttribute(True)> _
    'NOTE: the line continuation characters (_) will disappear form the code view after they have been typed!
    '------------------------------------------------------------------------------------------------------------------------------
    'Calling VB.NET from JavaScript
    'Add the following line to the Main.Load method:
    '  Me.WebBrowser1.ObjectForScripting = Me
    '------------------------------------------------------------------------------------------------------------------------------


#End Region 'Coding Notes ---------------------------------------------------------------------------------------------------------------------------------------------------------------------

#Region " Variable Declarations - All the variables used in this form and this application." '-------------------------------------------------------------------------------------------------

    'Declare Utility objects used to store information on the application, project, usage and display messages.
    Public WithEvents ApplicationInfo As New ADVL_Utilities_Library_1.ApplicationInfo 'This object is used to store application information.
    Public WithEvents Project As New ADVL_Utilities_Library_1.Project 'This object is used to store Project information.
    Public WithEvents Message As New ADVL_Utilities_Library_1.Message 'This object is used to display messages in the Messages window.
    Public WithEvents ApplicationUsage As New ADVL_Utilities_Library_1.Usage 'This object stores application usage information.

    'Declare Forms used by the application:
    Public WithEvents WebPageList As frmWebPageList

    Public WithEvents NewHtmlDisplay As frmHtmlDisplay
    Public HtmlDisplayFormList As New ArrayList 'Used for displaying multiple HtmlDisplay forms.

    Public WithEvents NewWebPage As frmWebPage
    Public WebPageFormList As New ArrayList 'Used for displaying multiple WebView forms.

    'Declare objects used to connect to the Application Network:
    Public client As ServiceReference1.MsgServiceClient
    Public WithEvents XMsg As New ADVL_Utilities_Library_1.XMessage
    Dim XDoc As New System.Xml.XmlDocument
    Public Status As New System.Collections.Specialized.StringCollection
    Dim ClientAppName As String = "" 'The name of the client requesting service
    Dim ClientAppNetName As String = "" 'The name of thge client Application Network requesting service. 
    Dim ClientConnName As String = "" 'The name of the client connection requesting service
    Dim MessageXDoc As System.Xml.Linq.XDocument
    Dim xmessage As XElement 'This will contain the message. It will be added to MessageXDoc.
    Dim xlocns As New List(Of XElement) 'A list of locations. Each location forms part of the reply message. The information in the reply message will be sent to the specified location in the client application.
    Dim MessageText As String = "" 'The text of a message sent through the Application Network.

    Public ConnectionName As String = "" 'The name of the connection used to connect this application to the AppNet.
    Public AppNetName As String = ""

    Public MsgServiceAppPath As String = "" 'The application path of the Message Service application (ComNet). This is where the "Application.Lock" file will be while ComNet is running
    Public MsgServiceExePath As String = "" 'The executable path of the Message Service.
    '----------------------------------------------------------------------------------------------------------------------------------

    'Variable for local processing of an XMessage:
    Public WithEvents XMsgLocal As New ADVL_Utilities_Library_1.XMessage
    Dim XDocLocal As New System.Xml.XmlDocument
    Public StatusLocal As New System.Collections.Specialized.StringCollection

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

    Public WithEvents CompanyList As frmCompanyList 'Form used to select a list of companies and display the Share Price chart.

    Public WithEvents Sequence As frmSequence

    Public WithEvents DesignPointChartQuery As frmDesignQuery

    Dim WithEvents Zip As ADVL_Utilities_Library_1.ZipComp

    Dim dsInput As DataSet = New DataSet 'The input dataset for calculations.
    Dim dsOutput As DataSet = New DataSet 'The output dataset for calculations.
    Dim outputQuery As String
    Dim outputConnString As String
    Dim outputConnection As OleDb.OleDbConnection = New OleDb.OleDbConnection
    Dim outputDa As OleDb.OleDbDataAdapter

    Public WithEvents XSeq As New ADVL_Utilities_Library_1.XSequence 'This is used to run a set of XML Sequence statements. These are used for data processing.

    Dim cboFieldSelections As New DataGridViewComboBoxColumn 'Used for selecting Y Value fields in the Charts: Share Prices tab

    Dim StockChartDefaults As New XDocument 'Default stock chart settings.
    Dim PointChartDefaults As New XDocument 'Default point chart (Cross Plot) settings.

    'Main.Load variables:
    Dim ProjectSelected As Boolean = False 'If True, a project has been selected using Command Arguments. Used in Main.Load.
    Dim StartupConnectionName As String = "" 'If not "" the application will be connected to the AppNet using this connection name in  Main.Load.

    'The following variables are used to run JavaScript in Web Pages: -------------------
    Public WithEvents WebXSeq As New ADVL_Utilities_Library_1.XSequence 'This is used to run a set of XML Sequence statements - for restoring Web page settings.
    'To run a Web XSequence:
    '  WebXSeq.RunXSequence(xDoc, Status) 'ImportStatus in Import
    '    Handle events:
    '      WebXSeq.ErrorMsg
    '      WebXSeq.Instruction(Info, Locn)

    Private XStatus As New System.Collections.Specialized.StringCollection

    'Variables used to restore Item values on a web page.
    Private FormName As String
    Private ItemName As String
    Private SelectId As String

    'StartProject variables:
    Private StartProject_AppName As String  'The application name
    Private StartProject_ConnName As String 'The connection name
    Private StartProject_ProjID As String   'The project ID

    'Get AppList in Add Application form
    Dim NewAppName As String 'The new application name that is being added to the the AppList dictionary in the Add Application form.

    'ShowSharePriceTable variable:
    Dim SharePricesFormNo As Integer = -1 'The number of the form used to display the Share Price Data.

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

    Private _connectedToComNet As Boolean = False  'True if the application is connected to the Communication Network (Message Service).
    Property ConnectedToComNet As Boolean
        Get
            Return _connectedToComNet
        End Get
        Set(value As Boolean)
            _connectedToComNet = value
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
                ProcessInstructions(_instrReceived)
            End If
        End Set
    End Property

    Private Sub ProcessInstructions(ByVal Instructions As String)
        'Process the XMessage instructions.

        'Add the message header to the XMessages window:
        Message.XAddText("Message received: " & vbCrLf, "XmlReceivedNotice")

        If Instructions.StartsWith("<XMsg>") Then 'This is an XMessage set of instructions.
            Try
                'Inititalise the reply message:
                Dim Decl As New XDeclaration("1.0", "utf-8", "yes")
                MessageXDoc = New XDocument(Decl, Nothing) 'Reply message - this will be sent to the Client App.
                xmessage = New XElement("XMsg")
                xlocns.Add(New XElement("Main")) 'Initially set the location in the Client App to Main.

                'Run the received message:
                Dim XmlHeader As String = "<?xml version=""1.0"" encoding=""utf-8"" standalone=""yes""?>"

                XDoc.LoadXml(XmlHeader & vbCrLf & Instructions.Replace("&", "&amp;")) 'Replace "&" with "&amp:" before loading the XML text.
                Message.XAddXml(XDoc)   'Add the message to the XMessages window.
                Message.XAddText(vbCrLf, "Normal") 'Add extra line
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

            If ClientConnName = "" Then
                'No client to send a message to - process the message locally.
                Message.XAddText("Message processed locally:" & vbCrLf, "XmlSentNotice")
                Message.XAddXml(MessageText)
                Message.XAddText(vbCrLf, "Normal") 'Add extra line
                ProcessLocalInstructions(MessageText)
            Else
                Message.XAddText("Message sent to " & ClientConnName & ":" & vbCrLf, "XmlSentNotice")
                Message.XAddXml(MessageText)
                Message.XAddText(vbCrLf, "Normal") 'Add extra line
                SendMessage() 'This subroutine triggers the timer to send the message after a short delay.
            End If
        Else 'This is not an XMessage!
            Message.XAddText("The message is not an XMessage: " & Instructions & vbCrLf, "Normal")
        End If
    End Sub

    Private Sub ProcessLocalInstructions(ByVal Instructions As String)
        'Process the XMessage instructions locally.

        If Instructions.StartsWith("<XMsg>") Then 'This is an XMessage set of instructions.
            'Run the received message:
            Dim XmlHeader As String = "<?xml version=""1.0"" encoding=""utf-8"" standalone=""yes""?>"
            XDocLocal.LoadXml(XmlHeader & vbCrLf & Instructions)
            XMsgLocal.Run(XDocLocal, StatusLocal)
        Else 'This is not an XMessage!
            Message.XAddText("The message is not an XMessage: " & Instructions & vbCrLf, "Normal")
        End If
    End Sub

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

    Private _startPageFileName As String = "" 'The file name of the html document displayed in the Start Page tab.
    Public Property StartPageFileName As String
        Get
            Return _startPageFileName
        End Get
        Set(value As String)
            _startPageFileName = value
        End Set
    End Property

#End Region 'Properties -----------------------------------------------------------------------------------------------------------------------------------------------------------------------

#Region " Process XML Files - Read and write XML files." '-------------------------------------------------------------------------------------------------------------------------------------

    Private Sub SaveFormSettings()
        'Save the form settings in an XML document.

        Dim settingsData = <?xml version="1.0" encoding="utf-8"?>
                           <!---->
                           <!--Form settings for Main form.-->
                           <FormSettings>
                               <Left><%= Me.Left %></Left>
                               <Top><%= Me.Top %></Top>
                               <Width><%= Me.Width %></Width>
                               <Height><%= Me.Height %></Height>
                               <MsgServiceAppPath><%= MsgServiceAppPath %></MsgServiceAppPath>
                               <MsgServiceExePath><%= MsgServiceExePath %></MsgServiceExePath>
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
                               <%= If(cmbCopyDataInputDb.SelectedIndex = -1,
                                   <CopyDataInputDb></CopyDataInputDb>,
                                   <CopyDataInputDb><%= cmbCopyDataInputDb.SelectedItem.ToString %></CopyDataInputDb>) %>
                               <%= If(cmbCopyDataInputData.SelectedIndex = -1,
                                   <CopyDataInputData></CopyDataInputData>,
                                   <CopyDataInputData><%= cmbCopyDataInputData.SelectedItem.ToString %></CopyDataInputData>) %>
                               <%= If(cmbCopyDataOutputDb.SelectedIndex = -1,
                                      <CopyDataOutputDb></CopyDataOutputDb>,
                                      <CopyDataOutputDb><%= cmbCopyDataOutputDb.SelectedItem.ToString %></CopyDataOutputDb>) %>
                               <%= If(cmbCopyDataOutputData.SelectedIndex = -1,
                                   <CopyDataOutputTable></CopyDataOutputTable>,
                                   <CopyDataOutputTable><%= cmbCopyDataOutputData.SelectedItem.ToString %></CopyDataOutputTable>) %>
                               <CopyDataSettingsFile><%= CopyDataSettingsFile %></CopyDataSettingsFile>
                               <!--Select Data-->
                               <%= If(cmbSelectDataInputDb.SelectedIndex = -1,
                                   <SelectDataInputDb></SelectDataInputDb>,
                                   <SelectDataInputDb><%= cmbSelectDataInputDb.SelectedItem.ToString %></SelectDataInputDb>) %>
                               <%= If(cmbSelectDataInputData.SelectedIndex = -1,
                                   <SelectDataInputData></SelectDataInputData>,
                                   <SelectDataInputData><%= cmbSelectDataInputData.SelectedItem.ToString %></SelectDataInputData>) %>
                               <%= If(cmbSelectDataOutputDb.SelectedIndex = -1,
                                   <SelectDataOutputDb></SelectDataOutputDb>,
                                   <SelectDataOutputDb><%= cmbSelectDataOutputDb.SelectedItem.ToString %></SelectDataOutputDb>) %>
                               <%= If(cmbSelectDataOutputData.SelectedIndex = -1,
                                   <SelectDataOutputTable></SelectDataOutputTable>,
                                   <SelectDataOutputTable><%= cmbSelectDataOutputData.SelectedItem.ToString %></SelectDataOutputTable>) %>
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
                               <%= If(cmbDailyPriceInputDb.SelectedIndex = -1,
                                   <DailyPricesInputDb></DailyPricesInputDb>,
                                   <DailyPricesInputDb><%= cmbDailyPriceInputDb.SelectedItem.ToString %></DailyPricesInputDb>) %>
                               <%= If(cmbDailyPriceInputTable.SelectedIndex = -1,
                                   <DailyPricesInputTable></DailyPricesInputTable>,
                                   <DailyPricesInputTable><%= cmbDailyPriceInputTable.SelectedItem.ToString %></DailyPricesInputTable>) %>
                               <%= If(cmbDailyPriceOutputDb.SelectedIndex = -1,
                                   <DailyPricesOutoutDb></DailyPricesOutoutDb>,
                                   <DailyPricesOutoutDb><%= cmbDailyPriceOutputDb.SelectedItem.ToString %></DailyPricesOutoutDb>) %>
                               <%= If(cmbDailyPriceOutputTable.SelectedIndex = -1,
                                   <DailyPricesOutputTable></DailyPricesOutputTable>,
                                   <DailyPricesOutputTable><%= cmbDailyPriceOutputTable.SelectedItem.ToString %></DailyPricesOutputTable>) %>
                               <%= If(cmbDailyPriceCalcType.SelectedIndex = -1,
                                   <DailyPricesCalculationType></DailyPricesCalculationType>,
                                   <DailyPricesCalculationType><%= cmbDailyPriceCalcType.SelectedItem.ToString %></DailyPricesCalculationType>) %>
                               <!--Utilities-->
                               <%= If(cmbUtilTablesDatabase.SelectedIndex = -1,
                                   <UtilitiesTablesDatabase></UtilitiesTablesDatabase>,
                                   <UtilitiesTablesDatabase><%= cmbUtilTablesDatabase.SelectedItem.ToString %></UtilitiesTablesDatabase>) %>
                               <!--Share Price Charts-->
                               <%= If(cmbSPChartDb.SelectedIndex = -1,
                                   <SPChartDatabase></SPChartDatabase>,
                                   <SPChartDatabase><%= cmbSPChartDb.SelectedItem.ToString %></SPChartDatabase>) %>
                               <%= If(cmbChartDataTable.SelectedIndex = -1,
                                   <SPChartDataTable></SPChartDataTable>,
                                   <SPChartDataTable><%= cmbChartDataTable.SelectedItem.ToString %></SPChartDataTable>) %>
                               <%= If(cmbCompanyCodeCol.SelectedIndex = -1,
                                   <SPChartCompanyCodeColumn></SPChartCompanyCodeColumn>,
                                   <SPChartCompanyCodeColumn><%= cmbCompanyCodeCol.SelectedItem.ToString %></SPChartCompanyCodeColumn>) %>
                               <SPChartCompanyCode><%= txtSPChartCompanyCode.Text %></SPChartCompanyCode>
                               <SPChartSeriesName><%= txtSeriesName.Text %></SPChartSeriesName>
                               <%= If(cmbXValues.SelectedIndex = -1,
                                   <SPChartXValues></SPChartXValues>,
                                   <SPChartXValues><%= cmbXValues.SelectedItem.ToString %></SPChartXValues>) %>
                               <%= If(DataGridView1.RowCount = 4,
                                   <SPChartHighPrice><%= DataGridView1.Rows(0).Cells(1).Value %></SPChartHighPrice>,
                                   <SPChartHighPrice></SPChartHighPrice>) %>
                               <%= If(DataGridView1.RowCount = 4,
                                   <SPChartLowPrice><%= DataGridView1.Rows(1).Cells(1).Value %></SPChartLowPrice>,
                                   <SPChartLowPrice></SPChartLowPrice>) %>
                               <%= If(DataGridView1.RowCount = 4,
                                   <SPChartOpenPrice><%= DataGridView1.Rows(2).Cells(1).Value %></SPChartOpenPrice>,
                                  <SPChartOpenPrice></SPChartOpenPrice>) %>
                               <%= If(DataGridView1.RowCount = 4,
                                  <SPChartClosePrice><%= DataGridView1.Rows(3).Cells(1).Value %></SPChartClosePrice>,
                                  <SPChartClosePrice></SPChartClosePrice>) %>
                               <SPChartTitleText><%= txtChartTitle.Text %></SPChartTitleText>
                               <SPChartTitleFontName><%= txtChartTitle.Font.Name %></SPChartTitleFontName>
                               <SPChartTitleColor><%= txtChartTitle.ForeColor %></SPChartTitleColor>
                               <SPChartTitleSize><%= txtChartTitle.Font.Size %></SPChartTitleSize>
                               <SPChartTitleBold><%= txtChartTitle.Font.Bold %></SPChartTitleBold>
                               <SPChartTitleItalic><%= txtChartTitle.Font.Italic %></SPChartTitleItalic>
                               <SPChartTitleUnderline><%= txtChartTitle.Font.Underline %></SPChartTitleUnderline>
                               <SPChartTitleStrikeout><%= txtChartTitle.Font.Strikeout %></SPChartTitleStrikeout>
                               <%= If(cmbAlignment.SelectedIndex = -1,
                                   <SPChartTitleAlignment></SPChartTitleAlignment>,
                                   <SPChartTitleAlignment><%= cmbAlignment.SelectedItem.ToString %></SPChartTitleAlignment>) %>
                               <SPChartSettingsFile><%= txtStockChartSettings.Text %></SPChartSettingsFile>
                               <SPChartUseDefaults><%= chkUseStockChartDefaults.Checked %></SPChartUseDefaults>
                               <SPChartUseDateRange><%= chkSPChartUseDateRange.Checked %></SPChartUseDateRange>
                               <SPChartFromDate><%= dtpSPChartFromDate.Value %></SPChartFromDate>
                               <SPChartToDate><%= dtpSPChartToDate.Value %></SPChartToDate>
                               <!--Cross Plot Charts-->
                               <%= If(cmbPointChartDb.SelectedIndex = -1,
                                   <CrossPlotDatabase></CrossPlotDatabase>,
                                   <CrossPlotDatabase><%= cmbPointChartDb.SelectedItem.ToString %></CrossPlotDatabase>) %>
                               <CrossPlotQuery><%= txtPointChartQuery.Text %></CrossPlotQuery>
                               <CrossPlotSeriesName><%= txtPointSeriesName.Text %></CrossPlotSeriesName>
                               <%= If(cmbPointXValues.SelectedIndex = -1,
                                   <CrossPlotXValues></CrossPlotXValues>,
                                   <CrossPlotXValues><%= cmbPointXValues.SelectedItem.ToString %></CrossPlotXValues>) %>
                               <%= If(cmbPointYValues.SelectedIndex = -1,
                                   <CrossPlotYValues></CrossPlotYValues>,
                                   <CrossPlotYValues><%= cmbPointYValues.SelectedItem.ToString %></CrossPlotYValues>) %>
                               <CrossPlotTitleText><%= txtPointChartTitle.Text %></CrossPlotTitleText>
                               <CrossPlotTitleFontName><%= txtPointChartTitle.Font.Name %></CrossPlotTitleFontName>
                               <CrossPlotTitleColor><%= txtPointChartTitle.ForeColor %></CrossPlotTitleColor>
                               <CrossPlotTitleSize><%= txtPointChartTitle.Font.Size %></CrossPlotTitleSize>
                               <CrossPlotTitleBold><%= txtPointChartTitle.Font.Bold %></CrossPlotTitleBold>
                               <CrossPlotTitleItalic><%= txtPointChartTitle.Font.Italic %></CrossPlotTitleItalic>
                               <CrossPlotTitleUnderline><%= txtPointChartTitle.Font.Underline %></CrossPlotTitleUnderline>
                               <CrossPlotTitleStrikeout><%= txtPointChartTitle.Font.Strikeout %></CrossPlotTitleStrikeout>
                               <%= If(cmbPointChartAlignment.SelectedIndex = -1,
                                   <CrossPlotTitleAlignment></CrossPlotTitleAlignment>,
                                   <CrossPlotTitleAlignment><%= cmbPointChartAlignment.SelectedItem.ToString %></CrossPlotTitleAlignment>) %>
                               <CrossPlotSettingsFile><%= txtPointChartSettings.Text %></CrossPlotSettingsFile>
                               <CrossPlotUseDefaults><%= chkUsePointChartDefaults.Checked %></CrossPlotUseDefaults>
                               <CrossPlotAutoXRange><%= chkAutoXRange.Checked %></CrossPlotAutoXRange>
                               <CrossPlotXMin><%= txtPointXMin.Text %></CrossPlotXMin>
                               <CrossPlotXMax><%= txtPointXMax.Text %></CrossPlotXMax>
                               <CrossPlotAutoYRange><%= chkAutoYRange.Checked %></CrossPlotAutoYRange>
                               <CrossPlotYMin><%= txtPointYMin.Text %></CrossPlotYMin>
                               <CrossPlotYMax><%= txtPointYMax.Text %></CrossPlotYMax>
                           </FormSettings>

        Dim SettingsFileName As String = "FormSettings_" & ApplicationInfo.Name & "_" & Me.Text & ".xml"
        Debug.Print("Writing settings file: " & SettingsFileName)
        Project.SaveXmlSettings(SettingsFileName, settingsData)

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

            If Settings.<FormSettings>.<MsgServiceAppPath>.Value <> Nothing Then MsgServiceAppPath = Settings.<FormSettings>.<MsgServiceAppPath>.Value
            If Settings.<FormSettings>.<MsgServiceExePath>.Value <> Nothing Then MsgServiceExePath = Settings.<FormSettings>.<MsgServiceExePath>.Value

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
            Else
                SharePriceDataViewList = "Default.SPDataList"
                SharePricesSettings.ListFileName = SharePriceDataViewList 'Set the file name in SharePricesSettings List.
                txtSharePricesDataList.Text = SharePriceDataViewList
            End If

            'Restore View Data - Financials settings
            If Settings.<FormSettings>.<FinancialsDbPath>.Value <> Nothing Then FinancialsDbPath = Settings.<FormSettings>.<FinancialsDbPath>.Value
            If Settings.<FormSettings>.<FinancialsDataViewList>.Value <> Nothing Then
                FinancialsDataViewList = Settings.<FormSettings>.<FinancialsDataViewList>.Value
                FinancialsSettings.ListFileName = FinancialsDataViewList 'Set the file name in FinancialsSettingsList.
                txtFinancialsDataList.Text = FinancialsDataViewList
                FinancialsSettings.LoadFile() 'Load the settings list in FinancialsSettingsList.
                DisplayFinancialsList()
            Else
                FinancialsDataViewList = "Default.FinDataList"
                FinancialsSettings.ListFileName = FinancialsDataViewList 'Set the file name in FinancialsSettingsList.
                txtFinancialsDataList.Text = FinancialsDataViewList
            End If

            'Restore View Data - Calculations settings
            If Settings.<FormSettings>.<CalculationsDbPath>.Value <> Nothing Then CalculationsDbPath = Settings.<FormSettings>.<CalculationsDbPath>.Value
            If Settings.<FormSettings>.<CalculationsDataViewList>.Value <> Nothing Then
                CalculationsDataViewList = Settings.<FormSettings>.<CalculationsDataViewList>.Value
                CalculationsSettings.ListFileName = CalculationsDataViewList 'Set the file name in CalculationsSettingsList.
                txtCalcsDataList.Text = CalculationsDataViewList
                CalculationsSettings.LoadFile() 'Load the settings list in CalculationsSettingsList.
                DisplayCalculationsList()
            Else
                CalculationsDataViewList = "Default.CalcDataList"
                CalculationsSettings.ListFileName = CalculationsDataViewList 'Set the file name in CalculationsSettingsList.
                txtCalcsDataList.Text = CalculationsDataViewList
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

            'Utilities
            If Settings.<FormSettings>.<UtilitiesTablesDatabase>.Value <> Nothing Then
                cmbUtilTablesDatabase.SelectedIndex = cmbUtilTablesDatabase.FindStringExact(Settings.<FormSettings>.<UtilitiesTablesDatabase>.Value)
                UtilTablesDatabaseChanged()
            End If

            'Share Price Charts
            If Settings.<FormSettings>.<SPChartDatabase>.Value <> Nothing Then cmbSPChartDb.SelectedIndex = cmbSPChartDb.FindStringExact(Settings.<FormSettings>.<SPChartDatabase>.Value)

            If Settings.<FormSettings>.<SPChartDataTable>.Value <> Nothing Then cmbChartDataTable.SelectedIndex = cmbChartDataTable.FindStringExact(Settings.<FormSettings>.<SPChartDataTable>.Value)
            If Settings.<FormSettings>.<SPChartCompanyCodeColumn>.Value <> Nothing Then cmbCompanyCodeCol.SelectedIndex = cmbCompanyCodeCol.FindStringExact(Settings.<FormSettings>.<SPChartCompanyCodeColumn>.Value)
            If Settings.<FormSettings>.<SPChartCompanyCode>.Value <> Nothing Then txtSPChartCompanyCode.Text = Settings.<FormSettings>.<SPChartCompanyCode>.Value
            UpdateSPChartQuery()
            UpdateChartSharePricesTab()

            If Settings.<FormSettings>.<SPChartSeriesName>.Value <> Nothing Then txtSeriesName.Text = Settings.<FormSettings>.<SPChartSeriesName>.Value
            If Settings.<FormSettings>.<SPChartXValues>.Value <> Nothing Then cmbXValues.SelectedIndex = cmbXValues.FindStringExact(Settings.<FormSettings>.<SPChartXValues>.Value)
            If Settings.<FormSettings>.<SPChartHighPrice>.Value <> Nothing Then DataGridView1.Rows(0).Cells(1).Value = Settings.<FormSettings>.<SPChartHighPrice>.Value
            If Settings.<FormSettings>.<SPChartLowPrice>.Value <> Nothing Then DataGridView1.Rows(1).Cells(1).Value = Settings.<FormSettings>.<SPChartLowPrice>.Value
            If Settings.<FormSettings>.<SPChartOpenPrice>.Value <> Nothing Then DataGridView1.Rows(2).Cells(1).Value = Settings.<FormSettings>.<SPChartOpenPrice>.Value
            If Settings.<FormSettings>.<SPChartClosePrice>.Value <> Nothing Then DataGridView1.Rows(3).Cells(1).Value = Settings.<FormSettings>.<SPChartClosePrice>.Value

            If Settings.<FormSettings>.<SPChartTitleText>.Value <> Nothing Then txtChartTitle.Text = Settings.<FormSettings>.<SPChartTitleText>.Value
            Dim myFontStyle As FontStyle = FontStyle.Regular
            Dim myFontSize As Single = 10
            Dim myFontName As String = "Arial"

            If Settings.<FormSettings>.<SPChartTitleFontName>.Value <> Nothing Then
                myFontName = Settings.<FormSettings>.<SPChartTitleFontName>.Value
            End If

            If Settings.<FormSettings>.<SPChartTitleColor>.Value <> Nothing Then txtChartTitle.ForeColor = Color.FromName(Settings.<FormSettings>.<SPChartTitleColor>.Value)

            If Settings.<FormSettings>.<SPChartTitleSize>.Value <> Nothing Then
                myFontSize = Settings.<FormSettings>.<SPChartTitleSize>.Value
            End If

            If Settings.<FormSettings>.<SPChartTitleBold>.Value <> Nothing Then
                If Settings.<FormSettings>.<SPChartTitleBold>.Value = "true" Then
                    myFontStyle = myFontStyle Or FontStyle.Bold
                End If
            End If

            If Settings.<FormSettings>.<SPChartTitleItalic>.Value <> Nothing Then
                If Settings.<FormSettings>.<SPChartTitleItalic>.Value = "true" Then
                    myFontStyle = myFontStyle Or FontStyle.Italic
                End If
            End If

            If Settings.<FormSettings>.<SPChartTitleUnderline>.Value <> Nothing Then
                If Settings.<FormSettings>.<SPChartTitleUnderline>.Value = "true" Then
                    myFontStyle = myFontStyle Or FontStyle.Underline
                End If
            End If

            If Settings.<FormSettings>.<SPChartTitleStrikeout>.Value <> Nothing Then
                If Settings.<FormSettings>.<SPChartTitleStrikeout>.Value = "true" Then
                    myFontStyle = myFontStyle Or FontStyle.Strikeout
                End If
            End If

            txtChartTitle.Font = New Font(myFontName, myFontSize, myFontStyle)
            If Settings.<FormSettings>.<SPChartTitleAlignment>.Value <> Nothing Then cmbAlignment.SelectedIndex = cmbAlignment.FindStringExact(Settings.<FormSettings>.<SPChartTitleAlignment>.Value)

            If Settings.<FormSettings>.<SPChartSettingsFile>.Value <> Nothing Then
                txtStockChartSettings.Text = Settings.<FormSettings>.<SPChartSettingsFile>.Value
                Project.ReadXmlData(txtStockChartSettings.Text, StockChartDefaults)

                If StockChartDefaults Is Nothing Then

                Else
                    rtbStockChartDefaults.Text = StockChartDefaults.ToString
                    FormatXmlText(rtbStockChartDefaults)
                End If
            End If

            If Settings.<FormSettings>.<SPChartUseDefaults>.Value <> Nothing Then chkUseStockChartDefaults.Checked = Settings.<FormSettings>.<SPChartUseDefaults>.Value
            If Settings.<FormSettings>.<SPChartUseDateRange>.Value <> Nothing Then chkSPChartUseDateRange.Checked = Settings.<FormSettings>.<SPChartUseDateRange>.Value
            If Settings.<FormSettings>.<SPChartFromDate>.Value <> Nothing Then dtpSPChartFromDate.Value = Settings.<FormSettings>.<SPChartFromDate>.Value
            If Settings.<FormSettings>.<SPChartToDate>.Value <> Nothing Then dtpSPChartToDate.Value = Settings.<FormSettings>.<SPChartToDate>.Value
            UpdateSPChartQuery()

            'Cross Plot Charts
            If Settings.<FormSettings>.<CrossPlotDatabase>.Value <> Nothing Then cmbPointChartDb.SelectedIndex = cmbPointChartDb.FindStringExact(Settings.<FormSettings>.<CrossPlotDatabase>.Value)
            If Settings.<FormSettings>.<CrossPlotQuery>.Value <> Nothing Then txtPointChartQuery.Text = Settings.<FormSettings>.<CrossPlotQuery>.Value
            UpdateChartCrossPlotsTab()
            If Settings.<FormSettings>.<CrossPlotSeriesName>.Value <> Nothing Then txtPointSeriesName.Text = Settings.<FormSettings>.<CrossPlotSeriesName>.Value
            If Settings.<FormSettings>.<CrossPlotXValues>.Value <> Nothing Then cmbPointXValues.SelectedIndex = cmbPointXValues.FindStringExact(Settings.<FormSettings>.<CrossPlotXValues>.Value)
            If Settings.<FormSettings>.<CrossPlotYValues>.Value <> Nothing Then cmbPointYValues.SelectedIndex = cmbPointYValues.FindStringExact(Settings.<FormSettings>.<CrossPlotYValues>.Value)
            If Settings.<FormSettings>.<CrossPlotTitleText>.Value <> Nothing Then txtPointChartTitle.Text = Settings.<FormSettings>.<CrossPlotTitleText>.Value

            If Settings.<FormSettings>.<CrossPlotTitleFontName>.Value <> Nothing Then
                myFontName = Settings.<FormSettings>.<CrossPlotTitleFontName>.Value
            End If

            If Settings.<FormSettings>.<CrossPlotTitleColor>.Value <> Nothing Then txtPointChartTitle.ForeColor = Color.FromName(Settings.<FormSettings>.<CrossPlotTitleColor>.Value)

            If Settings.<FormSettings>.<CrossPlotTitleSize>.Value <> Nothing Then
                myFontSize = Settings.<FormSettings>.<CrossPlotTitleSize>.Value
            End If

            If Settings.<FormSettings>.<CrossPlotTitleBold>.Value <> Nothing Then
                If Settings.<FormSettings>.<CrossPlotTitleBold>.Value = "true" Then
                    myFontStyle = myFontStyle Or FontStyle.Bold
                End If
            End If

            If Settings.<FormSettings>.<CrossPlotTitleItalic>.Value <> Nothing Then
                If Settings.<FormSettings>.<CrossPlotTitleItalic>.Value = "true" Then
                    myFontStyle = myFontStyle Or FontStyle.Italic
                End If
            End If

            If Settings.<FormSettings>.<CrossPlotTitleUnderline>.Value <> Nothing Then
                If Settings.<FormSettings>.<CrossPlotTitleUnderline>.Value = "true" Then
                    myFontStyle = myFontStyle Or FontStyle.Underline
                End If
            End If

            If Settings.<FormSettings>.<CrossPlotTitleStrikeout>.Value <> Nothing Then
                If Settings.<FormSettings>.<CrossPlotTitleStrikeout>.Value = "true" Then
                    myFontStyle = myFontStyle Or FontStyle.Strikeout
                End If
            End If

            txtPointChartTitle.Font = New Font(myFontName, myFontSize, myFontStyle)

            If Settings.<FormSettings>.<CrossPlotTitleAlignment>.Value <> Nothing Then cmbPointChartAlignment.SelectedIndex = cmbPointChartAlignment.FindStringExact(Settings.<FormSettings>.<CrossPlotTitleAlignment>.Value)

            If Settings.<FormSettings>.<CrossPlotSettingsFile>.Value <> Nothing Then
                txtPointChartSettings.Text = Settings.<FormSettings>.<CrossPlotSettingsFile>.Value
                Project.ReadXmlData(txtPointChartSettings.Text, PointChartDefaults)
                If PointChartDefaults Is Nothing Then

                Else
                    rtbPointChartDefaults.Text = PointChartDefaults.ToString
                    FormatXmlText(rtbPointChartDefaults)
                End If
            End If

            If Settings.<FormSettings>.<CrossPlotUseDefaults>.Value <> Nothing Then chkUsePointChartDefaults.Checked = Settings.<FormSettings>.<CrossPlotUseDefaults>.Value

            If Settings.<FormSettings>.<CrossPlotAutoXRange>.Value <> Nothing Then chkAutoXRange.Checked = Settings.<FormSettings>.<CrossPlotAutoXRange>.Value

            If Settings.<FormSettings>.<CrossPlotXMin>.Value <> Nothing Then
                txtPointXMin.Text = Settings.<FormSettings>.<CrossPlotXMin>.Value
            Else
                txtPointXMin.Text = "-100"
            End If

            If Settings.<FormSettings>.<CrossPlotXMax>.Value <> Nothing Then
                txtPointXMax.Text = Settings.<FormSettings>.<CrossPlotXMax>.Value
            Else
                txtPointXMax.Text = "100"
            End If

            If Settings.<FormSettings>.<CrossPlotAutoYRange>.Value <> Nothing Then chkAutoYRange.Checked = Settings.<FormSettings>.<CrossPlotAutoYRange>.Value

            If Settings.<FormSettings>.<CrossPlotYMin>.Value <> Nothing Then
                txtPointYMin.Text = Settings.<FormSettings>.<CrossPlotYMin>.Value
            Else
                txtPointYMin.Text = "-100"
            End If

            If Settings.<FormSettings>.<CrossPlotYMax>.Value <> Nothing Then
                txtPointYMax.Text = Settings.<FormSettings>.<CrossPlotYMax>.Value
            Else
                txtPointYMax.Text = "100"
            End If

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
        'Loading the Main form.

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
                Exit Sub
            End If
        End If

        ReadApplicationInfo()

        'Read the Application Usage information: --------------------------------------------
        ApplicationUsage.StartTime = Now
        ApplicationUsage.SaveLocn.Type = ADVL_Utilities_Library_1.FileLocation.Types.Directory
        ApplicationUsage.SaveLocn.Path = Project.ApplicationDir
        ApplicationUsage.RestoreUsageInfo()

        'Restore Project information: -------------------------------------------------------
        Project.ApplicationName = ApplicationInfo.Name

        'Set up Message object:
        Message.ApplicationName = ApplicationInfo.Name

        'Set up a temporary initial settings location:
        Dim TempLocn As New ADVL_Utilities_Library_1.FileLocation
        TempLocn.Type = ADVL_Utilities_Library_1.FileLocation.Types.Directory
        TempLocn.Path = ApplicationInfo.ApplicationDir
        Message.SettingsLocn = TempLocn

        Me.Show() 'Show this form before showing the Message form

        'Start showing messages here - Message system is set up.
        Message.AddText("------------------- Starting Application: ADVL Shares ----------------------------------- " & vbCrLf, "Heading")
        Message.AddText("Application usage: Total duration = " & ApplicationUsage.TotalDuration.TotalHours & " hours" & vbCrLf, "Normal")

        'https://msdn.microsoft.com/en-us/library/z2d603cy(v=vs.80).aspx#Y550
        'Process any command line arguments:
        Try
            For Each s As String In My.Application.CommandLineArgs
                Message.Add("Command line argument: " & vbCrLf)
                Message.AddXml(s & vbCrLf & vbCrLf)
                InstrReceived = s
            Next
        Catch ex As Exception
            Message.AddWarning("Error processing command line arguments: " & ex.Message & vbCrLf)
        End Try

        If ProjectSelected = False Then
            'Read the Settings Location for the last project used:
            Project.ReadLastProjectInfo()
            'The Last_Project_Info.xml file contains:
            '  Project Name and Description. Settings Location Type and Settings Location Path.
            Message.Add("Last project info has been read." & vbCrLf)
            Message.Add("Project.Type.ToString  " & Project.Type.ToString & vbCrLf)
            Message.Add("Project.Path  " & Project.Path & vbCrLf)

            'At this point read the application start arguments, if any.
            'The selected project may be changed here.

            'Check if the project is locked:
            If Project.ProjectLocked Then
                Message.AddWarning("The project is locked: " & Project.Name & vbCrLf)
                Dim dr As System.Windows.Forms.DialogResult
                dr = MessageBox.Show("Press 'Yes' to unlock the project", "Notice", MessageBoxButtons.YesNo)
                If dr = System.Windows.Forms.DialogResult.Yes Then
                    Project.UnlockProject()
                    Message.AddWarning("The project has been unlocked: " & Project.Name & vbCrLf)
                    'Read the Project Information file: -------------------------------------------------
                    Message.Add("Reading project info." & vbCrLf)
                    Project.ReadProjectInfoFile()   'Read the file in the SettingsLocation: ADVL_Project_Info.xml

                    Project.ReadParameters()
                    Project.ReadParentParameters()
                    If Project.ParentParameterExists("AppNetName") Then
                        Project.AddParameter("AppNetName", Project.ParentParameter("AppNetName").Value, Project.ParentParameter("AppNetName").Description) 'AddParameter will update the parameter if it already exists.
                        AppNetName = Project.Parameter("AppNetName").Value
                    Else
                        AppNetName = Project.GetParameter("AppNetName")
                    End If

                    Project.LockProject() 'Lock the project while it is open in this application.
                    'Set the project start time. This is used to track project usage.
                    Project.Usage.StartTime = Now
                    ApplicationInfo.SettingsLocn = Project.SettingsLocn
                    'Set up the Message object:
                    Message.SettingsLocn = Project.SettingsLocn
                Else
                    'Continue without any project selected.
                    Project.Name = ""
                    Project.Type = ADVL_Utilities_Library_1.Project.Types.None
                    Project.Description = ""
                    Project.SettingsLocn.Path = ""
                    Project.DataLocn.Path = ""
                End If

            Else
                'Read the Project Information file: -------------------------------------------------
                Message.Add("Reading project info." & vbCrLf)
                Project.ReadProjectInfoFile()   'Read the file in the SettingsLocation: ADVL_Project_Info.xml

                Project.ReadParameters()
                Project.ReadParentParameters()
                If Project.ParentParameterExists("AppNetName") Then
                    Project.AddParameter("AppNetName", Project.ParentParameter("AppNetName").Value, Project.ParentParameter("AppNetName").Description) 'AddParameter will update the parameter if it already exists.
                    AppNetName = Project.Parameter("AppNetName").Value
                Else
                    AppNetName = Project.GetParameter("AppNetName")
                End If

                Project.LockProject() 'Lock the project while it is open in this application.
                'Set the project start time. This is used to track project usage.
                Project.Usage.StartTime = Now
                ApplicationInfo.SettingsLocn = Project.SettingsLocn
                'Set up the Message object:
                Message.SettingsLocn = Project.SettingsLocn
            End If
        Else 'Project has been opened using Command Line arguments.

            Project.ReadParameters()
            Project.ReadParentParameters()
            If Project.ParentParameterExists("AppNetName") Then
                Project.AddParameter("AppNetName", Project.ParentParameter("AppNetName").Value, Project.ParentParameter("AppNetName").Description) 'AddParameter will update the parameter if it already exists.
                AppNetName = Project.Parameter("AppNetName").Value
            Else
                AppNetName = Project.GetParameter("AppNetName")
            End If

            Project.LockProject() 'Lock the project while it is open in this application.
            ProjectSelected = False 'Reset the Project Selected flag.
        End If

        'START Initialise the form: ===============================================================

        'Set up Settings Lists:
        FinancialsSettings.FileLocation = Project.DataLocn
        SharePricesSettings.FileLocation = Project.DataLocn
        CalculationsSettings.FileLocation = Project.DataLocn

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

        cmbUtilTablesDatabase.Items.Add("Share Prices")
        cmbUtilTablesDatabase.Items.Add("Financials")
        cmbUtilTablesDatabase.Items.Add("Calculations")
        cmbUtilTablesDatabase.SelectedIndex = 0 'Select first item as default


        'Set up the Simple Calculations tab: -----------------------------------------------
        'dgvSimpleCalcsParameterList.
        'Use these default settings: (The last form setting are sometimes not restored correctly!)
        SplitContainer2.SplitterDistance = 640 'Vertical spliiter distance
        SplitContainer3.SplitterDistance = 320 'LHS horizontal splitter distance
        SplitContainer4.SplitterDistance = 320 'RHS horizontal splitter distance

        'Set up context menus:
        txtCopyDataSettings.ContextMenuStrip = ContextMenuStrip1
        txtSimpleCalcSettings.ContextMenuStrip = ContextMenuStrip1

        'Set up the Charts tab:
        SetUpChartSharePricesTab()
        SetUpChartCrossPlotsTab()

        Me.WebBrowser1.ObjectForScripting = Me

        InitialiseForm() 'Initialise the form for a new project.

        'END Initialise the form: ------------------------------------------------------------------------------------

        RestoreFormSettings() 'Restore the form settings
        RestoreProjectSettings() 'Restore the Project settings

        'Show the project information: ------------------------------------------------------
        ShowProjectInfo()

        Message.AddText("------------------- Started OK -------------------------------------------------------------------------- " & vbCrLf & vbCrLf, "Heading")

        'Me.Show() 'Show this form before showing the Message form

        If StartupConnectionName = "" Then

            If Project.ConnectOnOpen Then
                ConnectToComNet() 'The Project is set to connect when it is opened.
            ElseIf ApplicationInfo.ConnectOnStartup Then
                ConnectToComNet() 'The Application is set to connect when it is started.
            Else
                'Don't connect to ComNet.
            End If

        Else
            'Connect to ComNet using the connection name StartupConnectionName.
            ConnectToComNet(StartupConnectionName)
        End If

        'Start the timer to keep the connection awake:
        'Timer3.Interval = 10000 '10 seconds - for testing
        Timer3.Interval = TimeSpan.FromMinutes(55).TotalMilliseconds '55 minute interval
        Timer3.Enabled = True
        Timer3.Start()

    End Sub

    Private Sub InitialiseForm()
        'Initialise the form for a new project.
        OpenStartPage()
    End Sub

    Private Sub ShowProjectInfo()
        'Show the project information:

        txtParentProject.Text = Project.ParentProjectName
        txtAppNetName.Text = Project.GetParameter("AppNetName")

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
        txtProjectPath.Text = Project.Path

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

        Select Case Project.SystemLocn.Type
            Case ADVL_Utilities_Library_1.FileLocation.Types.Directory
                txtSystemLocationType.Text = "Directory"
            Case ADVL_Utilities_Library_1.FileLocation.Types.Archive
                txtSystemLocationType.Text = "Archive"
        End Select
        txtSystemLocationPath.Text = Project.SystemLocn.Path

        If Project.ConnectOnOpen Then
            chkConnect.Checked = True
        Else
            chkConnect.Checked = False
        End If

        txtTotalDuration.Text = Project.Usage.TotalDuration.Days.ToString.PadLeft(5, "0"c) & ":" &
                                Project.Usage.TotalDuration.Hours.ToString.PadLeft(2, "0"c) & ":" &
                                Project.Usage.TotalDuration.Minutes.ToString.PadLeft(2, "0"c) & ":" &
                                Project.Usage.TotalDuration.Seconds.ToString.PadLeft(2, "0"c)

        txtCurrentDuration.Text = Project.Usage.CurrentDuration.Days.ToString.PadLeft(5, "0"c) & ":" &
                                  Project.Usage.CurrentDuration.Hours.ToString.PadLeft(2, "0"c) & ":" &
                                  Project.Usage.CurrentDuration.Minutes.ToString.PadLeft(2, "0"c) & ":" &
                                  Project.Usage.CurrentDuration.Seconds.ToString.PadLeft(2, "0"c)

    End Sub

    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        'Exit the Application

        DisconnectFromComNet() 'Disconnect from the Communication Network.

        SaveProjectSettings() 'Save project settings.

        'Save the settings file used to for data views:
        SharePricesSettings.SaveFile()
        FinancialsSettings.SaveFile()
        CalculationsSettings.SaveFile()

        ApplicationInfo.WriteFile() 'Update the Application Information file.
        ApplicationInfo.UnlockApplication()

        Project.SaveLastProjectInfo() 'Save information about the last project used.
        Project.SaveParameters()

        'Project.SaveProjectInfoFile() 'Update the Project Information file. This is not required unless there is a change made to the project.

        Project.Usage.SaveUsageInfo() 'Save Project usage information.

        Project.UnlockProject() 'Unlock the project.

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

    'THIS CODE IS USED IF MULTIPLE SHARE PRICES FORMS ARE TO BE SHOWN:
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

        If SharePricesFormList.Count <Index + 1 Then
            'Insert null entries into SharePricesList then add a new form at the specified index position:
        Dim I As Integer
            For I = SharePricesFormList.Count To Index
                SharePricesFormList.Add(Nothing)
            Next
            SharePrices = New frmSharePrices
            SharePricesFormList(Index) = SharePrices
            SharePricesFormList(Index).FormNo = Index
            SharePricesFormList(Index).Show
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
            SharePricesFormList(0).DataSummary = "New Share Price Data View"
            SharePricesFormList(0).Version = "Version 1"
        Else
            Dim I As Integer
            Dim FormAdded As Boolean = False
            For I = 0 To SharePricesFormList.Count - 1 'Check if there are closed forms in SharePricesList. They can be re-used.
                If IsNothing(SharePricesFormList(I)) Then
                    SharePricesFormList(I) = SharePrices
                    SharePricesFormList(I).FormNo = I
                    SharePricesFormList(I).Show()
                    SharePricesFormList(I).DataSummary = "New Share Price Data View"
                    SharePricesFormList(I).Version = "Version 1"
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
                SharePricesFormList(FormNo).DataSummary = "New Share Price Data View"
                SharePricesFormList(FormNo).Version = "Version 1"
            End If
        End If
    End Sub

    'THIS CODE IS USED IF MULTIPLE SHARE PRICES FORMS ARE TO BE SHOWN:
    Public Sub SharePricesFormClosed()
        'This subroutine is called when the SharePrices form has been closed.
        'The subroutine is usually called from the FormClosed event of the SharePrices form.
        'The SharePrices form may have multiple instances.
        'The ClosedFormNumber property should contains the number of the instance of the SharePrices form.
        'This property should be updated by the SharePrices form when it is being closed.
        'The ClosedFormNumber property value is used to determine which element in SharePricesList should be set to Nothing.

        'ERROR: When application is closed with SharePricesList forms open: !
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
        Else  'Open a new Financials form:
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

            FinancialsFormList(Index) = New frmFinancials
            FinancialsFormList(Index).FormNo = Index
            FinancialsFormList(Index).Show

        ElseIf IsNothing(FinancialsFormList(Index)) Then
            'Add the new form at specified index position:
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
            FinancialsFormList(0).DataSummary = "New Financials Data View"
            FinancialsFormList(0).Version = "Version 1"
        Else
            Dim I As Integer
            Dim FormAdded As Boolean = False
            For I = 0 To FinancialsFormList.Count - 1 'Check if there are closed forms in FinancialsFormList. They can be re-used.
                If IsNothing(FinancialsFormList(I)) Then
                    FinancialsFormList(I) = Financials
                    FinancialsFormList(I).FormNo = I
                    FinancialsFormList(I).Show()
                    FinancialsFormList(I).DataSummary = "New Financials Data View"
                    FinancialsFormList(I).Version = "Version 1"
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
                FinancialsFormList(FormNo).DataSummary = "New Financials Data View"
                FinancialsFormList(FormNo).Version = "Version 1"
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
            CalculationsFormList(0).DataSummary = "New Calculations Data View"
            CalculationsFormList(0).Version = "Version 1"
        Else
            Dim I As Integer
            Dim FormAdded As Boolean = False
            For I = 0 To CalculationsFormList.Count - 1 'Check if there are closed forms in CalculationsFormList. They can be re-used.
                If IsNothing(CalculationsFormList(I)) Then
                    CalculationsFormList(I) = Calculations
                    CalculationsFormList(I).FormNo = I
                    CalculationsFormList(I).Show()
                    CalculationsFormList(I).DataSummary = "New Calculations Data View"
                    CalculationsFormList(I).Version = "Version 1"
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
                CalculationsFormList(FormNo).DataSummary = "New Calculations Data View"
                CalculationsFormList(FormNo).Version = "Version 1"
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

    Private Sub btnSPChartCompList_Click(sender As Object, e As EventArgs) Handles btnSPChartCompList.Click
        'Show the Company List form:
        If IsNothing(CompanyList) Then
            CompanyList = New frmCompanyList
            CompanyList.Show()
        Else
            CompanyList.Show()
        End If
    End Sub

    Private Sub CompanyList_FormClosed(sender As Object, e As FormClosedEventArgs) Handles CompanyList.FormClosed
        CompanyList = Nothing
    End Sub

    Private Sub btnDesignPointChartQuery_Click(sender As Object, e As EventArgs) Handles btnDesignPointChartQuery.Click
        'Open the Design Query form:

        If IsNothing(DesignPointChartQuery) Then
            DesignPointChartQuery = New frmDesignQuery
            DesignPointChartQuery.Text = "Design Cross Plot Chart Data Query"
            DesignPointChartQuery.Show()
            DesignPointChartQuery.DatabasePath = txtPointChartDbPath.Text
        Else
            DesignPointChartQuery.Show()
        End If
    End Sub

    Private Sub DesignPointChartQuery_FormClosed(sender As Object, e As FormClosedEventArgs) Handles DesignPointChartQuery.FormClosed
        DesignPointChartQuery = Nothing
    End Sub

    Private Sub btnWebPages_Click(sender As Object, e As EventArgs) Handles btnWebPages.Click
        'Open the Web Pages form.

        If IsNothing(WebPageList) Then
            WebPageList = New frmWebPageList
            WebPageList.Show()
        Else
            WebPageList.Show()
            WebPageList.BringToFront()
        End If
    End Sub

    Private Sub WebPageList_FormClosed(sender As Object, e As FormClosedEventArgs) Handles WebPageList.FormClosed
        WebPageList = Nothing
    End Sub

    Public Function OpenNewWebPage() As Integer
        'Open a new HTML Web View window, or reuse an existing one if avaiable.
        'The new forms index number in WebViewFormList is returned.

        NewWebPage = New frmWebPage
        If WebPageFormList.Count = 0 Then
            WebPageFormList.Add(NewWebPage)
            WebPageFormList(0).FormNo = 0
            WebPageFormList(0).Show
            Return 0 'The new HTML Display is at position 0 in WebViewFormList()
        Else
            Dim I As Integer
            Dim FormAdded As Boolean = False
            For I = 0 To WebPageFormList.Count - 1 'Check if there are closed forms in WebViewFormList. They can be re-used.
                If IsNothing(WebPageFormList(I)) Then
                    WebPageFormList(I) = NewWebPage
                    WebPageFormList(I).FormNo = I
                    WebPageFormList(I).Show
                    FormAdded = True
                    Return I 'The new Html Display is at position I in WebViewFormList()
                    Exit For
                End If
            Next
            If FormAdded = False Then 'Add a new form to WebViewFormList
                Dim FormNo As Integer
                WebPageFormList.Add(NewWebPage)
                FormNo = WebPageFormList.Count - 1
                WebPageFormList(FormNo).FormNo = FormNo
                WebPageFormList(FormNo).Show
                Return FormNo 'The new WebPage is at position FormNo in WebPageFormList()
            End If

        End If
    End Function

    Public Sub WebPageFormClosed()
        'This subroutine is called when the Web Page form has been closed.
        'The subroutine is usually called from the FormClosed event of the WebPage form.
        'The WebPage form may have multiple instances.
        'The ClosedFormNumber property should contains the number of the instance of the WebPage form.
        'This property should be updated by the WebPage form when it is being closed.
        'The ClosedFormNumber property value is used to determine which element in WebPageList should be set to Nothing.

        If WebPageFormList.Count < ClosedFormNo + 1 Then
            'ClosedFormNo is too large to exist in WebPageFormList
            Exit Sub
        End If

        If IsNothing(WebPageFormList(ClosedFormNo)) Then
            'The form is already set to nothing
        Else
            WebPageFormList(ClosedFormNo) = Nothing
        End If
    End Sub

    Public Function OpenNewHtmlDisplayPage() As Integer
        'Open a new HTML display window, or reuse an existing one if avaiable.
        'The new forms index number in HtmlDisplayFormList is returned.

        NewHtmlDisplay = New frmHtmlDisplay
        If HtmlDisplayFormList.Count = 0 Then
            HtmlDisplayFormList.Add(NewHtmlDisplay)
            HtmlDisplayFormList(0).FormNo = 0
            HtmlDisplayFormList(0).Show
            Return 0 'The new HTML Display is at position 0 in HtmlDisplayFormList()
        Else
            Dim I As Integer
            Dim FormAdded As Boolean = False
            For I = 0 To HtmlDisplayFormList.Count - 1 'Check if there are closed forms in HtmlDisplayFormList. They can be re-used.
                If IsNothing(HtmlDisplayFormList(I)) Then
                    HtmlDisplayFormList(I) = NewHtmlDisplay
                    HtmlDisplayFormList(I).FormNo = I
                    HtmlDisplayFormList(I).Show
                    FormAdded = True
                    Return I 'The new Html Display is at position I in HtmlDisplayFormList()
                    Exit For
                End If
            Next
            If FormAdded = False Then 'Add a new form to HtmlDisplayFormList
                Dim FormNo As Integer
                HtmlDisplayFormList.Add(NewHtmlDisplay)
                FormNo = HtmlDisplayFormList.Count - 1
                HtmlDisplayFormList(FormNo).FormNo = FormNo
                HtmlDisplayFormList(FormNo).Show
                Return FormNo 'The new HtmlDisplay is at position FormNo in HtmlDisplayFormList()
            End If

        End If

    End Function

#End Region 'Open and Close Forms -------------------------------------------------------------------------------------------------------------------------------------------------------------

#Region " Form Methods - The main actions performed by this form." '---------------------------------------------------------------------------------------------------------------------------

    Private Sub btnProject_Click(sender As Object, e As EventArgs) Handles btnProject.Click
        Project.SelectProject()
    End Sub

    Private Sub btnParameters_Click(sender As Object, e As EventArgs) Handles btnParameters.Click
        Project.ShowParameters()
    End Sub

    Private Sub TabPage1_Enter(sender As Object, e As EventArgs) Handles TabPage1.Enter
        'This tab page shows the project information.
        'The current project duration is updated when the page is entered.
        'Timer2 is also set so the duration is updated every 5 seconds.

        txtCurrentDuration.Text = Project.Usage.CurrentDuration.Days.ToString.PadLeft(5, "0"c) & ":" &
                                  Project.Usage.CurrentDuration.Hours.ToString.PadLeft(2, "0"c) & ":" &
                                  Project.Usage.CurrentDuration.Minutes.ToString.PadLeft(2, "0"c) & ":" &
                                  Project.Usage.CurrentDuration.Seconds.ToString.PadLeft(2, "0"c)

        Timer2.Interval = 5000 '5 seconds
        Timer2.Enabled = True
        Timer2.Start()
    End Sub

    Private Sub Timer2_Tick(sender As Object, e As EventArgs) Handles Timer2.Tick
        'Update the current duration:

        txtCurrentDuration.Text = Project.Usage.CurrentDuration.Days.ToString.PadLeft(5, "0"c) & ":" &
                           Project.Usage.CurrentDuration.Hours.ToString.PadLeft(2, "0"c) & ":" &
                           Project.Usage.CurrentDuration.Minutes.ToString.PadLeft(2, "0"c) & ":" &
                           Project.Usage.CurrentDuration.Seconds.ToString.PadLeft(2, "0"c)
    End Sub

    Private Sub TabPage1_Leave(sender As Object, e As EventArgs) Handles TabPage1.Leave
        Timer2.Enabled = False
    End Sub

    Private Sub btnAppInfo_Click(sender As Object, e As EventArgs) Handles btnAppInfo.Click
        ApplicationInfo.ShowInfo()
    End Sub

    Private Sub btnAndorville_Click(sender As Object, e As EventArgs) Handles btnAndorville.Click
        ApplicationInfo.ShowInfo()
    End Sub

    Private Sub btnOpenProject_Click(sender As Object, e As EventArgs) Handles btnOpenProject.Click
        If Project.Type = ADVL_Utilities_Library_1.Project.Types.Archive Then

        Else
            Process.Start(Project.Path)
        End If
    End Sub

    Private Sub btnOpenSettings_Click(sender As Object, e As EventArgs) Handles btnOpenSettings.Click
        If Project.SettingsLocn.Type = ADVL_Utilities_Library_1.FileLocation.Types.Directory Then
            Process.Start(Project.SettingsLocn.Path)
        End If
    End Sub

    Private Sub btnOpenData_Click(sender As Object, e As EventArgs) Handles btnOpenData.Click
        If Project.DataLocn.Type = ADVL_Utilities_Library_1.FileLocation.Types.Directory Then
            Process.Start(Project.DataLocn.Path)
        End If
    End Sub

    Private Sub btnOpenSystem_Click(sender As Object, e As EventArgs) Handles btnOpenSystem.Click
        If Project.SystemLocn.Type = ADVL_Utilities_Library_1.FileLocation.Types.Directory Then
            Process.Start(Project.SystemLocn.Path)
        End If
    End Sub

    Private Sub btnOpenAppDir_Click(sender As Object, e As EventArgs) Handles btnOpenAppDir.Click
        Process.Start(ApplicationInfo.ApplicationDir)
    End Sub

    Private Sub DeleteSettingsFileToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DeleteSettingsFileToolStripMenuItem.Click
        'Delete the Settings File corresponding to the text box clicked.
        'Right-click in the Settings File text box (eg in the Copy Data tab) and select Delete Settings File to delete the file.
        '  This is used to delete settings files that are no longer needed.

        Message.Add("ContextMenuStrip1.SourceControl.Name = " & ContextMenuStrip1.SourceControl.Name & vbCrLf)

        Select Case ContextMenuStrip1.SourceControl.Name
            Case "txtCopyDataSettings" 'Delete the selected Copy Data settings file.
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

    Public Sub UpdateWebPage(ByVal FileName As String)
        'Update the web page in WebPageFormList if the Web file name is FileName.

        Dim NPages As Integer = WebPageFormList.Count
        Dim I As Integer

        For I = 0 To NPages - 1
            If WebPageFormList(I) Is Nothing Then
                'Message.Add("Web page not displayed." & vbCrLf)
            Else
                If WebPageFormList(I).FileName = FileName Then
                    WebPageFormList(I).OpenDocument
                End If
            End If
        Next
    End Sub


#Region " Start Page Code" '=========================================================================================================================================

    Public Sub OpenStartPage()
        'Open the StartPage.html file and display in the Start Page tab.

        If Project.DataFileExists("StartPage.html") Then
            StartPageFileName = "StartPage.html"
            DisplayStartPage()
        Else
            CreateStartPage()
            StartPageFileName = "StartPage.html"
            DisplayStartPage()
        End If

    End Sub

    Public Sub DisplayStartPage()
        'Display the StartPage.html file in the Start Page tab.

        If Project.DataFileExists(StartPageFileName) Then
            Dim rtbData As New IO.MemoryStream
            Project.ReadData(StartPageFileName, rtbData)
            rtbData.Position = 0
            Dim sr As New IO.StreamReader(rtbData)
            WebBrowser1.DocumentText = sr.ReadToEnd()
        Else
            Message.AddWarning("Web page file not found: " & StartPageFileName & vbCrLf)
        End If
    End Sub

    Private Sub CreateStartPage()
        'Create a new default StartPage.html file.

        Dim htmData As New IO.MemoryStream
        Dim sw As New IO.StreamWriter(htmData)
        sw.Write(AppInfoHtmlString("Application Information")) 'Create a web page providing information about the application.
        sw.Flush()
        Project.SaveData("StartPage.html", htmData)
    End Sub

    Public Function AppInfoHtmlString(ByVal DocumentTitle As String) As String
        'Create an Application Information Web Page.

        'This function should be edited to provide a brief description of the Application.

        Dim sb As New System.Text.StringBuilder

        sb.Append("<!DOCTYPE html>" & vbCrLf)
        sb.Append("<html>" & vbCrLf)
        sb.Append("<head>" & vbCrLf)
        sb.Append("<title>" & DocumentTitle & "</title>" & vbCrLf)
        sb.Append("</head>" & vbCrLf)

        sb.Append("<body style=""font-family:arial;"">" & vbCrLf & vbCrLf)

        sb.Append("<h2>" & "Andorville&trade; Shares Application" & "</h2>" & vbCrLf & vbCrLf) 'Add the page title.
        sb.Append("<hr>" & vbCrLf) 'Add a horizontal divider line.
        sb.Append("<p>The Shares application stores historical data for publicly traded shares and applies data processing and analysis techniques aimed at optimizing share trading returns.</p>" & vbCrLf) 'Add an application description.
        sb.Append("<hr>" & vbCrLf & vbCrLf) 'Add a horizontal divider line.

        sb.Append(DefaultJavaScriptString)

        sb.Append("</body>" & vbCrLf)
        sb.Append("</html>" & vbCrLf)

        Return sb.ToString

    End Function

    Public Function DefaultJavaScriptString() As String
        'Generate the default JavaScript section of an Andorville(TM) Workflow Web Page.

        Dim sb As New System.Text.StringBuilder

        'Add JavaScript section:
        sb.Append("<script>" & vbCrLf & vbCrLf)

        'START: User defined JavaScript functions ==========================================================================
        'Add functions to implement the main actions performed by this web page.
        sb.Append("//START: User defined JavaScript functions ==========================================================================" & vbCrLf)
        sb.Append("//  Add functions to implement the main actions performed by this web page." & vbCrLf & vbCrLf)

        sb.Append("//END:   User defined JavaScript functions __________________________________________________________________________" & vbCrLf & vbCrLf & vbCrLf)
        'END:   User defined JavaScript functions --------------------------------------------------------------------------


        'START: User modified JavaScript functions ==========================================================================
        'Modify these function to save all required web page settings and process all expected XMessage instructions.
        sb.Append("//START: User modified JavaScript functions ==========================================================================" & vbCrLf)
        sb.Append("//  Modify these function to save all required web page settings and process all expected XMessage instructions." & vbCrLf & vbCrLf)

        'Add the SaveSettings function - This is used to save web page settings between sessions.
        sb.Append("//Save the web page settings." & vbCrLf)
        sb.Append("function SaveSettings() {" & vbCrLf)
        sb.Append("  var xSettings = ""<Settings>"" + "" \n"" ; //String containing the web page settings in XML format." & vbCrLf)
        sb.Append("  //Add xml lines to save each setting." & vbCrLf & vbCrLf)
        sb.Append("  xSettings +=    ""</Settings>"" + ""\n"" ; //End of the Settings element." & vbCrLf)
        sb.Append(vbCrLf)
        sb.Append("  //Save the settings as an XML file in the project." & vbCrLf)
        sb.Append("  window.external.SaveHtmlSettings(xSettings) ;" & vbCrLf)
        sb.Append("}" & vbCrLf)
        sb.Append(vbCrLf)

        'Process a single XMsg instruction (Information:Location pair)
        sb.Append("//Process an XMessage instruction:" & vbCrLf)
        sb.Append("function XMsgInstruction(Info, Locn) {" & vbCrLf)
        sb.Append("  switch(Locn) {" & vbCrLf)
        sb.Append("  //Insert case statements here." & vbCrLf)
        sb.Append("  default:" & vbCrLf)
        sb.Append("    window.external.AddWarning(""Unknown location: "" + Locn + ""\r\n"") ;" & vbCrLf)
        sb.Append("  }" & vbCrLf)
        sb.Append("}" & vbCrLf)
        sb.Append(vbCrLf)

        sb.Append("//END:   User modified JavaScript functions __________________________________________________________________________" & vbCrLf & vbCrLf & vbCrLf)
        'END:   User modified JavaScript functions --------------------------------------------------------------------------

        'START: Required Document Library Web Page JavaScript functions ==========================================================================
        sb.Append("//START: Required Document Library Web Page JavaScript functions ==========================================================================" & vbCrLf & vbCrLf)

        'Add the AddText function - This sends a message to the message window using a named text type.
        sb.Append("//Add text to the Message window using a named txt type:" & vbCrLf)
        sb.Append("function AddText(Msg, TextType) {" & vbCrLf)
        sb.Append("  window.external.AddText(Msg, TextType) ;" & vbCrLf)
        sb.Append("}" & vbCrLf)
        sb.Append(vbCrLf)

        'Add the AddMessage function - This sends a message to the message window using default black text.
        sb.Append("//Add a message to the Message window using the default black text:" & vbCrLf)
        sb.Append("function AddMessage(Msg) {" & vbCrLf)
        sb.Append("  window.external.AddMessage(Msg) ;" & vbCrLf)
        sb.Append("}" & vbCrLf)
        sb.Append(vbCrLf)

        'Add the AddWarning function - This sends a red, bold warning message to the message window.
        sb.Append("//Add a warning message to the Message window using bold red text:" & vbCrLf)
        sb.Append("function AddWarning(Msg) {" & vbCrLf)
        sb.Append("  window.external.AddWarning(Msg) ;" & vbCrLf)
        sb.Append("}" & vbCrLf)
        sb.Append(vbCrLf)

        'Add the RestoreSettings function - This is used to restore web page settings.
        sb.Append("//Restore the web page settings." & vbCrLf)
        sb.Append("function RestoreSettings() {" & vbCrLf)
        sb.Append("  window.external.RestoreHtmlSettings() " & vbCrLf)
        sb.Append("}" & vbCrLf)
        sb.Append(vbCrLf)

        'This line runs the RestoreSettings function when the web page is loaded.
        sb.Append("//Restore the web page settings when the page loads." & vbCrLf)
        sb.Append("window.onload = RestoreSettings; " & vbCrLf)
        sb.Append(vbCrLf)

        'Restores a single setting on the web page.
        sb.Append("//Restore a web page setting." & vbCrLf)
        sb.Append("  function RestoreSetting(FormName, ItemName, ItemValue) {" & vbCrLf)
        sb.Append("  document.forms[FormName][ItemName].value = ItemValue ;" & vbCrLf)
        sb.Append("}" & vbCrLf)
        sb.Append(vbCrLf)

        'Add the RestoreOption function - This is used to add an option to a Select list.
        sb.Append("//Restore a Select control Option." & vbCrLf)
        sb.Append("function RestoreOption(SelectId, OptionText) {" & vbCrLf)
        sb.Append("  var x = document.getElementById(SelectId) ;" & vbCrLf)
        sb.Append("  var option = document.createElement(""Option"") ;" & vbCrLf)
        sb.Append("  option.text = OptionText ;" & vbCrLf)
        sb.Append("  x.add(option) ;" & vbCrLf)
        sb.Append("}" & vbCrLf)
        sb.Append(vbCrLf)

        sb.Append("//END:   Required Document Library Web Page JavaScript functions __________________________________________________________________________" & vbCrLf & vbCrLf)
        'END:   Required Document Library Web Page JavaScript functions --------------------------------------------------------------------------

        sb.Append("</script>" & vbCrLf & vbCrLf)

        Return sb.ToString

    End Function


    Public Function DefaultHtmlString(ByVal DocumentTitle As String) As String
        'Create a blank HTML Web Page.

        Dim sb As New System.Text.StringBuilder

        sb.Append("<!DOCTYPE html>" & vbCrLf)
        sb.Append("<html>" & vbCrLf)
        sb.Append("<head>" & vbCrLf)
        sb.Append("<title>" & DocumentTitle & "</title>" & vbCrLf)
        sb.Append("</head>" & vbCrLf)

        sb.Append("<body style=""font-family:arial;"">" & vbCrLf & vbCrLf)

        sb.Append("<h1>" & DocumentTitle & "</h1>" & vbCrLf & vbCrLf)

        sb.Append(DefaultJavaScriptString)

        sb.Append("</body>" & vbCrLf)
        sb.Append("</html>" & vbCrLf)

        Return sb.ToString

    End Function

#End Region 'Start Page Code ------------------------------------------------------------------------------------------------------------------------------------------------------------------



#Region " Methods Called by JavaScript - A collection of methods that can be called by JavaScript in a web page shown in WebBrowser1" '========================================================
    'These methods are used to display HTML pages in the Document tab.
    'The same methods can be found in the WebView form, which displays web pages on seprate forms.
    'NOTE: ANY NEW METHODS SHOULD ALSO BE ADDED TO THE WebView FORM CODE!

    Public Sub JSMethodTest1()
        'Test method that is called from JavaScript.
        Message.Add("JSMethodTest1 called OK." & vbCrLf)
    End Sub

    Public Sub JSMethodTest2(ByVal Var1 As String, ByVal Var2 As String)
        'Test method that is called from JavaScript.
        Message.Add("Var1 = " & Var1 & " Var2 = " & Var2 & vbCrLf)
    End Sub

    Public Sub JSDisplayXml(ByRef XDoc As XDocument)
        Message.Add(XDoc.ToString & vbCrLf & vbCrLf)
    End Sub

    'Show a message in the application message window. USE LATER VERSION BELOW?
    Public Sub ShowMessage(ByVal Msg As String)
        Message.Add(Msg)
    End Sub

    Public Sub SaveHtmlSettings(ByVal xSettings As String, ByVal FileName As String)
        'Save the Html settings for a web page.

        'Convert the XSettings to XML format:

        Dim XmlHeader As String = "<?xml version=""1.0"" encoding=""utf-8"" standalone=""yes""?>"

        Dim XDocSettings As New System.Xml.Linq.XDocument

        Try
            XDocSettings = System.Xml.Linq.XDocument.Parse(XmlHeader & vbCrLf & xSettings)
        Catch ex As Exception
            Message.AddWarning("Error saving HTML settings file. " & ex.Message & vbCrLf)
        End Try

        Project.SaveXmlData(FileName, XDocSettings)

    End Sub


    Public Sub RestoreHtmlSettings()
        'Restore the Html settings for a web page.

        Dim SettingsFileName As String = StartPageFileName & "Settings"

        Dim XDocSettings As New System.Xml.Linq.XDocument
        Project.ReadXmlData(SettingsFileName, XDocSettings)

        If XDocSettings Is Nothing Then
            'Message.Add("No HTML Settings file : " & SettingsFileName & vbCrLf)
        Else
            Dim XSettings As New System.Xml.XmlDocument
            Try
                XSettings.LoadXml(XDocSettings.ToString)
                'Run the Settings file:
                WebXSeq.RunXSequence(XSettings, Status)
            Catch ex As Exception
                Message.AddWarning("Error restoring HTML settings. " & ex.Message & vbCrLf)
            End Try
        End If
    End Sub

    Private Sub WebXSeq_ErrorMsg(ErrMsg As String) Handles WebXSeq.ErrorMsg
        Message.AddWarning(ErrMsg & vbCrLf)
    End Sub


    Private Sub WebXSeq_Instruction(Info As String, Locn As String) Handles WebXSeq.Instruction
        'Execute each instruction produced by running the XSeq file.

        Select Case Locn
            Case "Settings:Form:Name"
                FormName = Info

            Case "Settings:Form:Item:Name"
                ItemName = Info

            Case "Settings:Form:Item:Value"
                RestoreSetting(FormName, ItemName, Info)

            Case "Settings:Form:SelectId"
                SelectId = Info

            Case "Settings:Form:OptionText"
                RestoreOption(SelectId, Info)


            Case "Settings"

            Case "EndOfSequence"
                'Main.Message.Add("End of processing sequence" & Info & vbCrLf)

            Case Else
                Message.AddWarning("Unknown location: " & Locn & "  Info: " & Info & vbCrLf)

        End Select
    End Sub


    Public Sub RestoreSetting(ByVal FormName As String, ByVal ItemName As String, ByVal ItemValue As String)
        'Restore the setting value with the specified Form Name and Item Name.

        Me.WebBrowser1.Document.InvokeScript("RestoreSetting", New String() {FormName, ItemName, ItemValue})

    End Sub

    Public Sub RestoreOption(ByVal SelectId As String, ByVal OptionText As String)
        'Restore the Option text in the Select control with the Id SelectId.

        Me.WebBrowser1.Document.InvokeScript("RestoreOption", New String() {SelectId, OptionText})
    End Sub

    Private Sub SaveWebPageSettings()
        'Call the SaveSettings JavaScript function:
        Try
            Me.WebBrowser1.Document.InvokeScript("SaveSettings")
        Catch ex As Exception
            Message.AddWarning("Web page settings not saved: " & ex.Message & vbCrLf)
        End Try

    End Sub

    Public Function GetFormNo() As String
        'Return FormNo.ToString
        Return "-1"
    End Function

    'Add text to the application message window with the specified text type.
    Public Sub AddText(ByVal Msg As String, ByVal TextType As String)
        Message.AddText(Msg, TextType)
    End Sub

    Public Sub AddMessage(ByVal Msg As String)
        Message.Add(Msg)
    End Sub

    Public Sub AddWarning(ByVal Msg As String)
        Message.AddWarning(Msg)
    End Sub


    Public Sub SendXMessage(ByVal ConnName As String, ByVal XMsg As String)
        'Send the XMessage to the application with the connection name ConnName.

    End Sub

    Public Sub RunXSequence(ByVal XSequence As String)
        'Run the XMSequence
        Dim XmlSeq As New System.Xml.XmlDocument
        XmlSeq.LoadXml(XSequence)
        XSeq.RunXSequence(XmlSeq, Status)

    End Sub

    Public Sub OpenWebPage(ByVal WebPageFileName As String)
        'Open a Web Page from the WebPageFileName.
        '  Pass the ParentName Property to the new web page. The is the name of this web page that is opening the new page.
        '  Pass the ParentWebPageFormNo Property to the new web page. This is the FormNo of this web page that is opening the new page.
        '    A hash code is generated from the ParentName. This is used to define a file name to save and restore the Web Page settings.
        '    The new web page can pass instructions back to the ParentWebPage using its ParentWebPageFormNo.

        Dim NewFormNo As Integer = OpenNewWebPage()

        WebPageFormList(NewFormNo).ParentWebPageFileName = StartPageFileName 'Set the Parent Web Page property.
        WebPageFormList(NewFormNo).ParentWebPageFormNo = -1 'Set the Parent Form Number property.
        WebPageFormList(NewFormNo).Description = ""             'The web page description can be blank.
        WebPageFormList(NewFormNo).FileDirectory = ""           'Only Web files in the Project directory can be opened from another Web Page Form.
        WebPageFormList(NewFormNo).FileName = WebPageFileName  'Set the web page file name to be opened.
        WebPageFormList(NewFormNo).OpenDocument                'Open the web page file name.

    End Sub

    Public Sub ParentProjectName(ByVal FormName As String, ByVal ItemName As String)
        'Return the Parent Project name:
        RestoreSetting(FormName, ItemName, Project.ParentProjectName)
    End Sub

    Public Sub ParentProjectPath(ByVal FormName As String, ByVal ItemName As String)
        'Return the Parent Project path:
        RestoreSetting(FormName, ItemName, Project.ParentProjectPath)
    End Sub

    Public Sub ParentProjectParameterValue(ByVal FormName As String, ByVal ItemName As String, ByVal ParameterName As String)
        'Return the specified Parent Project parameter value:
        RestoreSetting(FormName, ItemName, Project.ParentParameter(ParameterName).Value)
    End Sub

    Public Sub ProjectParameterValue(ByVal FormName As String, ByVal ItemName As String, ByVal ParameterName As String)
        'Return the specified Project parameter value:
        RestoreSetting(FormName, ItemName, Project.Parameter(ParameterName).Value)
    End Sub

    Public Sub ApplicationNetworkName(ByVal FormName As String, ByVal ItemName As String)
        'Return the name of the Application Network:
        RestoreSetting(FormName, ItemName, Project.Parameter("AppNetName").Value)
    End Sub


#End Region 'Methods Called by JavaScript -----------------------------------------------------------------------------------------------------------------------------------------------------



#Region " Project Events Code"

    Private Sub Project_Message(Msg As String) Handles Project.Message
        'Display the Project message:
        Message.Add(Msg & vbCrLf)
    End Sub

    Private Sub Project_ErrorMessage(Msg As String) Handles Project.ErrorMessage
        'Display the Project error message:
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

        Project.ReadParameters()
        Project.ReadParentParameters()
        If Project.ParentParameterExists("AppNetName") Then
            Project.AddParameter("AppNetName", Project.ParentParameter("AppNetName").Value, Project.ParentParameter("AppNetName").Description) 'AddParameter will update the parameter if it already exists.
            AppNetName = Project.Parameter("AppNetName").Value
        Else
            AppNetName = Project.GetParameter("AppNetName")
        End If

        Project.LockProject() 'Lock the project while it is open in this application.

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
        'Connect to or disconnect from the Communication Network (Message Service).
        If ConnectedToComNet = False Then
            ConnectToComNet()
        Else
            DisconnectFromComNet()
        End If
    End Sub


    Private Sub ConnectToComNet()
        'Connect to the Communication Network. (Message Service)

        If IsNothing(client) Then
            client = New ServiceReference1.MsgServiceClient(New System.ServiceModel.InstanceContext(New MsgServiceCallback))
        End If

        If ComNetRunning() Then
            'The Message Service is Running.
        Else  'The Message Service is NOT Running.
            'Start the Message Service:
            If System.IO.File.Exists(MsgServiceExePath) Then 'OK to start the Message Service application:
                Shell(Chr(34) & MsgServiceExePath & Chr(34), AppWinStyle.NormalFocus) 'Start Message Service application with no argument
            Else
                'Incorrect Message Service Executable path.
                Message.AddWarning("Message Service exe file not found. Service not started." & vbCrLf)
            End If
        End If

        If client.State = ServiceModel.CommunicationState.Faulted Then
            Message.AddWarning("Client state is faulted. Connection not made!" & vbCrLf)
        Else
            Try
                'client.Endpoint.Binding.SendTimeout = New System.TimeSpan(0, 0, 8) 'Temporarily set the send timeaout to 8 seconds
                client.Endpoint.Binding.SendTimeout = New System.TimeSpan(0, 0, 16) 'Temporarily set the send timeaout to 16 seconds
                ConnectionName = ApplicationInfo.Name 'This name will be modified if it is already used in an existing connection.
                ConnectionName = client.Connect(AppNetName, ApplicationInfo.Name, ConnectionName, Project.Name, Project.Description, Project.Type, Project.Path, False, False) 'UPDATED 2Feb19

                If ConnectionName <> "" Then
                    Message.Add("Connected to the Communication Network as " & ConnectionName & vbCrLf)
                    client.Endpoint.Binding.SendTimeout = New System.TimeSpan(1, 0, 0) 'Restore the send timeaout to 1 hour
                    btnOnline.Text = "Online"
                    btnOnline.ForeColor = Color.ForestGreen
                    ConnectedToComNet = True
                    SendApplicationInfo()
                    client.GetMessageServiceAppInfoAsync() 'Update the Exe Path in case it has changed. This path may be needed in the future to start the ComNet (Message Service).
                Else
                    Message.Add("Connection to the Communication Network failed!" & vbCrLf)
                    client.Endpoint.Binding.SendTimeout = New System.TimeSpan(1, 0, 0) 'Restore the send timeaout to 1 hour
                End If
            Catch ex As System.TimeoutException
                Message.Add("Timeout error. Check if the Communication Network is running." & vbCrLf)
            Catch ex As Exception
                Message.Add("Error message: " & ex.Message & vbCrLf)
                client.Endpoint.Binding.SendTimeout = New System.TimeSpan(1, 0, 0) 'Restore the send timeaout to 1 hour
            End Try
        End If
    End Sub

    Private Sub ConnectToComNet(ByVal ConnName As String)
        'Connect to the Communication Network with the connection name ConnName.

        'If ConnectedToAppnet = False Then
        If ConnectedToComNet = False Then
            Dim Result As Boolean

            If IsNothing(client) Then
                client = New ServiceReference1.MsgServiceClient(New System.ServiceModel.InstanceContext(New MsgServiceCallback))
            End If

            If client.State = ServiceModel.CommunicationState.Faulted Then
                Message.AddWarning("client state is faulted. Connection not made!" & vbCrLf)
            Else
                Try
                    'client.Endpoint.Binding.SendTimeout = New System.TimeSpan(0, 0, 8) 'Temporarily set the send timeaout to 8 seconds
                    client.Endpoint.Binding.SendTimeout = New System.TimeSpan(0, 0, 16) 'Temporarily set the send timeaout to 16 seconds
                    ConnectionName = ConnName 'This name will be modified if it is already used in an existing connection.
                    ConnectionName = client.Connect(AppNetName, ApplicationInfo.Name, ConnectionName, Project.Name, Project.Description, Project.Type, Project.Path, False, False) 'UPDATED 2Feb19

                    If ConnectionName <> "" Then
                        Message.Add("Connected to the Communication Network as " & ConnectionName & vbCrLf)
                        client.Endpoint.Binding.SendTimeout = New System.TimeSpan(1, 0, 0) 'Restore the send timeaout to 1 hour
                        btnOnline.Text = "Online"
                        btnOnline.ForeColor = Color.ForestGreen
                        ConnectedToComNet = True
                        SendApplicationInfo()
                    Else
                        Message.Add("Connection to the Communication Network failed!" & vbCrLf)
                        client.Endpoint.Binding.SendTimeout = New System.TimeSpan(1, 0, 0) 'Restore the send timeaout to 1 hour
                    End If
                Catch ex As System.TimeoutException
                    Message.Add("Timeout error. Check if the Communication Network is running." & vbCrLf)
                Catch ex As Exception
                    Message.Add("Error message: " & ex.Message & vbCrLf)
                    client.Endpoint.Binding.SendTimeout = New System.TimeSpan(1, 0, 0) 'Restore the send timeaout to 1 hour
                End Try
            End If
        Else
            Message.AddWarning("Already connected to the Communication Network." & vbCrLf)
        End If

    End Sub

    Private Sub DisconnectFromComNet()
        'Disconnect from the Communication Network.

        If ConnectedToComNet = True Then
            If IsNothing(client) Then
                Message.Add("Already disconnected from the Communication Network." & vbCrLf)
                btnOnline.Text = "Offline"
                btnOnline.ForeColor = Color.Red
                ConnectedToComNet = False
                ConnectionName = ""
            Else
                If client.State = ServiceModel.CommunicationState.Faulted Then
                    Message.Add("client state is faulted." & vbCrLf)
                    ConnectionName = ""
                Else
                    Try
                        client.Disconnect(AppNetName, ConnectionName)
                        btnOnline.Text = "Offline"
                        btnOnline.ForeColor = Color.Red
                        ConnectedToComNet = False
                        ConnectionName = ""
                        Message.Add("Disconnected from the Communication Network." & vbCrLf)
                    Catch ex As Exception
                        Message.AddWarning("Error disconnecting from Communication Network: " & ex.Message & vbCrLf)
                    End Try
                End If
            End If
        End If
    End Sub

    Private Function ComNetRunning() As Boolean
        'Return True if ComNet (Message Service) is running.
        If MsgServiceAppPath = "" Then
            Message.Add("Message Service application path is not known." & vbCrLf)
            Message.Add("Run the Message Service before connecting to update the path." & vbCrLf)
            Return False
        Else
            If System.IO.File.Exists(MsgServiceAppPath & "\Application.Lock") Then
                Message.Add("AppLock found - ComNet is running." & vbCrLf)
                Return True
            Else
                Message.Add("AppLock not found - ComNet is running." & vbCrLf)
                Return False
            End If
        End If

    End Function

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

                Dim text As New XElement("Text", "Share Analysis")
                applicationInfo.Add(text)

                Dim exePath As New XElement("ExecutablePath", Me.ApplicationInfo.ExecutablePath)
                applicationInfo.Add(exePath)

                Dim directory As New XElement("Directory", Me.ApplicationInfo.ApplicationDir)
                applicationInfo.Add(directory)
                Dim description As New XElement("Description", Me.ApplicationInfo.Description)
                applicationInfo.Add(description)
                xmessage.Add(applicationInfo)
                doc.Add(xmessage)

                'Show the message sent to AppNet:
                Message.XAddText("Message sent to " & "MessageService" & ":" & vbCrLf, "XmlSentNotice")
                Message.XAddXml(doc.ToString)
                Message.XAddText(vbCrLf, "Normal") 'Add extra line

                client.SendMessage("", "MessageService", doc.ToString) 'UPDATED 2Feb19
            End If
        End If

    End Sub

#End Region 'Online/Offline code

    Private Sub btnAdd_Click(sender As Object, e As EventArgs) Handles btnAdd.Click
        'Add the current project to the Message Service list.

        If Project.ParentProjectName <> "" Then
            Message.AddWarning("This project has a parent: " & Project.ParentProjectName & vbCrLf)
            Message.AddWarning("Child projects can not be added to the list." & vbCrLf)
            Exit Sub
        End If

        If ConnectedToComNet = False Then
            Message.AddWarning("The application is not connected to the Message Service." & vbCrLf)
        Else 'Connected to the Message Service (ComNet).
            If IsNothing(client) Then
                Message.Add("No client connection available!" & vbCrLf)
            Else
                If client.State = ServiceModel.CommunicationState.Faulted Then
                    Message.Add("Client state is faulted. Message not sent!" & vbCrLf)
                Else
                    'Construct the XMessage to send to AppNet:
                    Dim decl As New XDeclaration("1.0", "utf-8", "yes")
                    Dim doc As New XDocument(decl, Nothing) 'Create an XDocument to store the instructions.
                    Dim xmessage As New XElement("XMsg") 'This indicates the start of the message in the XMessage class
                    Dim projectInfo As New XElement("ProjectInfo")

                    Dim Path As New XElement("Path", Project.Path)
                    projectInfo.Add(Path)
                    xmessage.Add(projectInfo)
                    doc.Add(xmessage)

                    'Show the message sent to AppNet:
                    Message.XAddText("Message sent to " & "MessageService" & ":" & vbCrLf, "XmlSentNotice")
                    Message.XAddXml(doc.ToString)
                    Message.XAddText(vbCrLf, "Normal") 'Add extra line
                    client.SendMessage("", "MessageService", doc.ToString)
                End If
            End If
        End If
    End Sub

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

        If IsDBNull(Info) Then
            Info = ""
        End If

        Select Case Locn

            Case "ClientAppNetName"
                ClientAppNetName = Info 'The name of the Client Application Network requesting service. 

            Case "ClientName"
                ClientAppName = Info 'The name of the Client requesting service.

            Case "ClientConnectionName"
                ClientConnName = Info 'The name of the client requesting service.

            Case "ClientLocn" 'The Location within the Client requesting service.
                Dim statusOK As New XElement("Status", "OK") 'Add Status OK element when the Client Location is changed
                xlocns(xlocns.Count - 1).Add(statusOK)

                xmessage.Add(xlocns(xlocns.Count - 1)) 'Add the instructions for the last location to the reply xmessage
                xlocns.Add(New XElement(Info)) 'Start the new location instructions

            Case "Main"
                 'Blank message - do nothing.

            Case "Main:Status"
                Select Case Info
                    Case "OK"
                        'Main instructions completed OK
                End Select

            Case "StockChart"
                 'Blank message - do nothing.

            Case "StockChart:Status"
                Select Case Info
                    Case "OK"
                        'Stock Chart instructions completed OK
                End Select

            Case "PointChart"
                'Blank message - do nothing.

            Case "PointChart:Status"
                Select Case Info
                    Case "OK"
                        'Point Chart instructions completed OK
                End Select

           'Stock Chart instructions: ---------------------------------------------------------------------------------------------------

            Case "StockChart:Settings:Command"
                Select Case Info
                    Case "ClearChart"
                        ClearStockChartDefaults()
                    Case "OK"
                        'StockChartDefaults has been updated. Display in rtbStockChartDefaults.
                        rtbStockChartDefaults.Text = StockChartDefaults.ToString
                        FormatXmlText(rtbStockChartDefaults)
                End Select

            Case "StockChart:Settings:InputData:Type"
                StockChartDefaults.<StockChart>.<Settings>.<InputData>.<Type>.Value = Info
            Case "StockChart:Settings:InputData:DatabasePath"
                StockChartDefaults.<StockChart>.<Settings>.<InputData>.<DatabasePath>.Value = Info
            Case "StockChart:Settings:InputData:DataDescription"
                StockChartDefaults.<StockChart>.<Settings>.<InputData>.<DataDescription>.Value = Info
            Case "StockChart:Settings:InputData:DatabaseQuery"
                StockChartDefaults.<StockChart>.<Settings>.<InputData>.<DatabaseQuery>.Value = Info

            Case "StockChart:Settings:ChartProperties:SeriesName"
                StockChartDefaults.<StockChart>.<Settings>.<ChartProperties>.<SeriesName>.Value = Info
            Case "StockChart:Settings:ChartProperties:XValuesFieldName"
                StockChartDefaults.<StockChart>.<Settings>.<ChartProperties>.<XValuesFieldName>.Value = Info
            Case "StockChart:Settings:ChartProperties:YValuesHighFieldName"
                StockChartDefaults.<StockChart>.<Settings>.<ChartProperties>.<YValuesHighFieldName>.Value = Info
            Case "StockChart:Settings:ChartProperties:YValuesLowFieldName"
                StockChartDefaults.<StockChart>.<Settings>.<ChartProperties>.<YValuesLowFieldName>.Value = Info
            Case "StockChart:Settings:ChartProperties:YValuesOpenFieldName"
                StockChartDefaults.<StockChart>.<Settings>.<ChartProperties>.<YValuesOpenFieldName>.Value = Info
            Case "StockChart:Settings:ChartProperties:YValuesCloseFieldName"
                StockChartDefaults.<StockChart>.<Settings>.<ChartProperties>.<YValuesCloseFieldName>.Value = Info

            Case "StockChart:Settings:ChartTitle:LabelName"
                StockChartDefaults.<StockChart>.<Settings>.<ChartTitle>.<LabelName>.Value = Info
            Case "StockChart:Settings:ChartTitle:Text"
                StockChartDefaults.<StockChart>.<Settings>.<ChartTitle>.<Text>.Value = Info
            Case "StockChart:Settings:ChartTitle:FontName"
                StockChartDefaults.<StockChart>.<Settings>.<ChartTitle>.<FontName>.Value = Info
            Case "StockChart:Settings:ChartTitle:Color"
                StockChartDefaults.<StockChart>.<Settings>.<ChartTitle>.<Color>.Value = Info
            Case "StockChart:Settings:ChartTitle:Size"
                StockChartDefaults.<StockChart>.<Settings>.<ChartTitle>.<Size>.Value = Info
            Case "StockChart:Settings:ChartTitle:Bold"
                StockChartDefaults.<StockChart>.<Settings>.<ChartTitle>.<Bold>.Value = Info
            Case "StockChart:Settings:ChartTitle:Italic"
                StockChartDefaults.<StockChart>.<Settings>.<ChartTitle>.<Italic>.Value = Info
            Case "StockChart:Settings:ChartTitle:Underline"
                StockChartDefaults.<StockChart>.<Settings>.<ChartTitle>.<Underline>.Value = Info
            Case "StockChart:Settings:ChartTitle:Strikeout"
                StockChartDefaults.<StockChart>.<Settings>.<ChartTitle>.<Strikeout>.Value = Info

            Case "StockChart:Settings:XAxis:TitleText"
                StockChartDefaults.<StockChart>.<Settings>.<XAxis>.<TitleText>.Value = Info
            Case "StockChart:Settings:XAxis:TitleFontName"
                StockChartDefaults.<StockChart>.<Settings>.<XAxis>.<TitleFontName>.Value = Info
            Case "StockChart:Settings:XAxis:TitleColor"
                StockChartDefaults.<StockChart>.<Settings>.<XAxis>.<TitleColor>.Value = Info
            Case "StockChart:Settings:XAxis:TitleSize"
                StockChartDefaults.<StockChart>.<Settings>.<XAxis>.<TitleSize>.Value = Info
            Case "StockChart:Settings:XAxis:TitleBold"
                StockChartDefaults.<StockChart>.<Settings>.<XAxis>.<TitleBold>.Value = Info
            Case "StockChart:Settings:XAxis:TitleItalic"
                StockChartDefaults.<StockChart>.<Settings>.<XAxis>.<TitleItalic>.Value = Info
            Case "StockChart:Settings:XAxis:TitleUnderline"
                StockChartDefaults.<StockChart>.<Settings>.<XAxis>.<TitleUnderline>.Value = Info
            Case "StockChart:Settings:XAxis:TitleStrikeout"
                StockChartDefaults.<StockChart>.<Settings>.<XAxis>.<TitleStrikeout>.Value = Info
            Case "StockChart:Settings:XAxis:TitleAlignment"
                StockChartDefaults.<StockChart>.<Settings>.<XAxis>.<TitleAlignment>.Value = Info
            Case "StockChart:Settings:XAxis:AutoMinimum"
                StockChartDefaults.<StockChart>.<Settings>.<XAxis>.<AutoMinimum>.Value = Info
            Case "StockChart:Settings:XAxis:Minimum"
                StockChartDefaults.<StockChart>.<Settings>.<XAxis>.<Minimum>.Value = Info
            Case "StockChart:Settings:XAxis:AutoMaximum"
                StockChartDefaults.<StockChart>.<Settings>.<XAxis>.<AutoMaximum>.Value = Info
            Case "StockChart:Settings:XAxis:Maximum"
                StockChartDefaults.<StockChart>.<Settings>.<XAxis>.<Maximum>.Value = Info
            Case "StockChart:Settings:XAxis:AutoInterval"
                StockChartDefaults.<StockChart>.<Settings>.<XAxis>.<AutoInterval>.Value = Info
            Case "StockChart:Settings:XAxis:Interval"
                StockChartDefaults.<StockChart>.<Settings>.<XAxis>.<Interval>.Value = Info
            Case "StockChart:Settings:XAxis:AutoMajorGridInterval"
                StockChartDefaults.<StockChart>.<Settings>.<XAxis>.<AutoMajorGridInterval>.Value = Info
            Case "StockChart:Settings:XAxis:MajorGridInterval"
                StockChartDefaults.<StockChart>.<Settings>.<XAxis>.<MajorGridInterval>.Value = Info

            Case "StockChart:Settings:YAxis:TitleText"
                StockChartDefaults.<StockChart>.<Settings>.<YAxis>.<TitleText>.Value = Info
            Case "StockChart:Settings:YAxis:TitleFontName"
                StockChartDefaults.<StockChart>.<Settings>.<YAxis>.<TitleFontName>.Value = Info
            Case "StockChart:Settings:YAxis:TitleColor"
                StockChartDefaults.<StockChart>.<Settings>.<YAxis>.<TitleColor>.Value = Info
            Case "StockChart:Settings:YAxis:TitleSize"
                StockChartDefaults.<StockChart>.<Settings>.<YAxis>.<TitleSize>.Value = Info
            Case "StockChart:Settings:YAxis:TitleBold"
                StockChartDefaults.<StockChart>.<Settings>.<YAxis>.<TitleBold>.Value = Info
            Case "StockChart:Settings:YAxis:TitleItalic"
                StockChartDefaults.<StockChart>.<Settings>.<YAxis>.<TitleItalic>.Value = Info
            Case "StockChart:Settings:YAxis:TitleUnderline"
                StockChartDefaults.<StockChart>.<Settings>.<YAxis>.<TitleUnderline>.Value = Info
            Case "StockChart:Settings:YAxis:TitleStrikeout"
                StockChartDefaults.<StockChart>.<Settings>.<YAxis>.<TitleStrikeout>.Value = Info
            Case "StockChart:Settings:YAxis:TitleAlignment"
                StockChartDefaults.<StockChart>.<Settings>.<YAxis>.<TitleAlignment>.Value = Info
            Case "StockChart:Settings:YAxis:AutoMinimum"
                StockChartDefaults.<StockChart>.<Settings>.<YAxis>.<AutoMinimum>.Value = Info
            Case "StockChart:Settings:YAxis:Minimum"
                StockChartDefaults.<StockChart>.<Settings>.<YAxis>.<Minimum>.Value = Info
            Case "StockChart:Settings:YAxis:AutoMaximum"
                StockChartDefaults.<StockChart>.<Settings>.<YAxis>.<AutoMaximum>.Value = Info
            Case "StockChart:Settings:YAxis:Maximum"
                StockChartDefaults.<StockChart>.<Settings>.<YAxis>.<Maximum>.Value = Info

            Case "StockChart:Settings:YAxis:AutoInterval"
                StockChartDefaults.<StockChart>.<Settings>.<YAxis>.<AutoInterval>.Value = Info
            Case "StockChart:Settings:YAxis:Interval"
                StockChartDefaults.<StockChart>.<Settings>.<YAxis>.<Interval>.Value = Info
            Case "StockChart:Settings:YAxis:AutoMajorGridInterval"
                StockChartDefaults.<StockChart>.<Settings>.<YAxis>.<AutoMajorGridInterval>.Value = Info
            Case "StockChart:Settings:YAxis:MajorGridInterval"
                StockChartDefaults.<StockChart>.<Settings>.<YAxis>.<MajorGridInterval>.Value = Info

           'Point Chart instructions: ---------------------------------------------------------------------------------------------------
            Case "PointChart:Settings:Command"
                Select Case Info
                    Case "ClearChart"
                        ClearPointChartDefaults()
                    Case "OK"
                        'PointChartDefaults has been updated. Display in rtbPointChartDefaults.
                        rtbPointChartDefaults.Text = PointChartDefaults.ToString
                        FormatXmlText(rtbPointChartDefaults)
                End Select

            Case "PointChart:Settings:InputData:Type"
                PointChartDefaults.<PointChart>.<Settings>.<InputData>.<Type>.Value = Info
            Case "PointChart:Settings:InputData:DatabasePath"
                PointChartDefaults.<PointChart>.<Settings>.<InputData>.<DatabasePath>.Value = Info
            Case "PointChart:Settings:InputData:DataDescription"
                PointChartDefaults.<PointChart>.<Settings>.<InputData>.<DataDescription>.Value = Info
            Case "PointChart:Settings:InputData:DatabaseQuery"
                PointChartDefaults.<PointChart>.<Settings>.<InputData>.<DatabaseQuery>.Value = Info

            Case "PointChart:Settings:ChartProperties:SeriesName"
                PointChartDefaults.<StockChart>.<Settings>.<ChartProperties>.<SeriesName>.Value = Info

            Case "PointChart:Settings:ChartProperties:XValuesFieldName"
                PointChartDefaults.<StockChart>.<Settings>.<ChartProperties>.<XValuesFieldName>.Value = Info

            Case "PointChart:Settings:ChartProperties:YValuesFieldName"
                PointChartDefaults.<StockChart>.<Settings>.<ChartProperties>.<YValuesFieldName>.Value = Info

            Case "PointChart:Settings:ChartTitle:LabelName"
                PointChartDefaults.<PointChart>.<Settings>.<ChartTitle>.<LabelName>.Value = Info
            Case "PointChart:Settings:ChartTitle:Text"
                PointChartDefaults.<PointChart>.<Settings>.<ChartTitle>.<Text>.Value = Info
            Case "PointChart:Settings:ChartTitle:FontName"
                PointChartDefaults.<PointChart>.<Settings>.<ChartTitle>.<FontName>.Value = Info
            Case "PointChart:Settings:ChartTitle:Color"
                PointChartDefaults.<PointChart>.<Settings>.<ChartTitle>.<Color>.Value = Info
            Case "PointChart:Settings:ChartTitle:Size"
                PointChartDefaults.<PointChart>.<Settings>.<ChartTitle>.<Size>.Value = Info
            Case "PointChart:Settings:ChartTitle:Bold"
                PointChartDefaults.<PointChart>.<Settings>.<ChartTitle>.<Bold>.Value = Info
            Case "PointChart:Settings:ChartTitle:Italic"
                PointChartDefaults.<PointChart>.<Settings>.<ChartTitle>.<Italic>.Value = Info
            Case "PointChart:Settings:ChartTitle:Underline"
                PointChartDefaults.<PointChart>.<Settings>.<ChartTitle>.<Underline>.Value = Info
            Case "PointChart:Settings:ChartTitle:Strikeout"
                PointChartDefaults.<PointChart>.<Settings>.<ChartTitle>.<Strikeout>.Value = Info
            Case "PointChart:Settings:ChartTitle:Alignment"
                PointChartDefaults.<PointChart>.<Settings>.<ChartTitle>.<Alignment>.Value = Info

            Case "PointChart:Settings:XAxis:TitleText"
                PointChartDefaults.<PointChart>.<Settings>.<XAxis>.<TitleText>.Value = Info
            Case "PointChart:Settings:XAxis:TitleFontName"
                PointChartDefaults.<PointChart>.<Settings>.<XAxis>.<TitleFontName>.Value = Info
            Case "PointChart:Settings:XAxis:TitleColor"
                PointChartDefaults.<PointChart>.<Settings>.<XAxis>.<TitleColor>.Value = Info
            Case "PointChart:Settings:XAxis:TitleSize"
                PointChartDefaults.<PointChart>.<Settings>.<XAxis>.<TitleSize>.Value = Info
            Case "PointChart:Settings:XAxis:TitleBold"
                PointChartDefaults.<PointChart>.<Settings>.<XAxis>.<TitleBold>.Value = Info
            Case "PointChart:Settings:XAxis:TitleItalic"
                PointChartDefaults.<PointChart>.<Settings>.<XAxis>.<TitleItalic>.Value = Info
            Case "PointChart:Settings:XAxis:TitleUnderline"
                PointChartDefaults.<PointChart>.<Settings>.<XAxis>.<TitleUnderline>.Value = Info
            Case "PointChart:Settings:XAxis:TitleStrikeout"
                PointChartDefaults.<PointChart>.<Settings>.<XAxis>.<TitleStrikeout>.Value = Info
            Case "PointChart:Settings:XAxis:TitleAlignment"
                PointChartDefaults.<PointChart>.<Settings>.<XAxis>.<TitleAlignment>.Value = Info
            Case "PointChart:Settings:XAxis:AutoMinimum"
                PointChartDefaults.<PointChart>.<Settings>.<XAxis>.<AutoMinimum>.Value = Info
            Case "PointChart:Settings:XAxis:Minimum"
                PointChartDefaults.<PointChart>.<Settings>.<XAxis>.<Minimum>.Value = Info
            Case "PointChart:Settings:XAxis:AutoMaximum"
                PointChartDefaults.<PointChart>.<Settings>.<XAxis>.<AutoMaximum>.Value = Info
            Case "PointChart:Settings:XAxis:Maximum"
                PointChartDefaults.<PointChart>.<Settings>.<XAxis>.<Maximum>.Value = Info
            Case "PointChart:Settings:XAxis:AutoInterval"
                PointChartDefaults.<PointChart>.<Settings>.<XAxis>.<AutoInterval>.Value = Info
            Case "PointChart:Settings:XAxis:Interval"
                PointChartDefaults.<PointChart>.<Settings>.<XAxis>.<Interval>.Value = Info
            Case "PointChart:Settings:XAxis:AutoMajorGridInterval"
                PointChartDefaults.<PointChart>.<Settings>.<XAxis>.<AutoMajorGridInterval>.Value = Info
            Case "PointChart:Settings:XAxis:MajorGridInterval"
                PointChartDefaults.<PointChart>.<Settings>.<XAxis>.<MajorGridInterval>.Value = Info

            Case "PointChart:Settings:YAxis:TitleText"
                PointChartDefaults.<PointChart>.<Settings>.<YAxis>.<TitleText>.Value = Info
            Case "PointChart:Settings:YAxis:TitleFontName"
                PointChartDefaults.<PointChart>.<Settings>.<YAxis>.<TitleFontName>.Value = Info
            Case "PointChart:Settings:YAxis:TitleColor"
                PointChartDefaults.<PointChart>.<Settings>.<YAxis>.<TitleColor>.Value = Info
            Case "PointChart:Settings:YAxis:TitleSize"
                PointChartDefaults.<PointChart>.<Settings>.<YAxis>.<TitleSize>.Value = Info
            Case "PointChart:Settings:YAxis:TitleBold"
                PointChartDefaults.<PointChart>.<Settings>.<YAxis>.<TitleBold>.Value = Info
            Case "PointChart:Settings:YAxis:TitleItalic"
                PointChartDefaults.<PointChart>.<Settings>.<YAxis>.<TitleItalic>.Value = Info
            Case "PointChart:Settings:YAxis:TitleUnderline"
                PointChartDefaults.<PointChart>.<Settings>.<YAxis>.<TitleUnderline>.Value = Info
            Case "PointChart:Settings:YAxis:TitleStrikeout"
                PointChartDefaults.<PointChart>.<Settings>.<YAxis>.<TitleStrikeout>.Value = Info
            Case "PointChart:Settings:YAxis:TitleAlignment"
                PointChartDefaults.<PointChart>.<Settings>.<YAxis>.<TitleAlignment>.Value = Info
            Case "PointChart:Settings:YAxis:AutoMinimum"
                PointChartDefaults.<PointChart>.<Settings>.<YAxis>.<AutoMinimum>.Value = Info
            Case "PointChart:Settings:YAxis:Minimum"
                PointChartDefaults.<PointChart>.<Settings>.<YAxis>.<Minimum>.Value = Info
            Case "PointChart:Settings:YAxis:AutoMaximum"
                PointChartDefaults.<PointChart>.<Settings>.<YAxis>.<AutoMaximum>.Value = Info
            Case "PointChart:Settings:YAxis:Maximum"
                PointChartDefaults.<PointChart>.<Settings>.<YAxis>.<Maximum>.Value = Info
            Case "PointChart:Settings:YAxis:AutoInterval"
                PointChartDefaults.<PointChart>.<Settings>.<YAxis>.<AutoInterval>.Value = Info
            Case "PointChart:Settings:YAxis:Interval"
                PointChartDefaults.<PointChart>.<Settings>.<YAxis>.<Interval>.Value = Info
            Case "PointChart:Settings:YAxis:AutoMajorGridInterval"
                PointChartDefaults.<PointChart>.<Settings>.<YAxis>.<AutoMajorGridInterval>.Value = Info
            Case "PointChart:Settings:YAxis:MajorGridInterval"
                PointChartDefaults.<PointChart>.<Settings>.<YAxis>.<MajorGridInterval>.Value = Info

           'Get Share Information ====================================================

           'Get GICS List
            Case "GetGicsList"
                If Info = "OK" Then
                    GetGicsList()
                End If

          'Get company list in GICS group. Info contains the GICS code.
            Case "GetGicsCompanyList"
                GetGicsCompanyList(Info)

           'Get company name. Info contains the ASX code.
            Case "GetCompanyName"
                GetCompanyName(Info)



           'Startup Command Arguments ================================================
            Case "ProjectName"
                If Project.OpenProject(Info) = True Then
                    ProjectSelected = True 'Project has been opened OK.
                Else
                    ProjectSelected = False 'Project could not be opened.
                End If

            Case "ProjectID"
                Message.AddWarning("Add code to handle ProjectID parameter at StartUp!" & vbCrLf)
                'Note the AppNet will usually select a project using ProjectPath.

            Case "ProjectPath"
                If Project.OpenProjectPath(Info) = True Then
                    ProjectSelected = True 'Project has been opened OK.

                Else
                    ProjectSelected = False 'Project could not be opened.
                End If

            Case "ConnectionName"
                StartupConnectionName = Info
            '--------------------------------------------------------------------------

            'Application Information  =================================================
            'returned by client.GetMessageServiceAppInfoAsync()
            Case "MessageServiceAppInfo:Name"
                'The name of the Message Service Application. (Not used.)

            Case "MessageServiceAppInfo:ExePath"
                'The executable file path of the Message Service Application.
                MsgServiceExePath = Info

            Case "MessageServiceAppInfo:Path"
                'The path of the Message Service Application (ComNet). (This is where an Application.Lock file will be found while ComNet is running.)
                MsgServiceAppPath = Info
            '---------------------------------------------------------------------------

            'Show Share Price Table ====================================================
            Case "ShowSharePriceTable:Command"
                Select Case Info
                    Case "Apply"
                        If SharePricesFormNo = -1 Then
                            Message.AddWarning("The Share Prices Form Number is not known." & vbCrLf)
                        Else
                            SharePricesFormList(SharePricesFormNo).ApplyQuery
                        End If
                    Case "OpenNewForm"
                        SharePricesFormNo = AppendSharePricesView()

                End Select

            Case "ShowSharePriceTable:Query"
                If SharePricesFormNo = -1 Then
                    'Message.AddWarning("The Share Prices Form Number is not known." & vbCrLf)
                    SharePricesFormNo = AppendSharePricesView()
                    SharePricesFormList(SharePricesFormNo).Query = Info
                Else
                    SharePricesFormList(SharePricesFormNo).Query = Info
                End If

            '---------------------------------------------------------------------------

            'Show Share Price Chart ====================================================
            Case "ShowSharePriceChart:Query"
                txtSPChartQuery.Text = Info 'Specify the Query used to extract the data to chart.
                UpdateChartSharePricesTab()

            Case "ShowSharePriceChart:SeriesName"
                txtSeriesName.Text = Info 'Set the Series Name -  the name of the series of points being charted.

            Case "ShowSharePriceChart:ChartTitle"
                txtChartTitle.Text = Info 'Set the Chart Title.

            Case "ShowSharePriceChart:Command"
                Select Case Info
                    Case "Apply"
                        DisplayStockChart()
                    Case Else
                        Message.AddWarning("Unknown ShowSharePriceChart command: " & Info & vbCrLf)
                End Select

            '---------------------------------------------------------------------------

            Case "EndOfSequence"
                'End of Information Vector Sequence reached.
                'Add Status OK element at the end of the sequence:
                Dim statusOK As New XElement("Status", "OK") 'Add Status OK element at the end of the sequence
                xlocns(xlocns.Count - 1).Add(statusOK)

            Case Else
                Message.AddWarning("Unknown location: " & Locn & vbCrLf)
                Message.AddWarning("            info: " & Info & vbCrLf)
        End Select

    End Sub

    Private Sub ClearStockChartDefaults()
        'Clear the settings in the StockChartDefaults XDocument.

        StockChartDefaults = <?xml version="1.0" encoding="utf-8"?>
                             <!---->
                             <StockChart>
                                 <Settings>
                                     <InputData>
                                         <Type>Database</Type>
                                         <DatabasePath></DatabasePath>
                                         <DataDescription></DataDescription>
                                         <DatabaseQuery></DatabaseQuery>
                                     </InputData>
                                     <ChartProperties>
                                         <SeriesName></SeriesName>
                                         <XValuesFieldName></XValuesFieldName>
                                         <YValuesHighFieldName></YValuesHighFieldName>
                                         <YValuesLowFieldName></YValuesLowFieldName>
                                         <YValuesOpenFieldName></YValuesOpenFieldName>
                                         <YValuesCloseFieldName></YValuesCloseFieldName>
                                     </ChartProperties>
                                     <ChartTitle>
                                         <LabelName>Label1</LabelName>
                                         <Text></Text>
                                         <FontName>Arial</FontName>
                                         <Color>Black</Color>
                                         <Size>12</Size>
                                         <Bold>true</Bold>
                                         <Italic>false</Italic>
                                         <Underline>false</Underline>
                                         <Strikeout>false</Strikeout>
                                         <Alignment>TopCenter</Alignment>
                                     </ChartTitle>
                                     <XAxis>
                                         <TitleText></TitleText>
                                         <TitleFontName>Arial</TitleFontName>
                                         <TitleColor>Black</TitleColor>
                                         <TitleSize>14</TitleSize>
                                         <TitleBold>true</TitleBold>
                                         <TitleItalic>false</TitleItalic>
                                         <TitleUnderline>false</TitleUnderline>
                                         <TitleStrikeout>false</TitleStrikeout>
                                         <TitleAlignment>Center</TitleAlignment>
                                         <AutoMinimum>true</AutoMinimum>
                                         <Minimum>0</Minimum>
                                         <AutoMaximum>true</AutoMaximum>
                                         <Maximum>1</Maximum>
                                         <AutoInterval>true</AutoInterval>
                                         <Interval>0</Interval>
                                         <AutoMajorGridInterval>true</AutoMajorGridInterval>
                                         <MajorGridInterval>0</MajorGridInterval>
                                     </XAxis>
                                     <YAxis>
                                         <TitleText></TitleText>
                                         <TitleFontName>Arial</TitleFontName>
                                         <TitleColor>Black</TitleColor>
                                         <TitleSize>14</TitleSize>
                                         <TitleBold>true</TitleBold>
                                         <TitleItalic>false</TitleItalic>
                                         <TitleUnderline>false</TitleUnderline>
                                         <TitleStrikeout>false</TitleStrikeout>
                                         <TitleAlignment>Center</TitleAlignment>
                                         <AutoMinimum>true</AutoMinimum>
                                         <Minimum>0</Minimum>
                                         <AutoMaximum>true</AutoMaximum>
                                         <Maximum>1</Maximum>
                                         <AutoInterval>true</AutoInterval>
                                         <Interval>0</Interval>
                                         <AutoMajorGridInterval>true</AutoMajorGridInterval>
                                         <MajorGridInterval>0</MajorGridInterval>
                                     </YAxis>
                                 </Settings>
                             </StockChart>


    End Sub

    Private Sub ClearPointChartDefaults()
        'Clear the settings in the PointChartDefaults XDocument.

        PointChartDefaults = <?xml version="1.0" encoding="utf-8"?>
                             <!---->
                             <PointChart>
                                 <Settings>
                                     <InputData>
                                         <Type>Database</Type>
                                         <DatabasePath></DatabasePath>
                                         <DataDescription></DataDescription>
                                         <DatabaseQuery></DatabaseQuery>
                                     </InputData>
                                     <ChartProperties>
                                         <SeriesName></SeriesName>
                                         <XValuesFieldName></XValuesFieldName>
                                         <YValuesFieldName></YValuesFieldName>
                                     </ChartProperties>
                                     <ChartTitle>
                                         <LabelName>Label1</LabelName>
                                         <Text></Text>
                                         <FontName>Arial</FontName>
                                         <Color>Black</Color>
                                         <Size>12</Size>
                                         <Bold>true</Bold>
                                         <Italic>false</Italic>
                                         <Underline>false</Underline>
                                         <Strikeout>false</Strikeout>
                                         <Alignment>TopCenter</Alignment>
                                     </ChartTitle>
                                     <XAxis>
                                         <TitleText></TitleText>
                                         <TitleFontName>Arial</TitleFontName>
                                         <TitleColor>Black</TitleColor>
                                         <TitleSize>14</TitleSize>
                                         <TitleBold>true</TitleBold>
                                         <TitleItalic>false</TitleItalic>
                                         <TitleUnderline>false</TitleUnderline>
                                         <TitleStrikeout>false</TitleStrikeout>
                                         <TitleAlignment>Center</TitleAlignment>
                                         <AutoMinimum>true</AutoMinimum>
                                         <Minimum>0</Minimum>
                                         <AutoMaximum>true</AutoMaximum>
                                         <Maximum>1</Maximum>
                                         <AutoInterval>true</AutoInterval>
                                         <Interval>0</Interval>
                                         <AutoMajorGridInterval>true</AutoMajorGridInterval>
                                         <MajorGridInterval>0</MajorGridInterval>
                                     </XAxis>
                                     <YAxis>
                                         <TitleText></TitleText>
                                         <TitleFontName>Arial</TitleFontName>
                                         <TitleColor>Black</TitleColor>
                                         <TitleSize>14</TitleSize>
                                         <TitleBold>true</TitleBold>
                                         <TitleItalic>false</TitleItalic>
                                         <TitleUnderline>false</TitleUnderline>
                                         <TitleStrikeout>false</TitleStrikeout>
                                         <TitleAlignment>Center</TitleAlignment>
                                         <AutoMinimum>true</AutoMinimum>
                                         <Minimum>0</Minimum>
                                         <AutoMaximum>true</AutoMaximum>
                                         <Maximum>1</Maximum>
                                         <AutoInterval>true</AutoInterval>
                                         <Interval>0</Interval>
                                         <AutoMajorGridInterval>true</AutoMajorGridInterval>
                                         <MajorGridInterval>0</MajorGridInterval>
                                     </YAxis>
                                 </Settings>
                             </PointChart>
    End Sub

    Private Sub GetGicsList()
        'Return the GICS list in an XMessage.

        If SharePriceDbPath = "" Then
            Message.AddWarning("A share price database has not been selected!" & vbCrLf)
            Exit Sub
        End If

        Dim connString As String
        Dim myConnection As OleDb.OleDbConnection = New OleDb.OleDbConnection
        Dim Query As String = "Select Distinct GICS_Industry_Group From ASX_Company_List"
        Dim ds As DataSet = New DataSet
        Dim da As OleDb.OleDbDataAdapter

        connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source =" & SharePriceDbPath 'DatabasePath
        myConnection.ConnectionString = connString
        myConnection.Open()

        da = New OleDb.OleDbDataAdapter(Query, myConnection)
        da.MissingSchemaAction = MissingSchemaAction.AddWithKey

        Try
            da.Fill(ds, "myData")

            Dim list As New XElement("GICS_List")
            Dim I As Integer
            For I = 0 To ds.Tables(0).Rows.Count - 1
                Dim GicsGroupName As New XElement("GicsGroupName", ds.Tables(0).Rows(I).Item(0))
                list.Add(GicsGroupName)
            Next
            xlocns(xlocns.Count - 1).Add(list)

        Catch ex As Exception
            Message.AddWarning("Error getting GICS list: " & ex.Message & vbCrLf)
        End Try
    End Sub

    Private Sub GetGicsCompanyList(ByVal GicsGroup As String)
        'Return the company list in the GICS group an XMessage.

        If SharePriceDbPath = "" Then
            Message.AddWarning("A share price database has not been selected!" & vbCrLf)
            Exit Sub
        End If

        Dim connString As String
        Dim myConnection As OleDb.OleDbConnection = New OleDb.OleDbConnection
        Dim Query As String = "Select ASX_Code From ASX_Company_List Where GICS_Industry_Group = '" & GicsGroup & "'"
        Dim ds As DataSet = New DataSet
        Dim da As OleDb.OleDbDataAdapter

        connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source =" & SharePriceDbPath 'DatabasePath
        myConnection.ConnectionString = connString
        myConnection.Open()

        da = New OleDb.OleDbDataAdapter(Query, myConnection)
        da.MissingSchemaAction = MissingSchemaAction.AddWithKey

        Try
            da.Fill(ds, "myData")
            Dim list As New XElement("Company_List")
            Dim I As Integer
            For I = 0 To ds.Tables(0).Rows.Count - 1
                Dim AsxCode As New XElement("AsxCode", ds.Tables(0).Rows(I).Item(0))
                list.Add(AsxCode)
            Next
            xlocns(xlocns.Count - 1).Add(list)
        Catch ex As Exception
            Message.AddWarning("Error getting Company List: " & ex.Message & vbCrLf)
        End Try
    End Sub

    Private Sub GetCompanyName(ByVal AsxCode As String)
        'Return the company name corresponding to the AsxCode in an XMessage.

        If SharePriceDbPath = "" Then
            Message.AddWarning("A share price database has not been selected!" & vbCrLf)
            Exit Sub
        End If

        Dim connString As String
        Dim myConnection As OleDb.OleDbConnection = New OleDb.OleDbConnection
        Dim Query As String = "Select Company_Name From ASX_Company_List Where ASX_Code = '" & AsxCode & "'"
        Dim ds As DataSet = New DataSet
        Dim da As OleDb.OleDbDataAdapter

        connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source =" & SharePriceDbPath 'DatabasePath
        myConnection.ConnectionString = connString
        myConnection.Open()

        da = New OleDb.OleDbDataAdapter(Query, myConnection)
        da.MissingSchemaAction = MissingSchemaAction.AddWithKey

        Try
            da.Fill(ds, "myData")

            If ds.Tables(0).Rows.Count = 0 Then
                Message.AddWarning("No company found with the ASX Code: " & AsxCode & vbCrLf)
                Dim CompanyName As New XElement("CompanyName", "")
                xlocns(xlocns.Count - 1).Add(CompanyName)
            Else
                If ds.Tables(0).Rows.Count = 1 Then
                    Dim CompanyName As New XElement("CompanyName", ds.Tables(0).Rows(0).Item(0))
                    xlocns(xlocns.Count - 1).Add(CompanyName)
                Else
                    Message.AddWarning(ds.Tables(0).Rows.Count & " companies found with the ASX Code: " & AsxCode & vbCrLf)
                    Dim CompanyName As New XElement("CompanyName", "")
                    xlocns(xlocns.Count - 1).Add(CompanyName)
                End If
            End If

        Catch ex As Exception
            Message.AddWarning("Error getting Company Name: " & ex.Message & vbCrLf)
        End Try

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
                    client.SendMessage(ClientAppNetName, ClientConnName, MessageText) 'Added 2Feb19
                    MessageText = "" 'Clear the message after it has been sent.
                    ClientAppName = "" 'Clear the Client Application Name after the message has been sent.
                    ClientConnName = "" 'Clear the Client Application Name after the message has been sent.
                    xlocns.Clear()
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
            'Set the corresponding Project Parameter:
            Project.AddParameter("SharePriceDatabasePath", SharePriceDbPath, "The path of the Share Price database.")
            Project.SaveParameters()
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
            'Set the corresponding Project Parameter:
            Project.AddParameter("FinancialsDatabasePath", FinancialsDbPath, "The path of the Historical Financials database.")
            Project.SaveParameters()
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
            'Set the corresponding Project Parameter:
            Project.AddParameter("CalculationsDatabasePath", CalculationsDbPath, "The path of the Calculations database.")
            Project.SaveParameters()
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


    Private Sub btnInsertViewSPAfter_Click(sender As Object, e As EventArgs) Handles btnInsertViewSPAfter.Click
        'Insert a new Share Prices view after the item selected in the share prices list.
        'If no item is selected, insert the new view at the end of the list.
        'This button was labelled "Insert After"
        'It is now labelled "New"

        Dim SelectedIndex As Integer = lstSharePrices.SelectedIndex 'The index of the selected view.
        Dim NViews As Integer = lstSharePrices.Items.Count 'The number of views in the Share Prices list.
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
            SharePricesFormList(1).DataSummary = "New Share Price Data View"
            SharePricesFormList(1).Version = "Version 1"
        Else
            If SelectedIndex >= 0 Then
                If SelectedIndex = NViews - 1 Then 'Last item selected
                    'Add the new View to the end of the list.
                    lstSharePrices.Items.Add("")
                    Dim NewSettings As New DataViewSettings
                    SharePricesSettings.List.Add(NewSettings)
                    OpenSharePricesFormNo(NViews)
                    SharePricesFormList(NViews).DataSummary = "New Share Price Data View"
                    SharePricesFormList(NViews).Version = "Version 1"
                Else
                    lstSharePrices.Items.Insert(SelectedIndex + 1, "")
                    'Insert a new Settings entry in SharePricesSettings:
                    Dim NewSettings As New DataViewSettings
                    SharePricesSettings.List.Insert(SelectedIndex + 1, NewSettings)
                    OpenSharePricesFormNo(SelectedIndex + 1)
                    SharePricesFormList(SelectedIndex + 1).DataSummary = "New Share Price Data View"
                    SharePricesFormList(SelectedIndex + 1).Version = "Version 1"
                End If
            Else
                'No item selected
                'Add the new View to the end of the list.
                lstSharePrices.Items.Add("")
                Dim NewSettings As New DataViewSettings
                SharePricesSettings.List.Add(NewSettings)
                OpenSharePricesFormNo(NViews + 1)
                SharePricesFormList(NViews + 1).DataSummary = "New Share Price Data View"
                SharePricesFormList(NViews + 1).Version = "Version 1"
            End If
        End If
    End Sub

    Private Function AppendSharePricesView() As Integer
        'Append a temporary Share Prices view to the list.
        'Return the Form Number of the Data View

        Dim NViews As Integer = lstSharePrices.Items.Count 'The number of views in the Share Prices list.
        lstSharePrices.Items.Add("") 'Add an entry with a blank name. - Blank entries can be removed when the application is closed.
        Dim NewSettings As New DataViewSettings
        SharePricesSettings.List.Add(NewSettings)
        OpenSharePricesFormNo(NViews + 1)
        SharePricesFormList(NViews + 1).DataSummary = "New Share Price Data View"
        SharePricesFormList(NViews + 1).Version = "Version 1"
        Return NViews + 1
    End Function

    Private Sub btnDeleteViewSP_Click(sender As Object, e As EventArgs) Handles btnDeleteViewSP.Click
        'Delete selected view

        Dim SelectedIndex As Integer = lstSharePrices.SelectedIndex 'The index of the selected view.
        Dim NViews As Integer = lstSharePrices.Items.Count 'The number of views in the Share Prices list.

        If SharePricesFormList.Count < SelectedIndex + 1 Then 'The Share Price data view is not being displayed.
            lstSharePrices.Items.RemoveAt(SelectedIndex) 'Remove the entry from the list displayed on the form.
            SharePricesSettings.List.RemoveAt(SelectedIndex)  'Delete the entry in SharePricesSettings
            Exit Sub
        End If

        If NViews = 0 Then
            'No Views to delete.
        Else
            lstSharePrices.Items.RemoveAt(SelectedIndex) 'Remove the selected View in lstSharePrices
            'Close the form if it is open:
            If IsNothing(SharePricesFormList(SelectedIndex)) Then
            Else
                SharePricesFormList(SelectedIndex).CloseForm
            End If
            'Delete the entry in SharePricesSettings
            SharePricesSettings.List.RemoveAt(SelectedIndex)
        End If
    End Sub

    Private Sub btnViewSPUp_Click(sender As Object, e As EventArgs) Handles btnViewSPUp.Click

    End Sub

    Private Sub btnViewSPDwn_Click(sender As Object, e As EventArgs) Handles btnViewSPDwn.Click

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
            FinancialsFormList(1).DataSummary = "New Financials Data View"
            FinancialsFormList(1).Version = "Version 1"
        Else
            If SelectedIndex >= 0 Then
                If SelectedIndex = NViews - 1 Then 'Last item selected
                    'Add the new View to the end of the list.
                    lstFinancials.Items.Add("")
                    Dim NewSettings As New DataViewSettings
                    FinancialsSettings.List.Add(NewSettings)
                    OpenFinancialsFormNo(NViews)
                    FinancialsFormList(NViews).DataSummary = "New Financials Data View"
                    FinancialsFormList(NViews).Version = "Version 1"
                Else
                    lstFinancials.Items.Insert(SelectedIndex + 1, "")
                    'Insert a new Settings entry in FinancialSettings:
                    Dim NewSettings As New DataViewSettings
                    FinancialsSettings.List.Insert(SelectedIndex + 1, NewSettings)
                    OpenFinancialsFormNo(SelectedIndex + 1)
                    FinancialsFormList(SelectedIndex + 1).DataSummary = "New Financials Data View"
                    FinancialsFormList(SelectedIndex + 1).Version = "Version 1"
                End If
            Else
                'No item selected
                'Add the new View to the end of the list.
                lstFinancials.Items.Add("")
                Dim NewSettings As New DataViewSettings
                FinancialsSettings.List.Add(NewSettings)
                OpenFinancialsFormNo(NViews + 1)
                FinancialsFormList(NViews + 1).DataSummary = "New Financials Data View"
                FinancialsFormList(NViews + 1).Version = "Version 1"
            End If
        End If
    End Sub



    Private Sub btnDeleteViewFin_Click(sender As Object, e As EventArgs) Handles btnDeleteViewFin.Click
        'Delete selected view

        Dim SelectedIndex As Integer = lstFinancials.SelectedIndex 'The index of the selected view.
        Dim NViews As Integer = lstFinancials.Items.Count 'The number of views in the Financials list.

        If FinancialsFormList.Count < SelectedIndex + 1 Then
            lstFinancials.Items.RemoveAt(SelectedIndex)
            Exit Sub
        End If

        If NViews = 0 Then
            'No Views to delete.
        Else
            lstFinancials.Items.RemoveAt(SelectedIndex) 'Remove the selected View in lstFinancials
            'Close the form if it is open:
            If IsNothing(FinancialsFormList(SelectedIndex)) Then
            Else
                FinancialsFormList(SelectedIndex).CloseForm
            End If
            'Delete the entry in FinancialsSettings
            FinancialsSettings.List.RemoveAt(SelectedIndex)
        End If
    End Sub

    Private Sub btnViewFinUp_Click(sender As Object, e As EventArgs) Handles btnViewFinUp.Click

    End Sub

    Private Sub btnViewFinDwn_Click(sender As Object, e As EventArgs) Handles btnViewFinDwn.Click

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
        Dim NViews As Integer = lstCalculations.Items.Count 'The number of views in the Calculations list.

        'If CalculationsDataViewList.Count < SelectedIndex + 1 Then
        If CalculationsFormList.Count < SelectedIndex + 1 Then
            lstCalculations.Items.RemoveAt(SelectedIndex)
            Exit Sub
        End If

        If NViews = 0 Then
            'No Views to delete.
        Else
            lstCalculations.Items.RemoveAt(SelectedIndex) 'Remove the selected View in lstCalculations
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
            CalculationsFormList(1).DataSummary = "New Calculations Data View"
            CalculationsFormList(1).Version = "Version 1"
        Else
            If SelectedIndex >= 0 Then
                If SelectedIndex = NViews - 1 Then 'Last item selected
                    'Add the new View to the end of the list.
                    lstCalculations.Items.Add("")
                    Dim NewSettings As New DataViewSettings
                    CalculationsSettings.List.Add(NewSettings)
                    OpenCalculationsFormNo(NViews)
                    CalculationsFormList(NViews).DataSummary = "New Calculations Data View"
                    CalculationsFormList(NViews).Version = "Version 1"
                Else
                    'Insert a new Settings entry in FinancialSettings:
                    lstCalculations.Items.Insert(SelectedIndex + 1, "")
                    Dim NewSettings As New DataViewSettings
                    CalculationsSettings.List.Insert(SelectedIndex + 1, NewSettings)
                    OpenCalculationsFormNo(SelectedIndex + 1)
                    CalculationsFormList(SelectedIndex + 1).DataSummary = "New Calculations Data View"
                    CalculationsFormList(SelectedIndex + 1).Version = "Version 1"
                End If
            Else
                'No item selected
                'Add the new View to the end of the list.
                lstCalculations.Items.Add("")
                Dim NewSettings As New DataViewSettings
                CalculationsSettings.List.Add(NewSettings)
                OpenCalculationsFormNo(NViews + 1)
                CalculationsFormList(NViews + 1).DataSummary = "New Calculations Data View"
                CalculationsFormList(NViews + 1).Version = "Version 1"
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
        txtCopyDataSettings.Text = CopyDataSettingsFile

        Dim ComboBoxCol0 As New DataGridViewComboBoxColumn
        dgvCopyData.Columns.Add(ComboBoxCol0)
        dgvCopyData.Columns(0).HeaderText = "Input Column"
        dgvCopyData.Columns(0).Width = 160

        If cmbCopyDataInputDb.SelectedIndex = -1 Then
        Else
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
        End If

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
            Message.AddWarning("No file name has been specified!" & vbCrLf)
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
        cmbCopyDataInputData.SelectedIndex = 0
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
                Query = txtCopyDataInputQuery.Text
                connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source =" & SharePriceDbPath
                myConnection.ConnectionString = connString
                myConnection.Open()
                da = New OleDb.OleDbDataAdapter(Query, myConnection)
                da.MissingSchemaAction = MissingSchemaAction.AddWithKey
                da.Fill(dsInput, "myData")
                myConnection.Close()
            Case "Financials"
                Query = txtCopyDataInputQuery.Text
                connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source =" & FinancialsDbPath
                myConnection.ConnectionString = connString
                myConnection.Open()
                da = New OleDb.OleDbDataAdapter(Query, myConnection)
                da.MissingSchemaAction = MissingSchemaAction.AddWithKey
                da.Fill(dsInput, "myData")
                myConnection.Close()
            Case "Calculations"
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

        If cmbSelectDataInputDb.SelectedIndex = -1 Then
        Else
            If cmbSelectDataInputData.SelectedIndex = -1 Then
            Else
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
            End If
        End If

        Dim ComboBoxCol1 As New DataGridViewComboBoxColumn
        dgvSelectData.Columns.Add(ComboBoxCol1)
        dgvSelectData.Columns(1).HeaderText = "Output Column"

        If cmbSelectDataOutputDb.SelectedIndex = -1 Then
            'cmbSelectDataOutputDb.SelectedIndex = 0
        End If

        If cmbSelectDataOutputData.SelectedIndex = -1 Then
            'cmbSelectDataOutputData.SelectedIndex = 0
        Else
            If cmbSelectDataOutputDb.SelectedIndex = -1 Then
            Else
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
            End If
        End If


        'Set up dgvSelectConstraints ---------------------------------------------
        dgvSelectConstraints.Columns.Clear()

        Dim ComboBoxCol20 As New DataGridViewComboBoxColumn
        dgvSelectConstraints.Columns.Add(ComboBoxCol20)
        dgvSelectConstraints.Columns(0).HeaderText = "WHERE Input Column"

        If cmbSelectDataInputDb.SelectedIndex = -1 Then
        Else
            If cmbSelectDataInputData.SelectedIndex = -1 Then
            Else
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
            End If
        End If

        Dim ComboBoxCol21 As New DataGridViewComboBoxColumn
        dgvSelectConstraints.Columns.Add(ComboBoxCol21)
        dgvSelectConstraints.Columns(1).HeaderText = "= Output Column"

        If cmbSelectDataOutputDb.SelectedIndex = -1 Then
        Else
            If cmbSelectDataOutputData.SelectedIndex = -1 Then
            Else
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
            End If
        End If

        RestoreSelectDataSelections()
    End Sub

    Private Sub RestoreSelectDataSelections()
        'Restore the Select Data settings. (Leave the input and output data selections unchanged.)

        If SelectDataSettingsFile = "" Then
            'Message.AddWarning("No Select Data settings file name has been specified!" & vbCrLf)
        Else
            Dim XDoc As System.Xml.Linq.XDocument
            Project.ReadXmlData(SelectDataSettingsFile, XDoc)

            dgvSelectData.Rows.Clear()

            Dim settings = From item In XDoc.<SelectDataSettings>.<SelectDataList>.<CopyColumn>

            For Each item In settings
                dgvSelectData.Rows.Add(item.<From>.Value, item.<To>.Value)
            Next

            dgvSelectData.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            dgvSelectData.AutoResizeColumns()

            dgvSelectConstraints.Rows.Clear()
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
            Dim settings = From item In XDoc.<SelectDataSettings>.<SelectDataList>.<CopyColumn>
            For Each item In settings
                dgvSelectData.Rows.Add(item.<From>.Value, item.<To>.Value)
            Next
            dgvSelectData.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            dgvSelectData.AutoResizeColumns()


            dgvSelectConstraints.Rows.Clear()
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
                Query = txtSelectDataInputQuery.Text
                Message.Add("Loading input Share Price data using query: " & Query & vbCrLf)
                connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source =" & SharePriceDbPath
                myConnection.ConnectionString = connString
                myConnection.Open()
                da = New OleDb.OleDbDataAdapter(Query, myConnection)
                da.MissingSchemaAction = MissingSchemaAction.AddWithKey
                da.Fill(dsInput, "myData")
                myConnection.Close()
            Case "Financials"
                Query = txtSelectDataInputQuery.Text
                Message.Add("Loading input Financial data using query: " & Query & vbCrLf)
                connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source =" & FinancialsDbPath
                myConnection.ConnectionString = connString
                myConnection.Open()
                da = New OleDb.OleDbDataAdapter(Query, myConnection)
                da.MissingSchemaAction = MissingSchemaAction.AddWithKey
                da.Fill(dsInput, "myData")
                myConnection.Close()
            Case "Calculations"
                Query = txtSelectDataInputQuery.Text
                Message.Add("Loading input Calculation data using query: " & Query & vbCrLf)
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
                outputQuery = txtSelectDataOutputQuery.Text
                Message.Add("Loading output Share Price data using query: " & outputQuery & vbCrLf)
                outputConnString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source =" & SharePriceDbPath
                outputConnection.ConnectionString = outputConnString
                outputConnection.Open()
                outputDa = New OleDb.OleDbDataAdapter(outputQuery, outputConnection)
                outputDa.MissingSchemaAction = MissingSchemaAction.AddWithKey
                Dim outputCmdBuilder As New OleDb.OleDbCommandBuilder(outputDa)
                outputDa.Fill(dsOutput, "myData")
                outputConnection.Close()
            Case "Financials"
                outputQuery = txtSelectDataOutputQuery.Text
                Message.Add("Loading output Financial data using query: " & outputQuery & vbCrLf)
                outputConnString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source =" & FinancialsDbPath
                outputConnection.ConnectionString = outputConnString
                outputConnection.Open()
                outputDa = New OleDb.OleDbDataAdapter(outputQuery, outputConnection)
                outputDa.MissingSchemaAction = MissingSchemaAction.AddWithKey
                Dim outputCmdBuilder As New OleDb.OleDbCommandBuilder(outputDa)
                outputDa.Fill(dsOutput, "myData")
                outputConnection.Close()
            Case "Calculations"
                outputQuery = txtSelectDataOutputQuery.Text
                Message.Add("Loading output Calculation data using query: " & outputQuery & vbCrLf)
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
            NRows = dgvSelectData.Rows.Count
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
        'ComboBoxCol1.Items.Add("Date") 'NOTE: This modification was tried but won't work becauser the Param dictionary can store only single values and not dates!!!

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

        If cmbSimpleCalcDb.SelectedIndex = -1 Then
            cmbSimpleCalcDb.SelectedIndex = 0
        End If
        If cmbSimpleCalcData.SelectedIndex = -1 Then

        Else
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
        End If

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

        If cmbSimpleCalcData.SelectedIndex = -1 Then
        Else
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
        End If

        dgvSimpleCalcsParameterList.AllowUserToAddRows = True 'Allow user to add rows again.
        dgvSimpleCalcsCalculations.AllowUserToAddRows = True
        dgvSimpleCalcsInputData.AllowUserToAddRows = True
        dgvSimpleCalcsOutputData.AllowUserToAddRows = True

        dgvSimpleCalcsParameterList.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
        dgvSimpleCalcsCalculations.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
        dgvSimpleCalcsInputData.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
        dgvSimpleCalcsOutputData.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

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

            dgvSimpleCalcsParameterList.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            dgvSimpleCalcsCalculations.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            dgvSimpleCalcsInputData.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            dgvSimpleCalcsOutputData.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

        End If
    End Sub

    Private Sub cmbSimpleCalcDb_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbSimpleCalcDb.SelectedIndexChanged
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
                If item.Cells(0).Value <> Nothing Then ComboBoxCol40.Items.Add(item.Cells(0).Value)
            Next
        End If

        Dim ComboBoxCol41 As New DataGridViewComboBoxColumn
        dgvSimpleCalcsOutputData.Columns.Add(ComboBoxCol41)
        dgvSimpleCalcsOutputData.Columns(1).HeaderText = "Output Column"
        dgvSimpleCalcsOutputData.Columns(1).Width = 140

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

        dgvSimpleCalcsParameterList.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
        dgvSimpleCalcsCalculations.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
        dgvSimpleCalcsInputData.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
        dgvSimpleCalcsOutputData.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

    End Sub

    Private Sub dgvSimpleCalcsParameterList_LostFocus(sender As Object, e As EventArgs) Handles dgvSimpleCalcsParameterList.LostFocus
        UpdateSimpleCalcsParams()
    End Sub

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
    End Function

    Private Function GetCalc(ByVal Input1 As String, ByVal Input2 As String, ByVal Operation As String, ByVal Output As String) As Calculation
        Dim NewCalculation As New Calculation

        NewCalculation.Input1 = Input1
        NewCalculation.Input2 = Input2
        NewCalculation.Operation = Operation
        NewCalculation.Output = Output
        Return NewCalculation
    End Function

    Private Function GetParamLocn(ByVal ParamName As String, ByVal ColName As String) As ParameterLocation
        Dim NewParamLocn As New ParameterLocation

        NewParamLocn.ParamName = ParamName
        NewParamLocn.ColName = ColName
        Return NewParamLocn
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
                Dim MonthCol As String = cmbDateCalcParam1.Text
                Dim YearCol As String = cmbDateCalcParam2.Text
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
                Dim MonthCol As String = cmbDateCalcParam1.Text
                Dim YearCol As String = cmbDateCalcParam2.Text
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
                Dim StartDateCol As String = cmbDateCalcParam1.Text
                Dim NDaysCol As String = cmbDateCalcParam2.Text
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
                Dim StartDateCol As String = cmbDateCalcParam1.Text
                Dim NDaysCol As String = cmbDateCalcParam2.Text
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
        GetDateSelectOutputDataList()
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
            Sequence.rtbSequence.SelectedText = "    <DateSelectionType>" & cmbDateSelectionType.Text & "</DateSelectionType>" & vbCrLf
            Sequence.rtbSequence.SelectedText = "    <InputDateColumn>" & cmbDateSelInputDateCol.Text & "</InputDateColumn>" & vbCrLf
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
                    myQuery = myQuery & " And " & InputDateCol & " = #" & Format(item(OutputDateCol), "MM-dd-yyyy") & "#" 'Dates in a query must have this format!

                    Dim myRecords = dsInput.Tables("myData").Select(myQuery)
                    If myRecords.Count = 0 Then
                        Message.AddWarning("No records found with this constraint: " & myQuery & vbCrLf)
                    ElseIf myRecords.Count = 1 Then
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
                    myQuery = myQuery & " And " & InputDateCol & " > #" & Format(item(OutputDateCol), "MM-dd-yyyy") & "#" 'Dates in a query must have this format!
                    mySort = InputDateCol & " ASC"

                    Dim myRecords = dsInput.Tables("myData").Select(myQuery, mySort)
                    If myRecords.Count = 0 Then
                        Message.AddWarning("No records found with this constraint: " & myQuery & vbCrLf)
                    ElseIf myRecords.Count = 1 Then
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
                    myQuery = myQuery & " And " & InputDateCol & " < #" & Format(item(OutputDateCol), "MM-dd-yyyy") & "#" 'Dates in a query must have this format!
                    mySort = InputDateCol & " DESC"

                    Dim myRecords = dsInput.Tables("myData").Select(myQuery, mySort)
                    If myRecords.Count = 0 Then
                        Message.AddWarning("No records found with this constraint: " & myQuery & vbCrLf)
                    ElseIf myRecords.Count = 1 Then
                        'Message.Add("One record found with this constraint: " & myQuery & vbCrLf)
                        For I = 1 To NCols
                            item(OutCols(I)) = myRecords(0).Item(InCols(I))
                        Next
                    Else
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
                Query = txtDateSelInputQuery.Text
                connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source =" & SharePriceDbPath
                myConnection.ConnectionString = connString
                myConnection.Open()
                da = New OleDb.OleDbDataAdapter(Query, myConnection)
                da.MissingSchemaAction = MissingSchemaAction.AddWithKey
                da.Fill(dsInput, "myData")
                myConnection.Close()
            Case "Financials"
                Query = txtDateSelInputQuery.Text
                connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source =" & FinancialsDbPath
                myConnection.ConnectionString = connString
                myConnection.Open()
                da = New OleDb.OleDbDataAdapter(Query, myConnection)
                da.MissingSchemaAction = MissingSchemaAction.AddWithKey
                da.Fill(dsInput, "myData")
                myConnection.Close()
            Case "Calculations"
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
                    Message.AddWarning("Calculations: Daily Prices: No input Share Prices database selected!" & vbCrLf)
                    Exit Sub
                Else
                    connectionString = "provider=Microsoft.ACE.OLEDB.12.0;data source = " + SharePriceDbPath
                End If
            Case "Financials"
                If FinancialsDbPath = "" Then
                    Message.AddWarning("Calculations: Daily Prices: No input Financials database selected!" & vbCrLf)
                    Exit Sub
                Else
                    connectionString = "provider=Microsoft.ACE.OLEDB.12.0;data source = " + FinancialsDbPath
                End If
            Case "Calculations"
                If CalculationsDbPath = "" Then
                    Message.AddWarning("Calculations: Daily Prices: No input Calculations database selected!" & vbCrLf)
                    Exit Sub
                Else
                    connectionString = "provider=Microsoft.ACE.OLEDB.12.0;data source = " + CalculationsDbPath
                End If
        End Select

        Dim conn As System.Data.OleDb.OleDbConnection 'Declare a connection for MS Access - used by the Data Adapter to connect to and disconnect from the database.
        Dim dt As DataTable

        cmbDailyPriceInputTable.Text = ""
        cmbDailyPriceInputTable.Items.Clear()

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
                    Message.AddWarning("Calculations: Daily Prices: No output Share Prices database selected!" & vbCrLf)
                    Exit Sub
                Else
                    connectionString = "provider=Microsoft.ACE.OLEDB.12.0;data source = " + SharePriceDbPath
                End If
            Case "Financials"
                If FinancialsDbPath = "" Then
                    Message.AddWarning("Calculations: Daily Prices: No output Financials database selected!" & vbCrLf)
                    Exit Sub
                Else
                    connectionString = "provider=Microsoft.ACE.OLEDB.12.0;data source = " + FinancialsDbPath
                End If
            Case "Calculations"
                If CalculationsDbPath = "" Then
                    Message.AddWarning("Calculations: Daily Prices: No output Calculations database selected!" & vbCrLf)
                    Exit Sub
                Else
                    connectionString = "provider=Microsoft.ACE.OLEDB.12.0;data source = " + CalculationsDbPath
                End If
        End Select

        Dim conn As System.Data.OleDb.OleDbConnection 'Declare a connection for MS Access - used by the Data Adapter to connect to and disconnect from the database.
        Dim dt As DataTable

        cmbDailyPriceOutputTable.Text = ""
        cmbDailyPriceOutputTable.Items.Clear()

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

        End If
    End Sub



#End Region 'Processing Sequence Tab ----------------------------------------------------------------------------------------------------------------------------------------------------------

#Region " Charts Tab" '========================================================================================================================================================================


    Private Sub DataGridView1_DataError(sender As Object, e As DataGridViewDataErrorEventArgs) Handles DataGridView1.DataError
        'This code stops an unnecessary error dialog appearing.
        If e.Context = DataGridViewDataErrorContexts.Formatting Or e.Context = DataGridViewDataErrorContexts.PreferredSize Then
            e.ThrowException = False
        End If

    End Sub

    Private Sub SetUpChartSharePricesTab()
        'Set up the Share Prices tab under the Charts tab.

        'Set up database selection options:
        cmbSPChartDb.Items.Clear()
        cmbSPChartDb.Items.Add("Share Prices")
        cmbSPChartDb.Items.Add("Financials")
        cmbSPChartDb.Items.Add("Calculations")

        '
        'Dim cboFieldSelections As New DataGridViewComboBoxColumn 'Used for selecting Y Value fields in the Chart Settings tab
        DataGridView1.ColumnCount = 1
        DataGridView1.RowCount = 1
        DataGridView1.Columns(0).HeaderText = "Y Value"
        DataGridView1.Columns(0).Width = 120
        DataGridView1.Columns.Insert(1, cboFieldSelections)
        DataGridView1.Columns(1).HeaderText = "Field"
        DataGridView1.Columns(1).Width = 360
        DataGridView1.AllowUserToResizeColumns = True

        'Set up Y Values selection options:
        DataGridView1.Rows.Clear()
        DataGridView1.Rows.Add(4)
        DataGridView1.Rows(0).Cells(0).Value = "High" 'First Y Value Parameter Name
        DataGridView1.Rows(1).Cells(0).Value = "Low" 'Second Y Value parameter name
        DataGridView1.Rows(2).Cells(0).Value = "Open" 'Third Y Value parameter name
        DataGridView1.Rows(3).Cells(0).Value = "Close" 'Fourth Y value parameter name

        'Get list of columns that the query would retreive.
        'Add the list to the selection options in Column 1.
        'https://stackoverflow.com/questions/7159524/get-column-names-from-a-query-without-data

        'Set up the Chart Title alignment options:
        cmbAlignment.Items.Clear()
        'Show the list of ContentAlignment enumerations in the cmbAlignment combobox:
        For Each item In System.Enum.GetValues(GetType(ContentAlignment))
            cmbAlignment.Items.Add(item)
        Next
        cmbAlignment.SelectedIndex = 1 'Top Center


    End Sub

    Private Sub SetUpChartCrossPlotsTab()
        'Set up the Cross Plots tab under the Charts tab.

        'Set up database selection options:
        cmbPointChartDb.Items.Clear()
        cmbPointChartDb.Items.Add("Share Prices")
        cmbPointChartDb.Items.Add("Financials")
        cmbPointChartDb.Items.Add("Calculations")

        'Set up the Chart Title alignment options:
        cmbPointChartAlignment.Items.Clear()
        'Show the list of ContentAlignment enumerations in the cmbAlignment combobox:
        For Each item In System.Enum.GetValues(GetType(ContentAlignment))
            cmbPointChartAlignment.Items.Add(item)
        Next
        cmbPointChartAlignment.SelectedIndex = 1 'Top Center

    End Sub

    Private Sub UpdateChartSharePricesTab()
        'Update the field selection options on the Chart Share Prices tab.

        If txtSPChartDbPath.Text = "" Then
            Message.AddWarning("Charts: Share Prices: No database has been selected." & vbCrLf)
            Exit Sub
        End If

        If txtSPChartQuery.Text = "" Then
            Message.AddWarning("No query has been specified." & vbCrLf)
        End If

        Dim connString As String
        connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source = " & txtSPChartDbPath.Text

        Dim myConnection As OleDb.OleDbConnection = New OleDb.OleDbConnection
        myConnection.ConnectionString = connString
        myConnection.Open()

        Dim Query As String = txtSPChartQuery.Text & " AND 1 = 2" 'This is used to get all the fields in the query. " AND 1 = 2" ensures no data rows are retrieved.

        Dim da As OleDb.OleDbDataAdapter
        da = New OleDb.OleDbDataAdapter(Query, myConnection)
        da.MissingSchemaAction = MissingSchemaAction.AddWithKey

        Dim ds As DataSet = New DataSet

        Try
            cboFieldSelections.Items.Clear()
            cmbXValues.Items.Clear()
            da.Fill(ds, "myData")

            If ds.Tables(0).Columns.Count > 0 Then
                Dim I As Integer 'Loop index
                Dim Name As String
                For I = 1 To ds.Tables(0).Columns.Count
                    Name = ds.Tables(0).Columns(I - 1).ColumnName
                    cmbXValues.Items.Add(ds.Tables(0).Columns(I - 1).ColumnName) 'ComboBox used to selected the XValues column.
                    cboFieldSelections.Items.Add(Name)
                    'Make default selections:
                    If LCase(Name).Contains("date") Then cmbXValues.Text = Name
                    If LCase(Name).Contains("high") Then DataGridView1.Rows(0).Cells(1).Value = Name
                    If LCase(Name).Contains("low") Then DataGridView1.Rows(1).Cells(1).Value = Name
                    If LCase(Name).Contains("open") Then DataGridView1.Rows(2).Cells(1).Value = Name
                    If LCase(Name).Contains("close") Then DataGridView1.Rows(3).Cells(1).Value = Name
                Next
            End If
        Catch ex As Exception
            Message.AddWarning("Error applying query: " & ex.Message & vbCrLf)
        End Try
    End Sub

    Private Sub cmbSPChartDb_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbSPChartDb.SelectedIndexChanged
        'The selected database has changed.

        Select Case cmbSPChartDb.SelectedItem.ToString
            Case "Share Prices"
                txtSPChartDbPath.Text = SharePriceDbPath
            Case "Financials"
                txtSPChartDbPath.Text = FinancialsDbPath
            Case "Calculations"
                txtSPChartDbPath.Text = CalculationsDbPath
        End Select

        FillChartDataTableList()
    End Sub

    Private Sub FillChartDataTableList()
        'Fill the list of tables in cmbChartDataTable

        'Database access for MS Access:
        Dim connectionString As String 'Declare a connection string for MS Access - defines the database or server to be used.
        Dim conn As System.Data.OleDb.OleDbConnection 'Declare a connection for MS Access - used by the Data Adapter to connect to and disconnect from the database.
        Dim dt As DataTable

        cmbChartDataTable.Items.Clear()

        If txtSPChartDbPath.Text = "" Then
            Exit Sub
        End If

        'Access 2007:
        connectionString = "provider=Microsoft.ACE.OLEDB.12.0;" +
        "data source = " + txtSPChartDbPath.Text

        'Connect to the Access database:
        conn = New System.Data.OleDb.OleDbConnection(connectionString)

        Try
            conn.Open()

            Dim restrictions As String() = New String() {Nothing, Nothing, Nothing, "TABLE"} 'This restriction removes system tables
            dt = conn.GetSchema("Tables", restrictions)

            'Fill lstSelectTable
            Dim dr As DataRow
            Dim I As Integer 'Loop index
            Dim MaxI As Integer

            MaxI = dt.Rows.Count
            For I = 0 To MaxI - 1
                dr = dt.Rows(0)
                'lstTables.Items.Add(dt.Rows(I).Item(2).ToString)
                cmbChartDataTable.Items.Add(dt.Rows(I).Item(2).ToString)
            Next I

            conn.Close()
        Catch ex As Exception
            Message.AddWarning("Error opening database: " & txtSPChartDbPath.Text & vbCrLf)
            Message.AddWarning(ex.Message & vbCrLf & vbCrLf)
        End Try
    End Sub

    Private Sub cmbChartDataTable_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbChartDataTable.SelectedIndexChanged

        FillCompanyCodeColumn()
    End Sub

    Private Sub FillCompanyCodeColumn()
        'Fill the list of available table columns in cmbCompanyCodeColumn

        If cmbChartDataTable.SelectedIndex = -1 Then
            'No item is selected
            cmbCompanyCodeCol.Items.Clear()
        Else
            'Database access for MS Access:
            Dim connectionString As String 'Declare a connection string for MS Access - defines the database or server to be used.
            Dim conn As System.Data.OleDb.OleDbConnection 'Declare a connection for MS Access - used by the Data Adapter to connect to and disconnect from the database.
            Dim commandString As String 'Declare a command string - contains the query to be passed to the database.
            Dim ds As DataSet 'Declate a Dataset.
            Dim dt As DataTable

            cmbCompanyCodeCol.Items.Clear()

            'Specify the connection string (Access 2007):
            connectionString = "provider=Microsoft.ACE.OLEDB.12.0;" +
            "data source = " + txtSPChartDbPath.Text

            'Connect to the Access database:
            conn = New System.Data.OleDb.OleDbConnection(connectionString)
            conn.Open()

            'Specify the commandString to query the database:
            commandString = "SELECT Top 500 * FROM " + cmbChartDataTable.SelectedItem.ToString
            Dim dataAdapter As New System.Data.OleDb.OleDbDataAdapter(commandString, conn)
            ds = New DataSet
            dataAdapter.Fill(ds, "SelTable") 'ds was defined earlier as a DataSet
            dt = ds.Tables("SelTable")

            Dim NFields As Integer
            NFields = dt.Columns.Count
            Dim I As Integer
            For I = 0 To NFields - 1
                cmbCompanyCodeCol.Items.Add(dt.Columns(I).ColumnName.ToString)
            Next

            conn.Close()

        End If

    End Sub

    Private Sub cmbCompanyCodeCol_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbCompanyCodeCol.SelectedIndexChanged
        'The company code column selection has changed

        If cmbCompanyCodeCol.SelectedIndex = -1 Then
            'No Company Code column has been selected.
            txtSPChartQuery.Text = ""
        Else
            UpdateSPChartQuery()
        End If

    End Sub

    Private Sub txtSPChartCompanyCode_LostFocus(sender As Object, e As EventArgs) Handles txtSPChartCompanyCode.LostFocus
        UpdateSPChartQuery()
        txtSeriesName.Text = txtSPChartCompanyCode.Text
    End Sub

    Private Sub UpdateSPChartQuery()

        If cmbChartDataTable.SelectedIndex = -1 Then
            Message.AddWarning("Charts: No data table has been selected." & vbCrLf)
            Exit Sub
        End If

        If cmbCompanyCodeCol.SelectedIndex = -1 Then
            Message.AddWarning("No company code column has been selected." & vbCrLf)
            Exit Sub
        End If
        If chkSPChartUseDateRange.Checked Then 'Include a date range in the query:
            txtSPChartQuery.Text = "SELECT * FROM " & cmbChartDataTable.SelectedItem.ToString & " WHERE " & cmbCompanyCodeCol.SelectedItem.ToString & " = '" & txtSPChartCompanyCode.Text & "'" & " AND " & cmbXValues.SelectedItem.ToString & " BETWEEN #" & Format(dtpSPChartFromDate.Value, "MM-dd-yyyy") & "# AND #" & Format(dtpSPChartToDate.Value, "MM-dd-yyyy") & "#"
        Else 'Dont include a date range in the query:
            txtSPChartQuery.Text = "SELECT * FROM " & cmbChartDataTable.SelectedItem.ToString & " WHERE " & cmbCompanyCodeCol.SelectedItem.ToString & " = '" & txtSPChartCompanyCode.Text & "'"
        End If

    End Sub

    Private Sub chkSPChartUseDateRange_CheckedChanged(sender As Object, e As EventArgs) Handles chkSPChartUseDateRange.CheckedChanged
        UpdateSPChartQuery()
    End Sub

    Private Sub txtSPChartQuery_LostFocus(sender As Object, e As EventArgs) Handles txtSPChartQuery.LostFocus
        UpdateChartSharePricesTab()
    End Sub

    Private Sub btnDisplayStockChart_Click(sender As Object, e As EventArgs) Handles btnDisplayStockChart.Click


        DisplayStockChart()
    End Sub

    Public Sub DisplayStockChartUsingDefaults()
        'Display Stock Chart.
        'Use the default parameters in StockChartDefaults
        'Send the instructions to the Chart application to display the stock chart.

        If StockChartDefaults Is Nothing Then
            Message.AddWarning("No Stock Chart default settings loaded." & vbCrLf)
            DisplayStockChartNoDefaults()
            Exit Sub
        End If

        'Create the xml instructions
        Dim decl As New XDeclaration("1.0", "utf-8", "yes")
        Dim doc As New XDocument(decl, Nothing) 'Create an XDocument to store the instructions.
        Dim xmessage As New XElement("XMsg") 'This indicates the start of the message in the XMessage class

        Dim clientAppNetName As New XElement("ClientAppNetName", AppNetName)
        xmessage.Add(clientAppNetName)

        Dim clientName As New XElement("ClientName", ApplicationInfo.Name) 'This tells the coordinate server the name of the client making the request.
        xmessage.Add(clientName)

        Dim clientLocn As New XElement("ClientLocn", "StockChart")
        xmessage.Add(clientLocn)

        Dim chartSettings As New XElement("StockChartSettings")

        Dim chartType As New XElement("ChartType", "Stock")
        chartSettings.Add(chartType)

        Dim commandClearChart As New XElement("Command", "ClearChart")
        chartSettings.Add(commandClearChart)

        Dim inputData As New XElement("InputData")
        Dim dataType As New XElement("Type", "Database")
        inputData.Add(dataType)

        Dim databasePath As New XElement("DatabasePath", txtSPChartDbPath.Text)
        inputData.Add(databasePath)

        Dim dataDescription As New XElement("DataDescription", txtSeriesName.Text)
        inputData.Add(dataDescription)

        Dim databaseQuery As New XElement("DatabaseQuery", txtSPChartQuery.Text)
        inputData.Add(databaseQuery)

        chartSettings.Add(inputData)

        Dim chartProperties As New XElement("ChartProperties")
        Dim seriesName As New XElement("SeriesName", txtSeriesName.Text)
        chartProperties.Add(seriesName)
        Dim xValuesFieldName As New XElement("XValuesFieldName", cmbXValues.SelectedItem.ToString)
        chartProperties.Add(xValuesFieldName)
        Dim yValuesHighFieldName As New XElement("YValuesHighFieldName", DataGridView1.Rows(0).Cells(1).Value)
        chartProperties.Add(yValuesHighFieldName)
        Dim yValuesLowFieldName As New XElement("YValuesLowFieldName", DataGridView1.Rows(1).Cells(1).Value)
        chartProperties.Add(yValuesLowFieldName)
        Dim yValuesOpenFieldName As New XElement("YValuesOpenFieldName", DataGridView1.Rows(2).Cells(1).Value)
        chartProperties.Add(yValuesOpenFieldName)
        Dim yValuesCloseFieldName As New XElement("YValuesCloseFieldName", DataGridView1.Rows(3).Cells(1).Value)
        chartProperties.Add(yValuesCloseFieldName)
        chartSettings.Add(chartProperties)

        Dim chartTitle As New XElement("ChartTitle")
        Dim chartTitleLabelName As New XElement("LabelName", "Label1")
        chartTitle.Add(chartTitleLabelName)
        Dim chartTitleText As New XElement("Text", txtChartTitle.Text)
        chartTitle.Add(chartTitleText)

        If StockChartDefaults.<StockChart>.<Settings>.<ChartTitle>.<FontName>.Value <> Nothing Then
            Dim chartTitleFontName As New XElement("FontName", StockChartDefaults.<StockChart>.<Settings>.<ChartTitle>.<FontName>.Value)
            chartTitle.Add(chartTitleFontName)
        Else
            Message.AddWarning("Default Chart Title Font Name settings not found." & vbCrLf)
            Dim chartTitleFontName As New XElement("FontName", txtChartTitle.Font.Name)
            chartTitle.Add(chartTitleFontName)
        End If

        If StockChartDefaults.<StockChart>.<Settings>.<ChartTitle>.<Color>.Value <> Nothing Then
            Dim chartTitleColor As New XElement("Color", StockChartDefaults.<StockChart>.<Settings>.<ChartTitle>.<Color>.Value)
            chartTitle.Add(chartTitleColor)
        Else
            Message.AddWarning("Default Chart Title Color settings not found." & vbCrLf)
            Dim chartTitleColor As New XElement("Color", txtChartTitle.ForeColor)
            chartTitle.Add(chartTitleColor)
        End If

        If StockChartDefaults.<StockChart>.<Settings>.<ChartTitle>.<Size>.Value <> Nothing Then
            Dim chartTitleSize As New XElement("Size", StockChartDefaults.<StockChart>.<Settings>.<ChartTitle>.<Size>.Value)
            chartTitle.Add(chartTitleSize)
        Else
            Message.AddWarning("Default Chart Title Size settings not found." & vbCrLf)
            Dim chartTitleSize As New XElement("Size", txtChartTitle.Font.Size)
            chartTitle.Add(chartTitleSize)
        End If

        If StockChartDefaults.<StockChart>.<Settings>.<ChartTitle>.<Bold>.Value <> Nothing Then
            Dim chartTitleBold As New XElement("Bold", StockChartDefaults.<StockChart>.<Settings>.<ChartTitle>.<Bold>.Value)
            chartTitle.Add(chartTitleBold)
        Else
            Message.AddWarning("Default Chart Title Bold settings not found." & vbCrLf)
            Dim chartTitleBold As New XElement("Bold", txtChartTitle.Font.Bold)
            chartTitle.Add(chartTitleBold)
        End If

        If StockChartDefaults.<StockChart>.<Settings>.<ChartTitle>.<Italic>.Value <> Nothing Then
            Dim chartTitleItalic As New XElement("Italic", StockChartDefaults.<StockChart>.<Settings>.<ChartTitle>.<Italic>.Value)
            chartTitle.Add(chartTitleItalic)
        Else
            Message.AddWarning("Default Chart Title Italic settings not found." & vbCrLf)
            Dim chartTitleItalic As New XElement("Italic", txtChartTitle.Font.Italic)
            chartTitle.Add(chartTitleItalic)
        End If

        If StockChartDefaults.<StockChart>.<Settings>.<ChartTitle>.<Underline>.Value <> Nothing Then
            Dim chartTitleUnderline As New XElement("Underline", StockChartDefaults.<StockChart>.<Settings>.<ChartTitle>.<Underline>.Value)
            chartTitle.Add(chartTitleUnderline)
        Else
            Message.AddWarning("Default Chart Title Underline settings not found." & vbCrLf)
            Dim chartTitleUnderline As New XElement("Underline", txtChartTitle.Font.Underline)
            chartTitle.Add(chartTitleUnderline)
        End If

        If StockChartDefaults.<StockChart>.<Settings>.<ChartTitle>.<Strikeout>.Value <> Nothing Then
            Dim chartTitleStrikeout As New XElement("Strikeout", StockChartDefaults.<StockChart>.<Settings>.<ChartTitle>.<Strikeout>.Value)
            chartTitle.Add(chartTitleStrikeout)
        Else
            Message.AddWarning("Default Chart Title Strikeout settings not found." & vbCrLf)
            Dim chartTitleStrikeout As New XElement("Strikeout", txtChartTitle.Font.Strikeout)
            chartTitle.Add(chartTitleStrikeout)
        End If

        If StockChartDefaults.<StockChart>.<Settings>.<ChartTitle>.<Alignment>.Value <> Nothing Then
            Dim chartTitleAlignment As New XElement("Alignment", StockChartDefaults.<StockChart>.<Settings>.<ChartTitle>.<Alignment>.Value)
            chartTitle.Add(chartTitleAlignment)
        Else
            Message.AddWarning("Default Chart Title Alignment settings not found." & vbCrLf)
            Dim chartTitleAlignment As New XElement("Alignment", cmbAlignment.SelectedItem.ToString)
            chartTitle.Add(chartTitleAlignment)
        End If

        chartSettings.Add(chartTitle)

        Dim xAxis As New XElement("XAxis")

        If StockChartDefaults.<StockChart>.<Settings>.<XAxis>.<TitleText>.Value <> Nothing Then
            Dim xAxisTitleText As New XElement("TitleText", StockChartDefaults.<StockChart>.<Settings>.<XAxis>.<TitleText>.Value)
            xAxis.Add(xAxisTitleText)
        End If

        If StockChartDefaults.<StockChart>.<Settings>.<XAxis>.<TitleFontName>.Value <> Nothing Then
            Dim xAxisTitleFontName As New XElement("TitleFontName", StockChartDefaults.<StockChart>.<Settings>.<XAxis>.<TitleFontName>.Value)
            xAxis.Add(xAxisTitleFontName)
        End If

        If StockChartDefaults.<StockChart>.<Settings>.<XAxis>.<TitleColor>.Value <> Nothing Then
            Dim xAxisTitleColor As New XElement("TitleColor", StockChartDefaults.<StockChart>.<Settings>.<XAxis>.<TitleColor>.Value)
            xAxis.Add(xAxisTitleColor)
        End If

        If StockChartDefaults.<StockChart>.<Settings>.<XAxis>.<TitleSize>.Value <> Nothing Then
            Dim xAxisTitleSize As New XElement("TitleSize", StockChartDefaults.<StockChart>.<Settings>.<XAxis>.<TitleSize>.Value)
            xAxis.Add(xAxisTitleSize)
        End If

        If StockChartDefaults.<StockChart>.<Settings>.<XAxis>.<TitleBold>.Value <> Nothing Then
            Dim xAxisTitleBold As New XElement("TitleBold", StockChartDefaults.<StockChart>.<Settings>.<XAxis>.<TitleBold>.Value)
            xAxis.Add(xAxisTitleBold)
        End If

        If StockChartDefaults.<StockChart>.<Settings>.<XAxis>.<TitleItalic>.Value <> Nothing Then
            Dim xAxisTitleItalic As New XElement("TitleItalic", StockChartDefaults.<StockChart>.<Settings>.<XAxis>.<TitleItalic>.Value)
            xAxis.Add(xAxisTitleItalic)
        End If

        If StockChartDefaults.<StockChart>.<Settings>.<XAxis>.<TitleUnderline>.Value <> Nothing Then
            Dim xAxisTitleUnderline As New XElement("TitleUnderline", StockChartDefaults.<StockChart>.<Settings>.<XAxis>.<TitleUnderline>.Value)
            xAxis.Add(xAxisTitleUnderline)
        End If

        If StockChartDefaults.<StockChart>.<Settings>.<XAxis>.<TitleStrikeout>.Value <> Nothing Then
            Dim xAxisTitleStrikeout As New XElement("TitleStrikeout", StockChartDefaults.<StockChart>.<Settings>.<XAxis>.<TitleStrikeout>.Value)
            xAxis.Add(xAxisTitleStrikeout)
        End If

        If StockChartDefaults.<StockChart>.<Settings>.<XAxis>.<TitleAlignment>.Value <> Nothing Then
            Dim xAxisTitleAlignment As New XElement("TitleAlignment", StockChartDefaults.<StockChart>.<Settings>.<XAxis>.<TitleAlignment>.Value)
            xAxis.Add(xAxisTitleAlignment)
        End If

        If StockChartDefaults.<StockChart>.<Settings>.<XAxis>.<AutoMinimum>.Value <> Nothing Then
            Dim xAxisAutoMinimum As New XElement("AutoMinimum", StockChartDefaults.<StockChart>.<Settings>.<XAxis>.<AutoMinimum>.Value)
            xAxis.Add(xAxisAutoMinimum)
        End If

        If StockChartDefaults.<StockChart>.<Settings>.<XAxis>.<Minimum>.Value <> Nothing Then
            Dim xAxisMinimum As New XElement("Minimum", StockChartDefaults.<StockChart>.<Settings>.<XAxis>.<Minimum>.Value)
            xAxis.Add(xAxisMinimum)
        End If

        If StockChartDefaults.<StockChart>.<Settings>.<XAxis>.<AutoMaximum>.Value <> Nothing Then
            Dim xAxisAutoMaximum As New XElement("AutoMaximum", StockChartDefaults.<StockChart>.<Settings>.<XAxis>.<AutoMaximum>.Value)
            xAxis.Add(xAxisAutoMaximum)
        End If

        If StockChartDefaults.<StockChart>.<Settings>.<XAxis>.<Maximum>.Value <> Nothing Then
            Dim xAxisMaximum As New XElement("Maximum", StockChartDefaults.<StockChart>.<Settings>.<XAxis>.<Maximum>.Value)
            xAxis.Add(xAxisMaximum)
        End If

        If StockChartDefaults.<StockChart>.<Settings>.<XAxis>.<AutoInterval>.Value <> Nothing Then
            Dim xAxisAutoInterval As New XElement("AutoInterval", StockChartDefaults.<StockChart>.<Settings>.<XAxis>.<AutoInterval>.Value)
            xAxis.Add(xAxisAutoInterval)
        End If

        If StockChartDefaults.<StockChart>.<Settings>.<XAxis>.<MajorGridInterval>.Value <> Nothing Then
            Dim xAxisMajorGridInterval As New XElement("MajorGridInterval", StockChartDefaults.<StockChart>.<Settings>.<XAxis>.<MajorGridInterval>.Value)
            xAxis.Add(xAxisMajorGridInterval)
        End If

        If StockChartDefaults.<StockChart>.<Settings>.<XAxis>.<AutoMajorGridInterval>.Value <> Nothing Then
            Dim xAxisAutoMajorGridInterval As New XElement("AutoMajorGridInterval", StockChartDefaults.<StockChart>.<Settings>.<XAxis>.<AutoMajorGridInterval>.Value)
            xAxis.Add(xAxisAutoMajorGridInterval)
        End If

        chartSettings.Add(xAxis)

        Dim yAxis As New XElement("YAxis")

        If StockChartDefaults.<StockChart>.<Settings>.<YAxis>.<TitleText>.Value <> Nothing Then
            Dim yAxisTitleText As New XElement("TitleText", StockChartDefaults.<StockChart>.<Settings>.<YAxis>.<TitleText>.Value)
            yAxis.Add(yAxisTitleText)
        End If

        If StockChartDefaults.<StockChart>.<Settings>.<YAxis>.<TitleFontName>.Value <> Nothing Then
            Dim yAxisTitleFontName As New XElement("TitleFontName", StockChartDefaults.<StockChart>.<Settings>.<YAxis>.<TitleFontName>.Value)
            yAxis.Add(yAxisTitleFontName)
        End If

        If StockChartDefaults.<StockChart>.<Settings>.<YAxis>.<TitleColor>.Value <> Nothing Then
            Dim yAxisTitleColor As New XElement("TitleColor", StockChartDefaults.<StockChart>.<Settings>.<YAxis>.<TitleColor>.Value)
            yAxis.Add(yAxisTitleColor)
        End If

        If StockChartDefaults.<StockChart>.<Settings>.<YAxis>.<TitleSize>.Value <> Nothing Then
            Dim yAxisTitleSize As New XElement("TitleSize", StockChartDefaults.<StockChart>.<Settings>.<YAxis>.<TitleSize>.Value)
            yAxis.Add(yAxisTitleSize)
        End If

        If StockChartDefaults.<StockChart>.<Settings>.<YAxis>.<TitleBold>.Value <> Nothing Then
            Dim yAxisTitleBold As New XElement("TitleBold", StockChartDefaults.<StockChart>.<Settings>.<YAxis>.<TitleBold>.Value)
            yAxis.Add(yAxisTitleBold)
        End If

        If StockChartDefaults.<StockChart>.<Settings>.<YAxis>.<TitleItalic>.Value <> Nothing Then
            Dim yAxisTitleItalic As New XElement("TitleItalic", StockChartDefaults.<StockChart>.<Settings>.<YAxis>.<TitleItalic>.Value)
            yAxis.Add(yAxisTitleItalic)
        End If

        If StockChartDefaults.<StockChart>.<Settings>.<YAxis>.<TitleUnderline>.Value <> Nothing Then
            Dim yAxisTitleUnderline As New XElement("TitleUnderline", StockChartDefaults.<StockChart>.<Settings>.<YAxis>.<TitleUnderline>.Value)
            yAxis.Add(yAxisTitleUnderline)
        End If

        If StockChartDefaults.<StockChart>.<Settings>.<YAxis>.<TitleStrikeout>.Value <> Nothing Then
            Dim yAxisTitleStrikeout As New XElement("TitleStrikeout", StockChartDefaults.<StockChart>.<Settings>.<YAxis>.<TitleStrikeout>.Value)
            yAxis.Add(yAxisTitleStrikeout)
        End If

        If StockChartDefaults.<StockChart>.<Settings>.<YAxis>.<TitleAlignment>.Value <> Nothing Then
            Dim yAxisTitleAlignment As New XElement("TitleAlignment", StockChartDefaults.<StockChart>.<Settings>.<YAxis>.<TitleAlignment>.Value)
            yAxis.Add(yAxisTitleAlignment)
        End If

        If StockChartDefaults.<StockChart>.<Settings>.<YAxis>.<AutoMinimum>.Value <> Nothing Then
            Dim yAxisAutoMinimum As New XElement("AutoMinimum", StockChartDefaults.<StockChart>.<Settings>.<YAxis>.<AutoMinimum>.Value)
            yAxis.Add(yAxisAutoMinimum)
        End If

        If StockChartDefaults.<StockChart>.<Settings>.<YAxis>.<Minimum>.Value <> Nothing Then
            Dim yAxisMinimum As New XElement("Minimum", StockChartDefaults.<StockChart>.<Settings>.<YAxis>.<Minimum>.Value)
            yAxis.Add(yAxisMinimum)
        End If

        If StockChartDefaults.<StockChart>.<Settings>.<YAxis>.<AutoMaximum>.Value <> Nothing Then
            Dim yAxisAutoMaximum As New XElement("AutoMaximum", StockChartDefaults.<StockChart>.<Settings>.<YAxis>.<AutoMaximum>.Value)
            yAxis.Add(yAxisAutoMaximum)
        End If

        If StockChartDefaults.<StockChart>.<Settings>.<YAxis>.<Maximum>.Value <> Nothing Then
            Dim yAxisMaximum As New XElement("Maximum", StockChartDefaults.<StockChart>.<Settings>.<YAxis>.<Maximum>.Value)
            yAxis.Add(yAxisMaximum)
        End If

        If StockChartDefaults.<StockChart>.<Settings>.<YAxis>.<AutoInterval>.Value <> Nothing Then
            Dim yAxisAutoInterval As New XElement("AutoInterval", StockChartDefaults.<StockChart>.<Settings>.<YAxis>.<AutoInterval>.Value)
            yAxis.Add(yAxisAutoInterval)
        End If

        If StockChartDefaults.<StockChart>.<Settings>.<YAxis>.<MajorGridInterval>.Value <> Nothing Then
            Dim yAxisMajorGridInterval As New XElement("MajorGridInterval", StockChartDefaults.<StockChart>.<Settings>.<YAxis>.<MajorGridInterval>.Value)
            yAxis.Add(yAxisMajorGridInterval)
        End If

        chartSettings.Add(yAxis)


        Dim commandDrawChart As New XElement("Command", "DrawChart")
        chartSettings.Add(commandDrawChart)

        xmessage.Add(chartSettings)
        doc.Add(xmessage)

        If IsNothing(client) Then
            Message.AddWarning("No client connection available!" & vbCrLf)
            Beep()
        Else

            client.SendMessageAsync(AppNetName, "ADVL_Stock_Chart_1", doc.ToString) 'Added 3Feb19
            Message.XAddText("Message sent to " & "ADVL_Stock_Chart_1" & ":" & vbCrLf, "XmlSentNotice")
            Message.XAddXml(doc.ToString)
            Message.XAddText(vbCrLf, "Message") 'Add extra line
        End If

    End Sub

    Public Sub DisplayStockChart()
        'Display Stock Chart.

        'Check if connected to ComNet:
        If ConnectedToComNet = False Then
            ConnectToComNet()
        End If

        If chkUseStockChartDefaults.Checked Then
            DisplayStockChartUsingDefaults()
        Else
            DisplayStockChartNoDefaults()
        End If
    End Sub

    Private Sub DisplayStockChartNoDefaults()
        'Display the stock chart without using the default chart settings.

        'Send the instructions to the Chart application to display the stock chart.

        'Check that required selections have been made:
        If cmbXValues.SelectedItem Is Nothing Then
            Message.AddWarning("Select a field for the X Values." & vbCrLf)
            Exit Sub
        End If


        'Create the xml instructions
        Dim decl As New XDeclaration("1.0", "utf-8", "yes")
        Dim doc As New XDocument(decl, Nothing) 'Create an XDocument to store the instructions.
        Dim xmessage As New XElement("XMsg") 'This indicates the start of the message in the XMessage class

        Dim clientAppNetName As New XElement("ClientAppNetName", AppNetName)
        xmessage.Add(clientAppNetName)

        Dim clientName As New XElement("ClientName", ApplicationInfo.Name) 'This tells the coordinate server the name of the client making the request.
        xmessage.Add(clientName)

        Dim clientLocn As New XElement("ClientLocn", "StockChart")
        xmessage.Add(clientLocn)

        Dim chartSettings As New XElement("StockChartSettings")

        Dim chartType As New XElement("ChartType", "Stock")
        chartSettings.Add(chartType)

        Dim commandClearChart As New XElement("Command", "ClearChart")
        chartSettings.Add(commandClearChart)

        Dim inputData As New XElement("InputData")
        Dim dataType As New XElement("Type", "Database")
        inputData.Add(dataType)
        Dim databasePath As New XElement("DatabasePath", txtSPChartDbPath.Text)
        inputData.Add(databasePath)
        Dim dataDescription As New XElement("DataDescription", txtSeriesName.Text)
        inputData.Add(dataDescription)
        Dim databaseQuery As New XElement("DatabaseQuery", txtSPChartQuery.Text)
        inputData.Add(databaseQuery)
        chartSettings.Add(inputData)

        Dim chartProperties As New XElement("ChartProperties")
        Dim seriesName As New XElement("SeriesName", txtSeriesName.Text)
        chartProperties.Add(seriesName)
        Dim xValuesFieldName As New XElement("XValuesFieldName", cmbXValues.SelectedItem.ToString)
        chartProperties.Add(xValuesFieldName)
        Dim yValuesHighFieldName As New XElement("YValuesHighFieldName", DataGridView1.Rows(0).Cells(1).Value)
        chartProperties.Add(yValuesHighFieldName)
        Dim yValuesLowFieldName As New XElement("YValuesLowFieldName", DataGridView1.Rows(1).Cells(1).Value)
        chartProperties.Add(yValuesLowFieldName)
        Dim yValuesOpenFieldName As New XElement("YValuesOpenFieldName", DataGridView1.Rows(2).Cells(1).Value)
        chartProperties.Add(yValuesOpenFieldName)
        Dim yValuesCloseFieldName As New XElement("YValuesCloseFieldName", DataGridView1.Rows(3).Cells(1).Value)
        chartProperties.Add(yValuesCloseFieldName)
        chartSettings.Add(chartProperties)

        Dim chartTitle As New XElement("ChartTitle")
        Dim chartTitleLabelName As New XElement("LabelName", "Label1")
        chartTitle.Add(chartTitleLabelName)
        Dim chartTitleText As New XElement("Text", txtChartTitle.Text)
        chartTitle.Add(chartTitleText)
        Dim chartTitleFontName As New XElement("FontName", txtChartTitle.Font.Name)
        chartTitle.Add(chartTitleFontName)
        Dim chartTitleColor As New XElement("Color", txtChartTitle.ForeColor)
        chartTitle.Add(chartTitleColor)
        Dim chartTitleSize As New XElement("Size", txtChartTitle.Font.Size)
        chartTitle.Add(chartTitleSize)
        Dim chartTitleBold As New XElement("Bold", txtChartTitle.Font.Bold)
        chartTitle.Add(chartTitleBold)
        Dim chartTitleItalic As New XElement("Italic", txtChartTitle.Font.Italic)
        chartTitle.Add(chartTitleItalic)
        Dim chartTitleUnderline As New XElement("Underline", txtChartTitle.Font.Underline)
        chartTitle.Add(chartTitleUnderline)
        Dim chartTitleStrikeout As New XElement("Strikeout", txtChartTitle.Font.Strikeout)
        chartTitle.Add(chartTitleStrikeout)
        Dim chartTitleAlignment As New XElement("Alignment", cmbAlignment.SelectedItem.ToString)
        chartTitle.Add(chartTitleAlignment)
        chartSettings.Add(chartTitle)

        Dim commandDrawChart As New XElement("Command", "DrawChart")
        chartSettings.Add(commandDrawChart)

        xmessage.Add(chartSettings)
        doc.Add(xmessage)

        If IsNothing(client) Then
            Message.AddWarning("No client connection available!" & vbCrLf)
            Beep()
        ElseIf client.State = ServiceModel.CommunicationState.Faulted Then
            Message.AddWarning("Client state is faulted. Message not sent!" & vbCrLf)
        Else
            client.SendMessageAsync(AppNetName, "ADVL_Stock_Chart_1", doc.ToString) 'Added 3Feb19
            Message.XAddText("Message sent to " & "ADVL_Stock_Chart_1" & ":" & vbCrLf, "XmlSentNotice")
            Message.XAddXml(doc.ToString)
            Message.XAddText(vbCrLf, "Message") 'Add extra line
        End If
    End Sub

    Private Sub btnGetStockChartDefaults_Click(sender As Object, e As EventArgs) Handles btnGetStockChartDefaults.Click
        'Send a request to ADVL_Charts_1 for the current Stock Chart settings.

        'Create the xml instructions
        Dim decl As New XDeclaration("1.0", "utf-8", "yes")
        Dim doc As New XDocument(decl, Nothing) 'Create an XDocument to store the instructions.
        Dim xmessage As New XElement("XMsg") 'This indicates the start of the message in the XMessage class

        'ADDED 3Feb19:
        Dim clientAppNetName As New XElement("ClientAppNetName", AppNetName)
        xmessage.Add(clientAppNetName)

        Dim clientName As New XElement("ClientName", ApplicationInfo.Name) 'This tells the coordinate server the name of the client making the request.
        xmessage.Add(clientName)

        Dim clientLocn As New XElement("ClientLocn", "StockChart")
        xmessage.Add(clientLocn)

        Dim commandGetSettings As New XElement("Command", "GetStockChartSettings")
        xmessage.Add(commandGetSettings)

        doc.Add(xmessage)

        If IsNothing(client) Then
            Message.AddWarning("No client connection available!" & vbCrLf)
            Beep()
        Else
            client.SendMessageAsync(AppNetName, "ADVL_Stock_Chart_1", doc.ToString) 'Added 3Feb19
            Message.XAddText("Message sent to " & "ADVL_Stock_Chart_1" & ":" & vbCrLf, "XmlSentNotice")
            Message.XAddXml(doc.ToString)
            Message.XAddText(vbCrLf, "Message") 'Add extra line

        End If
    End Sub

    Private Sub UpdateChartCrossPlotsTab()
        'Update the field selection options on the Chart Cross Plots tab.

        If txtPointChartDbPath.Text = "" Then
            Message.AddWarning("Charts: Cross Plots: No database has been selected." & vbCrLf)
            Exit Sub
        End If

        If txtPointChartQuery.Text = "" Then
            Message.AddWarning("No query has been specified." & vbCrLf)
        End If

        Dim connString As String
        connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source = " & txtPointChartDbPath.Text

        Dim myConnection As OleDb.OleDbConnection = New OleDb.OleDbConnection
        myConnection.ConnectionString = connString
        myConnection.Open()

        Dim Query As String = txtPointChartQuery.Text & " AND 1 = 2" 'This is used to get all the fields in the query. " AND 1 = 2" ensures no data rows are retrieved.

        Dim da As OleDb.OleDbDataAdapter
        da = New OleDb.OleDbDataAdapter(Query, myConnection)
        da.MissingSchemaAction = MissingSchemaAction.AddWithKey

        Dim ds As DataSet = New DataSet

        Try
            cmbPointXValues.Items.Clear()
            cmbPointYValues.Items.Clear()
            da.Fill(ds, "myData")

            If ds.Tables(0).Columns.Count > 0 Then
                Dim I As Integer 'Loop index
                Dim Name As String
                For I = 1 To ds.Tables(0).Columns.Count
                    Name = ds.Tables(0).Columns(I - 1).ColumnName
                    cmbPointXValues.Items.Add(Name)
                    cmbPointYValues.Items.Add(Name)
                Next
            End If
        Catch ex As Exception
            Message.AddWarning("Error applying query: " & ex.Message & vbCrLf)
        End Try

    End Sub

    Private Sub txtPointChartQuery_LostFocus(sender As Object, e As EventArgs) Handles txtPointChartQuery.LostFocus
        UpdateChartCrossPlotsTab()
    End Sub

    Private Sub cmbPointChartDb_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbPointChartDb.SelectedIndexChanged
        'The selected Point Chart database has been changed.
        Select Case cmbPointChartDb.SelectedItem.ToString
            Case "Share Prices"
                txtPointChartDbPath.Text = SharePriceDbPath
            Case "Financials"
                txtPointChartDbPath.Text = FinancialsDbPath
            Case "Calculations"
                txtPointChartDbPath.Text = CalculationsDbPath
        End Select
        FillPointChartDataTableList()

    End Sub

    Private Sub FillPointChartDataTableList()
        'Fill the lstSelectTable listbox with the available tables in the selected database.

        'Database access for MS Access:
        Dim connectionString As String 'Declare a connection string for MS Access - defines the database or server to be used.
        Dim conn As System.Data.OleDb.OleDbConnection 'Declare a connection for MS Access - used by the Data Adapter to connect to and disconnect from the database.
        Dim dt As DataTable

        lstTables.Items.Clear()
        lstFields.Items.Clear()

        'Specify the connection string:
        'Access 2003
        'connectionString = "provider=Microsoft.Jet.OLEDB.4.0;" + _
        '"data source = " + txtDatabase.Text

        If txtPointChartDbPath.Text = "" Then
            Exit Sub
        End If

        'Access 2007:
        'connectionString = "provider=Microsoft.ACE.OLEDB.12.0;" +
        '"data source = " + txtDatabase.Text
        connectionString = "provider=Microsoft.ACE.OLEDB.12.0;" +
        "data source = " + txtPointChartDbPath.Text 'DatabasePath

        'Connect to the Access database:
        conn = New System.Data.OleDb.OleDbConnection(connectionString)

        Try
            conn.Open()

            Dim restrictions As String() = New String() {Nothing, Nothing, Nothing, "TABLE"} 'This restriction removes system tables
            dt = conn.GetSchema("Tables", restrictions)

            'Fill lstSelectTable
            Dim dr As DataRow
            Dim I As Integer 'Loop index
            Dim MaxI As Integer

            MaxI = dt.Rows.Count
            For I = 0 To MaxI - 1
                dr = dt.Rows(0)
                lstTables.Items.Add(dt.Rows(I).Item(2).ToString)
            Next I

            conn.Close()
        Catch ex As Exception
            Message.Add("Error opening database: " & txtPointChartDbPath.Text & vbCrLf)
            Message.Add(ex.Message & vbCrLf & vbCrLf)
        End Try
    End Sub

    Private Sub lstTables_Click(sender As Object, e As EventArgs) Handles lstTables.Click
        FillPointChartTableLstFields()
    End Sub

    Private Sub FillPointChartTableLstFields()
        'Fill the lstSelectField listbox with the availalble fields in the selected table.

        'Database access for MS Access:
        Dim connectionString As String 'Declare a connection string for MS Access - defines the database or server to be used.
        Dim conn As System.Data.OleDb.OleDbConnection 'Declare a connection for MS Access - used by the Data Adapter to connect to and disconnect from the database.
        Dim commandString As String 'Declare a command string - contains the query to be passed to the database.
        Dim ds As DataSet 'Declate a Dataset.
        Dim dt As DataTable

        If lstTables.SelectedIndex = -1 Then 'No item is selected
            lstFields.Items.Clear()

        Else 'A table has been selected. List its fields:
            lstFields.Items.Clear()

            'Specify the connection string (Access 2007):
            connectionString = "provider=Microsoft.ACE.OLEDB.12.0;" +
            "data source = " + txtPointChartDbPath.Text 'txtDatabase.Text

            'Connect to the Access database:
            conn = New System.Data.OleDb.OleDbConnection(connectionString)
            conn.Open()

            'Specify the commandString to query the database:
            commandString = "SELECT Top 500 * FROM " + lstTables.SelectedItem.ToString
            Dim dataAdapter As New System.Data.OleDb.OleDbDataAdapter(commandString, conn)
            ds = New DataSet
            dataAdapter.Fill(ds, "SelTable") 'ds was defined earlier as a DataSet
            dt = ds.Tables("SelTable")

            Dim NFields As Integer
            NFields = dt.Columns.Count
            Dim I As Integer
            For I = 0 To NFields - 1
                lstFields.Items.Add(dt.Columns(I).ColumnName.ToString)
            Next

            conn.Close()

        End If
    End Sub

    Public Sub FormatXmlText(ByRef rtbControl As RichTextBox)
        'Format the XML text in rtbSequence rich text box control:

        Dim Posn As Integer
        Dim SelLen As Integer
        Posn = rtbControl.SelectionStart
        SelLen = rtbControl.SelectionLength

        'Set colour of the start tag names (for a tag without attributes):
        Dim RegExString2 As String = "(?<=<)([A-Za-z\d]+)(?=>)"
        Dim myRegEx2 As New System.Text.RegularExpressions.Regex(RegExString2)
        Dim myMatches2 As System.Text.RegularExpressions.MatchCollection
        myMatches2 = myRegEx2.Matches(rtbControl.Text)
        For Each aMatch As System.Text.RegularExpressions.Match In myMatches2
            rtbControl.Select(aMatch.Index, aMatch.Length)
            rtbControl.SelectionColor = Color.Crimson
            Dim f As Font = rtbControl.SelectionFont
            rtbControl.SelectionFont = New Font(f.Name, f.Size, FontStyle.Regular)
        Next

        'Set colour of the start tag names (for a tag with attributes):
        Dim RegExString2b As String = "(?<=<)([A-Za-z\d]+)(?= [A-Za-z\d]+=""[A-Za-z\d ]+"">)"
        Dim myRegEx2b As New System.Text.RegularExpressions.Regex(RegExString2b)
        Dim myMatches2b As System.Text.RegularExpressions.MatchCollection
        myMatches2b = myRegEx2b.Matches(rtbControl.Text)
        For Each aMatch As System.Text.RegularExpressions.Match In myMatches2b
            rtbControl.Select(aMatch.Index, aMatch.Length)
            rtbControl.SelectionColor = Color.Crimson
            Dim f As Font = rtbControl.SelectionFont
            rtbControl.SelectionFont = New Font(f.Name, f.Size, FontStyle.Regular)
        Next

        'Set colour of the attribute names (for a tag with attributes):
        Dim RegExString2c As String = "(?<=<[A-Za-z\d]+ )([A-Za-z\d]+)(?==""[A-Za-z\d ]+"">)"
        Dim myRegEx2c As New System.Text.RegularExpressions.Regex(RegExString2c)
        Dim myMatches2c As System.Text.RegularExpressions.MatchCollection
        myMatches2c = myRegEx2c.Matches(rtbControl.Text)
        For Each aMatch As System.Text.RegularExpressions.Match In myMatches2c
            rtbControl.Select(aMatch.Index, aMatch.Length)
            rtbControl.SelectionColor = Color.Crimson
            Dim f As Font = rtbControl.SelectionFont
            rtbControl.SelectionFont = New Font(f.Name, f.Size, FontStyle.Regular)
        Next

        'Set colour of the attribute values (for a tag with attributes):
        Dim RegExString2d As String = "(?<=<[A-Za-z\d]+ [A-Za-z\d]+="")([A-Za-z\d ]+)(?="">)"
        Dim myRegEx2d As New System.Text.RegularExpressions.Regex(RegExString2d)
        Dim myMatches2d As System.Text.RegularExpressions.MatchCollection
        myMatches2d = myRegEx2d.Matches(rtbControl.Text)
        For Each aMatch As System.Text.RegularExpressions.Match In myMatches2d
            rtbControl.Select(aMatch.Index, aMatch.Length)
            rtbControl.SelectionColor = Color.Black
            Dim f As Font = rtbControl.SelectionFont
            rtbControl.SelectionFont = New Font(f.Name, f.Size, FontStyle.Bold)
        Next

        'Set colour of the end tag names:
        Dim RegExString3 As String = "(?<=</)([A-Za-z\d]+)(?=>)"
        Dim myRegEx3 As New System.Text.RegularExpressions.Regex(RegExString3)
        Dim myMatches3 As System.Text.RegularExpressions.MatchCollection
        myMatches3 = myRegEx3.Matches(rtbControl.Text)
        For Each aMatch As System.Text.RegularExpressions.Match In myMatches3
            rtbControl.Select(aMatch.Index, aMatch.Length)
            rtbControl.SelectionColor = Color.Crimson
            Dim f As Font = rtbControl.SelectionFont
            rtbControl.SelectionFont = New Font(f.Name, f.Size, FontStyle.Regular)
        Next

        'Set colour of comments:
        Dim RegExString4 As String = "(?<=<!--)([A-Za-z\d \.,_:]+)(?=-->)"
        Dim myRegEx4 As New System.Text.RegularExpressions.Regex(RegExString4)
        Dim myMatches4 As System.Text.RegularExpressions.MatchCollection
        myMatches4 = myRegEx4.Matches(rtbControl.Text)
        For Each aMatch As System.Text.RegularExpressions.Match In myMatches4
            rtbControl.Select(aMatch.Index, aMatch.Length)
            rtbControl.SelectionColor = Color.Gray
            Dim f As Font = rtbControl.SelectionFont
            rtbControl.SelectionFont = New Font(f.Name, f.Size, FontStyle.Regular)
        Next

        'Set colour of "<" and ">" characters to blue
        Dim RegExString As String = "</|<!--|-->|<|>"
        Dim myRegEx As New System.Text.RegularExpressions.Regex(RegExString)
        Dim myMatches As System.Text.RegularExpressions.MatchCollection
        myMatches = myRegEx.Matches(rtbControl.Text)
        For Each aMatch As System.Text.RegularExpressions.Match In myMatches
            rtbControl.Select(aMatch.Index, aMatch.Length)
            rtbControl.SelectionColor = Color.Blue
            Dim f As Font = rtbControl.SelectionFont
            rtbControl.SelectionFont = New Font(f.Name, f.Size, FontStyle.Regular)
        Next

        'Set tag contents (between ">" and "</") to black, bold
        Dim RegExString5 As String = "(?<=>)([A-Za-z\d \.,'\:\-\[\]\&\*\;\\=/+#_]+)(?=</)"
        Dim myRegEx5 As New System.Text.RegularExpressions.Regex(RegExString5)
        Dim myMatches5 As System.Text.RegularExpressions.MatchCollection
        myMatches5 = myRegEx5.Matches(rtbControl.Text)
        For Each aMatch As System.Text.RegularExpressions.Match In myMatches5
            rtbControl.Select(aMatch.Index, aMatch.Length)
            rtbControl.SelectionColor = Color.Black
            Dim f As Font = rtbControl.SelectionFont
            rtbControl.SelectionFont = New Font(f.Name, f.Size, FontStyle.Bold)
        Next

        'Remove blank lines
        Dim RegExString6 As String = "(?<=\n)\ *\n"
        Dim myRegEx6 As New System.Text.RegularExpressions.Regex(RegExString6)
        Dim myMatches6 As System.Text.RegularExpressions.MatchCollection
        myMatches6 = myRegEx6.Matches(rtbControl.Text)
        For Each aMatch As System.Text.RegularExpressions.Match In myMatches6
            rtbControl.Select(aMatch.Index, aMatch.Length)
            rtbControl.SelectedText = ""
        Next

        rtbControl.SelectionStart = Posn
        rtbControl.SelectionLength = SelLen

    End Sub

    Private Sub btnChartTitleFont_Click(sender As Object, e As EventArgs) Handles btnChartTitleFont.Click
        'Edit chart title font
        FontDialog1.Font = txtChartTitle.Font
        FontDialog1.ShowDialog()
        txtChartTitle.Font = FontDialog1.Font
    End Sub

    Private Sub btnSaveStockDefaults_Click(sender As Object, e As EventArgs) Handles btnSaveStockDefaults.Click
        'Save the Stock Chart Default Settings in a file:

        Try
            Dim xmlSeq As System.Xml.Linq.XDocument = XDocument.Parse(rtbStockChartDefaults.Text)

            Dim SettingsFileName As String = ""

            If Trim(txtStockChartSettings.Text).EndsWith(".SPChartDefaults") Then
                SettingsFileName = Trim(txtStockChartSettings.Text)
            Else
                SettingsFileName = Trim(txtStockChartSettings.Text) & ".SPChartDefaults"
                txtStockChartSettings.Text = SettingsFileName
            End If

            Project.SaveXmlData(SettingsFileName, xmlSeq)
            Message.Add("Stock Chart Default Settings saved OK" & vbCrLf)
        Catch ex As Exception
            Message.AddWarning(ex.Message & vbCrLf)
            Beep()
        End Try

    End Sub

    Private Sub btnOpenStockChartDefaults_Click(sender As Object, e As EventArgs) Handles btnOpenStockChartDefaults.Click
        'Open a Stock Chart Default Settings file:

        Dim SelectedFileName As String = ""

        SelectedFileName = Project.SelectDataFile("Share Price Chart Defaults", "SPChartDefaults")
        Message.Add("Selected Stock Chart Default Settings: " & SelectedFileName & vbCrLf)

        txtStockChartSettings.Text = SelectedFileName

        Project.ReadXmlData(SelectedFileName, StockChartDefaults)

        If StockChartDefaults Is Nothing Then
            Exit Sub
        End If

        rtbStockChartDefaults.Text = StockChartDefaults.ToString

        FormatXmlText(rtbStockChartDefaults)

    End Sub

    Private Sub DesignPointChartQuery_Apply(myQuery As String) Handles DesignPointChartQuery.Apply
        txtPointChartQuery.Text = myQuery
    End Sub

    Private Sub btnPointChartTitleFont_Click(sender As Object, e As EventArgs) Handles btnPointChartTitleFont.Click
        'Edit chart title font
        FontDialog1.Font = txtPointChartTitle.Font
        FontDialog1.ShowDialog()
        txtPointChartTitle.Font = FontDialog1.Font

    End Sub

    Private Sub btnDisplayPointChart_Click(sender As Object, e As EventArgs) Handles btnDisplayPointChart.Click
        DisplayPointChart()
    End Sub

    Private Sub DisplayPointChart()
        'Display Cross Plot Chart.
        If chkUsePointChartDefaults.Checked Then
            DisplayPointChartUsingDefaults()
        Else
            DisplayPointChartNoDefaults()
        End If
    End Sub

    Private Sub DisplayPointChartUsingDefaults()
        'Display Point Chart.
        'Use the default parameters in PointChartDefaults
        'Send the instructions to the Chart application to display the point chart (crossplot).

        If PointChartDefaults Is Nothing Then
            Message.AddWarning("No Point Chart default settings loaded." & vbCrLf)
            DisplayPointChartNoDefaults()
            Exit Sub
        End If

        'Create the xml instructions
        Dim decl As New XDeclaration("1.0", "utf-8", "yes")
        Dim doc As New XDocument(decl, Nothing) 'Create an XDocument to store the instructions.
        Dim xmessage As New XElement("XMsg") 'This indicates the start of the message in the XMessage class

        Dim clientAppNetName As New XElement("ClientAppNetName", AppNetName)
        xmessage.Add(clientAppNetName)

        Dim clientName As New XElement("ClientName", ApplicationInfo.Name) 'This tells the coordinate server the name of the client making the request.
        xmessage.Add(clientName)

        Dim clientLocn As New XElement("ClientLocn", "PointChart")
        xmessage.Add(clientLocn)

        Dim chartSettings As New XElement("PointChartSettings")

        Dim chartType As New XElement("ChartType", "Point")
        chartSettings.Add(chartType)

        Dim commandClearChart As New XElement("Command", "ClearChart")
        chartSettings.Add(commandClearChart)

        Dim inputData As New XElement("InputData")
        Dim dataType As New XElement("Type", "Database")
        inputData.Add(dataType)

        Dim databasePath As New XElement("DatabasePath", txtPointChartDbPath.Text)
        inputData.Add(databasePath)

        Dim dataDescription As New XElement("DataDescription", txtPointSeriesName.Text)
        inputData.Add(dataDescription)

        Dim databaseQuery As New XElement("DatabaseQuery", txtPointChartQuery.Text)
        inputData.Add(databaseQuery)

        chartSettings.Add(inputData)

        Dim chartProperties As New XElement("ChartProperties")
        Dim seriesName As New XElement("SeriesName", txtPointSeriesName.Text)
        chartProperties.Add(seriesName)
        Dim xValuesFieldName As New XElement("XValuesFieldName", cmbPointXValues.SelectedItem.ToString)
        chartProperties.Add(xValuesFieldName)
        Dim yValuesFieldName As New XElement("YValuesFieldName", cmbPointYValues.SelectedItem.ToString)
        chartProperties.Add(yValuesFieldName)
        chartSettings.Add(chartProperties)

        Dim chartTitle As New XElement("ChartTitle")
        Dim chartTitleLabelName As New XElement("LabelName", "Label1")
        chartTitle.Add(chartTitleLabelName)
        Dim chartTitleText As New XElement("Text", txtPointChartTitle.Text)
        chartTitle.Add(chartTitleText)

        If PointChartDefaults.<PointChart>.<Settings>.<ChartTitle>.<FontName>.Value <> Nothing Then
            Dim chartTitleFontName As New XElement("FontName", PointChartDefaults.<PointChart>.<Settings>.<ChartTitle>.<FontName>.Value)
            chartTitle.Add(chartTitleFontName)
        Else
            Message.AddWarning("Default Chart Title Font Name settings not found." & vbCrLf)
            Dim chartTitleFontName As New XElement("FontName", txtPointChartTitle.Font.Name)
            chartTitle.Add(chartTitleFontName)
        End If

        If PointChartDefaults.<PointChart>.<Settings>.<ChartTitle>.<Color>.Value <> Nothing Then
            Dim chartTitleColor As New XElement("Color", PointChartDefaults.<PointChart>.<Settings>.<ChartTitle>.<Color>.Value)
            chartTitle.Add(chartTitleColor)
        Else
            Message.AddWarning("Default Chart Title Color settings not found." & vbCrLf)
            Dim chartTitleColor As New XElement("Color", txtChartTitle.ForeColor)
            chartTitle.Add(chartTitleColor)
        End If

        If PointChartDefaults.<PointChart>.<Settings>.<ChartTitle>.<Size>.Value <> Nothing Then
            Dim chartTitleSize As New XElement("Size", PointChartDefaults.<PointChart>.<Settings>.<ChartTitle>.<Size>.Value)
            chartTitle.Add(chartTitleSize)
        Else
            Message.AddWarning("Default Chart Title Size settings not found." & vbCrLf)
            Dim chartTitleSize As New XElement("Size", txtChartTitle.Font.Size)
            chartTitle.Add(chartTitleSize)
        End If

        If PointChartDefaults.<PointChart>.<Settings>.<ChartTitle>.<Bold>.Value <> Nothing Then
            Dim chartTitleBold As New XElement("Bold", PointChartDefaults.<PointChart>.<Settings>.<ChartTitle>.<Bold>.Value)
            chartTitle.Add(chartTitleBold)
        Else
            Message.AddWarning("Default Chart Title Bold settings not found." & vbCrLf)
            Dim chartTitleBold As New XElement("Bold", txtChartTitle.Font.Bold)
            chartTitle.Add(chartTitleBold)
        End If

        If PointChartDefaults.<PointChart>.<Settings>.<ChartTitle>.<Italic>.Value <> Nothing Then
            Dim chartTitleItalic As New XElement("Italic", PointChartDefaults.<PointChart>.<Settings>.<ChartTitle>.<Italic>.Value)
            chartTitle.Add(chartTitleItalic)
        Else
            Message.AddWarning("Default Chart Title Italic settings not found." & vbCrLf)
            Dim chartTitleItalic As New XElement("Italic", txtChartTitle.Font.Italic)
            chartTitle.Add(chartTitleItalic)
        End If

        If PointChartDefaults.<PointChart>.<Settings>.<ChartTitle>.<Underline>.Value <> Nothing Then
            Dim chartTitleUnderline As New XElement("Underline", PointChartDefaults.<PointChart>.<Settings>.<ChartTitle>.<Underline>.Value)
            chartTitle.Add(chartTitleUnderline)
        Else
            Message.AddWarning("Default Chart Title Underline settings not found." & vbCrLf)
            Dim chartTitleUnderline As New XElement("Underline", txtChartTitle.Font.Underline)
            chartTitle.Add(chartTitleUnderline)
        End If

        If PointChartDefaults.<PointChart>.<Settings>.<ChartTitle>.<Strikeout>.Value <> Nothing Then
            Dim chartTitleStrikeout As New XElement("Strikeout", PointChartDefaults.<PointChart>.<Settings>.<ChartTitle>.<Strikeout>.Value)
            chartTitle.Add(chartTitleStrikeout)
        Else
            Message.AddWarning("Default Chart Title Strikeout settings not found." & vbCrLf)
            Dim chartTitleStrikeout As New XElement("Strikeout", txtChartTitle.Font.Strikeout)
            chartTitle.Add(chartTitleStrikeout)
        End If

        If PointChartDefaults.<PointChart>.<Settings>.<ChartTitle>.<Alignment>.Value <> Nothing Then
            Dim chartTitleAlignment As New XElement("Alignment", PointChartDefaults.<PointChart>.<Settings>.<ChartTitle>.<Alignment>.Value)
            chartTitle.Add(chartTitleAlignment)
        Else
            Message.AddWarning("Default Chart Title Alignment settings not found." & vbCrLf)
            Dim chartTitleAlignment As New XElement("Alignment", cmbAlignment.SelectedItem.ToString)
            chartTitle.Add(chartTitleAlignment)
        End If

        chartSettings.Add(chartTitle)

        Dim xAxis As New XElement("XAxis")

        If PointChartDefaults.<PointChart>.<Settings>.<XAxis>.<TitleText>.Value <> Nothing Then
            Dim xAxisTitleText As New XElement("TitleText", PointChartDefaults.<PointChart>.<Settings>.<XAxis>.<TitleText>.Value)
            xAxis.Add(xAxisTitleText)
        End If

        If PointChartDefaults.<PointChart>.<Settings>.<XAxis>.<TitleFontName>.Value <> Nothing Then
            Dim xAxisTitleFontName As New XElement("TitleFontName", PointChartDefaults.<PointChart>.<Settings>.<XAxis>.<TitleFontName>.Value)
            xAxis.Add(xAxisTitleFontName)
        End If

        If PointChartDefaults.<PointChart>.<Settings>.<XAxis>.<TitleColor>.Value <> Nothing Then
            Dim xAxisTitleColor As New XElement("TitleColor", PointChartDefaults.<PointChart>.<Settings>.<XAxis>.<TitleColor>.Value)
            xAxis.Add(xAxisTitleColor)
        End If

        If PointChartDefaults.<PointChart>.<Settings>.<XAxis>.<TitleSize>.Value <> Nothing Then
            Dim xAxisTitleSize As New XElement("TitleSize", PointChartDefaults.<PointChart>.<Settings>.<XAxis>.<TitleSize>.Value)
            xAxis.Add(xAxisTitleSize)
        End If

        If PointChartDefaults.<PointChart>.<Settings>.<XAxis>.<TitleBold>.Value <> Nothing Then
            Dim xAxisTitleBold As New XElement("TitleBold", PointChartDefaults.<PointChart>.<Settings>.<XAxis>.<TitleBold>.Value)
            xAxis.Add(xAxisTitleBold)
        End If

        If PointChartDefaults.<PointChart>.<Settings>.<XAxis>.<TitleItalic>.Value <> Nothing Then
            Dim xAxisTitleItalic As New XElement("TitleItalic", PointChartDefaults.<PointChart>.<Settings>.<XAxis>.<TitleItalic>.Value)
            xAxis.Add(xAxisTitleItalic)
        End If

        If PointChartDefaults.<PointChart>.<Settings>.<XAxis>.<TitleUnderline>.Value <> Nothing Then
            Dim xAxisTitleUnderline As New XElement("TitleUnderline", PointChartDefaults.<PointChart>.<Settings>.<XAxis>.<TitleUnderline>.Value)
            xAxis.Add(xAxisTitleUnderline)
        End If

        If PointChartDefaults.<PointChart>.<Settings>.<XAxis>.<TitleStrikeout>.Value <> Nothing Then
            Dim xAxisTitleStrikeout As New XElement("TitleStrikeout", PointChartDefaults.<PointChart>.<Settings>.<XAxis>.<TitleStrikeout>.Value)
            xAxis.Add(xAxisTitleStrikeout)
        End If

        If PointChartDefaults.<PointChart>.<Settings>.<XAxis>.<TitleAlignment>.Value <> Nothing Then
            Dim xAxisTitleAlignment As New XElement("TitleAlignment", PointChartDefaults.<PointChart>.<Settings>.<XAxis>.<TitleAlignment>.Value)
            xAxis.Add(xAxisTitleAlignment)
        End If

        If PointChartDefaults.<PointChart>.<Settings>.<XAxis>.<AutoMinimum>.Value <> Nothing Then
            Dim xAxisAutoMinimum As New XElement("AutoMinimum", PointChartDefaults.<PointChart>.<Settings>.<XAxis>.<AutoMinimum>.Value)
            xAxis.Add(xAxisAutoMinimum)
        End If

        If PointChartDefaults.<PointChart>.<Settings>.<XAxis>.<Minimum>.Value <> Nothing Then
            Dim xAxisMinimum As New XElement("Minimum", PointChartDefaults.<PointChart>.<Settings>.<XAxis>.<Minimum>.Value)
            xAxis.Add(xAxisMinimum)
        End If

        If PointChartDefaults.<PointChart>.<Settings>.<XAxis>.<AutoMaximum>.Value <> Nothing Then
            Dim xAxisAutoMaximum As New XElement("AutoMaximum", PointChartDefaults.<PointChart>.<Settings>.<XAxis>.<AutoMaximum>.Value)
            xAxis.Add(xAxisAutoMaximum)
        End If

        If PointChartDefaults.<PointChart>.<Settings>.<XAxis>.<Maximum>.Value <> Nothing Then
            Dim xAxisMaximum As New XElement("Maximum", PointChartDefaults.<PointChart>.<Settings>.<XAxis>.<Maximum>.Value)
            xAxis.Add(xAxisMaximum)
        End If

        If PointChartDefaults.<PointChart>.<Settings>.<XAxis>.<AutoInterval>.Value <> Nothing Then
            Dim xAxisAutoInterval As New XElement("AutoInterval", PointChartDefaults.<PointChart>.<Settings>.<XAxis>.<AutoInterval>.Value)
            xAxis.Add(xAxisAutoInterval)
        End If

        If PointChartDefaults.<PointChart>.<Settings>.<XAxis>.<MajorGridInterval>.Value <> Nothing Then
            Dim xAxisMajorGridInterval As New XElement("MajorGridInterval", PointChartDefaults.<PointChart>.<Settings>.<XAxis>.<MajorGridInterval>.Value)
            xAxis.Add(xAxisMajorGridInterval)
        End If

        If PointChartDefaults.<PointChart>.<Settings>.<XAxis>.<AutoMajorGridInterval>.Value <> Nothing Then
            Dim xAxisAutoMajorGridInterval As New XElement("AutoMajorGridInterval", PointChartDefaults.<PointChart>.<Settings>.<XAxis>.<AutoMajorGridInterval>.Value)
            xAxis.Add(xAxisAutoMajorGridInterval)
        End If

        chartSettings.Add(xAxis)


        Dim yAxis As New XElement("YAxis")

        If PointChartDefaults.<PointChart>.<Settings>.<YAxis>.<TitleText>.Value <> Nothing Then
            Dim yAxisTitleText As New XElement("TitleText", PointChartDefaults.<PointChart>.<Settings>.<YAxis>.<TitleText>.Value)
            yAxis.Add(yAxisTitleText)
        End If

        If PointChartDefaults.<PointChart>.<Settings>.<YAxis>.<TitleFontName>.Value <> Nothing Then
            Dim yAxisTitleFontName As New XElement("TitleFontName", PointChartDefaults.<PointChart>.<Settings>.<YAxis>.<TitleFontName>.Value)
            yAxis.Add(yAxisTitleFontName)
        End If

        If PointChartDefaults.<PointChart>.<Settings>.<YAxis>.<TitleColor>.Value <> Nothing Then
            Dim yAxisTitleColor As New XElement("TitleColor", PointChartDefaults.<PointChart>.<Settings>.<YAxis>.<TitleColor>.Value)
            yAxis.Add(yAxisTitleColor)
        End If

        If PointChartDefaults.<PointChart>.<Settings>.<YAxis>.<TitleSize>.Value <> Nothing Then
            Dim yAxisTitleSize As New XElement("TitleSize", PointChartDefaults.<PointChart>.<Settings>.<YAxis>.<TitleSize>.Value)
            yAxis.Add(yAxisTitleSize)
        End If

        If PointChartDefaults.<PointChart>.<Settings>.<YAxis>.<TitleBold>.Value <> Nothing Then
            Dim yAxisTitleBold As New XElement("TitleBold", PointChartDefaults.<PointChart>.<Settings>.<YAxis>.<TitleBold>.Value)
            yAxis.Add(yAxisTitleBold)
        End If

        If PointChartDefaults.<PointChart>.<Settings>.<YAxis>.<TitleItalic>.Value <> Nothing Then
            Dim yAxisTitleItalic As New XElement("TitleItalic", PointChartDefaults.<PointChart>.<Settings>.<YAxis>.<TitleItalic>.Value)
            yAxis.Add(yAxisTitleItalic)
        End If

        If PointChartDefaults.<PointChart>.<Settings>.<YAxis>.<TitleUnderline>.Value <> Nothing Then
            Dim yAxisTitleUnderline As New XElement("TitleUnderline", PointChartDefaults.<PointChart>.<Settings>.<YAxis>.<TitleUnderline>.Value)
            yAxis.Add(yAxisTitleUnderline)
        End If

        If PointChartDefaults.<PointChart>.<Settings>.<YAxis>.<TitleStrikeout>.Value <> Nothing Then
            Dim yAxisTitleStrikeout As New XElement("TitleStrikeout", PointChartDefaults.<PointChart>.<Settings>.<YAxis>.<TitleStrikeout>.Value)
            yAxis.Add(yAxisTitleStrikeout)
        End If

        If PointChartDefaults.<PointChart>.<Settings>.<YAxis>.<TitleAlignment>.Value <> Nothing Then
            Dim yAxisTitleAlignment As New XElement("TitleAlignment", PointChartDefaults.<PointChart>.<Settings>.<YAxis>.<TitleAlignment>.Value)
            yAxis.Add(yAxisTitleAlignment)
        End If

        If PointChartDefaults.<PointChart>.<Settings>.<YAxis>.<AutoMinimum>.Value <> Nothing Then
            Dim yAxisAutoMinimum As New XElement("AutoMinimum", PointChartDefaults.<PointChart>.<Settings>.<YAxis>.<AutoMinimum>.Value)
            yAxis.Add(yAxisAutoMinimum)
        End If

        If PointChartDefaults.<PointChart>.<Settings>.<YAxis>.<Minimum>.Value <> Nothing Then
            Dim yAxisMinimum As New XElement("Minimum", PointChartDefaults.<PointChart>.<Settings>.<YAxis>.<Minimum>.Value)
            yAxis.Add(yAxisMinimum)
        End If

        If PointChartDefaults.<PointChart>.<Settings>.<YAxis>.<AutoMaximum>.Value <> Nothing Then
            Dim yAxisAutoMaximum As New XElement("AutoMaximum", PointChartDefaults.<PointChart>.<Settings>.<YAxis>.<AutoMaximum>.Value)
            yAxis.Add(yAxisAutoMaximum)
        End If

        If PointChartDefaults.<PointChart>.<Settings>.<YAxis>.<Maximum>.Value <> Nothing Then
            Dim yAxisMaximum As New XElement("Maximum", PointChartDefaults.<PointChart>.<Settings>.<YAxis>.<Maximum>.Value)
            yAxis.Add(yAxisMaximum)
        End If

        If PointChartDefaults.<PointChart>.<Settings>.<YAxis>.<AutoInterval>.Value <> Nothing Then
            Dim yAxisAutoInterval As New XElement("AutoInterval", PointChartDefaults.<PointChart>.<Settings>.<YAxis>.<AutoInterval>.Value)
            yAxis.Add(yAxisAutoInterval)
        End If

        If PointChartDefaults.<PointChart>.<Settings>.<YAxis>.<MajorGridInterval>.Value <> Nothing Then
            Dim yAxisMajorGridInterval As New XElement("MajorGridInterval", PointChartDefaults.<PointChart>.<Settings>.<YAxis>.<MajorGridInterval>.Value)
            yAxis.Add(yAxisMajorGridInterval)
        End If

        chartSettings.Add(yAxis)


        Dim commandDrawChart As New XElement("Command", "DrawChart")
        chartSettings.Add(commandDrawChart)

        xmessage.Add(chartSettings)
        doc.Add(xmessage)

        If IsNothing(client) Then
            Message.AddWarning("No client connection available!" & vbCrLf)
            Beep()
        Else
            client.SendMessageAsync(AppNetName, "ADVL_Chart_1", doc.ToString)
            Message.XAddText("Message sent to " & "ADVL_Chart_1" & ":" & vbCrLf, "XmlSentNotice")
            Message.XAddXml(doc.ToString)
            Message.XAddText(vbCrLf, "Message") 'Add extra line
        End If
    End Sub

    Private Sub DisplayPointChartNoDefaults()
        'Display the point chart without using the default chart settings.

        'Send the instructions to the Chart application to display the stock chart.
        Dim decl As New XDeclaration("1.0", "utf-8", "yes")
        Dim doc As New XDocument(decl, Nothing) 'Create an XDocument to store the instructions.

        Try
            'Create the xml instructions

            Dim xmessage As New XElement("XMsg") 'This indicates the start of the message in the XMessage class

            'ADDED 3Feb19:
            Dim clientAppNetName As New XElement("ClientAppNetName", AppNetName)
            xmessage.Add(clientAppNetName)

            Dim clientName As New XElement("ClientName", ApplicationInfo.Name) 'This tells the coordinate server the name of the client making the request.
            xmessage.Add(clientName)

            Dim clientLocn As New XElement("ClientLocn", "PointChart")
            xmessage.Add(clientLocn)

            Dim chartSettings As New XElement("PointChartSettings")

            Dim chartType As New XElement("ChartType", "Point")
            chartSettings.Add(chartType)

            Dim commandClearChart As New XElement("Command", "ClearChart")
            chartSettings.Add(commandClearChart)

            Dim inputData As New XElement("InputData")
            Dim dataType As New XElement("Type", "Database")
            inputData.Add(dataType)
            Dim databasePath As New XElement("DatabasePath", txtPointChartDbPath.Text)
            inputData.Add(databasePath)
            Dim dataDescription As New XElement("DataDescription", txtPointSeriesName.Text)
            inputData.Add(dataDescription)
            Dim databaseQuery As New XElement("DatabaseQuery", txtPointChartQuery.Text)
            inputData.Add(databaseQuery)
            chartSettings.Add(inputData)

            Dim chartProperties As New XElement("ChartProperties")
            Dim seriesName As New XElement("SeriesName", txtPointSeriesName.Text)
            chartProperties.Add(seriesName)
            Dim xValuesFieldName As New XElement("XValuesFieldName", cmbPointXValues.SelectedItem.ToString)
            chartProperties.Add(xValuesFieldName)
            Dim yValuesFieldName As New XElement("YValuesFieldName", cmbPointYValues.SelectedItem.ToString)
            chartProperties.Add(yValuesFieldName)
            chartSettings.Add(chartProperties)

            Dim chartTitle As New XElement("ChartTitle")
            Dim chartTitleLabelName As New XElement("LabelName", "Label1")
            chartTitle.Add(chartTitleLabelName)
            Dim chartTitleText As New XElement("Text", txtPointChartTitle.Text)
            chartTitle.Add(chartTitleText)
            Dim chartTitleFontName As New XElement("FontName", txtPointChartTitle.Font.Name)
            chartTitle.Add(chartTitleFontName)
            If txtPointChartTitle.ForeColor.Name.Contains("Color [WindowText]") Then
                Dim chartTitleColor As New XElement("Color", "Color [Black]")
                chartTitle.Add(chartTitleColor)
            Else
                Dim chartTitleColor As New XElement("Color", txtPointChartTitle.ForeColor)
                chartTitle.Add(chartTitleColor)
            End If

            Dim chartTitleSize As New XElement("Size", txtPointChartTitle.Font.Size)
            chartTitle.Add(chartTitleSize)
            Dim chartTitleBold As New XElement("Bold", txtPointChartTitle.Font.Bold)
            chartTitle.Add(chartTitleBold)
            Dim chartTitleItalic As New XElement("Italic", txtPointChartTitle.Font.Italic)
            chartTitle.Add(chartTitleItalic)
            Dim chartTitleUnderline As New XElement("Underline", txtPointChartTitle.Font.Underline)
            chartTitle.Add(chartTitleUnderline)
            Dim chartTitleStrikeout As New XElement("Strikeout", txtPointChartTitle.Font.Strikeout)
            chartTitle.Add(chartTitleStrikeout)
            Dim chartTitleAlignment As New XElement("Alignment", cmbPointChartAlignment.SelectedItem.ToString)
            chartTitle.Add(chartTitleAlignment)
            chartSettings.Add(chartTitle)

            Dim xAxis As New XElement("XAxis")
            If chkAutoXRange.Checked Then
                Dim autoMinimum As New XElement("AutoMinimum", "true")
                xAxis.Add(autoMinimum)
                chartSettings.Add(xAxis)
            Else
                Dim autoMinimum As New XElement("AutoMinimum", "false")
                xAxis.Add(autoMinimum)
                Dim minimum As New XElement("Minimum", txtPointXMin.Text)
                xAxis.Add(minimum)
                Dim autoMaximum As New XElement("AutoMaximum", "false")
                xAxis.Add(autoMaximum)
                Dim maximum As New XElement("Maximum", txtPointXMax.Text)
                xAxis.Add(maximum)
                chartSettings.Add(xAxis)
            End If

            Dim yAxis As New XElement("YAxis")
            If chkAutoYRange.Checked Then
                Dim autoMinimum As New XElement("AutoMinimum", "true")
                yAxis.Add(autoMinimum)
                chartSettings.Add(yAxis)
            Else
                Dim autoMinimum As New XElement("AutoMinimum", "false")
                yAxis.Add(autoMinimum)
                Dim minimum As New XElement("Minimum", txtPointYMin.Text)
                yAxis.Add(minimum)
                Dim autoMaximum As New XElement("AutoMaximum", "false")
                yAxis.Add(autoMaximum)
                Dim maximum As New XElement("Maximum", txtPointYMax.Text)
                yAxis.Add(maximum)
                chartSettings.Add(yAxis)
            End If

            Dim commandDrawChart As New XElement("Command", "DrawChart")
            chartSettings.Add(commandDrawChart)

            xmessage.Add(chartSettings)
            doc.Add(xmessage)
        Catch ex As Exception
            Message.AddWarning("Error creating Point Chart instructions: " & ex.Message & vbCrLf)
        End Try

        If IsNothing(client) Then
            Message.AddWarning("No client connection available!" & vbCrLf)
            Beep()
        Else
            client.SendMessageAsync(AppNetName, "ADVL_Chart_1", doc.ToString)
            Message.XAddText("Message sent to " & "ADVL_Chart_1" & ":" & vbCrLf, "XmlSentNotice")
            Message.XAddXml(doc.ToString)
            Message.XAddText(vbCrLf, "Message") 'Add extra line
        End If
    End Sub

    Private Sub btnGetPointChartDefaults_Click(sender As Object, e As EventArgs) Handles btnGetPointChartDefaults.Click
        'Send a request to ADVL_Charts_1 for the current Point Chart settings.

        'Create the xml instructions
        Dim decl As New XDeclaration("1.0", "utf-8", "yes")
        Dim doc As New XDocument(decl, Nothing) 'Create an XDocument to store the instructions.
        Dim xmessage As New XElement("XMsg") 'This indicates the start of the message in the XMessage class

        Dim clientAppNetName As New XElement("ClientAppNetName", AppNetName)
        xmessage.Add(clientAppNetName)

        Dim clientName As New XElement("ClientName", ApplicationInfo.Name) 'This tells the coordinate server the name of the client making the request.
        xmessage.Add(clientName)

        Dim clientLocn As New XElement("ClientLocn", "PointChart")
        xmessage.Add(clientLocn)

        Dim commandGetSettings As New XElement("Command", "GetPointChartSettings")
        xmessage.Add(commandGetSettings)

        doc.Add(xmessage)

        If IsNothing(client) Then
            Message.AddWarning("No client connection available!" & vbCrLf)
            Beep()
        Else
            client.SendMessageAsync(AppNetName, "ADVL_Chart_1", doc.ToString)
            Message.XAddText("Message sent to " & "ADVL_Chart_1" & ":" & vbCrLf, "XmlSentNotice")
            Message.XAddXml(doc.ToString)
            Message.XAddText(vbCrLf, "Message") 'Add extra line
        End If
    End Sub

    Private Sub btnSavePointDefaults_Click(sender As Object, e As EventArgs) Handles btnSavePointDefaults.Click
        'Save the Point Chart Default Settings in a file:

        Try
            Dim xmlSeq As System.Xml.Linq.XDocument = XDocument.Parse(rtbPointChartDefaults.Text)

            Dim SettingsFileName As String = ""

            If Trim(txtPointChartSettings.Text).EndsWith(".PointChartDefaults") Then
                SettingsFileName = Trim(txtPointChartSettings.Text)
            Else
                SettingsFileName = Trim(txtPointChartSettings.Text) & ".PointChartDefaults"
                txtPointChartSettings.Text = SettingsFileName
            End If

            Project.SaveXmlData(SettingsFileName, xmlSeq)
            Message.Add("Point Chart Default Settings saved OK" & vbCrLf)
        Catch ex As Exception
            Message.AddWarning(ex.Message & vbCrLf)
            Beep()
        End Try

    End Sub

    Private Sub btnOpenPointChartDefaults_Click(sender As Object, e As EventArgs) Handles btnOpenPointChartDefaults.Click
        'Open a Point Chart Default Settings file:

        Dim SelectedFileName As String = ""

        SelectedFileName = Project.SelectDataFile("Cross Plot Chart Defaults", "PointChartDefaults")
        Message.Add("Selected Cross Plot Chart Default Settings: " & SelectedFileName & vbCrLf)

        txtPointChartSettings.Text = SelectedFileName

        Project.ReadXmlData(SelectedFileName, PointChartDefaults)

        If PointChartDefaults Is Nothing Then
            Exit Sub
        End If

        rtbPointChartDefaults.Text = PointChartDefaults.ToString

        FormatXmlText(rtbPointChartDefaults)

    End Sub


#End Region 'Charts Tab -----------------------------------------------------------------------------------------------------------------------------------------------------------------------

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
                Dim NDays As Double = Val(txtNDays.Text)
                Dim calcDate As DateTime = StartDate.AddDays(NDays)
                txtStartDateNDaysCalc.Text = calcDate.ToLongDateString

            Case "Date of Start Date - N Days"
                Dim StartDate As DateTime = DateTime.ParseExact(txtStartDate.Text, "dd MMM yyyy", Nothing)
                Dim NDays As Double = Val(txtNDays.Text)
                Dim calcDate As DateTime = StartDate.AddDays(-NDays)
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

#Region "Data Views Sub Tab" '=================================================================================================================================================================

#End Region 'Data Views Sub Tab ---------------------------------------------------------------------------------------------------------------------------------------------------------------

#Region "Database Tables Sub Tab" '============================================================================================================================================================

    Private Sub cmbUtilTablesDatabase_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbUtilTablesDatabase.SelectedIndexChanged
        'The selected database has changed.
        UtilTablesDatabaseChanged()

    End Sub

    Private Sub UtilTablesDatabaseChanged()
        'The Utilities Database Tables selection has changed.
        'Update the tab.

        FillCmbSelectTable()

        Select Case cmbUtilTablesDatabase.SelectedItem.ToString
            Case "Share Prices"
                txtUtilTablesPath.Text = SharePriceDbPath
            Case "Financials"
                txtUtilTablesPath.Text = FinancialsDbPath
            Case "Calculations"
                txtUtilTablesPath.Text = CalculationsDbPath
        End Select
    End Sub

    Public Sub FillCmbSelectTable()
        'Fill the cmbSelectTable listbox with the available tables in the selected database.

        Dim DatabasePath As String = ""

        Select Case cmbUtilTablesDatabase.SelectedItem.ToString
            Case "Share Prices"
                DatabasePath = SharePriceDbPath
            Case "Financials"
                DatabasePath = FinancialsDbPath
            Case "Calculations"
                DatabasePath = CalculationsDbPath
        End Select

        If DatabasePath = "" Then

            Exit Sub
        End If

        'Database access for MS Access:
        Dim connectionString As String 'Declare a connection string for MS Access - defines the database or server to be used.
        Dim conn As System.Data.OleDb.OleDbConnection 'Declare a connection for MS Access - used by the Data Adapter to connect to and disconnect from the database.
        Dim dt As DataTable

        Dim ds As DataSet = New DataSet

        cmbSelectTable.Text = ""
        cmbSelectTable.Items.Clear()
        ds.Clear()
        ds.Reset()

        'Specify the connection string:
        'Access 2003
        'connectionString = "provider=Microsoft.Jet.OLEDB.4.0;" + _
        '"data source = " + txtDatabase.Text

        'Access 2007:
        connectionString = "provider=Microsoft.ACE.OLEDB.12.0;" +
        "data source = " + DatabasePath

        'Connect to the Access database:
        conn = New System.Data.OleDb.OleDbConnection(connectionString)
        conn.Open()

        'This error occurs on the above line (conn.Open()):
        'Additional information: The 'Microsoft.ACE.OLEDB.12.0' provider is not registered on the local machine.
        'Fix attempt: 
        'http://www.microsoft.com/en-us/download/confirmation.aspx?id=23734
        'Download AccessDatabaseEngine.exe
        'Run the file to install the 2007 Office System Driver: Data Connectivity Components.

        Dim restrictions As String() = New String() {Nothing, Nothing, Nothing, "TABLE"} 'This restriction removes system tables
        dt = conn.GetSchema("Tables", restrictions)

        'Fill lstSelectTable
        Dim dr As DataRow
        Dim I As Integer 'Loop index
        Dim MaxI As Integer

        MaxI = dt.Rows.Count
        For I = 0 To MaxI - 1
            dr = dt.Rows(0)
            cmbSelectTable.Items.Add(dt.Rows(I).Item(2).ToString)
        Next I

        conn.Close()

    End Sub

    Private Sub btnDeleteTable_Click(sender As Object, e As EventArgs) Handles btnDeleteTable.Click
        'Delete the selected table.
        DeleteTable()
    End Sub

    Private Sub DeleteTable()
        'Dete the specified Database Table.

        Dim DatabasePath As String = ""

        Select Case cmbUtilTablesDatabase.SelectedItem.ToString
            Case "Share Prices"
                DatabasePath = SharePriceDbPath
            Case "Financials"
                DatabasePath = FinancialsDbPath
            Case "Calculations"
                DatabasePath = CalculationsDbPath
        End Select

        If DatabasePath = "" Then
            Message.AddWarning("No database selected!" & vbCrLf)
            Exit Sub
        End If

        'Database access for MS Access:
        Dim connectionString As String 'Declare a connection string for MS Access - defines the database or server to be used.
        Dim conn As System.Data.OleDb.OleDbConnection 'Declare a connection for MS Access - used by the Data Adapter to connect to and disconnect from the database.

        'Access 2007:
        connectionString = "provider=Microsoft.ACE.OLEDB.12.0;" +
        "data source = " + DatabasePath

        'Connect to the Access database:
        conn = New System.Data.OleDb.OleDbConnection(connectionString)
        conn.Open()

        Dim myQuery As String = "DROP TABLE [" & cmbSelectTable.SelectedItem.ToString & "]"
        Message.Add("myQuery: " & myQuery & vbCrLf)

        Dim cmd As New OleDb.OleDbCommand
        cmd.CommandText = myQuery
        cmd.Connection = conn

        Try
            cmd.ExecuteNonQuery()
        Catch ex As Exception
            Message.AddWarning("Error creating new table: " & ex.Message & vbCrLf)
        End Try
        conn.Close()

        FillCmbSelectTable()
    End Sub

    Private Sub btnDeleteRecords_Click(sender As Object, e As EventArgs) Handles btnDeleteRecords.Click
        'Delete all the records from the selected table
        DeleteRecords()
    End Sub

    Private Sub DeleteRecords()
        'Delete all the records in the specified table.

        Dim DatabasePath As String = ""

        Select Case cmbUtilTablesDatabase.SelectedItem.ToString
            Case "Share Prices"
                DatabasePath = SharePriceDbPath
            Case "Financials"
                DatabasePath = FinancialsDbPath
            Case "Calculations"
                DatabasePath = CalculationsDbPath
        End Select

        If DatabasePath = "" Then
            Message.AddWarning("No database selected!" & vbCrLf)
            Exit Sub
        End If

        'Database access for MS Access:
        Dim connectionString As String 'Declare a connection string for MS Access - defines the database or server to be used.
        Dim conn As System.Data.OleDb.OleDbConnection 'Declare a connection for MS Access - used by the Data Adapter to connect to and disconnect from the database.

        'Access 2007:
        connectionString = "provider=Microsoft.ACE.OLEDB.12.0;" +
        "data source = " + DatabasePath

        'Connect to the Access database:
        conn = New System.Data.OleDb.OleDbConnection(connectionString)
        conn.Open()

        Dim myQuery As String = "DELETE FROM [" & cmbSelectTable.SelectedItem.ToString & "]"
        Message.Add("myQuery: " & myQuery & vbCrLf)

        Dim cmd As New OleDb.OleDbCommand
        cmd.CommandText = myQuery
        cmd.Connection = conn

        Try
            cmd.ExecuteNonQuery()
        Catch ex As Exception
            Message.AddWarning("Error creating new table: " & ex.Message & vbCrLf)
        End Try
        conn.Close()
    End Sub

    Private Sub btnCopyTableCols_Click(sender As Object, e As EventArgs) Handles btnCopyTableCols.Click
        'Copy the columns in the selected table into a new table with the specified name.
        CopyTableColumns()
    End Sub


    Private Sub CopyTableColumns()
        'Copy the columns (but not the data) in the specified table into a new table.

        Dim NewTableName As String = Trim(txtNewTableName.Text)

        Dim DatabasePath As String = ""

        Select Case cmbUtilTablesDatabase.SelectedItem.ToString
            Case "Share Prices"
                DatabasePath = SharePriceDbPath
            Case "Financials"
                DatabasePath = FinancialsDbPath
            Case "Calculations"
                DatabasePath = CalculationsDbPath
        End Select

        If DatabasePath = "" Then
            Message.AddWarning("No database selected!" & vbCrLf)
            Exit Sub
        End If

        'Database access for MS Access:
        Dim connectionString As String 'Declare a connection string for MS Access - defines the database or server to be used.
        Dim conn As System.Data.OleDb.OleDbConnection 'Declare a connection for MS Access - used by the Data Adapter to connect to and disconnect from the database.

        'Access 2007:
        connectionString = "provider=Microsoft.ACE.OLEDB.12.0;" +
        "data source = " + DatabasePath

        'Connect to the Access database:
        conn = New System.Data.OleDb.OleDbConnection(connectionString)
        conn.Open()

        Dim myQuery As String = "SELECT  * INTO [" & txtNewTableName.Text & "] FROM [" & cmbSelectTable.SelectedItem.ToString & "]" & " WHERE 1 = 2"
        Message.Add("myQuery: " & myQuery & vbCrLf)

        Dim cmd As New OleDb.OleDbCommand
        cmd.CommandText = myQuery
        cmd.Connection = conn

        Dim mySchema As DataTable = conn.GetSchema("Indexes", New String() {Nothing, Nothing, Nothing, Nothing, cmbSelectTable.SelectedItem.ToString})
        'Restrictions: TABLE_CATALOG TABLE_SCHEMA INDEX_NAME TYPE TABLE_NAME

        Try
            cmd.ExecuteNonQuery()

            'Add the primary key:
            Dim I As Integer
            Dim NRows As Integer = mySchema.Rows.Count
            If NRows = 0 Then
                Message.AddWarning("No primary keys in the table." & vbCrLf)
            ElseIf NRows = 1 Then
                Dim sb As New System.Text.StringBuilder
                If mySchema.Rows(0).Item("PRIMARY_KEY") = True Then
                    sb.Append("CREATE INDEX idxPrimaryKey ON " & txtNewTableName.Text & " ([" & mySchema.Rows(0).Item("COLUMN_NAME") & "]) WITH PRIMARY")
                    Dim pkCommand As New OleDb.OleDbCommand
                    pkCommand.CommandText = sb.ToString
                    pkCommand.Connection = conn
                    Try
                        pkCommand.ExecuteNonQuery()
                        Message.Add("Single column primary key added OK." & vbCrLf)
                    Catch ex As Exception
                        Message.AddWarning(ex.Message & vbCrLf)
                    End Try
                Else
                    Message.AddWarning("The table index is not a primary key." & vbCrLf)
                End If

            Else
                Dim PKColNo As Integer = 0
                Dim sb As New System.Text.StringBuilder
                For I = 1 To NRows
                    If mySchema.Rows(I - 1).Item("PRIMARY_KEY") = True Then
                        PKColNo = PKColNo + 1
                        If PKColNo = 1 Then
                            sb.Append("CREATE INDEX idxPrimaryKey ON " & txtNewTableName.Text & " ([" & mySchema.Rows(I - 1).Item("COLUMN_NAME") & "]")
                        Else
                            sb.Append(", [" & mySchema.Rows(I - 1).Item("COLUMN_NAME") & "]")
                        End If
                    End If
                Next
                If PKColNo = 0 Then
                    Message.AddWarning("The table indexes do not contain a primary key." & vbCrLf)
                Else
                    sb.Append(") WITH PRIMARY")
                    Dim pkCommand As New OleDb.OleDbCommand
                    pkCommand.CommandText = sb.ToString
                    pkCommand.Connection = conn
                    Try
                        pkCommand.ExecuteNonQuery()
                        Message.Add("Multiple column primary key added OK." & vbCrLf)
                    Catch ex As Exception
                        Message.AddWarning(ex.Message & vbCrLf)
                    End Try
                End If
            End If

        Catch ex As Exception
            Message.AddWarning("Error creating New table: " & ex.Message & vbCrLf)
        End Try
        conn.Close()
        FillCmbSelectTable()

    End Sub

    Private Sub CopyTableColumns_Old()
        'Copy the columns (but not the data) in the specified table into a new table.

        Dim NewTableName As String = Trim(txtNewTableName.Text)

        Dim DatabasePath As String = ""

        Select Case cmbUtilTablesDatabase.SelectedItem.ToString
            Case "Share Prices"
                DatabasePath = SharePriceDbPath
            Case "Financials"
                DatabasePath = FinancialsDbPath
            Case "Calculations"
                DatabasePath = CalculationsDbPath
        End Select

        If DatabasePath = "" Then
            Message.AddWarning("No database selected!" & vbCrLf)
            Exit Sub
        End If

        'Database access for MS Access:
        Dim connectionString As String 'Declare a connection string for MS Access - defines the database or server to be used.
        Dim conn As System.Data.OleDb.OleDbConnection 'Declare a connection for MS Access - used by the Data Adapter to connect to and disconnect from the database.

        'Access 2007:
        connectionString = "provider=Microsoft.ACE.OLEDB.12.0;" +
        "data source = " + DatabasePath

        'Connect to the Access database:
        conn = New System.Data.OleDb.OleDbConnection(connectionString)
        conn.Open()

        Dim myQuery As String = "SELECT  * INTO [" & txtNewTableName.Text & "] FROM [" & cmbSelectTable.SelectedItem.ToString & "]" & " WHERE 1 = 2"
        Message.Add("myQuery: " & myQuery & vbCrLf)

        Dim cmd As New OleDb.OleDbCommand
        cmd.CommandText = myQuery
        cmd.Connection = conn

        Try
            cmd.ExecuteNonQuery()
        Catch ex As Exception
            Message.AddWarning("Error creating New table: " & ex.Message & vbCrLf)
        End Try
        conn.Close()
        FillCmbSelectTable()

    End Sub

    Private Sub btnCopyTable_Click(sender As Object, e As EventArgs) Handles btnCopyTable.Click
        'Copy the selected table (and data) into a new table with the specified name.
        CopyTable()

    End Sub

    Private Sub CopyTable()
        'Copy the selected table (and data) into a new table with the specified name.

        Dim NewTableName As String = Trim(txtNewTableName.Text)

        Dim DatabasePath As String = ""

        Select Case cmbUtilTablesDatabase.SelectedItem.ToString
            Case "Share Prices"
                DatabasePath = SharePriceDbPath
            Case "Financials"
                DatabasePath = FinancialsDbPath
            Case "Calculations"
                DatabasePath = CalculationsDbPath
        End Select

        If DatabasePath = "" Then
            Message.AddWarning("No database selected!" & vbCrLf)
            Exit Sub
        End If

        'Database access for MS Access:
        Dim connectionString As String 'Declare a connection string for MS Access - defines the database or server to be used.
        Dim conn As System.Data.OleDb.OleDbConnection 'Declare a connection for MS Access - used by the Data Adapter to connect to and disconnect from the database.

        'Access 2007:
        connectionString = "provider=Microsoft.ACE.OLEDB.12.0;" +
        "data source = " + DatabasePath

        'Connect to the Access database:
        conn = New System.Data.OleDb.OleDbConnection(connectionString)
        conn.Open()

        Dim myQuery As String = "SELECT * INTO [" & txtNewTableName.Text & "] FROM [" & cmbSelectTable.SelectedItem.ToString & "]"
        Message.Add("myQuery: " & myQuery & vbCrLf)

        Dim cmd As New OleDb.OleDbCommand
        cmd.CommandText = myQuery
        cmd.Connection = conn

        Dim mySchema As DataTable = conn.GetSchema("Indexes", New String() {Nothing, Nothing, Nothing, Nothing, cmbSelectTable.SelectedItem.ToString})
        'Restrictions: TABLE_CATALOG TABLE_SCHEMA INDEX_NAME TYPE TABLE_NAME

        Try
            cmd.ExecuteNonQuery()

            'Add the primary key:
            Dim I As Integer
            Dim NRows As Integer = mySchema.Rows.Count
            If NRows = 0 Then
                Message.AddWarning("No primary keys in the table." & vbCrLf)
            ElseIf NRows = 1 Then
                Dim sb As New System.Text.StringBuilder
                If mySchema.Rows(0).Item("PRIMARY_KEY") = True Then
                    sb.Append("CREATE INDEX idxPrimaryKey ON " & txtNewTableName.Text & " ([" & mySchema.Rows(0).Item("COLUMN_NAME") & "]) WITH PRIMARY")
                    Dim pkCommand As New OleDb.OleDbCommand
                    pkCommand.CommandText = sb.ToString
                    pkCommand.Connection = conn
                    Try
                        pkCommand.ExecuteNonQuery()
                        Message.Add("Single column primary key added OK." & vbCrLf)
                    Catch ex As Exception
                        Message.AddWarning(ex.Message & vbCrLf)
                    End Try
                Else
                    Message.AddWarning("The table index is not a primary key." & vbCrLf)
                End If

            Else
                Dim PKColNo As Integer = 0
                Dim sb As New System.Text.StringBuilder
                For I = 1 To NRows
                    If mySchema.Rows(I - 1).Item("PRIMARY_KEY") = True Then
                        PKColNo = PKColNo + 1
                        If PKColNo = 1 Then
                            sb.Append("CREATE INDEX idxPrimaryKey ON " & txtNewTableName.Text & " ([" & mySchema.Rows(I - 1).Item("COLUMN_NAME") & "]")
                        Else
                            sb.Append(", [" & mySchema.Rows(I - 1).Item("COLUMN_NAME") & "]")
                        End If
                    End If
                Next
                If PKColNo = 0 Then
                    Message.AddWarning("The table indexes do not contain a primary key." & vbCrLf)
                Else
                    sb.Append(") WITH PRIMARY")
                    Dim pkCommand As New OleDb.OleDbCommand
                    pkCommand.CommandText = sb.ToString
                    pkCommand.Connection = conn
                    Try
                        pkCommand.ExecuteNonQuery()
                        Message.Add("Multiple column primary key added OK." & vbCrLf)
                    Catch ex As Exception
                        Message.AddWarning(ex.Message & vbCrLf)
                    End Try
                End If
            End If
        Catch ex As Exception
            Message.AddWarning("Error creating New table: " & ex.Message & vbCrLf)
        End Try
        conn.Close()
        FillCmbSelectTable()
    End Sub

    Private Sub btnAddDeleteTableToSequence_Click(sender As Object, e As EventArgs) Handles btnAddDeleteTableToSequence.Click
        'Add the Delete Table sequence to the Processing Sequence.

        If IsNothing(Sequence) Then
            Message.AddWarning("The Processing Sequence form Is Not open." & vbCrLf)
        Else
            'Write new instructions to the Sequence text.
            Dim Posn As Integer
            Posn = Sequence.rtbSequence.SelectionStart

            Sequence.rtbSequence.SelectedText = vbCrLf & "<!--Delete Table: Settings used to delete a table from a database :-->" & vbCrLf
            Sequence.rtbSequence.SelectedText = "  <DeleteTable>" & vbCrLf

            'Input data parameters:
            Sequence.rtbSequence.SelectedText = "    <Database>" & cmbUtilTablesDatabase.SelectedItem.ToString & "</Database>" & vbCrLf
            Select Case cmbUtilTablesDatabase.SelectedItem.ToString
                Case "Share Prices"
                    Sequence.rtbSequence.SelectedText = "    <DatabasePath>" & SharePriceDbPath & "</DatabasePath>" & vbCrLf
                Case "Financials"
                    Sequence.rtbSequence.SelectedText = "    <DatabasePath>" & FinancialsDbPath & "</DatabasePath>" & vbCrLf
                Case "Calculations"
                    Sequence.rtbSequence.SelectedText = "    <DatabasePath>" & CalculationsDbPath & "</DatabasePath>" & vbCrLf
            End Select
            Sequence.rtbSequence.SelectedText = "    <TableToDelete>" & cmbSelectTable.SelectedItem.ToString & "</TableToDelete>" & vbCrLf

            Sequence.rtbSequence.SelectedText = "    <Command>Apply</Command>" & vbCrLf
            Sequence.rtbSequence.SelectedText = "  </DeleteTable>" & vbCrLf
            Sequence.rtbSequence.SelectedText = "<!---->" & vbCrLf

            Sequence.FormatXmlText()

        End If

    End Sub

    Private Sub btnAddDeleteRecordsToSequence_Click(sender As Object, e As EventArgs) Handles btnAddDeleteRecordsToSequence.Click
        'Add the Delete Records sequence to the Processing Sequence.

        If IsNothing(Sequence) Then
            Message.AddWarning("The Processing Sequence form is not open." & vbCrLf)
        Else
            'Write new instructions to the Sequence text.
            Dim Posn As Integer
            Posn = Sequence.rtbSequence.SelectionStart

            Sequence.rtbSequence.SelectedText = vbCrLf & "<!--Delete Records: Settings used to delete the records from a table :-->" & vbCrLf
            Sequence.rtbSequence.SelectedText = "  <DeleteRecords>" & vbCrLf

            'Input data parameters:
            Sequence.rtbSequence.SelectedText = "    <Database>" & cmbUtilTablesDatabase.SelectedItem.ToString & "</Database>" & vbCrLf
            Select Case cmbUtilTablesDatabase.SelectedItem.ToString
                Case "Share Prices"
                    Sequence.rtbSequence.SelectedText = "    <DatabasePath>" & SharePriceDbPath & "</DatabasePath>" & vbCrLf
                Case "Financials"
                    Sequence.rtbSequence.SelectedText = "    <DatabasePath>" & FinancialsDbPath & "</DatabasePath>" & vbCrLf
                Case "Calculations"
                    Sequence.rtbSequence.SelectedText = "    <DatabasePath>" & CalculationsDbPath & "</DatabasePath>" & vbCrLf
            End Select
            Sequence.rtbSequence.SelectedText = "    <Table>" & cmbSelectTable.SelectedItem.ToString & "</Table>" & vbCrLf

            Sequence.rtbSequence.SelectedText = "    <Command>Apply</Command>" & vbCrLf
            Sequence.rtbSequence.SelectedText = "  </DeleteRecords>" & vbCrLf
            Sequence.rtbSequence.SelectedText = "<!---->" & vbCrLf

            Sequence.FormatXmlText()
        End If

    End Sub

    Private Sub btnAddCopyColsToSequence_Click(sender As Object, e As EventArgs) Handles btnAddCopyColsToSequence.Click
        'Add the Copy Table Columns sequence to the Processing Sequence.

        If IsNothing(Sequence) Then
            Message.AddWarning("The Processing Sequence form is not open." & vbCrLf)
        Else
            'Write new instructions to the Sequence text.
            Dim Posn As Integer
            Posn = Sequence.rtbSequence.SelectionStart

            Sequence.rtbSequence.SelectedText = vbCrLf & "<!--Copy Table Columns: Settings used to copy the columns in a table to a new table :-->" & vbCrLf
            Sequence.rtbSequence.SelectedText = "  <CopyTableColumns>" & vbCrLf

            'Input data parameters:
            Sequence.rtbSequence.SelectedText = "    <Database>" & cmbUtilTablesDatabase.SelectedItem.ToString & "</Database>" & vbCrLf
            Select Case cmbUtilTablesDatabase.SelectedItem.ToString
                Case "Share Prices"
                    Sequence.rtbSequence.SelectedText = "    <DatabasePath>" & SharePriceDbPath & "</DatabasePath>" & vbCrLf
                Case "Financials"
                    Sequence.rtbSequence.SelectedText = "    <DatabasePath>" & FinancialsDbPath & "</DatabasePath>" & vbCrLf
                Case "Calculations"
                    Sequence.rtbSequence.SelectedText = "    <DatabasePath>" & CalculationsDbPath & "</DatabasePath>" & vbCrLf
            End Select
            Sequence.rtbSequence.SelectedText = "    <Table>" & cmbSelectTable.SelectedItem.ToString & "</Table>" & vbCrLf
            Sequence.rtbSequence.SelectedText = "    <NewTable>" & txtNewTableName.Text & "</NewTable>" & vbCrLf

            Sequence.rtbSequence.SelectedText = "    <Command>Apply</Command>" & vbCrLf
            Sequence.rtbSequence.SelectedText = "  </CopyTableColumns>" & vbCrLf
            Sequence.rtbSequence.SelectedText = "<!---->" & vbCrLf

            Sequence.FormatXmlText()
        End If

    End Sub

    Private Sub btnAddCopyTableToSequence_Click(sender As Object, e As EventArgs) Handles btnAddCopyTableToSequence.Click
        'Add the Copy Table sequence to the Processing Sequence.

        If IsNothing(Sequence) Then
            Message.AddWarning("The Processing Sequence form is not open." & vbCrLf)
        Else
            'Write new instructions to the Sequence text.
            Dim Posn As Integer
            Posn = Sequence.rtbSequence.SelectionStart

            Sequence.rtbSequence.SelectedText = vbCrLf & "<!--Copy Table: Settings used to copy the columns and data in a table to a new table :-->" & vbCrLf
            Sequence.rtbSequence.SelectedText = "  <CopyTable>" & vbCrLf

            'Input data parameters:
            Sequence.rtbSequence.SelectedText = "    <Database>" & cmbUtilTablesDatabase.SelectedItem.ToString & "</Database>" & vbCrLf
            Select Case cmbUtilTablesDatabase.SelectedItem.ToString
                Case "Share Prices"
                    Sequence.rtbSequence.SelectedText = "    <DatabasePath>" & SharePriceDbPath & "</DatabasePath>" & vbCrLf
                Case "Financials"
                    Sequence.rtbSequence.SelectedText = "    <DatabasePath>" & FinancialsDbPath & "</DatabasePath>" & vbCrLf
                Case "Calculations"
                    Sequence.rtbSequence.SelectedText = "    <DatabasePath>" & CalculationsDbPath & "</DatabasePath>" & vbCrLf
            End Select
            Sequence.rtbSequence.SelectedText = "    <Table>" & cmbSelectTable.SelectedItem.ToString & "</Table>" & vbCrLf
            Sequence.rtbSequence.SelectedText = "    <NewTable>" & txtNewTableName.Text & "</NewTable>" & vbCrLf

            Sequence.rtbSequence.SelectedText = "    <Command>Apply</Command>" & vbCrLf
            Sequence.rtbSequence.SelectedText = "  </CopyTable>" & vbCrLf
            Sequence.rtbSequence.SelectedText = "<!---->" & vbCrLf

            Sequence.FormatXmlText()
        End If

    End Sub


#End Region 'Database Tables Sub Tab ----------------------------------------------------------------------------------------------------------------------------------------------------------


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

            'Utilities - Database Tables Code: -----------------------------------------------------------------------------------------------------------
            Case "DeleteTable:Database"
                cmbUtilTablesDatabase.SelectedIndex = cmbUtilTablesDatabase.FindStringExact(Info)
            Case "DeleteTable:DatabasePath"
                 'Not used. Database path determined from Database selection.
            Case "DeleteTable:TableToDelete"
                cmbSelectTable.SelectedIndex = cmbSelectTable.FindStringExact(Info)
            Case "DeleteTable:TableToDelete:ReadParameter"
                If XSeq.Parameter.ContainsKey(Info) Then
                    cmbSelectTable.SelectedIndex = cmbSelectTable.FindStringExact(XSeq.Parameter(Info).Value)
                Else
                    Message.AddWarning("Instruction error - DeleteTable:TableToDelete:ReadParameter - The following parameter was not found: " & Info & vbCrLf)
                End If
            Case "DeleteTable:Command"
                Select Case Info
                    Case "Apply"
                        DeleteTable()
                    Case Else
                        Message.AddWarning("Unknown DeleteTable:Command Information Value: " & Info & vbCrLf)
                End Select

            Case "DeleteRecords:Database"
                cmbUtilTablesDatabase.SelectedIndex = cmbUtilTablesDatabase.FindStringExact(Info)
            Case "DeleteRecords:DatabasePath"
                 'Not used. Database path determined from Database selection.
            Case "DeleteRecords:Table"
                cmbSelectTable.SelectedIndex = cmbSelectTable.FindStringExact(Info)
            Case "DeleteRecords:Table:ReadParameter"
                If XSeq.Parameter.ContainsKey(Info) Then
                    cmbSelectTable.SelectedIndex = cmbSelectTable.FindStringExact(XSeq.Parameter(Info).Value)
                Else
                    Message.AddWarning("Instruction error - DeleteRecords:Table:ReadParameter - The following parameter was not found: " & Info & vbCrLf)
                End If

            Case "DeleteRecords:Command"
                Select Case Info
                    Case "Apply"
                        DeleteRecords()
                    Case Else
                        Message.AddWarning("Unknown DeleteRecords:Command Information Value: " & Info & vbCrLf)
                End Select

            Case "CopyTableColumns:Database"
                cmbUtilTablesDatabase.SelectedIndex = cmbUtilTablesDatabase.FindStringExact(Info)
            Case "CopyTableColumns:DatabasePath"
                'Not used. Database path determined from Database selection.
            Case "CopyTableColumns:Table"
                cmbSelectTable.SelectedIndex = cmbSelectTable.FindStringExact(Info)
            Case "CopyTableColumns:Table:ReadParameter"
                If XSeq.Parameter.ContainsKey(Info) Then
                    cmbSelectTable.SelectedIndex = cmbSelectTable.FindStringExact(XSeq.Parameter(Info).Value)
                Else
                    Message.AddWarning("Instruction error - CopyTableColumns:Table:ReadParameter - The following parameter was not found: " & Info & vbCrLf)
                End If

            Case "CopyTableColumns:NewTable"
                txtNewTableName.Text = Info
            Case "CopyTableColumns:NewTable:ReadParameter"
                If XSeq.Parameter.ContainsKey(Info) Then
                    txtNewTableName.Text = XSeq.Parameter(Info).Value
                Else
                    Message.AddWarning("Instruction error - CopyTableColumns:Table:ReadParameter - The following parameter was not found: " & Info & vbCrLf)
                End If

            Case "CopyTableColumns:Command"
                Select Case Info
                    Case "Apply"
                        CopyTableColumns()
                    Case Else
                        Message.AddWarning("Unknown CopyTableColumns:Command Information Value: " & Info & vbCrLf)
                End Select

            Case "CopyTable:Database"
                cmbUtilTablesDatabase.SelectedIndex = cmbUtilTablesDatabase.FindStringExact(Info)
            Case "CopyTable:DatabasePath"
                 'Not used. Database path determined from Database selection.
            Case "CopyTable:Table"
                cmbSelectTable.SelectedIndex = cmbSelectTable.FindStringExact(Info)
            Case "CopyTable:Table:ReadParameter"
                If XSeq.Parameter.ContainsKey(Info) Then
                    cmbSelectTable.SelectedIndex = cmbSelectTable.FindStringExact(XSeq.Parameter(Info).Value)
                Else
                    Message.AddWarning("Instruction error - CopyTable:Table:ReadParameter - The following parameter was not found: " & Info & vbCrLf)
                End If

            Case "CopyTable:NewTable"
                txtNewTableName.Text = Info
            Case "CopyTable:NewTable:ReadParameter"
                If XSeq.Parameter.ContainsKey(Info) Then
                    txtNewTableName.Text = XSeq.Parameter(Info).Value
                Else
                    Message.AddWarning("Instruction error - CopyTable:NewTable:ReadParameter - The following parameter was not found: " & Info & vbCrLf)
                End If

            Case "CopyTable:Command"
                Select Case Info
                    Case "Apply"
                        CopyTable()
                    Case Else
                        Message.AddWarning("Unknown CopyTable:Command Information Value: " & Info & vbCrLf)
                End Select


            'End of Sequence Code: -----------------------------------------------------------------------------------------------------------------------
            Case "EndOfSequence"
                XSeq.Parameter.Clear() 'Clear the Parameter dictionary.
                Message.Add("Processing sequence has completed." & vbCrLf)

            Case Else
                Message.AddWarning("Unknown Information Location: " & Locn & vbCrLf)
        End Select

    End Sub




#End Region 'Run XSequence Code ---------------------------------------------------------------------------------------------------------------------------------------------------------------

    Private Sub ApplicationInfo_UpdateExePath() Handles ApplicationInfo.UpdateExePath
        'Update the Executable Path.
        ApplicationInfo.ExecutablePath = Application.ExecutablePath
    End Sub

    Private Sub Zip_FileSelected(FileName As String) Handles Zip.FileSelected

    End Sub

    Private Sub XMsgLocal_Instruction(Info As String, Locn As String) Handles XMsgLocal.Instruction
        'Process an XMessage instruction locally.

        If IsDBNull(Info) Then
            Info = ""
        End If

        'Intercept and instructions with the prefix "WebPage_"
        If Locn.StartsWith("WebPage_") Then 'Send the Info, Location data to the correct Web Page:
            Dim EndOfWebPageNoString As Integer = Locn.IndexOf(":")
            Dim PageNoLen As Integer = EndOfWebPageNoString - 8
            Dim WebPageNoString As String = Locn.Substring(8, PageNoLen)
            Dim WebPageNo As Integer = CInt(WebPageNoString)
            Dim WebPageInfo As String = Info
            Dim WebPageLocn As String = Locn.Substring(EndOfWebPageNoString + 1)

            WebPageFormList(WebPageNo).XMsgInstruction(WebPageInfo, WebPageLocn)

        Else

            Select Case Locn
                Case "ClientName"
                    ClientAppName = Info 'The name of the Client requesting service.

                Case "Main"
                 'Blank message - do nothing.

                Case "Main:Status"
                    Select Case Info
                        Case "OK"
                            'Main instructions completed OK
                    End Select

                Case "EndOfSequence"
                    'End of Information Vector Sequence reached.

                Case Else
                    Message.AddWarning("Unknown location: " & Locn & vbCrLf)
                    Message.AddWarning("            info: " & Info & vbCrLf)
            End Select
        End If
    End Sub


    Private Sub chkConnect_LostFocus(sender As Object, e As EventArgs) Handles chkConnect.LostFocus
        If chkConnect.Checked Then
            Project.ConnectOnOpen = True
        Else
            Project.ConnectOnOpen = False
        End If
        Project.SaveProjectInfoFile()

    End Sub

    Private Sub Timer3_Tick(sender As Object, e As EventArgs) Handles Timer3.Tick
        'Keet the connection awake with each tick:

        If ConnectedToComNet = True Then
            Try
                If client.IsAlive() Then
                    Message.Add(Format(Now, "HH:mm:ss") & " Connection OK." & vbCrLf)
                    Timer3.Interval = TimeSpan.FromMinutes(55).TotalMilliseconds '55 minute interval
                Else
                    Message.Add(Format(Now, "HH:mm:ss") & " Connection Fault." & vbCrLf)
                    Timer3.Interval = TimeSpan.FromMinutes(55).TotalMilliseconds '55 minute interval
                End If
            Catch ex As Exception
                Message.AddWarning(ex.Message & vbCrLf)
                'Set interval to five minutes - try again in five minutes:
                Timer3.Interval = TimeSpan.FromMinutes(5).TotalMilliseconds '5 minute interval
            End Try
        Else
            Message.Add(Format(Now, "HH:mm:ss") & " Not connected." & vbCrLf)
        End If

    End Sub

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


