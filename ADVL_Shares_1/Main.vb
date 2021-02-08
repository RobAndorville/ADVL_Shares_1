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

Imports System.ComponentModel
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
    'Enter the address: http://localhost:8734/ADVLService
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

    Public WithEvents DesignPointChartQuery As frmDesignQuery
    Public WithEvents DesignLineChartQuery As frmDesignQuery

    'Declare objects used to connect to the Application Network:
    Public client As ServiceReference1.MsgServiceClient
    Public WithEvents XMsg As New ADVL_Utilities_Library_1.XMessage
    Dim XDoc As New System.Xml.XmlDocument
    Public Status As New System.Collections.Specialized.StringCollection
    Dim ClientAppName As String = "" 'The name of the client requesting service 
    Dim ClientProNetName As String = "" 'The name of the client Project Network requesting service. 
    Dim ClientConnName As String = "" 'The name of the client connection requesting service
    Dim MessageXDoc As System.Xml.Linq.XDocument
    Dim xmessage As XElement 'This will contain the message. It will be added to MessageXDoc.
    Dim xlocns As New List(Of XElement) 'A list of locations. Each location forms part of the reply message. The information in the reply message will be sent to the specified location in the client application.
    Dim MessageText As String = "" 'The text of a message sent through the Application Network.

    Public OnCompletionInstruction As String = "Stop" 'The last instruction returned on completion of the processing of an XMessage.
    Public EndInstruction As String = "Stop" 'Another method of specifying the last instruction. This is processed in the EndOfSequence section of XMsg.Instructions.

    Public ConnectionName As String = "" 'The name of the connection used to connect this application to the AppNet.
    Public ProNetName As String = "" 'The name of the Project Network
    Public ProNetPath As String = "" 'The path of the Project Network

    Public AdvlNetworkAppPath As String = "" 'The application path of the ADVL Network application (ComNet). This is where the "Application.Lock" file will be while ComNet is running
    Public AdvlNetworkExePath As String = "" 'The executable path of the ADVL Network.

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

    Dim WithEvents Zip As ADVL_Utilities_Library_1.ZipComp

    Dim dsInput As DataSet = New DataSet 'The input dataset for calculations.
    Dim dsOutput As DataSet = New DataSet 'The output dataset for calculations.
    Dim outputQuery As String
    Dim outputConnString As String
    Dim outputConnection As OleDb.OleDbConnection = New OleDb.OleDbConnection
    Dim outputDa As OleDb.OleDbDataAdapter

    Public WithEvents XSeq As New ADVL_Utilities_Library_1.XSequence 'This is used to run a set of XML Sequence statements. These are used for data processing.

    Dim cboFieldSelections As New DataGridViewComboBoxColumn 'Used for selecting Y Value fields in the Charts: Share Prices tab

    Dim StockChartSettingsList As New XDocument 'Stock chart settings list.
    Dim PointChartSettingsList As New XDocument 'Point chart (Cross Plot) settings list.
    Dim LineChartSettingsList As New XDocument  'Line chart settings list.

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

    Dim CalculationsFormNo As Integer = -1 'The number of the form used to display the Calculations Data.

    Private WithEvents bgwComCheck As New System.ComponentModel.BackgroundWorker 'Used to perform communication checks on a separate thread.

    'Private WithEvents bgwSendMessage As New System.ComponentModel.BackgroundWorker 'Used to send a message through the Message Service.
    Public WithEvents bgwSendMessage As New System.ComponentModel.BackgroundWorker 'Used to send a message through the Message Service.
    Public SendMessageParams As New clsSendMessageParams 'This hold the Send Message parameters: .ProjectNetworkName, .ConnectionName & .Message

    'Alternative SendMessage background worker - needed to send a message while instructions are being processed.
    Public WithEvents bgwSendMessageAlt As New System.ComponentModel.BackgroundWorker 'Used to send a message through the Message Service - alternative backgound worker.
    Dim SendMessageParamsAlt As New clsSendMessageParams 'This holds the Send Message parameters: .ProjectNetworkName, .ConnectionName & .Message - for the alternative background worker.

    Public WithEvents bgwRunInstruction As New System.ComponentModel.BackgroundWorker 'Used to run a single instruction
    Dim InstructionParams As New clsInstructionParams 'This holds the Info and Locn parameters of an instruction.


    Public Proj As New Proj 'Proj contains a list of all projects. Proj also contains methods to read, add and save the list.
    'This is used to select projects for use by this application, such as displaying share charts.
    'Each Project entry contains:
    '  Name, ProNetname, ID, Type, Path, Description, ApplicationName, ParentProjectName, ParentProjectID

    Dim ProjListNo As Integer 'The current Project List number. This is used to load project information from an XMessage.

    Public ShareChartProj As New Proj 'List of Share Chart projects.
    Public SelShareChartProjNo As Integer = -1 'The selected Share Chart project
    Public PointChartProj As New Proj ' List of Point Chart projects
    Public SelPointChartProjNo As Integer = -1 'The selected Point HCart project
    Public LineChartProj As New Proj 'List of Line Chart projects.
    Public SelLineChartProjNo As Integer = -1 'The selected LineChart project.

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

        Dim MsgType As String
        If Instructions.StartsWith("<XMsg>") Then
            MsgType = "XMsg"
            If ShowXMessages Then
                'Add the message header to the XMessages window:
                Message.XAddText("Message received: " & vbCrLf, "XmlReceivedNotice")
            End If
        ElseIf Instructions.StartsWith("<XSys>") Then
            MsgType = "XSys"
            If ShowSysMessages Then
                'Add the message header to the XMessages window:
                Message.XAddText("System Message received: " & vbCrLf, "XmlReceivedNotice")
            End If
        Else
            MsgType = "Unknown"

        End If

        If MsgType = "XMsg" Or MsgType = "XSys" Then 'This is an XMessage or XSystem set of instructions.
                Try
                    ClientProNetName = ""
                    ClientConnName = ""
                    ClientAppName = ""
                    'Inititalise the reply message:
                    Dim Decl As New XDeclaration("1.0", "utf-8", "yes")
                MessageXDoc = New XDocument(Decl, Nothing) 'Reply message - this will be sent to the Client App.
                xmessage = New XElement(MsgType)
                xlocns.Clear() 'Clear the list of locations in the reply message. 
                    xlocns.Add(New XElement("Main")) 'Initially set the location in the Client App to Main.

                    'Run the received message:
                    Dim XmlHeader As String = "<?xml version=""1.0"" encoding=""utf-8"" standalone=""yes""?>"

                XDoc.LoadXml(XmlHeader & vbCrLf & Instructions.Replace("&", "&amp;")) 'Replace "&" with "&amp:" before loading the XML text.
                If (MsgType = "XMsg") And ShowXMessages Then
                    Message.XAddXml(XDoc)  'Add the message to the XMessages window.
                    Message.XAddText(vbCrLf, "Normal") 'Add extra line
                ElseIf (MsgType = "XSys") And ShowSysMessages Then
                    Message.XAddXml(XDoc)  'Add the message to the XMessages window.
                    Message.XAddText(vbCrLf, "Normal") 'Add extra line
                End If
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
                If (MsgType = "XMsg") And ShowXMessages Then
                    Message.XAddText("Message processed locally:" & vbCrLf, "XmlSentNotice")
                    Message.XAddXml(MessageText)
                    Message.XAddText(vbCrLf, "Normal") 'Add extra line
                ElseIf (MsgType = "XSys") And ShowSysMessages Then
                    Message.XAddText("System Message processed locally:" & vbCrLf, "XmlSentNotice")
                    Message.XAddXml(MessageText)
                    Message.XAddText(vbCrLf, "Normal") 'Add extra line
                End If
                ProcessLocalInstructions(MessageText)
                Else
                'No client to send a message to - process the message locally.
                If (MsgType = "XMsg") And ShowXMessages Then
                    Message.XAddText("Message sent to [" & ClientProNetName & "]." & ClientConnName & ":" & vbCrLf, "XmlSentNotice")   'NOTE: There is no SendMessage code in the Message Service application!
                    Message.XAddXml(MessageText)
                    Message.XAddText(vbCrLf, "Normal") 'Add extra line
                ElseIf (MsgType = "XSys") And ShowSysMessages Then
                    Message.XAddText("System Message sent to [" & ClientProNetName & "]." & ClientConnName & ":" & vbCrLf, "XmlSentNotice")   'NOTE: There is no SendMessage code in the Message Service application!
                    Message.XAddXml(MessageText)
                    Message.XAddText(vbCrLf, "Normal") 'Add extra line
                End If

                'Send Message on a new thread:
                SendMessageParams.ProjectNetworkName = ClientProNetName
                    SendMessageParams.ConnectionName = ClientConnName
                    SendMessageParams.Message = MessageText
                    If bgwSendMessage.IsBusy Then
                        Message.AddWarning("Send Message backgroundworker is busy." & vbCrLf)
                    Else
                        bgwSendMessage.RunWorkerAsync(SendMessageParams)
                    End If
                End If
            Else 'This is not an XMessage!
                If Instructions.StartsWith("<XMsgBlk>") Then 'This is an XMessageBlock.
                'Process the received message:
                Dim XmlHeader As String = "<?xml version=""1.0"" encoding=""utf-8"" standalone=""yes""?>"
                XDoc.LoadXml(XmlHeader & vbCrLf & Instructions.Replace("&", "&amp;")) 'Replace "&" with "&amp:" before loading the XML text.
                If ShowXMessages Then
                    Message.XAddXml(XDoc)   'Add the message to the XMessages window.
                    Message.XAddText(vbCrLf, "Normal") 'Add extra line
                End If

                'Process the XMessageBlock:
                Dim XMsgBlkLocn As String
                XMsgBlkLocn = XDoc.GetElementsByTagName("ClientLocn")(0).InnerText
                Select Case XMsgBlkLocn
                    Case "StockChart"
                        Dim XData As Xml.XmlNodeList = XDoc.GetElementsByTagName("XInfo")
                        Dim ChartXDoc As New Xml.XmlDocument
                        Try
                            ChartXDoc.LoadXml("<?xml version=""1.0"" encoding=""utf-8"" standalone=""yes""?>" & XData(0).InnerXml)
                            XmlStockChartSettingsList.Rtf = XmlStockChartSettingsList.XmlToRtf(ChartXDoc, False)
                        Catch ex As Exception
                            Message.Add(ex.Message & vbCrLf)
                        End Try

                    Case "PointChart"
                        Dim XData As Xml.XmlNodeList = XDoc.GetElementsByTagName("XInfo")
                        Dim ChartXDoc As New Xml.XmlDocument
                        Try
                            ChartXDoc.LoadXml("<?xml version=""1.0"" encoding=""utf-8"" standalone=""yes""?>" & XData(0).InnerXml)
                            XmlPointChartSettingsList.Rtf = XmlStockChartSettingsList.XmlToRtf(ChartXDoc, False)
                        Catch ex As Exception
                            Message.Add(ex.Message & vbCrLf)
                        End Try

                    Case "LineChart"
                        Dim XData As Xml.XmlNodeList = XDoc.GetElementsByTagName("XInfo")
                        Dim ChartXDoc As New Xml.XmlDocument
                        Try
                            ChartXDoc.LoadXml("<?xml version=""1.0"" encoding=""utf-8"" standalone=""yes""?>" & XData(0).InnerXml)
                            XmlLineChartSettingsList.Rtf = XmlLineChartSettingsList.XmlToRtf(ChartXDoc, False)
                        Catch ex As Exception
                            Message.Add(ex.Message & vbCrLf)
                        End Try

                    Case Else
                        Message.AddWarning("Unknown XInfo Message location: " & XMsgBlkLocn & vbCrLf)
                End Select
            Else 'This is not an XMessage or an XMessageBlock!
                Message.XAddText("The message is not an XMessage or XMessageBlock: " & vbCrLf & Instructions & vbCrLf & vbCrLf, "Normal")
            End If
        End If
    End Sub

    Private Sub ProcessLocalInstructions(ByVal Instructions As String)
        'Process the XMessage instructions locally.

        'If Instructions.StartsWith("<XMsg>") Then 'This is an XMessage set of instructions.
        If Instructions.StartsWith("<XMsg>") Or Instructions.StartsWith("<XSys>") Then 'This is an XMessage set of instructions.
                'Run the received message:
                Dim XmlHeader As String = "<?xml version=""1.0"" encoding=""utf-8"" standalone=""yes""?>"
                XDocLocal.LoadXml(XmlHeader & vbCrLf & Instructions)
                XMsgLocal.Run(XDocLocal, StatusLocal)
            Else 'This is not an XMessage!
                Message.XAddText("The message is not an XMessage: " & Instructions & vbCrLf, "Normal")
        End If
    End Sub

    Private _showXMessages As Boolean = True 'If True, XMessages that are sent or received will be shown in the Messages window.
    Property ShowXMessages As Boolean
        Get
            Return _showXMessages
        End Get
        Set(value As Boolean)
            _showXMessages = value
        End Set
    End Property

    Private _showSysMessages As Boolean = True 'If True, System messages that are sent or received will be shown in the messages window.
    Property ShowSysMessages As Boolean
        Get
            Return _showSysMessages
        End Get
        Set(value As Boolean)
            _showSysMessages = value
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

    Private _workflowFileName As String = "" 'The file name of the html document displayed in the Workflow tab.
    Public Property WorkflowFileName As String
        Get
            Return _workflowFileName
        End Get
        Set(value As String)
            _workflowFileName = value
        End Set
    End Property

    'Other selected Database properties:

    Private _sharePriceDbName As String 'The name assigned to the selected Share Price database. This may be different from the file name and is used to identify the database.
    Property SharePriceDbName As String
        Get
            Return _sharePriceDbName
        End Get
        Set(value As String)
            _sharePriceDbName = value
            txtSPDatabaseName.Text = _sharePriceDbName
        End Set
    End Property

    Private _sharePriceDbDescription As String 'A description of the selected Share Price database.
    Property SharePriceDbDescription As String
        Get
            Return _sharePriceDbDescription
        End Get
        Set(value As String)
            _sharePriceDbDescription = value
            txtSPDatabaseDescr.Text = _sharePriceDbDescription
        End Set
    End Property

    Private _sharePriceDbFileName As String 'The selected Share Price database file name.
    Property SharePriceDbFileName As String
        Get
            Return _sharePriceDbFileName
        End Get
        Set(value As String)
            _sharePriceDbFileName = value
            txtSPDatabaseFileName.Text = _sharePriceDbFileName
        End Set
    End Property

    Private _sharePriceDbLocation 'The location type of the Share Price database: Project, Settings, Data, System or External. This allows the file path to be restored if the project is moved.
    Property SharePriceDbLocation As String
        Get
            Return _sharePriceDbLocation
        End Get
        Set(value As String)
            _sharePriceDbLocation = value
            txtSPDatabaseLocn.Text = _sharePriceDbLocation
        End Set
    End Property

    Private _financialsDbName As String 'The name assigned to the selected Financials database. This may be different from the file name and is used to identify the database.
    Property FinancialsDbName As String
        Get
            Return _financialsDbName
        End Get
        Set(value As String)
            _financialsDbName = value
            txtFinDbName.Text = _financialsDbName
        End Set
    End Property

    Private _financialsDbDescription As String 'A description of the selected Financials database.
    Property FinancialsDbDescription As String
        Get
            Return _financialsDbDescription
        End Get
        Set(value As String)
            _financialsDbDescription = value
            txtFinDbDescr.Text = _financialsDbDescription
        End Set
    End Property

    Private _financialsDbFileName As String 'The selected Financials database file name.
    Property FinancialsDbFileName As String
        Get
            Return _financialsDbFileName
        End Get
        Set(value As String)
            _financialsDbFileName = value
            txtFinDbFileName.Text = _financialsDbFileName
        End Set
    End Property

    Private _financialsDbLocation 'The location type of the Financials database: Project, Settings, Data, System or External. This allows the file path to be restored if the project is moved.
    Property FinancialsDbLocation As String
        Get
            Return _financialsDbLocation
        End Get
        Set(value As String)
            _financialsDbLocation = value
            txtFinDbLocn.Text = _financialsDbLocation
        End Set
    End Property

    Private _calculationsDbName As String 'The name assigned to the selected SCalculations database. This may be different from the file name and is used to identify the database.
    Property CalculationsDbName As String
        Get
            Return _calculationsDbName
        End Get
        Set(value As String)
            _calculationsDbName = value
            txtCalcDbName.Text = _calculationsDbName
        End Set
    End Property

    Private _calculationsDbDescription As String 'A description of the selected Calculations database.
    Property CalculationsDbDescription As String
        Get
            Return _calculationsDbDescription
        End Get
        Set(value As String)
            _calculationsDbDescription = value
            txtCalcDbDescr.Text = _calculationsDbDescription
        End Set
    End Property

    Private _calculationsDbFileName As String 'The selected Calculations database file name.
    Property CalculationsDbFileName As String
        Get
            Return _calculationsDbFileName
        End Get
        Set(value As String)
            _calculationsDbFileName = value
            txtCalcDbFileName.Text = _calculationsDbFileName
        End Set
    End Property

    Private _calculationsDbLocation 'The location type of the SCalculations database: Project, Settings, Data, System or External. This allows the file path to be restored if the project is moved.
    Property CalculationsDbLocation As String
        Get
            Return _calculationsDbLocation
        End Get
        Set(value As String)
            _calculationsDbLocation = value
            txtCalcDbLocn.Text = _calculationsDbLocation
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
                               <AdvlNetworkAppPath><%= AdvlNetworkAppPath %></AdvlNetworkAppPath>
                               <AdvlNetworkExePath><%= AdvlNetworkExePath %></AdvlNetworkExePath>
                               <ShowXMessages><%= ShowXMessages %></ShowXMessages>
                               <ShowSysMessages><%= ShowSysMessages %></ShowSysMessages>
                               <!---->
                               <SelectedMainTabIndex><%= TabControl1.SelectedIndex %></SelectedMainTabIndex>
                               <SelectedViewDataTabIndex><%= TabControl3.SelectedIndex %></SelectedViewDataTabIndex>
                               <!---->
                               <SharePriceDbName><%= SharePriceDbName %></SharePriceDbName>
                               <SharePriceDbDescription><%= SharePriceDbDescription %></SharePriceDbDescription>
                               <SharePriceDbFileName><%= SharePriceDbFileName %></SharePriceDbFileName>
                               <SharePriceDbLocation><%= SharePriceDbLocation %></SharePriceDbLocation>
                               <SharePriceDbPath><%= SharePriceDbPath %></SharePriceDbPath>
                               <SharePriceDataViewList><%= SharePriceDataViewList %></SharePriceDataViewList>
                               <!---->
                               <FinancialsDbName><%= FinancialsDbName %></FinancialsDbName>
                               <FinancialsDbDescription><%= FinancialsDbDescription %></FinancialsDbDescription>
                               <FinancialsDbFileName><%= FinancialsDbFileName %></FinancialsDbFileName>
                               <FinancialsDbLocation><%= FinancialsDbLocation %></FinancialsDbLocation>
                               <FinancialsDbPath><%= FinancialsDbPath %></FinancialsDbPath>
                               <FinancialsDataViewList><%= FinancialsDataViewList %></FinancialsDataViewList>
                               <!---->
                               <CalculationsDbName><%= CalculationsDbName %></CalculationsDbName>
                               <CalculationsDbDescription><%= CalculationsDbDescription %></CalculationsDbDescription>
                               <CalculationsDbFileName><%= CalculationsDbFileName %></CalculationsDbFileName>
                               <CalculationsDbLocation><%= CalculationsDbLocation %></CalculationsDbLocation>
                               <CalculationsDbPath><%= CalculationsDbPath %></CalculationsDbPath>
                               <CalculationsDataViewList><%= CalculationsDataViewList %></CalculationsDataViewList>
                               <!---->
                               <NewsDbPath><%= NewsDbPath %></NewsDbPath>
                               <NewsDataViewList><%= NewsDataViewList %></NewsDataViewList>
                               <!---->
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
                               <SharePriceChartProjectNo><%= SelShareChartProjNo %></SharePriceChartProjectNo>
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
                               <SPChartTitleColor><%= txtChartTitle.ForeColor.ToArgb.ToString %></SPChartTitleColor>
                               <SPChartTitleSize><%= txtChartTitle.Font.Size %></SPChartTitleSize>
                               <SPChartTitleBold><%= txtChartTitle.Font.Bold %></SPChartTitleBold>
                               <SPChartTitleItalic><%= txtChartTitle.Font.Italic %></SPChartTitleItalic>
                               <SPChartTitleUnderline><%= txtChartTitle.Font.Underline %></SPChartTitleUnderline>
                               <SPChartTitleStrikeout><%= txtChartTitle.Font.Strikeout %></SPChartTitleStrikeout>
                               <%= If(cmbAlignment.SelectedIndex = -1,
                                   <SPChartTitleAlignment></SPChartTitleAlignment>,
                                   <SPChartTitleAlignment><%= cmbAlignment.SelectedItem.ToString %></SPChartTitleAlignment>) %>
                               <SPChartSettingsFile><%= txtStockChartSettings.Text %></SPChartSettingsFile>
                               <SPChartUseDefaults><%= chkUseStockChartSettingsList.Checked %></SPChartUseDefaults>
                               <SPChartUseDateRange><%= chkSPChartUseDateRange.Checked %></SPChartUseDateRange>
                               <SPChartFromDate><%= dtpSPChartFromDate.Value %></SPChartFromDate>
                               <SPChartToDate><%= dtpSPChartToDate.Value %></SPChartToDate>
                               <!--Cross Plot Charts-->
                               <PointChartProjectNo><%= SelPointChartProjNo %></PointChartProjectNo>
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
                               <CrossPlotTitleColor><%= txtPointChartTitle.ForeColor.ToArgb.ToString %></CrossPlotTitleColor>
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
                               <!--Line Charts-->
                               <LineChartProjectNo><%= SelLineChartProjNo %></LineChartProjectNo>
                               <%= If(cmbLineChartDb.SelectedIndex = -1,
                                   <LineChartDatabase></LineChartDatabase>,
                                   <LineChartDatabase><%= cmbLineChartDb.SelectedItem.ToString %></LineChartDatabase>) %>
                               <LineChartQuery><%= txtLineChartQuery.Text %></LineChartQuery>
                               <LineChartSeriesName><%= txtLineSeriesName.Text %></LineChartSeriesName>
                               <%= If(cmbLineXValues.SelectedIndex = -1,
                                   <LineChartXValues></LineChartXValues>,
                                   <LineChartXValues><%= cmbLineXValues.SelectedItem.ToString %></LineChartXValues>) %>
                               <%= If(cmbLineYValues.SelectedIndex = -1,
                                   <LineChartYValues></LineChartYValues>,
                                   <LineChartYValues><%= cmbLineYValues.SelectedItem.ToString %></LineChartYValues>) %>
                               <LineChartTitleText><%= txtLineChartTitle.Text %></LineChartTitleText>
                               <LineChartTitleFontName><%= txtLineChartTitle.Font.Name %></LineChartTitleFontName>
                               <LineChartTitleColor><%= txtLineChartTitle.ForeColor.ToArgb.ToString %></LineChartTitleColor>
                               <LineChartTitleSize><%= txtLineChartTitle.Font.Size %></LineChartTitleSize>
                               <LineChartTitleBold><%= txtLineChartTitle.Font.Bold %></LineChartTitleBold>
                               <LineChartTitleItalic><%= txtLineChartTitle.Font.Italic %></LineChartTitleItalic>
                               <LineChartTitleUnderline><%= txtLineChartTitle.Font.Underline %></LineChartTitleUnderline>
                               <LineChartTitleStrikeout><%= txtLineChartTitle.Font.Strikeout %></LineChartTitleStrikeout>
                               <%= If(cmbLineChartAlignment.SelectedIndex = -1,
                                   <LineChartTitleAlignment></LineChartTitleAlignment>,
                                   <LineChartTitleAlignment><%= cmbLineChartAlignment.SelectedItem.ToString %></LineChartTitleAlignment>) %>
                               <LineChartSettingsFile><%= txtLineChartSettings.Text %></LineChartSettingsFile>
                               <LineChartUseDefaults><%= chkUseLineChartDefaults.Checked %></LineChartUseDefaults>
                               <LineChartAutoXRange><%= chkLineAutoXRange.Checked %></LineChartAutoXRange>
                               <LineChartXMin><%= txtLineXMin.Text %></LineChartXMin>
                               <LineChartXMax><%= txtLineXMax.Text %></LineChartXMax>
                               <LineChartAutoYRange><%= chkLineAutoYRange.Checked %></LineChartAutoYRange>
                               <LineChartYMin><%= txtLineYMin.Text %></LineChartYMin>
                               <LineChartYMax><%= txtLineYMax.Text %></LineChartYMax>
                           </FormSettings>

        '<SPChartTitleColor><%= txtChartTitle.ForeColor %></SPChartTitleColor>
        '<CrossPlotTitleColor><%= txtPointChartTitle.ForeColor %></CrossPlotTitleColor>
        '<LineChartTitleColor><%= txtLineChartTitle.ForeColor %></LineChartTitleColor>


        '<MsgServiceAppPath><%= MsgServiceAppPath %></MsgServiceAppPath>
        '<MsgServiceExePath><%= MsgServiceExePath %></MsgServiceExePath>

        'Dim SettingsFileName As String = "FormSettings_" & ApplicationInfo.Name & "_" & Me.Text & ".xml"
        Dim SettingsFileName As String = "FormSettings_" & ApplicationInfo.Name & " - Main.xml"
        Debug.Print("Writing settings file: " & SettingsFileName)
        Project.SaveXmlSettings(SettingsFileName, settingsData)

    End Sub

    Private Sub RestoreFormSettings()
        'Read the form settings from an XML document.

        'Dim SettingsFileName As String = "FormSettings_" & ApplicationInfo.Name & "_" & Me.Text & ".xml"
        Dim SettingsFileName As String = "FormSettings_" & ApplicationInfo.Name & " - Main.xml"
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

            'If Settings.<FormSettings>.<MsgServiceAppPath>.Value <> Nothing Then MsgServiceAppPath = Settings.<FormSettings>.<MsgServiceAppPath>.Value
            'If Settings.<FormSettings>.<MsgServiceExePath>.Value <> Nothing Then MsgServiceExePath = Settings.<FormSettings>.<MsgServiceExePath>.Value
            If Settings.<FormSettings>.<AdvlNetworkAppPath>.Value <> Nothing Then AdvlNetworkAppPath = Settings.<FormSettings>.<AdvlNetworkAppPath>.Value
            If Settings.<FormSettings>.<AdvlNetworkExePath>.Value <> Nothing Then AdvlNetworkExePath = Settings.<FormSettings>.<AdvlNetworkExePath>.Value

            If Settings.<FormSettings>.<ShowXMessages>.Value <> Nothing Then ShowXMessages = Settings.<FormSettings>.<ShowXMessages>.Value
            If Settings.<FormSettings>.<ShowSysMessages>.Value <> Nothing Then ShowSysMessages = Settings.<FormSettings>.<ShowSysMessages>.Value

            'Add code to read other saved setting here:

            'Restore Main Tab selection (Project Information - Settings - View Data - Calculations - Charts)
            If Settings.<FormSettings>.<SelectedMainTabIndex>.Value <> Nothing Then TabControl1.SelectedIndex = Settings.<FormSettings>.<SelectedMainTabIndex>.Value

            'Restore View Data Sub Tab selection (Share Prices - Financials - Calculations - News)
            If Settings.<FormSettings>.<SelectedViewDataTabIndex>.Value <> Nothing Then TabControl3.SelectedIndex = Settings.<FormSettings>.<SelectedViewDataTabIndex>.Value

            'Restore View Data - Share Prices settings
            If Settings.<FormSettings>.<SharePriceDbName>.Value <> Nothing Then SharePriceDbName = Settings.<FormSettings>.<SharePriceDbName>.Value
            If Settings.<FormSettings>.<SharePriceDbDescription>.Value <> Nothing Then SharePriceDbDescription = Settings.<FormSettings>.<SharePriceDbDescription>.Value
            If Settings.<FormSettings>.<SharePriceDbFileName>.Value <> Nothing Then SharePriceDbFileName = Settings.<FormSettings>.<SharePriceDbFileName>.Value
            If Settings.<FormSettings>.<SharePriceDbLocation>.Value <> Nothing Then SharePriceDbLocation = Settings.<FormSettings>.<SharePriceDbLocation>.Value
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
            If Settings.<FormSettings>.<FinancialsDbName>.Value <> Nothing Then FinancialsDbName = Settings.<FormSettings>.<FinancialsDbName>.Value
            If Settings.<FormSettings>.<FinancialsDbDescription>.Value <> Nothing Then FinancialsDbDescription = Settings.<FormSettings>.<FinancialsDbDescription>.Value
            If Settings.<FormSettings>.<FinancialsDbFileName>.Value <> Nothing Then FinancialsDbFileName = Settings.<FormSettings>.<FinancialsDbFileName>.Value
            If Settings.<FormSettings>.<FinancialsDbLocation>.Value <> Nothing Then FinancialsDbLocation = Settings.<FormSettings>.<FinancialsDbLocation>.Value
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
            If Settings.<FormSettings>.<CalculationsDbName>.Value <> Nothing Then CalculationsDbName = Settings.<FormSettings>.<CalculationsDbName>.Value
            If Settings.<FormSettings>.<CalculationsDbDescription>.Value <> Nothing Then CalculationsDbDescription = Settings.<FormSettings>.<CalculationsDbDescription>.Value
            If Settings.<FormSettings>.<CalculationsDbFileName>.Value <> Nothing Then CalculationsDbFileName = Settings.<FormSettings>.<CalculationsDbFileName>.Value
            If Settings.<FormSettings>.<CalculationsDbLocation>.Value <> Nothing Then CalculationsDbLocation = Settings.<FormSettings>.<CalculationsDbLocation>.Value
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
            If Settings.<FormSettings>.<SharePriceChartProjectNo>.Value <> Nothing Then SelShareChartProjNo = Settings.<FormSettings>.<SharePriceChartProjectNo>.Value

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

            'If Settings.<FormSettings>.<SPChartTitleColor>.Value <> Nothing Then txtChartTitle.ForeColor = Color.FromName(Settings.<FormSettings>.<SPChartTitleColor>.Value)
            'If Settings.<FormSettings>.<SPChartTitleColor>.Value <> Nothing Then txtChartTitle.ForeColor = Color.FromArgb(Settings.<FormSettings>.<SPChartTitleColor>.Value)
            If Settings.<FormSettings>.<SPChartTitleColor>.Value <> Nothing Then
                If IsNumeric(Settings.<FormSettings>.<SPChartTitleColor>.Value) Then
                    txtChartTitle.ForeColor = Color.FromArgb(Settings.<FormSettings>.<SPChartTitleColor>.Value)
                Else
                    txtChartTitle.ForeColor = Color.FromName(Settings.<FormSettings>.<SPChartTitleColor>.Value)
                End If
            End If

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
                'Project.ReadXmlData(txtStockChartSettings.Text, StockChartDefaults)
                Project.ReadXmlData(txtStockChartSettings.Text, StockChartSettingsList)

                'If StockChartDefaults Is Nothing Then
                If StockChartSettingsList Is Nothing Then

                Else
                    '    rtbStockChartDefaults.Text = StockChartDefaults.ToString
                    'FormatXmlText(rtbStockChartDefaults)
                    'XmlStockChartDefaults.Rtf = XmlStockChartDefaults.XmlToRtf(StockChartDefaults.ToString, True)
                    'XmlStockChartDefaults.Rtf = XmlStockChartDefaults.XmlToRtf(StockChartDefaults.ToString, False)
                    XmlStockChartSettingsList.Rtf = XmlStockChartSettingsList.XmlToRtf(StockChartSettingsList.ToString, False)
                End If
            End If

            If Settings.<FormSettings>.<SPChartUseDefaults>.Value <> Nothing Then chkUseStockChartSettingsList.Checked = Settings.<FormSettings>.<SPChartUseDefaults>.Value
            If Settings.<FormSettings>.<SPChartUseDateRange>.Value <> Nothing Then chkSPChartUseDateRange.Checked = Settings.<FormSettings>.<SPChartUseDateRange>.Value
            If Settings.<FormSettings>.<SPChartFromDate>.Value <> Nothing Then dtpSPChartFromDate.Value = Settings.<FormSettings>.<SPChartFromDate>.Value
            If Settings.<FormSettings>.<SPChartToDate>.Value <> Nothing Then dtpSPChartToDate.Value = Settings.<FormSettings>.<SPChartToDate>.Value
            UpdateSPChartQuery()

            'Cross Plot Charts
            If Settings.<FormSettings>.<PointChartProjectNo>.Value <> Nothing Then SelPointChartProjNo = Settings.<FormSettings>.<PointChartProjectNo>.Value

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

            'If Settings.<FormSettings>.<CrossPlotTitleColor>.Value <> Nothing Then txtPointChartTitle.ForeColor = Color.FromName(Settings.<FormSettings>.<CrossPlotTitleColor>.Value)
            'If Settings.<FormSettings>.<CrossPlotTitleColor>.Value <> Nothing Then txtPointChartTitle.ForeColor = Color.FromArgb(Settings.<FormSettings>.<CrossPlotTitleColor>.Value)
            If Settings.<FormSettings>.<CrossPlotTitleColor>.Value <> Nothing Then
                If IsNumeric(Settings.<FormSettings>.<CrossPlotTitleColor>.Value) Then
                    txtPointChartTitle.ForeColor = Color.FromArgb(Settings.<FormSettings>.<CrossPlotTitleColor>.Value)
                Else
                    txtPointChartTitle.ForeColor = Color.FromName(Settings.<FormSettings>.<CrossPlotTitleColor>.Value)
                End If
            End If

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
                'Project.ReadXmlData(txtPointChartSettings.Text, PointChartDefaults)
                Project.ReadXmlData(txtPointChartSettings.Text, PointChartSettingsList)
                'If PointChartDefaults Is Nothing Then
                If PointChartSettingsList Is Nothing Then

                Else
                    'rtbPointChartDefaults.Text = PointChartSettingsList.ToString
                    'FormatXmlText(rtbPointChartDefaults)

                    XmlPointChartSettingsList.Rtf = XmlPointChartSettingsList.XmlToRtf(PointChartSettingsList.ToString, False)
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

            'Line Charts
            If Settings.<FormSettings>.<LineChartProjectNo>.Value <> Nothing Then SelLineChartProjNo = Settings.<FormSettings>.<LineChartProjectNo>.Value
            If Settings.<FormSettings>.<LineChartDatabase>.Value <> Nothing Then cmbLineChartDb.SelectedIndex = cmbLineChartDb.FindStringExact(Settings.<FormSettings>.<LineChartDatabase>.Value)
            If Settings.<FormSettings>.<LineChartQuery>.Value <> Nothing Then txtLineChartQuery.Text = Settings.<FormSettings>.<LineChartQuery>.Value
            UpdateChartLineTab()
            If Settings.<FormSettings>.<LineChartSeriesName>.Value <> Nothing Then txtLineSeriesName.Text = Settings.<FormSettings>.<LineChartSeriesName>.Value
            If Settings.<FormSettings>.<LineChartXValues>.Value <> Nothing Then cmbLineXValues.SelectedIndex = cmbLineXValues.FindStringExact(Settings.<FormSettings>.<LineChartXValues>.Value)
            If Settings.<FormSettings>.<LineChartYValues>.Value <> Nothing Then cmbLineYValues.SelectedIndex = cmbLineYValues.FindStringExact(Settings.<FormSettings>.<LineChartYValues>.Value)
            If Settings.<FormSettings>.<LineChartTitleText>.Value <> Nothing Then txtLineChartTitle.Text = Settings.<FormSettings>.<LineChartTitleText>.Value
            If Settings.<FormSettings>.<LineChartTitleFontName>.Value <> Nothing Then myFontName = Settings.<FormSettings>.<LineChartTitleFontName>.Value

            'If Settings.<FormSettings>.<LineChartTitleColor>.Value <> Nothing Then txtLineChartTitle.ForeColor = Color.FromArgb(Settings.<FormSettings>.<LineChartTitleColor>.Value)
            If Settings.<FormSettings>.<LineChartTitleColor>.Value <> Nothing Then
                If IsNumeric(Settings.<FormSettings>.<LineChartTitleColor>.Value) Then
                    txtLineChartTitle.ForeColor = Color.FromArgb(Settings.<FormSettings>.<LineChartTitleColor>.Value)
                Else
                    txtLineChartTitle.ForeColor = Color.FromName(Settings.<FormSettings>.<LineChartTitleColor>.Value)
                End If
            End If

            If Settings.<FormSettings>.<LineChartTitleSize>.Value <> Nothing Then myFontSize = Settings.<FormSettings>.<LineChartTitleSize>.Value
            If Settings.<FormSettings>.<LineChartTitleBold>.Value <> Nothing Then
                If Settings.<FormSettings>.<LineChartTitleBold>.Value = "true" Then
                    myFontStyle = myFontStyle Or FontStyle.Bold
                End If
            End If

            If Settings.<FormSettings>.<LineChartTitleItalic>.Value <> Nothing Then
                If Settings.<FormSettings>.<LineChartTitleItalic>.Value = "true" Then
                    myFontStyle = myFontStyle Or FontStyle.Italic
                End If
            End If

            If Settings.<FormSettings>.<LineChartTitleUnderline>.Value <> Nothing Then
                If Settings.<FormSettings>.<LineChartTitleUnderline>.Value = "true" Then
                    myFontStyle = myFontStyle Or FontStyle.Underline
                End If
            End If

            If Settings.<FormSettings>.<LineChartTitleStrikeout>.Value <> Nothing Then
                If Settings.<FormSettings>.<LineChartTitleStrikeout>.Value = "true" Then
                    myFontStyle = myFontStyle Or FontStyle.Strikeout
                End If
            End If

            txtLineChartTitle.Font = New Font(myFontName, myFontSize, myFontStyle)

            If Settings.<FormSettings>.<LineChartTitleAlignment>.Value <> Nothing Then cmbLineChartAlignment.SelectedIndex = cmbLineChartAlignment.FindStringExact(Settings.<FormSettings>.<LineChartTitleAlignment>.Value)

            If Settings.<FormSettings>.<LineChartSettingsFile>.Value <> Nothing Then
                txtLineChartSettings.Text = Settings.<FormSettings>.<LineChartSettingsFile>.Value
                Project.ReadXmlData(txtLineChartSettings.Text, LineChartSettingsList)
                If LineChartSettingsList Is Nothing Then

                Else
                    XmlLineChartSettingsList.Rtf = XmlLineChartSettingsList.XmlToRtf(LineChartSettingsList.ToString, False)
                End If
            End If

            If Settings.<FormSettings>.<LineChartUseDefaults>.Value <> Nothing Then chkUseLineChartDefaults.Checked = Settings.<FormSettings>.<LineChartUseDefaults>.Value

            If Settings.<FormSettings>.<LineChartAutoXRange>.Value <> Nothing Then chkLineAutoXRange.Checked = Settings.<FormSettings>.<LineChartAutoXRange>.Value

            If Settings.<FormSettings>.<LineChartXMin>.Value <> Nothing Then
                txtLineXMin.Text = Settings.<FormSettings>.<LineChartXMin>.Value
            Else
                txtLineXMin.Text = "-100"
            End If

            If Settings.<FormSettings>.<LineChartXMax>.Value <> Nothing Then
                txtLineXMax.Text = Settings.<FormSettings>.<LineChartXMax>.Value
            Else
                txtLineXMax.Text = "100"
            End If

            If Settings.<FormSettings>.<LineChartAutoYRange>.Value <> Nothing Then chkLineAutoYRange.Checked = Settings.<FormSettings>.<LineChartAutoYRange>.Value

            If Settings.<FormSettings>.<LineChartYMin>.Value <> Nothing Then
                txtLineYMin.Text = Settings.<FormSettings>.<LineChartYMin>.Value
            Else
                txtLineYMin.Text = "-100"
            End If

            If Settings.<FormSettings>.<LineChartYMax>.Value <> Nothing Then
                txtLineYMax.Text = Settings.<FormSettings>.<LineChartYMax>.Value
            Else
                txtLineYMax.Text = "100"
            End If



            CheckFormPos()
        Else
            Debug.Print("Settings file not found.")
        End If
    End Sub

    Private Sub CheckFormPos()
        'Check that the form can be seen on a screen.

        Dim MinWidthVisible As Integer = 192 'Minimum number of X pixels visible. The form will be moved if this many form pixels are not visible.
        Dim MinHeightVisible As Integer = 64 'Minimum number of Y pixels visible. The form will be moved if this many form pixels are not visible.

        Dim FormRect As New Rectangle(Me.Left, Me.Top, Me.Width, Me.Height)
        Dim WARect As Rectangle = Screen.GetWorkingArea(FormRect) 'The Working Area rectangle - the usable area of the screen containing the form.

        ''Check if the top of the form is less than zero:
        'If Me.Top < 0 Then Me.Top = 0

        'Check if the top of the form is above the top of the Working Area:
        If Me.Top < WARect.Top Then
            Me.Top = WARect.Top
        End If

        'Check if the top of the form is too close to the bottom of the Working Area:
        If (Me.Top + MinHeightVisible) > (WARect.Top + WARect.Height) Then
            Me.Top = WARect.Top + WARect.Height - MinHeightVisible
        End If

        'Check if the left edge of the form is too close to the right edge of the Working Area:
        If (Me.Left + MinWidthVisible) > (WARect.Left + WARect.Width) Then
            Me.Left = WARect.Left + WARect.Width - MinWidthVisible
        End If

        'Check if the right edge of the form is too close to the left edge of the Working Area:
        If (Me.Left + Me.Width - MinWidthVisible) < WARect.Left Then
            Me.Left = WARect.Left - Me.Width + MinWidthVisible
        End If
    End Sub

    Private Sub ReadApplicationInfo()
        'Read the Application Information.

        If ApplicationInfo.FileExists Then
            ApplicationInfo.ReadFile()
        Else
            'There is no Application_Info_ADVL_2.xml file.
            DefaultAppProperties() 'Create a new Application Info file with default application properties:
            ApplicationInfo.WriteFile() 'Write the file now. The file information may be used by other applications.
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
        ''Get the Application Version Information:
        ApplicationInfo.Version.Major = My.Application.Info.Version.Major
        ApplicationInfo.Version.Minor = My.Application.Info.Version.Minor
        ApplicationInfo.Version.Build = My.Application.Info.Version.Build
        ApplicationInfo.Version.Revision = My.Application.Info.Version.Revision

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
        Project.Application.Name = ApplicationInfo.Name

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
                    If Project.ParentParameterExists("ProNetName") Then
                        Project.AddParameter("ProNetName", Project.ParentParameter("ProNetName").Value, Project.ParentParameter("ProNetName").Description) 'AddParameter will update the parameter if it already exists.
                        ProNetName = Project.Parameter("ProNetName").Value
                    Else
                        ProNetName = Project.GetParameter("ProNetName")
                    End If
                    If Project.ParentParameterExists("ProNetPath") Then 'Get the parent parameter value - it may have been updated.
                        Project.AddParameter("ProNetPath", Project.ParentParameter("ProNetPath").Value, Project.ParentParameter("ProNetPath").Description) 'AddParameter will update the parameter if it already exists.
                        ProNetPath = Project.Parameter("ProNetPath").Value
                    Else
                        ProNetPath = Project.GetParameter("ProNetPath") 'If the parameter does not exist, the value is set to ""
                    End If
                    Project.SaveParameters() 'These should be saved now - child projects look for parent parameters in the parameter file.


                    Project.LockProject() 'Lock the project while it is open in this application.
                    'Set the project start time. This is used to track project usage.
                    Project.Usage.StartTime = Now
                    ApplicationInfo.SettingsLocn = Project.SettingsLocn
                    'Set up the Message object:
                    Message.SettingsLocn = Project.SettingsLocn
                    Message.Show()
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
                If Project.ParentParameterExists("ProNetName") Then
                    Project.AddParameter("ProNetName", Project.ParentParameter("ProNetName").Value, Project.ParentParameter("ProNetName").Description) 'AddParameter will update the parameter if it already exists.
                    ProNetName = Project.Parameter("ProNetName").Value
                Else
                    ProNetName = Project.GetParameter("ProNetName")
                End If
                If Project.ParentParameterExists("ProNetPath") Then 'Get the parent parameter value - it may have been updated.
                    Project.AddParameter("ProNetPath", Project.ParentParameter("ProNetPath").Value, Project.ParentParameter("ProNetPath").Description) 'AddParameter will update the parameter if it already exists.
                    ProNetPath = Project.Parameter("ProNetPath").Value
                Else
                    ProNetPath = Project.GetParameter("ProNetPath") 'If the parameter does not exist, the value is set to ""
                End If
                Project.SaveParameters() 'These should be saved now - child projects look for parent parameters in the parameter file.

                Project.LockProject() 'Lock the project while it is open in this application.
                'Set the project start time. This is used to track project usage.
                Project.Usage.StartTime = Now
                ApplicationInfo.SettingsLocn = Project.SettingsLocn
                'Set up the Message object:
                Message.SettingsLocn = Project.SettingsLocn
                Message.Show() 'Added 18May19
            End If
        Else 'Project has been opened using Command Line arguments.

            Project.ReadParameters()
            Project.ReadParentParameters()
            If Project.ParentParameterExists("ProNetName") Then
                Project.AddParameter("ProNetName", Project.ParentParameter("ProNetName").Value, Project.ParentParameter("ProNetName").Description) 'AddParameter will update the parameter if it already exists.
                ProNetName = Project.Parameter("ProNetName").Value
            Else
                ProNetName = Project.GetParameter("ProNetName")
            End If
            If Project.ParentParameterExists("ProNetPath") Then 'Get the parent parameter value - it may have been updated.
                Project.AddParameter("ProNetPath", Project.ParentParameter("ProNetPath").Value, Project.ParentParameter("ProNetPath").Description) 'AddParameter will update the parameter if it already exists.
                ProNetPath = Project.Parameter("ProNetPath").Value
            Else
                ProNetPath = Project.GetParameter("ProNetPath") 'If the parameter does not exist, the value is set to ""
            End If
            Project.SaveParameters() 'These should be saved now - child projects look for parent parameters in the parameter file.

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
        SetupChartLinePlotTab()

        'Set up the Share Price Database data grid view:
        Dim TextBoxCol0 As New DataGridViewTextBoxColumn
        dgvSPDatabase.Columns.Add(TextBoxCol0)
        dgvSPDatabase.Columns(0).HeaderText = "Name"
        dgvSPDatabase.Columns(0).Width = 240
        Dim TextBoxCol1 As New DataGridViewTextBoxColumn
        dgvSPDatabase.Columns.Add(TextBoxCol1)
        dgvSPDatabase.Columns(1).HeaderText = "Description"
        dgvSPDatabase.Columns(1).Width = 280
        dgvSPDatabase.Columns(1).DefaultCellStyle.WrapMode = DataGridViewTriState.True
        Dim TextBoxCol2 As New DataGridViewTextBoxColumn
        dgvSPDatabase.Columns.Add(TextBoxCol2)
        dgvSPDatabase.Columns(2).HeaderText = "File Name"
        dgvSPDatabase.Columns(2).Width = 240
        Dim TextBoxCol3 As New DataGridViewTextBoxColumn
        dgvSPDatabase.Columns.Add(TextBoxCol3)
        dgvSPDatabase.Columns(3).HeaderText = "Location" 'External, Project, Settings, Data, System
        dgvSPDatabase.Columns(3).Width = 60
        Dim TextBoxCol4 As New DataGridViewTextBoxColumn
        dgvSPDatabase.Columns.Add(TextBoxCol4)
        dgvSPDatabase.Columns(4).HeaderText = "Path"
        dgvSPDatabase.Columns(4).Width = 280
        dgvSPDatabase.Columns(4).DefaultCellStyle.WrapMode = DataGridViewTriState.True

        dgvSPDatabase.AllowUserToAddRows = False
        dgvSPDatabase.SelectionMode = DataGridViewSelectionMode.FullRowSelect

        dgvSPDatabase.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
        dgvSPDatabase.AutoResizeRows()

        'Set up the Financials Database data grid view:
        Dim TextBoxCol10 As New DataGridViewTextBoxColumn
        dgvFinDatabase.Columns.Add(TextBoxCol10)
        dgvFinDatabase.Columns(0).HeaderText = "Name"
        dgvFinDatabase.Columns(0).Width = 240
        Dim TextBoxCol11 As New DataGridViewTextBoxColumn
        dgvFinDatabase.Columns.Add(TextBoxCol11)
        dgvFinDatabase.Columns(1).HeaderText = "Description"
        dgvFinDatabase.Columns(1).Width = 280
        dgvFinDatabase.Columns(1).DefaultCellStyle.WrapMode = DataGridViewTriState.True
        Dim TextBoxCol12 As New DataGridViewTextBoxColumn
        dgvFinDatabase.Columns.Add(TextBoxCol12)
        dgvFinDatabase.Columns(2).HeaderText = "File Name"
        dgvFinDatabase.Columns(2).Width = 240
        Dim TextBoxCol13 As New DataGridViewTextBoxColumn
        dgvFinDatabase.Columns.Add(TextBoxCol13)
        dgvFinDatabase.Columns(3).HeaderText = "Location" 'External, Project, Settings, Data, System
        dgvFinDatabase.Columns(3).Width = 60
        Dim TextBoxCol14 As New DataGridViewTextBoxColumn
        dgvFinDatabase.Columns.Add(TextBoxCol14)
        dgvFinDatabase.Columns(4).HeaderText = "Path"
        dgvFinDatabase.Columns(4).Width = 280
        dgvFinDatabase.Columns(4).DefaultCellStyle.WrapMode = DataGridViewTriState.True

        dgvFinDatabase.AllowUserToAddRows = False
        dgvFinDatabase.SelectionMode = DataGridViewSelectionMode.FullRowSelect

        dgvFinDatabase.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
        dgvFinDatabase.AutoResizeRows()

        'Set up the Calculations Database data grid view:
        Dim TextBoxCol20 As New DataGridViewTextBoxColumn
        dgvCalcDatabase.Columns.Add(TextBoxCol20)
        dgvCalcDatabase.Columns(0).HeaderText = "Name"
        dgvCalcDatabase.Columns(0).Width = 240
        Dim TextBoxCol21 As New DataGridViewTextBoxColumn
        dgvCalcDatabase.Columns.Add(TextBoxCol21)
        dgvCalcDatabase.Columns(1).HeaderText = "Description"
        dgvCalcDatabase.Columns(1).Width = 280
        dgvCalcDatabase.Columns(1).DefaultCellStyle.WrapMode = DataGridViewTriState.True
        Dim TextBoxCol22 As New DataGridViewTextBoxColumn
        dgvCalcDatabase.Columns.Add(TextBoxCol22)
        dgvCalcDatabase.Columns(2).HeaderText = "File Name"
        dgvCalcDatabase.Columns(2).Width = 240
        Dim TextBoxCol23 As New DataGridViewTextBoxColumn
        dgvCalcDatabase.Columns.Add(TextBoxCol23)
        dgvCalcDatabase.Columns(3).HeaderText = "Location" 'External, Project, Settings, Data, System
        dgvCalcDatabase.Columns(3).Width = 60
        Dim TextBoxCol24 As New DataGridViewTextBoxColumn
        dgvCalcDatabase.Columns.Add(TextBoxCol24)
        dgvCalcDatabase.Columns(4).HeaderText = "Path"
        dgvCalcDatabase.Columns(4).Width = 280
        dgvCalcDatabase.Columns(4).DefaultCellStyle.WrapMode = DataGridViewTriState.True

        dgvCalcDatabase.AllowUserToAddRows = False
        dgvCalcDatabase.SelectionMode = DataGridViewSelectionMode.FullRowSelect

        dgvCalcDatabase.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
        dgvCalcDatabase.AutoResizeRows()

        RestoreDatabaseList() 'Restore the lists of Share Price, Financials and Calculations databases
        ReadProjectList()
        'UpdateChartProjLists() 'Wait until RestoreFormSettings() has run!

        bgwSendMessage.WorkerReportsProgress = True
        bgwSendMessage.WorkerSupportsCancellation = True

        bgwSendMessageAlt.WorkerReportsProgress = True
        bgwSendMessageAlt.WorkerSupportsCancellation = True

        bgwRunInstruction.WorkerReportsProgress = True
        bgwRunInstruction.WorkerSupportsCancellation = True

        Me.WebBrowser1.ObjectForScripting = Me

        XmlStockChartSettingsList.WordWrap = False
        XmlStockChartSettingsList.Settings.ClearAllTextTypes()
        'Default message display settings:
        XmlStockChartSettingsList.Settings.AddNewTextType("Warning")
        XmlStockChartSettingsList.Settings.TextType("Warning").FontName = "Arial"
        XmlStockChartSettingsList.Settings.TextType("Warning").Bold = True
        XmlStockChartSettingsList.Settings.TextType("Warning").Color = Color.Red
        XmlStockChartSettingsList.Settings.TextType("Warning").PointSize = 12

        XmlStockChartSettingsList.Settings.AddNewTextType("Default")
        XmlStockChartSettingsList.Settings.TextType("Default").FontName = "Arial"
        XmlStockChartSettingsList.Settings.TextType("Default").Bold = False
        XmlStockChartSettingsList.Settings.TextType("Default").Color = Color.Black
        XmlStockChartSettingsList.Settings.TextType("Default").PointSize = 10

        'XML formatting adjustments:
        XmlStockChartSettingsList.Settings.IndentSpaces = 4
        XmlStockChartSettingsList.Settings.Value.Bold = True
        XmlStockChartSettingsList.Settings.Comment.Color = System.Drawing.Color.Gray

        XmlStockChartSettingsList.Settings.UpdateFontIndexes()
        XmlStockChartSettingsList.Settings.UpdateColorIndexes()

        XmlPointChartSettingsList.WordWrap = False
        XmlPointChartSettingsList.Settings.ClearAllTextTypes()
        'Default message display settings:
        XmlPointChartSettingsList.Settings.AddNewTextType("Warning")
        XmlPointChartSettingsList.Settings.TextType("Warning").FontName = "Arial"
        XmlPointChartSettingsList.Settings.TextType("Warning").Bold = True
        XmlPointChartSettingsList.Settings.TextType("Warning").Color = Color.Red
        XmlPointChartSettingsList.Settings.TextType("Warning").PointSize = 12

        XmlPointChartSettingsList.Settings.AddNewTextType("Default")
        XmlPointChartSettingsList.Settings.TextType("Default").FontName = "Arial"
        XmlPointChartSettingsList.Settings.TextType("Default").Bold = False
        XmlPointChartSettingsList.Settings.TextType("Default").Color = Color.Black
        XmlPointChartSettingsList.Settings.TextType("Default").PointSize = 10

        'XML formatting adjustments:
        XmlPointChartSettingsList.Settings.IndentSpaces = 4
        XmlPointChartSettingsList.Settings.Value.Bold = True
        XmlPointChartSettingsList.Settings.Comment.Color = System.Drawing.Color.Gray

        XmlPointChartSettingsList.Settings.UpdateFontIndexes()
        XmlPointChartSettingsList.Settings.UpdateColorIndexes()

        XmlLineChartSettingsList.WordWrap = False
        XmlLineChartSettingsList.Settings.ClearAllTextTypes()
        'Default message display settings:
        XmlLineChartSettingsList.Settings.AddNewTextType("Warning")
        XmlLineChartSettingsList.Settings.TextType("Warning").FontName = "Arial"
        XmlLineChartSettingsList.Settings.TextType("Warning").Bold = True
        XmlLineChartSettingsList.Settings.TextType("Warning").Color = Color.Red
        XmlLineChartSettingsList.Settings.TextType("Warning").PointSize = 12

        XmlLineChartSettingsList.Settings.AddNewTextType("Default")
        XmlLineChartSettingsList.Settings.TextType("Default").FontName = "Arial"
        XmlLineChartSettingsList.Settings.TextType("Default").Bold = False
        XmlLineChartSettingsList.Settings.TextType("Default").Color = Color.Black
        XmlLineChartSettingsList.Settings.TextType("Default").PointSize = 10

        'XML formatting adjustments:
        XmlLineChartSettingsList.Settings.IndentSpaces = 4
        XmlLineChartSettingsList.Settings.Value.Bold = True
        XmlLineChartSettingsList.Settings.Comment.Color = System.Drawing.Color.Gray

        XmlLineChartSettingsList.Settings.UpdateFontIndexes()
        XmlLineChartSettingsList.Settings.UpdateColorIndexes()

        InitialiseForm() 'Initialise the form for a new project.

        'END Initialise the form: ------------------------------------------------------------------------------------

        RestoreFormSettings() 'Restore the form settings
        Message.ShowXMessages = ShowXMessages
        Message.ShowSysMessages = ShowSysMessages
        RestoreProjectSettings() 'Restore the Project settings
        UpdateChartProjLists()

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

    End Sub

    Private Sub InitialiseForm()
        'Initialise the form for a new project.
        OpenStartPage()
    End Sub

    Private Sub ShowProjectInfo()
        'Show the project information:

        txtParentProject.Text = Project.ParentProjectName
        txtProNetName.Text = Project.GetParameter("ProNetName")

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

        SaveDatabaseList() 'Save the lists of Share Price, Financials and Calculations databases

        WriteProjectList() 'Save the list of Projects on the Network - including Charting projects.

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
        Message.ShowXMessages = ShowXMessages
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
            'SharePricesFormList(0).DataSummary = "New Share Price Data View"
            SharePricesFormList(0).DataName = "New Share Price Data View"
            SharePricesFormList(0).Version = "Version 1"
        Else
            Dim I As Integer
            Dim FormAdded As Boolean = False
            For I = 0 To SharePricesFormList.Count - 1 'Check if there are closed forms in SharePricesList. They can be re-used.
                If IsNothing(SharePricesFormList(I)) Then
                    SharePricesFormList(I) = SharePrices
                    SharePricesFormList(I).FormNo = I
                    SharePricesFormList(I).Show()
                    'SharePricesFormList(I).DataSummary = "New Share Price Data View"
                    SharePricesFormList(I).DataName = "New Share Price Data View"
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
                'SharePricesFormList(FormNo).DataSummary = "New Share Price Data View"
                SharePricesFormList(FormNo).DataName = "New Share Price Data View"
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

        If SharePricesFormNo = ClosedFormNo Then 'The current Share Prices form number has been closed.
            SharePricesFormNo = -1
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
            'FinancialsFormList(0).DataSummary = "New Financials Data View"
            FinancialsFormList(0).DataName = "New Financials Data View"
            FinancialsFormList(0).Version = "Version 1"
        Else
            Dim I As Integer
            Dim FormAdded As Boolean = False
            For I = 0 To FinancialsFormList.Count - 1 'Check if there are closed forms in FinancialsFormList. They can be re-used.
                If IsNothing(FinancialsFormList(I)) Then
                    FinancialsFormList(I) = Financials
                    FinancialsFormList(I).FormNo = I
                    FinancialsFormList(I).Show()
                    'FinancialsFormList(I).DataSummary = "New Financials Data View"
                    FinancialsFormList(I).DataName = "New Financials Data View"
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
                'FinancialsFormList(FormNo).DataSummary = "New Financials Data View"
                FinancialsFormList(FormNo).DataName = "New Financials Data View"
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

        'If FinancialsFormNo = ClosedFormNo Then 'The current Financials form number has been closed.
        '    FinancialsFormNo = -1
        'End If

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
            'CalculationsFormList(0).DataSummary = "New Calculations Data View"
            CalculationsFormList(0).DataName = "New Calculations Data View"
            CalculationsFormList(0).Version = "Version 1"
        Else
            Dim I As Integer
            Dim FormAdded As Boolean = False
            For I = 0 To CalculationsFormList.Count - 1 'Check if there are closed forms in CalculationsFormList. They can be re-used.
                If IsNothing(CalculationsFormList(I)) Then
                    CalculationsFormList(I) = Calculations
                    CalculationsFormList(I).FormNo = I
                    CalculationsFormList(I).Show()
                    'CalculationsFormList(I).DataSummary = "New Calculations Data View"
                    CalculationsFormList(I).DataName = "New Calculations Data View"
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
                'CalculationsFormList(FormNo).DataSummary = "New Calculations Data View"
                CalculationsFormList(FormNo).DataName = "New Calculations Data View"
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

        If CalculationsFormNo = ClosedFormNo Then 'The current Calculations form number has been closed.
            CalculationsFormNo = -1
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

    Private Sub btnDesignLineChartQuery_Click(sender As Object, e As EventArgs) Handles btnDesignLineChartQuery.Click
        'Open the Design Query form:

        If IsNothing(DesignLineChartQuery) Then
            DesignLineChartQuery = New frmDesignQuery
            DesignLineChartQuery.Text = "Design Line Chart Data Query"
            DesignLineChartQuery.Show()
            DesignLineChartQuery.DatabasePath = txtLineChartDbPath.Text
        Else
            DesignLineChartQuery.Show()
        End If
    End Sub

    Private Sub DesignLineChartQuery_FormClosed(sender As Object, e As FormClosedEventArgs) Handles DesignLineChartQuery.FormClosed
        DesignLineChartQuery = Nothing
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

    Public Sub HTMLDisplayPageFormClosed()
        'This subroutine is called when the Web Page form has been closed.
        'The subroutine is usually called from the FormClosed event of the WebPage form.
        'The WebPage form may have multiple instances.
        'The ClosedFormNumber property should contains the number of the instance of the WebPage form.
        'This property should be updated by the WebPage form when it is being closed.
        'The ClosedFormNumber property value is used to determine which element in WebPageList should be set to Nothing.

        If HtmlDisplayFormList.Count < ClosedFormNo + 1 Then
            'ClosedFormNo is too large to exist in WebPageFormList
            Exit Sub
        End If

        If IsNothing(HtmlDisplayFormList(ClosedFormNo)) Then
            'The form is already set to nothing
        Else
            HtmlDisplayFormList(ClosedFormNo) = Nothing
        End If
    End Sub




#End Region 'Open and Close Forms -------------------------------------------------------------------------------------------------------------------------------------------------------------

#Region " Form Methods - The main actions performed by this form." '---------------------------------------------------------------------------------------------------------------------------

    Private Sub SaveDatabaseList()
        'Save the lists of Share Price, Financials and Calculations databases:

        Dim XDoc = <?xml version="1.0" encoding="utf-8"?>
                   <Databases>
                       <SharePriceDatabaseList>
                           <!---->
                           <!--List of Share Price databases-->
                           <%= From item In dgvSPDatabase.Rows
                               Select
                               <Database>
                                   <Name><%= item.Cells(0).Value %></Name>
                                   <Description><%= item.Cells(1).Value %></Description>
                                   <FileName><%= item.Cells(2).Value %></FileName>
                                   <Location><%= item.Cells(3).Value %></Location>
                                   <Path><%= item.Cells(4).Value %></Path>
                               </Database> %>
                       </SharePriceDatabaseList>
                       <!---->
                       <FinancialsDatabaseList>
                           <!---->
                           <!--List of Financials databases-->
                           <%= From item In dgvFinDatabase.Rows
                               Select
                               <Database>
                                   <Name><%= item.Cells(0).Value %></Name>
                                   <Description><%= item.Cells(1).Value %></Description>
                                   <FileName><%= item.Cells(2).Value %></FileName>
                                   <Location><%= item.Cells(3).Value %></Location>
                                   <Path><%= item.Cells(4).Value %></Path>
                               </Database> %>
                       </FinancialsDatabaseList>
                       <!---->
                       <CalculationsDatabaseList>
                           <!---->
                           <!--List of Calculations databases-->
                           <%= From item In dgvCalcDatabase.Rows
                               Select
                               <Database>
                                   <Name><%= item.Cells(0).Value %></Name>
                                   <Description><%= item.Cells(1).Value %></Description>
                                   <FileName><%= item.Cells(2).Value %></FileName>
                                   <Location><%= item.Cells(3).Value %></Location>
                                   <Path><%= item.Cells(4).Value %></Path>
                               </Database> %>
                       </CalculationsDatabaseList>
                   </Databases>
        Project.SaveXmlData("DatabaseList.xml", XDoc)

    End Sub

    Private Sub RestoreDatabaseList()
        'Restore the lists of Share Price, Financials and Calculations databases:

        dgvSPDatabase.Rows.Clear()

        If Project.DataFileExists("DatabaseList.xml") Then
            Dim XDoc As System.Xml.Linq.XDocument
            Project.ReadXmlData("DatabaseList.xml", XDoc)

            'Restore list of Share Price databases:
            Dim SPDatabases = From item In XDoc.<Databases>.<SharePriceDatabaseList>.<Database>
            For Each item In SPDatabases
                dgvSPDatabase.Rows.Add(item.<Name>.Value, item.<Description>.Value, item.<FileName>.Value, item.<Location>.Value, item.<Path>.Value)
            Next

            'Restore list of Financials databases:
            Dim FinDatabases = From item In XDoc.<Databases>.<FinancialsDatabaseList>.<Database>
            For Each item In FinDatabases
                dgvFinDatabase.Rows.Add(item.<Name>.Value, item.<Description>.Value, item.<FileName>.Value, item.<Location>.Value, item.<Path>.Value)
            Next

            'Restore list of Calculations databases:
            Dim CalcDatabases = From item In XDoc.<Databases>.<CalculationsDatabaseList>.<Database>
            For Each item In CalcDatabases
                dgvCalcDatabase.Rows.Add(item.<Name>.Value, item.<Description>.Value, item.<FileName>.Value, item.<Location>.Value, item.<Path>.Value)
            Next
        Else
            Message.AddWarning("Database list file, DatabaseList.xml not found." & vbCrLf)
        End If

    End Sub

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
                    Message.XAddText("Message sent to " & "Message Service" & ":" & vbCrLf, "XmlSentNotice")
                    Message.XAddXml(doc.ToString)
                    Message.XAddText(vbCrLf, "Normal") 'Add extra line
                    client.SendMessage("", "MessageService", doc.ToString)
                End If
            End If
        End If
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
        'Open the StartPage.html file and display in the Workflow tab.

        If Project.DataFileExists("StartPage.html") Then
            WorkflowFileName = "StartPage.html"
            DisplayWorkflow()
        Else
            CreateStartPage()
            WorkflowFileName = "StartPage.html"
            DisplayWorkflow()
        End If
    End Sub

    Public Sub DisplayWorkflow()
        'Display the StartPage.html file in the Start Page tab.

        If Project.DataFileExists(WorkflowFileName) Then
            Dim rtbData As New IO.MemoryStream
            Project.ReadData(WorkflowFileName, rtbData)
            rtbData.Position = 0
            Dim sr As New IO.StreamReader(rtbData)
            WebBrowser1.DocumentText = sr.ReadToEnd()
        Else
            Message.AddWarning("Web page file not found: " & WorkflowFileName & vbCrLf)
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
        sb.Append("< html > " & vbCrLf)
        sb.Append(" < head > " & vbCrLf)
        sb.Append(" < title > " & DocumentTitle & "</title>" & vbCrLf)
        sb.Append("<meta name=""description"" content=""Application information."">" & vbCrLf)
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

        'Add the Start Up code section.
        sb.Append("//Code to execute on Start Up:" & vbCrLf)
        sb.Append("function StartUpCode() {" & vbCrLf)
        sb.Append("  RestoreSettings() ;" & vbCrLf)
        sb.Append("}" & vbCrLf & vbCrLf)

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
        sb.Append(vbCrLf)

        'sb.Append(vbCrLf)
        'sb.Append("  case ""Status"" :" & vbCrLf)
        'sb.Append("    if (Info = ""OK"") { " & vbCrLf)
        'sb.Append("      //Instruction processing completed OK:" & vbCrLf)
        'sb.Append("      } else {" & vbCrLf)
        'sb.Append("      window.external.AddWarning(""Error: Unknown Status information: "" + "" Info: "" + Info + ""\r\n"") ;" & vbCrLf)
        'sb.Append("     }" & vbCrLf)
        'sb.Append("    break ;" & vbCrLf)
        'sb.Append(vbCrLf)

        'sb.Append("  case ""OnCompletion"" :" & vbCrLf)
        sb.Append("  case ""EndInstruction"" :" & vbCrLf)
        sb.Append("    switch(Info) {" & vbCrLf)
        sb.Append("      case ""Stop"" :" & vbCrLf)
        sb.Append("        //Do nothing." & vbCrLf)
        sb.Append("        break ;" & vbCrLf)
        sb.Append(vbCrLf)
        sb.Append("      default:" & vbCrLf)
        'sb.Append("        window.external.AddWarning(""Error: Unknown OnCompletion information:  "" + "" Info: "" + Info + ""\r\n"") ;" & vbCrLf)
        sb.Append("        window.external.AddWarning(""Error: Unknown EndInstruction information:  "" + "" Info: "" + Info + ""\r\n"") ;" & vbCrLf)
        sb.Append("        break ;" & vbCrLf)
        sb.Append("    }" & vbCrLf)
        sb.Append("    break ;" & vbCrLf)
        sb.Append(vbCrLf)

        sb.Append("  case ""Status"" :" & vbCrLf)
        sb.Append("    switch(Info) {" & vbCrLf)
        sb.Append("      case ""OK"" :" & vbCrLf)
        sb.Append("        //Instruction processing completed OK." & vbCrLf)
        sb.Append("        break ;" & vbCrLf)
        sb.Append(vbCrLf)
        sb.Append("      default:" & vbCrLf)
        sb.Append("        window.external.AddWarning(""Error: Unknown Status information:  "" + "" Info: "" + Info + ""\r\n"") ;" & vbCrLf)
        sb.Append("        break ;" & vbCrLf)
        sb.Append("    }" & vbCrLf)
        sb.Append("    break ;" & vbCrLf)
        sb.Append(vbCrLf)

        sb.Append("  default:" & vbCrLf)
        'sb.Append("    window.external.AddWarning(""Unknown location: "" + Locn + ""\r\n"") ;" & vbCrLf)
        sb.Append("    window.external.AddWarning(""Workflow WebPage XMessage: "" + ""\r\n"" + ""Unknown location: "" + Locn + ""\r\n"") ;" & vbCrLf)
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
        'sb.Append("window.onload = RestoreSettings; " & vbCrLf)
        sb.Append("window.onload = StartUpCode ; " & vbCrLf)
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
        sb.Append("<!-- Andorville(TM) Workflow File -->" & vbCrLf)
        sb.Append("<!-- Application Name:    " & ApplicationInfo.Name & " -->" & vbCrLf)
        sb.Append("<!-- Application Version: " & My.Application.Info.Version.ToString & " -->" & vbCrLf)
        sb.Append("<!-- Creation Date:          " & Format(Now, "dd MMMM yyyy") & " -->" & vbCrLf)
        sb.Append("<head>" & vbCrLf)
        sb.Append("<title>" & DocumentTitle & "</title>" & vbCrLf)
        sb.Append("<meta name=""description"" content=""Workflow description."">" & vbCrLf)
        sb.Append("</head>" & vbCrLf)

        sb.Append("<body style=""font-family:arial;"">" & vbCrLf & vbCrLf)

        sb.Append("<h2>" & DocumentTitle & "</h2>" & vbCrLf & vbCrLf)

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


    'Display Messages ==============================================================================================

    Public Sub AddMessage(ByVal Msg As String)
        'Add a normal text message to the Message window.
        Message.Add(Msg)
    End Sub

    Public Sub AddWarning(ByVal Msg As String)
        'Add a warning text message to the Message window.
        Message.AddWarning(Msg)
    End Sub

    Public Sub AddTextTypeMessage(ByVal Msg As String, ByVal TextType As String)
        'Add a message with the specified Text Type to the Message window.
        Message.AddText(Msg, TextType)
    End Sub

    Public Sub AddXmlMessage(ByVal XmlText As String)
        'Add an Xml message to the Message window.
        Message.AddXml(XmlText)
    End Sub

    'END Display Messages ------------------------------------------------------------------------------------------


    'Run an XSequence ==============================================================================================

    Public Sub RunClipboardXSeq()
        'Run the XSequence instructions in the clipboard.

        Dim XDocSeq As System.Xml.Linq.XDocument
        Try
            XDocSeq = XDocument.Parse(My.Computer.Clipboard.GetText)
        Catch ex As Exception
            Message.AddWarning("Error reading Clipboard data. " & ex.Message & vbCrLf)
            Exit Sub
        End Try

        If IsNothing(XDocSeq) Then
            Message.Add("No XSequence instructions were found in the clipboard.")
        Else
            Dim XmlSeq As New System.Xml.XmlDocument
            Try
                XmlSeq.LoadXml(XDocSeq.ToString) 'Convert XDocSeq to an XmlDocument to process with XSeq.
                'Run the sequence:
                XSeq.RunXSequence(XmlSeq, Status)
            Catch ex As Exception
                Message.AddWarning("Error restoring HTML settings. " & ex.Message & vbCrLf)
            End Try
        End If
    End Sub

    Public Sub RunXSequence(ByVal XSequence As String)
        'Run the XMSequence
        Dim XmlSeq As New System.Xml.XmlDocument
        XmlSeq.LoadXml(XSequence)
        XSeq.RunXSequence(XmlSeq, Status)
    End Sub

    'NOTE WebXSeq NOT USED. USE XSeq.
    'Private Sub WebXSeq_ErrorMsg(ErrMsg As String) Handles WebXSeq.ErrorMsg
    '    Message.AddWarning(ErrMsg & vbCrLf)
    'End Sub


    'Private Sub WebXSeq_Instruction(Info As String, Locn As String) Handles WebXSeq.Instruction
    '    'Execute each instruction produced by running the XSeq file.

    '    Select Case Locn
    '        Case "Settings:Form:Name"
    '            FormName = Info

    '        Case "Settings:Form:Item:Name"
    '            ItemName = Info

    '        Case "Settings:Form:Item:Value"
    '            RestoreSetting(FormName, ItemName, Info)

    '        Case "Settings:Form:SelectId"
    '            SelectId = Info

    '        Case "Settings:Form:OptionText"
    '            RestoreOption(SelectId, Info)


    '        Case "Settings"

    '        Case "EndOfSequence"
    '            'Main.Message.Add("End of processing sequence" & Info & vbCrLf)

    '        Case Else
    '            Message.AddWarning("Web XSequence: " & Locn & vbCrLf)
    '            Message.AddWarning("Unknown location: " & Locn & "  Info: " & Info & vbCrLf)

    '    End Select
    'End Sub

    'END Run an XSequence ------------------------------------------------------------------------------------------


    'Run an XMessage ===============================================================================================

    Public Sub RunXMessage(ByVal XMsg As String)
        'Run the XMessage by sending it to InstrReceived.
        InstrReceived = XMsg
    End Sub

    Public Sub SendXMessage(ByVal ConnName As String, ByVal XMsg As String)
        'Send the XMessage to the application with the connection name ConnName.
        If IsNothing(client) Then
            Message.Add("No client connection available!" & vbCrLf)
        Else
            If client.State = ServiceModel.CommunicationState.Faulted Then
                Message.Add("client state is faulted. Message not sent!" & vbCrLf)
            Else
                If bgwSendMessage.IsBusy Then
                    Message.AddWarning("Send Message backgroundworker is busy." & vbCrLf)
                Else
                    Dim SendMessageParams As New Main.clsSendMessageParams
                    SendMessageParams.ProjectNetworkName = ProNetName
                    SendMessageParams.ConnectionName = ConnName
                    SendMessageParams.Message = XMsg
                    bgwSendMessage.RunWorkerAsync(SendMessageParams)
                    If ShowXMessages Then
                        Message.XAddText("Message sent to " & "[" & ProNetName & "]." & ConnName & ":" & vbCrLf, "XmlSentNotice")
                        Message.XAddXml(XMsg)
                        Message.XAddText(vbCrLf, "Normal") 'Add extra line
                    End If
                End If
            End If
        End If
    End Sub

    Public Sub SendXMessageExt(ByVal ProNetName As String, ByVal ConnName As String, ByVal XMsg As String)
        'Send the XMsg to the application with the connection name ConnName and Project Network Name ProNetname.
        'This version can send the XMessage to a connection external to the current Project Network.
        If IsNothing(client) Then
            Message.Add("No client connection available!" & vbCrLf)
        Else
            If client.State = ServiceModel.CommunicationState.Faulted Then
                Message.Add("client state is faulted. Message not sent!" & vbCrLf)
            Else
                If bgwSendMessage.IsBusy Then
                    Message.AddWarning("Send Message backgroundworker is busy." & vbCrLf)
                Else
                    Dim SendMessageParams As New Main.clsSendMessageParams
                    SendMessageParams.ProjectNetworkName = ProNetName
                    SendMessageParams.ConnectionName = ConnName
                    SendMessageParams.Message = XMsg
                    bgwSendMessage.RunWorkerAsync(SendMessageParams)
                    If ShowXMessages Then
                        Message.XAddText("Message sent to " & "[" & ProNetName & "]." & ConnName & ":" & vbCrLf, "XmlSentNotice")
                        Message.XAddXml(XMsg)
                        Message.XAddText(vbCrLf, "Normal") 'Add extra line
                    End If
                End If
            End If
        End If
    End Sub

    Public Sub SendXMessageWait(ByVal ConnName As String, ByVal XMsg As String)
        'Send the XMsg to the application with the connection name ConnName.
        'Wait for the connection to be made.
        If IsNothing(client) Then
            Message.Add("No client connection available!" & vbCrLf)
        Else
            Try
                'Application.DoEvents() 'TRY THE METHOD WITHOUT THE DOEVENTS
                If client.State = ServiceModel.CommunicationState.Faulted Then
                    Message.Add("client state is faulted. Message not sent!" & vbCrLf)
                Else
                    Dim StartTime As Date = Now
                    Dim Duration As TimeSpan
                    'Wait up to 16 seconds for the connection ConnName to be established
                    While client.ConnectionExists(ProNetName, ConnName) = False 'Wait until the required connection is made.
                        System.Threading.Thread.Sleep(1000) 'Pause for 1000ms
                        Duration = Now - StartTime
                        If Duration.Seconds > 16 Then Exit While
                    End While

                    If client.ConnectionExists(ProNetName, ConnName) = False Then
                        Message.AddWarning("Connection not available: " & ConnName & " in application network: " & ProNetName & vbCrLf)
                    Else
                        If bgwSendMessage.IsBusy Then
                            Message.AddWarning("Send Message backgroundworker is busy." & vbCrLf)
                        Else
                            Dim SendMessageParams As New Main.clsSendMessageParams
                            SendMessageParams.ProjectNetworkName = ProNetName
                            SendMessageParams.ConnectionName = ConnName
                            SendMessageParams.Message = XMsg
                            bgwSendMessage.RunWorkerAsync(SendMessageParams)
                            If ShowXMessages Then
                                Message.XAddText("Message sent to " & "[" & ProNetName & "]." & ConnName & ":" & vbCrLf, "XmlSentNotice")
                                Message.XAddXml(XMsg)
                                Message.XAddText(vbCrLf, "Normal") 'Add extra line
                            End If
                        End If
                    End If
                End If
            Catch ex As Exception
                Message.AddWarning(ex.Message & vbCrLf)
            End Try
        End If
    End Sub

    Public Sub SendXMessageExtWait(ByVal ProNetName As String, ByVal ConnName As String, ByVal XMsg As String)
        'Send the XMsg to the application with the connection name ConnName and Project Network Name ProNetName.
        'Wait for the connection to be made.
        'This version can send the XMessage to a connection external to the current Project Network.
        If IsNothing(client) Then
            Message.Add("No client connection available!" & vbCrLf)
        Else
            If client.State = ServiceModel.CommunicationState.Faulted Then
                Message.Add("client state is faulted. Message not sent!" & vbCrLf)
            Else
                Dim StartTime As Date = Now
                Dim Duration As TimeSpan
                'Wait up to 16 seconds for the connection ConnName to be established
                While client.ConnectionExists(ProNetName, ConnName) = False
                    System.Threading.Thread.Sleep(1000) 'Pause for 1000ms
                    Duration = Now - StartTime
                    If Duration.Seconds > 16 Then Exit While
                End While

                If client.ConnectionExists(ProNetName, ConnName) = False Then
                    Message.AddWarning("Connection not available: " & ConnName & " in application network: " & ProNetName & vbCrLf)
                Else
                    If bgwSendMessage.IsBusy Then
                        Message.AddWarning("Send Message backgroundworker is busy." & vbCrLf)
                    Else
                        Dim SendMessageParams As New Main.clsSendMessageParams
                        SendMessageParams.ProjectNetworkName = ProNetName
                        SendMessageParams.ConnectionName = ConnName
                        SendMessageParams.Message = XMsg
                        bgwSendMessage.RunWorkerAsync(SendMessageParams)
                        If ShowXMessages Then
                            Message.XAddText("Message sent to " & "[" & ProNetName & "]." & ConnName & ":" & vbCrLf, "XmlSentNotice")
                            Message.XAddXml(XMsg)
                            Message.XAddText(vbCrLf, "Normal") 'Add extra line
                        End If
                    End If
                End If
            End If
        End If
    End Sub

    Public Sub XMsgInstruction(ByVal Info As String, ByVal Locn As String)
        'Send the XMessage Instruction to the JavaScript function XMsgInstruction for processing.
        Me.WebBrowser1.Document.InvokeScript("XMsgInstruction", New String() {Info, Locn})
    End Sub

    'END Run an XMessage -------------------------------------------------------------------------------------------


    'Get Information ===============================================================================================
    Public Function GetFormNo() As String
        'Return the Form Number of the current instance of the WebPage form.
        'Return FormNo.ToString
        Return "-1"
    End Function

    Public Function GetParentFormNo() As String
        'Return the Form Number of the Parent Form (that called this form).
        'Return ParentWebPageFormNo.ToString
        Return "-1" 'The Main Form does not have a Parent Web Page.
    End Function

    Public Function GetConnectionName() As String
        'Return the Connection Name of the Project.
        Return ConnectionName
    End Function

    Public Function GetProNetName() As String
        'Return the Project Network Name of the Project.
        Return ProNetName
    End Function

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
        'RestoreSetting(FormName, ItemName, Project.ParentParameter(ParameterName).Value)
        RestoreSetting(FormName, ItemName, Project.GetParentParameter(ParameterName))
    End Sub

    Public Sub ProjectParameterValue(ByVal FormName As String, ByVal ItemName As String, ByVal ParameterName As String)
        'Return the specified Project parameter value:
        'RestoreSetting(FormName, ItemName, Project.Parameter(ParameterName).Value)
        RestoreSetting(FormName, ItemName, Project.GetParameter(ParameterName))
    End Sub

    Public Sub ProjectNetworkName(ByVal FormName As String, ByVal ItemName As String)
        'Return the name of the Project Network:
        'RestoreSetting(FormName, ItemName, Project.Parameter("AppNetName").Value)
        RestoreSetting(FormName, ItemName, Project.GetParameter("ProNetName"))
    End Sub

    Public Sub CalcValsTable(ByVal TableName As String, ByVal FormName As String, ByVal ItemName As String)
        'Search for TableName in the Calculated Values database. 
        'If Found return TableName to FormName/ItemName.
        'If not found return "" to FormName/ItemName.

        If TableExists(Project.GetParameter("CalculationsDatabasePath"), TableName) Then
            RestoreSetting(FormName, ItemName, TableName)
        Else
            RestoreSetting(FormName, ItemName, "")
        End If
    End Sub

    'END Get Information -------------------------------------------------------------------------------------------


    'Open a Web Page ===============================================================================================

    Public Sub OpenWebPage(ByVal FileName As String)
        'Open the web page with the specified File Name.

        If FileName = "" Then

        Else
            'First check if the HTML file is already open:
            Dim FileFound As Boolean = False
            If WebPageFormList.Count = 0 Then

            Else
                Dim I As Integer
                For I = 0 To WebPageFormList.Count - 1
                    If WebPageFormList(I) Is Nothing Then

                    Else
                        If WebPageFormList(I).FileName = FileName Then
                            FileFound = True
                            WebPageFormList(I).BringToFront
                        End If
                    End If
                Next
            End If

            If FileFound = False Then
                Dim FormNo As Integer = OpenNewWebPage()
                WebPageFormList(FormNo).FileName = FileName
                WebPageFormList(FormNo).OpenDocument
                WebPageFormList(FormNo).BringToFront
            End If
        End If
    End Sub

    'END Open a Web Page -------------------------------------------------------------------------------------------


    'Open and Close Projects =======================================================================================

    Public Sub OpenProjectAtRelativePath(ByVal RelativePath As String, ByVal ConnectionName As String)
        'Open the Project at the specified Relative Path using the specified Connection Name.

        Dim ProjectPath As String
        If RelativePath.StartsWith("\") Then
            ProjectPath = Project.Path & RelativePath
            client.StartProjectAtPath(ProjectPath, ConnectionName)
        Else
            ProjectPath = Project.Path & "\" & RelativePath
            client.StartProjectAtPath(ProjectPath, ConnectionName)
        End If
    End Sub

    Public Sub CheckOpenProjectAtRelativePath(ByVal RelativePath As String, ByVal ConnectionName As String)
        'Check if the project at the specified Relative Path is open.
        'Open it if it is not already open.
        'Open the Project at the specified Relative Path using the specified Connection Name.

        Dim ProjectPath As String
        If RelativePath.StartsWith("\") Then
            ProjectPath = Project.Path & RelativePath
            If client.ProjectOpen(ProjectPath) Then
                'Project is already open.
            Else
                client.StartProjectAtPath(ProjectPath, ConnectionName)
            End If
        Else
            ProjectPath = Project.Path & "\" & RelativePath
            If client.ProjectOpen(ProjectPath) Then
                'Project is already open.
            Else
                client.StartProjectAtPath(ProjectPath, ConnectionName)
            End If
        End If
    End Sub

    Public Sub OpenProjectAtProNetPath(ByVal RelativePath As String, ByVal ConnectionName As String)
        'Open the Project at the specified Path (relative to the ProNet Path) using the specified Connection Name.

        Dim ProjectPath As String
        If RelativePath.StartsWith("\") Then
            If Project.ParameterExists("ProNetPath") Then
                ProjectPath = Project.GetParameter("ProNetPath") & RelativePath
                client.StartProjectAtPath(ProjectPath, ConnectionName)
            Else
                Message.AddWarning("The Project Network Path is not known." & vbCrLf)
            End If
        Else
            If Project.ParameterExists("ProNetPath") Then
                ProjectPath = Project.GetParameter("ProNetPath") & "\" & RelativePath
                client.StartProjectAtPath(ProjectPath, ConnectionName)
            Else
                Message.AddWarning("The Project Network Path is not known." & vbCrLf)
            End If
        End If
    End Sub

    Public Sub CheckOpenProjectAtProNetPath(ByVal RelativePath As String, ByVal ConnectionName As String)
        'Check if the project at the specified Path (relative to the ProNet Path) is open.
        'Open it if it is not already open.
        'Open the Project at the specified Path using the specified Connection Name.

        Dim ProjectPath As String
        If RelativePath.StartsWith("\") Then
            If Project.ParameterExists("ProNetPath") Then
                ProjectPath = Project.GetParameter("ProNetPath") & RelativePath
                If client.ProjectOpen(ProjectPath) Then
                    'Project is already open.
                Else
                    client.StartProjectAtPath(ProjectPath, ConnectionName)
                End If
            Else
                Message.AddWarning("The Project Network Path is not known." & vbCrLf)
            End If
        Else
            If Project.ParameterExists("ProNetPath") Then
                ProjectPath = Project.GetParameter("ProNetPath") & "\" & RelativePath
                If client.ProjectOpen(ProjectPath) Then
                    'Project is already open.
                Else
                    client.StartProjectAtPath(ProjectPath, ConnectionName)
                End If
            Else
                Message.AddWarning("The Project Network Path is not known." & vbCrLf)
            End If
        End If
    End Sub

    Public Sub CloseProjectAtConnection(ByVal ProNetName As String, ByVal ConnectionName As String)
        'Close the Project at the specified connection.

        If IsNothing(client) Then
            Message.Add("No client connection available!" & vbCrLf)
        Else
            If client.State = ServiceModel.CommunicationState.Faulted Then
                Message.Add("client state is faulted. Message not sent!" & vbCrLf)
            Else
                'Create the XML instructions to close the Project at the connection.
                Dim decl As New XDeclaration("1.0", "utf-8", "yes")
                Dim doc As New XDocument(decl, Nothing) 'Create an XDocument to store the instructions.
                Dim xmessage As New XElement("XMsg") 'This indicates the start of the message in the XMessage class

                'NOTE: No reply expected. No need to provide the following client information(?)
                'Dim clientConnName As New XElement("ClientConnectionName", Me.ConnectionName)
                'xmessage.Add(clientConnName)

                Dim command As New XDocument("Command", "Close")
                xmessage.Add(command)

                doc.Add(xmessage)

                'Show the message sent:
                Message.XAddText("Message sent to [" & ProNetName & "]." & ConnectionName & ":" & vbCrLf, "XmlSentNotice")
                Message.XAddXml(doc.ToString)
                Message.XAddText(vbCrLf, "Normal") 'Add extra line

                client.SendMessage(ProNetName, ConnectionName, doc.ToString)
            End If
        End If
    End Sub

    'END Open and Close Projects -----------------------------------------------------------------------------------


    'System Methods ================================================================================================

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

        Dim SettingsFileName As String = WorkflowFileName & "Settings"

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

    'END System Methods --------------------------------------------------------------------------------------------


    'Legacy Code (These methods should no longer be used) ==========================================================

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

    'Add text to the application message window with the specified text type.
    Public Sub AddText(ByVal Msg As String, ByVal TextType As String)
        Message.AddText(Msg, TextType)
    End Sub

    'END Legacy Code -----------------------------------------------------------------------------------------------


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
        Project.Usage.SaveUsageInfo()  'Save the current project usage information.
        Project.UnlockProject() 'Unlock the current project before it Is closed.
        If ConnectedToComNet Then DisconnectFromComNet()
    End Sub

    Private Sub Project_Selected() Handles Project.Selected
        'A new project has been selected.

        RestoreFormSettings()
        Project.ReadProjectInfoFile()

        Project.ReadParameters()
        Project.ReadParentParameters()
        If Project.ParentParameterExists("ProNetName") Then
            Project.AddParameter("ProNetName", Project.ParentParameter("ProNetName").Value, Project.ParentParameter("ProNetName").Description) 'AddParameter will update the parameter if it already exists.
            ProNetName = Project.Parameter("ProNetName").Value
        Else
            ProNetName = Project.GetParameter("ProNetName")
        End If
        If Project.ParentParameterExists("ProNetPath") Then 'Get the parent parameter value - it may have been updated.
            Project.AddParameter("ProNetPath", Project.ParentParameter("ProNetPath").Value, Project.ParentParameter("ProNetPath").Description) 'AddParameter will update the parameter if it already exists.
            ProNetPath = Project.Parameter("ProNetPath").Value
        Else
            ProNetPath = Project.GetParameter("ProNetPath") 'If the parameter does not exist, the value is set to ""
        End If
        Project.SaveParameters() 'These should be saved now - child projects look for parent parameters in the parameter file.

        Project.LockProject() 'Lock the project while it is open in this application.

        Project.Usage.StartTime = Now

        ApplicationInfo.SettingsLocn = Project.SettingsLocn
        Message.SettingsLocn = Project.SettingsLocn
        Message.Show() 'Added 18May19

        'Restore the new project settings:
        RestoreProjectSettings() 'Update this subroutine if project settings need to be restored.

        ShowProjectInfo()

        ''Show the project information:
        'txtProjectName.Text = Project.Name
        'txtProjectDescription.Text = Project.Description
        'Select Case Project.Type
        '    Case ADVL_Utilities_Library_1.Project.Types.Directory
        '        txtProjectType.Text = "Directory"
        '    Case ADVL_Utilities_Library_1.Project.Types.Archive
        '        txtProjectType.Text = "Archive"
        '    Case ADVL_Utilities_Library_1.Project.Types.Hybrid
        '        txtProjectType.Text = "Hybrid"
        '    Case ADVL_Utilities_Library_1.Project.Types.None
        '        txtProjectType.Text = "None"
        'End Select

        'txtCreationDate.Text = Format(Project.CreationDate, "d-MMM-yyyy H:mm:ss")
        'txtLastUsed.Text = Format(Project.Usage.LastUsed, "d-MMM-yyyy H:mm:ss")
        'Select Case Project.SettingsLocn.Type
        '    Case ADVL_Utilities_Library_1.FileLocation.Types.Directory
        '        txtSettingsLocationType.Text = "Directory"
        '    Case ADVL_Utilities_Library_1.FileLocation.Types.Archive
        '        txtSettingsLocationType.Text = "Archive"
        'End Select
        'txtSettingsLocationPath.Text = Project.SettingsLocn.Path
        'Select Case Project.DataLocn.Type
        '    Case ADVL_Utilities_Library_1.FileLocation.Types.Directory
        '        txtDataLocationType.Text = "Directory"
        '    Case ADVL_Utilities_Library_1.FileLocation.Types.Archive
        '        txtDataLocationType.Text = "Archive"
        'End Select
        'txtDataLocationPath.Text = Project.DataLocn.Path

        If Project.ConnectOnOpen Then
            ConnectToComNet() 'The Project is set to connect when it is opened.
        ElseIf ApplicationInfo.ConnectOnStartup Then
            ConnectToComNet() 'The Application is set to connect when it is started.
        Else
            'Don't connect to ComNet.
        End If

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
        'Connect to the Message Service. (ComNet)

        If IsNothing(client) Then
            client = New ServiceReference1.MsgServiceClient(New System.ServiceModel.InstanceContext(New MsgServiceCallback))
        End If

        If ComNetRunning() Then
            'The Message Service is Running.
            'The Application.Lock file has been found at AdvlNetworkAppPath
        Else  'The Message Service is NOT Running.
            'Start the Message Service:
            If AdvlNetworkAppPath = "" Then
                Message.AddWarning("Andorville™ Network application path is unknown." & vbCrLf)
            Else
                If System.IO.File.Exists(AdvlNetworkExePath) Then 'OK to start the Message Service application:
                    Shell(Chr(34) & AdvlNetworkExePath & Chr(34), AppWinStyle.NormalFocus) 'Start Message Service application with no argument
                Else
                    'Incorrect Message Service Executable path.
                    Message.AddWarning("Andorville™ Network exe file not found. Service not started." & vbCrLf)
                End If
            End If
        End If

        'Try to fix a faulted client state:
        If client.State = ServiceModel.CommunicationState.Faulted Then
            client = Nothing
            client = New ServiceReference1.MsgServiceClient(New System.ServiceModel.InstanceContext(New MsgServiceCallback))
        End If

        If client.State = ServiceModel.CommunicationState.Faulted Then
            Message.AddWarning("Client state is faulted. Connection not made!" & vbCrLf)
        Else
            Try
                'client.Endpoint.Binding.SendTimeout = New System.TimeSpan(0, 0, 8) 'Temporarily set the send timeaout to 8 seconds
                client.Endpoint.Binding.SendTimeout = New System.TimeSpan(0, 0, 16) 'Temporarily set the send timeaout to 16 seconds
                ConnectionName = ApplicationInfo.Name 'This name will be modified if it is already used in an existing connection.
                ConnectionName = client.Connect(ProNetName, ApplicationInfo.Name, ConnectionName, Project.Name, Project.Description, Project.Type, Project.Path, False, False)

                If ConnectionName <> "" Then
                    Message.Add("Connected to the Andorville™ Network with Connection Name: [" & ProNetName & "]." & ConnectionName & vbCrLf)
                    client.Endpoint.Binding.SendTimeout = New System.TimeSpan(1, 0, 0) 'Restore the send timeaout to 1 hour
                    btnOnline.Text = "Online"
                    btnOnline.ForeColor = Color.ForestGreen
                    ConnectedToComNet = True
                    SendApplicationInfo()
                    SendProjectInfo()
                    client.GetAdvlNetworkAppInfoAsync() 'Update the Exe Path in case it has changed. This path may be needed in the future to start the ComNet (Message Service).

                    bgwComCheck.WorkerReportsProgress = True
                    bgwComCheck.WorkerSupportsCancellation = True
                    If bgwComCheck.IsBusy Then
                        'The ComCheck thread is already running.
                    Else
                        bgwComCheck.RunWorkerAsync() 'Start the ComCheck thread.
                    End If
                Else
                    Message.Add("Connection to the Andorville™ Network failed!" & vbCrLf)
                    client.Endpoint.Binding.SendTimeout = New System.TimeSpan(1, 0, 0) 'Restore the send timeaout to 1 hour
                End If
            Catch ex As System.TimeoutException
                Message.Add("Timeout error. Check if the Andorville™ Network (Message Service) is running." & vbCrLf)
            Catch ex As Exception
                Message.Add("Error message: " & ex.Message & vbCrLf)
                client.Endpoint.Binding.SendTimeout = New System.TimeSpan(1, 0, 0) 'Restore the send timeaout to 1 hour
            End Try
        End If
    End Sub

    Private Sub ConnectToComNet(ByVal ConnName As String)
        'Connect to the Message Service (ComNet) with the connection name ConnName.

        If ConnectedToComNet = False Then
            Dim Result As Boolean

            If IsNothing(client) Then
                client = New ServiceReference1.MsgServiceClient(New System.ServiceModel.InstanceContext(New MsgServiceCallback))
            End If

            'Try to fix a faulted client state:
            If client.State = ServiceModel.CommunicationState.Faulted Then
                client = Nothing
                client = New ServiceReference1.MsgServiceClient(New System.ServiceModel.InstanceContext(New MsgServiceCallback))
            End If

            If client.State = ServiceModel.CommunicationState.Faulted Then
                Message.AddWarning("client state is faulted. Connection not made!" & vbCrLf)
            Else
                Try
                    'client.Endpoint.Binding.SendTimeout = New System.TimeSpan(0, 0, 8) 'Temporarily set the send timeaout to 8 seconds
                    client.Endpoint.Binding.SendTimeout = New System.TimeSpan(0, 0, 16) 'Temporarily set the send timeaout to 16 seconds
                    ConnectionName = ConnName 'This name will be modified if it is already used in an existing connection.
                    'ConnectionName = client.Connect(AppNetName, ApplicationInfo.Name, ConnectionName, Project.Name, Project.Description, Project.Type, Project.Path, False, False) 'UPDATED 2Feb19
                    ConnectionName = client.Connect(ProNetName, ApplicationInfo.Name, ConnectionName, Project.Name, Project.Description, Project.Type, Project.Path, False, False)

                    If ConnectionName <> "" Then
                        'Message.Add("Connected to the Communication Network as " & ConnectionName & vbCrLf)
                        'Message.Add("Connected to the Andorville™ Network with Connection Name: [" & AppNetName & "]." & ConnectionName & vbCrLf)
                        Message.Add("Connected to the Andorville™ Network with Connection Name: [" & ProNetName & "]." & ConnectionName & vbCrLf)
                        client.Endpoint.Binding.SendTimeout = New System.TimeSpan(1, 0, 0) 'Restore the send timeaout to 1 hour
                        'client.Endpoint.Binding.SendTimeout = New System.TimeSpan(2, 0, 0) 'Restore the send timeaout to 2 hours
                        btnOnline.Text = "Online"
                        btnOnline.ForeColor = Color.ForestGreen
                        ConnectedToComNet = True
                        SendApplicationInfo()
                        SendProjectInfo()
                        client.GetAdvlNetworkAppInfoAsync() 'Update the Exe Path in case it has changed. This path may be needed in the future to start the ComNet (Message Service).

                        bgwComCheck.WorkerReportsProgress = True
                        bgwComCheck.WorkerSupportsCancellation = True
                        If bgwComCheck.IsBusy Then
                            'The ComCheck thread is already running.
                        Else
                            bgwComCheck.RunWorkerAsync() 'Start the ComCheck thread.
                        End If
                    Else
                        'Message.Add("Connection to the Communication Network failed!" & vbCrLf)
                        Message.Add("Connection to the Andorville™ Network failed!" & vbCrLf)
                        client.Endpoint.Binding.SendTimeout = New System.TimeSpan(1, 0, 0) 'Restore the send timeaout to 1 hour
                        'client.Endpoint.Binding.SendTimeout = New System.TimeSpan(2, 0, 0) 'Restore the send timeaout to 2 hours
                    End If
                Catch ex As System.TimeoutException
                    'Message.Add("Timeout error. Check if the Communication Network is running." & vbCrLf)
                    Message.Add("Timeout error. Check if the Andorville™ Network (Message Service) is running." & vbCrLf)
                Catch ex As Exception
                    Message.Add("Error message: " & ex.Message & vbCrLf)
                    client.Endpoint.Binding.SendTimeout = New System.TimeSpan(1, 0, 0) 'Restore the send timeaout to 1 hour
                    'client.Endpoint.Binding.SendTimeout = New System.TimeSpan(2, 0, 0) 'Restore the send timeaout to 2 hours
                End Try
            End If
        Else
            'Message.AddWarning("Already connected to the Communication Network." & vbCrLf)
            Message.AddWarning("Already connected to the Andorville™ Network (Message Service)." & vbCrLf)
        End If

    End Sub

    Private Sub DisconnectFromComNet()
        'Disconnect from the Communication Network (Message Service).

        If ConnectedToComNet = True Then
            If IsNothing(client) Then
                'Message.Add("Already disconnected from the Communication Network." & vbCrLf)
                Message.Add("Already disconnected from the Andorville™ Network (Message Service)." & vbCrLf)
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
                        'client.Disconnect(AppNetName, ConnectionName)
                        client.Disconnect(ProNetName, ConnectionName)
                        btnOnline.Text = "Offline"
                        btnOnline.ForeColor = Color.Red
                        ConnectedToComNet = False
                        ConnectionName = ""
                        'Message.Add("Disconnected from the Communication Network." & vbCrLf)
                        Message.Add("Disconnected from the Andorville™ Network (Message Service)." & vbCrLf)

                        If bgwComCheck.IsBusy Then
                            bgwComCheck.CancelAsync()
                        End If
                    Catch ex As Exception
                        'Message.AddWarning("Error disconnecting from Communication Network: " & ex.Message & vbCrLf)
                        Message.AddWarning("Error disconnecting from Andorville™ Network (Message Service): " & ex.Message & vbCrLf)
                    End Try
                End If
            End If
        End If
    End Sub

    Private Function ComNetRunning() As Boolean
        'Return True if ComNet (Message Service) is running.

        If AdvlNetworkAppPath = "" Then
            Message.Add("Andorville™ Network application path is not known." & vbCrLf)
            Message.Add("Run the Andorville™ Network before connecting to update the path." & vbCrLf)
            Return False
        Else
            If System.IO.File.Exists(AdvlNetworkAppPath & "\Application.Lock") Then
                Return True
            Else
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
                Message.XAddText("Message sent to " & "Message Service" & ":" & vbCrLf, "XmlSentNotice")
                Message.XAddXml(doc.ToString)
                Message.XAddText(vbCrLf, "Normal") 'Add extra line

                client.SendMessage("", "MessageService", doc.ToString)
            End If
        End If

    End Sub

    Private Sub SendProjectInfo()
        'Send the project information to the Network application.

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

                    'Show the message sent to the Message Service:
                    Message.XAddText("Message sent to " & "Message Service" & ":" & vbCrLf, "XmlSentNotice")
                    Message.XAddXml(doc.ToString)
                    Message.XAddText(vbCrLf, "Normal") 'Add extra line
                    client.SendMessage("", "MessageService", doc.ToString)
                End If
            End If
        End If
    End Sub

    Public Sub SendProjectInfo(ByVal ProjectPath As String)
        'Send the project information to the Network application.
        'This version of SendProjectInfo uses the ProjectPath argument.

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

                    'Dim Path As New XElement("Path", Project.Path)
                    Dim Path As New XElement("Path", ProjectPath)
                    projectInfo.Add(Path)
                    xmessage.Add(projectInfo)
                    doc.Add(xmessage)

                    'Show the message sent to the Message Service:
                    Message.XAddText("Message sent to " & "Message Service" & ":" & vbCrLf, "XmlSentNotice")
                    Message.XAddXml(doc.ToString)
                    Message.XAddText(vbCrLf, "Normal") 'Add extra line
                    client.SendMessage("", "MessageService", doc.ToString)
                End If
            End If
        End If
    End Sub

#End Region 'Online/Offline code



#Region " Process XMessages" '=========================================================================================================================================================

    Private Sub XMsg_Instruction(Data As String, Locn As String) Handles XMsg.Instruction
        'Process an XMessage instruction.
        'An XMessage is a simplified XSequence. It is used to exchange information between Andorville (TM) applications.
        '
        'An XSequence file is an AL-H7 (TM) Information Vector Sequence stored in an XML format.
        'AL-H7(TM) is the name of a programming system that uses sequences of information and location value pairs to store data items or processing steps.
        'A single information and location value pair is called a knowledge element (or noxel).
        'Any program, mathematical expression or data set can be expressed as an Information Vector Sequence.

        'Add code here to process the XMessage instructions.
        'See other Andorville(TM) applications for examples.

        If IsDBNull(Data) Then
            Data = ""
        End If

        'For Debugging:
        'Message.Add("XMsg.Instruction - Data = " & Data & "  Locn = " & Locn & vbCrLf)

        'Intercept instructions with the prefix "WebPage_"
        If Locn.StartsWith("WebPage_") Then 'Send the Data, Location data to the correct Web Page:
            'Message.Add("Web Page Location: " & Locn & vbCrLf)
            If Locn.Contains(":") Then
                Dim EndOfWebPageNoString As Integer = Locn.IndexOf(":")
                If Locn.Contains("-") Then
                    Dim HyphenLocn As Integer = Locn.IndexOf("-")
                    If HyphenLocn < EndOfWebPageNoString Then 'Web Page Location contains a sub-location in the web page - WebPage_1-SubLocn:Locn - SubLocn:Locn will be sent to Web page 1
                        EndOfWebPageNoString = HyphenLocn
                    End If
                End If
                Dim PageNoLen As Integer = EndOfWebPageNoString - 8
                Dim WebPageNoString As String = Locn.Substring(8, PageNoLen)
                Dim WebPageNo As Integer = CInt(WebPageNoString)
                Dim WebPageData As String = Data
                Dim WebPageLocn As String = Locn.Substring(EndOfWebPageNoString + 1)

                'Message.Add("WebPageData = " & WebPageData & "  WebPageLocn = " & WebPageLocn & vbCrLf)

                WebPageFormList(WebPageNo).XMsgInstruction(WebPageData, WebPageLocn)
            Else
                Message.AddWarning("XMessage instruction location is not complete: " & Locn & vbCrLf)
            End If
        Else

            Select Case Locn

                'Case "ClientAppNetName"
                '    ClientAppNetName = Data 'The name of the Client Application Network requesting service. 

                Case "ClientProNetName"
                    ClientProNetName = Data 'The name of the Client Application Network requesting service. 

                Case "ClientName"
                    ClientAppName = Data 'The name of the Client requesting service.

                Case "ClientConnectionName"
                    ClientConnName = Data 'The name of the client requesting service.

                Case "ClientLocn" 'The Location within the Client requesting service.
                    Dim statusOK As New XElement("Status", "OK") 'Add Status OK element when the Client Location is changed
                    xlocns(xlocns.Count - 1).Add(statusOK)
                    xmessage.Add(xlocns(xlocns.Count - 1)) 'Add the instructions for the last location to the reply xmessage
                    xlocns.Add(New XElement(Data)) 'Start the new location instructions

                ''DEPRECATED:
                'Case "OnCompletion" 'Specify the last instruction to be returned on completion of the XMessage processing.
                '    'CompletionInstruction = Data
                '    XCompletionInstruction = Data

                'UPDATE:
                Case "OnCompletion"
                    OnCompletionInstruction = Data

                Case "TestInstruction"
                    'Test - do nothing.

                ''REPLACED BY:
                'Case "EndInstructionData"
                '    EndInstructionData = Data
                'Case "EndInstructionLocn"
                '    EndInstructionLocn = Data

                'Case "EndInstructionInfoReply"
                '    Dim endInstructionInfo As New XElement("EndInstructionInfo", Data)
                '    xlocns(xlocns.Count - 1).Add(endInstructionInfo)
                'Case "EndInstructionLocnReply"
                '    Dim endInstructionLocn As New XElement("EndInstructionLocn", Data)
                '    xlocns(xlocns.Count - 1).Add(endInstructionLocn)

                ''TESTING: - OK
                'Case "EndInstruction"
                '    EndInstruction = Data


                Case "Command"
                    Select Case Data
                        Case "Close"

                            'NOTE: The following does not work:
                            'Me.btnExit.PerformClick() 'Press the Exit button

                            'Use Timer4 to delay the click:
                            Timer4.Interval = 100 '100ms delay
                            Timer4.Enabled = True 'Start the timer

                        Case "AppComCheck"
                            'Add the Appplication Communication info to the reply message:
                            Dim clientProNetName As New XElement("ClientProNetName", ProNetName) 'The Project Network Name
                            xlocns(xlocns.Count - 1).Add(clientProNetName)
                            Dim clientName As New XElement("ClientName", "ADVL_Shares_1") 'The name of this application.
                            xlocns(xlocns.Count - 1).Add(clientName)
                            Dim clientConnectionName As New XElement("ClientConnectionName", ConnectionName)
                            xlocns(xlocns.Count - 1).Add(clientConnectionName)
                            '<Status>OK</Status> will be automatically appended to the XMessage before it is sent.

                        Case Else
                            Message.AddWarning("Unknown application command: " & Data & vbCrLf)
                    End Select

                Case "Main"
                 'Blank message - do nothing.

                'Case "Main:OnCompletion"
                '    'Select Case "Stop"
                '    Select Case Data
                '        Case "Stop"
                '            'Stop on completion of the instruction sequence.
                '    End Select

                'Case "Main:EndInstructionInfo"
                '    Select Case Data
                '        Case "Stop"
                '            'Stop on completion of the instruction sequence.

                '            'Add other cases here:
                '    End Select

                Case "Main:EndInstruction"
                    Select Case Data
                        Case "Stop"
                            'Stop at the end of the instruction sequence.

                            'Add other cases here:
                        Case "UpdateStockChartProjects"
                            UpdateStockChartProjects()
                            Message.Add("Stock Chart projects updated!" & vbCrLf)
                        Case "UpdateChartProjLists"
                            UpdateChartProjLists()
                        Case Else
                            Message.AddWarning("Unknown End Instruction: " & Data & vbCrLf)
                    End Select




                Case "Main:Status"
                    Select Case Data
                        Case "OK"
                            'Main instructions completed OK
                    End Select

                Case "StockChart"
                 'Blank message - do nothing.

                Case "StockChart:Status"
                    Select Case Data
                        Case "OK"
                            'Stock Chart instructions completed OK
                    End Select

                Case "PointChart"
                'Blank message - do nothing.

                Case "PointChart:Status"
                    Select Case Data
                        Case "OK"
                            'Point Chart instructions completed OK
                    End Select

           'Stock Chart instructions: ---------------------------------------------------------------------------------------------------

                Case "StockChart:Settings:Command"
                    Select Case Data
                        Case "ClearChart"
                            'ClearStockChartDefaults()
                            ClearStockChartSettingsList()
                        Case "OK"
                            'StockChartDefaults has been updated. Display in rtbStockChartDefaults.
                            'rtbStockChartDefaults.Text = StockChartDefaults.ToString
                            'FormatXmlText(rtbStockChartDefaults)
                            'XmlStockChartDefaults.Rtf = XmlStockChartDefaults.XmlToRtf(StockChartDefaults.ToString, True)
                            'XmlStockChartSettingsList.Rtf = XmlStockChartSettingsList.XmlToRtf(StockChartDefaults.ToString, False)
                            XmlStockChartSettingsList.Rtf = XmlStockChartSettingsList.XmlToRtf(StockChartSettingsList.ToString, False)
                    End Select

                Case "StockChart:Settings:InputData:Type"
                    'StockChartDefaults.<StockChart>.<Settings>.<InputData>.<Type>.Value = Data
                    StockChartSettingsList.<StockChart>.<Settings>.<InputData>.<Type>.Value = Data
                Case "StockChart:Settings:InputData:DatabasePath"
                    'StockChartDefaults.<StockChart>.<Settings>.<InputData>.<DatabasePath>.Value = Data
                    StockChartSettingsList.<StockChart>.<Settings>.<InputData>.<DatabasePath>.Value = Data
                Case "StockChart:Settings:InputData:DataDescription"
                    StockChartSettingsList.<StockChart>.<Settings>.<InputData>.<DataDescription>.Value = Data
                Case "StockChart:Settings:InputData:DatabaseQuery"
                    StockChartSettingsList.<StockChart>.<Settings>.<InputData>.<DatabaseQuery>.Value = Data

                Case "StockChart:Settings:ChartProperties:SeriesName"
                    StockChartSettingsList.<StockChart>.<Settings>.<ChartProperties>.<SeriesName>.Value = Data
                Case "StockChart:Settings:ChartProperties:XValuesFieldName"
                    StockChartSettingsList.<StockChart>.<Settings>.<ChartProperties>.<XValuesFieldName>.Value = Data
                Case "StockChart:Settings:ChartProperties:YValuesHighFieldName"
                    StockChartSettingsList.<StockChart>.<Settings>.<ChartProperties>.<YValuesHighFieldName>.Value = Data
                Case "StockChart:Settings:ChartProperties:YValuesLowFieldName"
                    StockChartSettingsList.<StockChart>.<Settings>.<ChartProperties>.<YValuesLowFieldName>.Value = Data
                Case "StockChart:Settings:ChartProperties:YValuesOpenFieldName"
                    StockChartSettingsList.<StockChart>.<Settings>.<ChartProperties>.<YValuesOpenFieldName>.Value = Data
                Case "StockChart:Settings:ChartProperties:YValuesCloseFieldName"
                    StockChartSettingsList.<StockChart>.<Settings>.<ChartProperties>.<YValuesCloseFieldName>.Value = Data

                Case "StockChart:Settings:ChartTitle:LabelName"
                    StockChartSettingsList.<StockChart>.<Settings>.<ChartTitle>.<LabelName>.Value = Data
                Case "StockChart:Settings:ChartTitle:Text"
                    StockChartSettingsList.<StockChart>.<Settings>.<ChartTitle>.<Text>.Value = Data
                Case "StockChart:Settings:ChartTitle:FontName"
                    StockChartSettingsList.<StockChart>.<Settings>.<ChartTitle>.<FontName>.Value = Data
                Case "StockChart:Settings:ChartTitle:Color"
                    StockChartSettingsList.<StockChart>.<Settings>.<ChartTitle>.<Color>.Value = Data
                Case "StockChart:Settings:ChartTitle:Size"
                    StockChartSettingsList.<StockChart>.<Settings>.<ChartTitle>.<Size>.Value = Data
                Case "StockChart:Settings:ChartTitle:Bold"
                    StockChartSettingsList.<StockChart>.<Settings>.<ChartTitle>.<Bold>.Value = Data
                Case "StockChart:Settings:ChartTitle:Italic"
                    StockChartSettingsList.<StockChart>.<Settings>.<ChartTitle>.<Italic>.Value = Data
                Case "StockChart:Settings:ChartTitle:Underline"
                    StockChartSettingsList.<StockChart>.<Settings>.<ChartTitle>.<Underline>.Value = Data
                Case "StockChart:Settings:ChartTitle:Strikeout"
                    StockChartSettingsList.<StockChart>.<Settings>.<ChartTitle>.<Strikeout>.Value = Data

                Case "StockChart:Settings:XAxis:TitleText"
                    StockChartSettingsList.<StockChart>.<Settings>.<XAxis>.<TitleText>.Value = Data
                Case "StockChart:Settings:XAxis:TitleFontName"
                    StockChartSettingsList.<StockChart>.<Settings>.<XAxis>.<TitleFontName>.Value = Data
                Case "StockChart:Settings:XAxis:TitleColor"
                    StockChartSettingsList.<StockChart>.<Settings>.<XAxis>.<TitleColor>.Value = Data
                Case "StockChart:Settings:XAxis:TitleSize"
                    StockChartSettingsList.<StockChart>.<Settings>.<XAxis>.<TitleSize>.Value = Data
                Case "StockChart:Settings:XAxis:TitleBold"
                    StockChartSettingsList.<StockChart>.<Settings>.<XAxis>.<TitleBold>.Value = Data
                Case "StockChart:Settings:XAxis:TitleItalic"
                    StockChartSettingsList.<StockChart>.<Settings>.<XAxis>.<TitleItalic>.Value = Data
                Case "StockChart:Settings:XAxis:TitleUnderline"
                    StockChartSettingsList.<StockChart>.<Settings>.<XAxis>.<TitleUnderline>.Value = Data
                Case "StockChart:Settings:XAxis:TitleStrikeout"
                    StockChartSettingsList.<StockChart>.<Settings>.<XAxis>.<TitleStrikeout>.Value = Data
                Case "StockChart:Settings:XAxis:TitleAlignment"
                    StockChartSettingsList.<StockChart>.<Settings>.<XAxis>.<TitleAlignment>.Value = Data
                Case "StockChart:Settings:XAxis:AutoMinimum"
                    StockChartSettingsList.<StockChart>.<Settings>.<XAxis>.<AutoMinimum>.Value = Data
                Case "StockChart:Settings:XAxis:Minimum"
                    StockChartSettingsList.<StockChart>.<Settings>.<XAxis>.<Minimum>.Value = Data
                Case "StockChart:Settings:XAxis:AutoMaximum"
                    StockChartSettingsList.<StockChart>.<Settings>.<XAxis>.<AutoMaximum>.Value = Data
                Case "StockChart:Settings:XAxis:Maximum"
                    StockChartSettingsList.<StockChart>.<Settings>.<XAxis>.<Maximum>.Value = Data
                Case "StockChart:Settings:XAxis:AutoInterval"
                    StockChartSettingsList.<StockChart>.<Settings>.<XAxis>.<AutoInterval>.Value = Data
                Case "StockChart:Settings:XAxis:Interval"
                    StockChartSettingsList.<StockChart>.<Settings>.<XAxis>.<Interval>.Value = Data
                Case "StockChart:Settings:XAxis:AutoMajorGridInterval"
                    StockChartSettingsList.<StockChart>.<Settings>.<XAxis>.<AutoMajorGridInterval>.Value = Data
                Case "StockChart:Settings:XAxis:MajorGridInterval"
                    StockChartSettingsList.<StockChart>.<Settings>.<XAxis>.<MajorGridInterval>.Value = Data

                Case "StockChart:Settings:YAxis:TitleText"
                    StockChartSettingsList.<StockChart>.<Settings>.<YAxis>.<TitleText>.Value = Data
                Case "StockChart:Settings:YAxis:TitleFontName"
                    StockChartSettingsList.<StockChart>.<Settings>.<YAxis>.<TitleFontName>.Value = Data
                Case "StockChart:Settings:YAxis:TitleColor"
                    StockChartSettingsList.<StockChart>.<Settings>.<YAxis>.<TitleColor>.Value = Data
                Case "StockChart:Settings:YAxis:TitleSize"
                    StockChartSettingsList.<StockChart>.<Settings>.<YAxis>.<TitleSize>.Value = Data
                Case "StockChart:Settings:YAxis:TitleBold"
                    StockChartSettingsList.<StockChart>.<Settings>.<YAxis>.<TitleBold>.Value = Data
                Case "StockChart:Settings:YAxis:TitleItalic"
                    StockChartSettingsList.<StockChart>.<Settings>.<YAxis>.<TitleItalic>.Value = Data
                Case "StockChart:Settings:YAxis:TitleUnderline"
                    StockChartSettingsList.<StockChart>.<Settings>.<YAxis>.<TitleUnderline>.Value = Data
                Case "StockChart:Settings:YAxis:TitleStrikeout"
                    StockChartSettingsList.<StockChart>.<Settings>.<YAxis>.<TitleStrikeout>.Value = Data
                Case "StockChart:Settings:YAxis:TitleAlignment"
                    StockChartSettingsList.<StockChart>.<Settings>.<YAxis>.<TitleAlignment>.Value = Data
                Case "StockChart:Settings:YAxis:AutoMinimum"
                    StockChartSettingsList.<StockChart>.<Settings>.<YAxis>.<AutoMinimum>.Value = Data
                Case "StockChart:Settings:YAxis:Minimum"
                    StockChartSettingsList.<StockChart>.<Settings>.<YAxis>.<Minimum>.Value = Data
                Case "StockChart:Settings:YAxis:AutoMaximum"
                    StockChartSettingsList.<StockChart>.<Settings>.<YAxis>.<AutoMaximum>.Value = Data
                Case "StockChart:Settings:YAxis:Maximum"
                    StockChartSettingsList.<StockChart>.<Settings>.<YAxis>.<Maximum>.Value = Data

                Case "StockChart:Settings:YAxis:AutoInterval"
                    StockChartSettingsList.<StockChart>.<Settings>.<YAxis>.<AutoInterval>.Value = Data
                Case "StockChart:Settings:YAxis:Interval"
                    StockChartSettingsList.<StockChart>.<Settings>.<YAxis>.<Interval>.Value = Data
                Case "StockChart:Settings:YAxis:AutoMajorGridInterval"
                    StockChartSettingsList.<StockChart>.<Settings>.<YAxis>.<AutoMajorGridInterval>.Value = Data
                Case "StockChart:Settings:YAxis:MajorGridInterval"
                    StockChartSettingsList.<StockChart>.<Settings>.<YAxis>.<MajorGridInterval>.Value = Data

           'Point Chart instructions: ---------------------------------------------------------------------------------------------------
                Case "PointChart:Settings:Command"
                    Select Case Data
                        Case "ClearChart"
                            'ClearPointChartDefaults()
                            ClearPointChartSettingsList()
                        Case "OK"
                            'PointChartDefaults has been updated. Display in rtbPointChartDefaults.
                            'rtbPointChartDefaults.Text = PointChartSettingsList.ToString
                            'FormatXmlText(rtbPointChartDefaults)
                            XmlPointChartSettingsList.Rtf = XmlPointChartSettingsList.XmlToRtf(PointChartSettingsList.ToString, False)
                    End Select

                Case "PointChart:Settings:InputData:Type"
                    'PointChartDefaults.<PointChart>.<Settings>.<InputData>.<Type>.Value = Data
                    PointChartSettingsList.<PointChart>.<Settings>.<InputData>.<Type>.Value = Data
                Case "PointChart:Settings:InputData:DatabasePath"
                    'PointChartDefaults.<PointChart>.<Settings>.<InputData>.<DatabasePath>.Value = Data
                    PointChartSettingsList.<PointChart>.<Settings>.<InputData>.<DatabasePath>.Value = Data
                Case "PointChart:Settings:InputData:DataDescription"
                    PointChartSettingsList.<PointChart>.<Settings>.<InputData>.<DataDescription>.Value = Data
                Case "PointChart:Settings:InputData:DatabaseQuery"
                    PointChartSettingsList.<PointChart>.<Settings>.<InputData>.<DatabaseQuery>.Value = Data

                Case "PointChart:Settings:ChartProperties:SeriesName"
                    PointChartSettingsList.<StockChart>.<Settings>.<ChartProperties>.<SeriesName>.Value = Data

                Case "PointChart:Settings:ChartProperties:XValuesFieldName"
                    PointChartSettingsList.<StockChart>.<Settings>.<ChartProperties>.<XValuesFieldName>.Value = Data

                Case "PointChart:Settings:ChartProperties:YValuesFieldName"
                    PointChartSettingsList.<StockChart>.<Settings>.<ChartProperties>.<YValuesFieldName>.Value = Data

                Case "PointChart:Settings:ChartTitle:LabelName"
                    PointChartSettingsList.<PointChart>.<Settings>.<ChartTitle>.<LabelName>.Value = Data
                Case "PointChart:Settings:ChartTitle:Text"
                    PointChartSettingsList.<PointChart>.<Settings>.<ChartTitle>.<Text>.Value = Data
                Case "PointChart:Settings:ChartTitle:FontName"
                    PointChartSettingsList.<PointChart>.<Settings>.<ChartTitle>.<FontName>.Value = Data
                Case "PointChart:Settings:ChartTitle:Color"
                    PointChartSettingsList.<PointChart>.<Settings>.<ChartTitle>.<Color>.Value = Data
                Case "PointChart:Settings:ChartTitle:Size"
                    PointChartSettingsList.<PointChart>.<Settings>.<ChartTitle>.<Size>.Value = Data
                Case "PointChart:Settings:ChartTitle:Bold"
                    PointChartSettingsList.<PointChart>.<Settings>.<ChartTitle>.<Bold>.Value = Data
                Case "PointChart:Settings:ChartTitle:Italic"
                    PointChartSettingsList.<PointChart>.<Settings>.<ChartTitle>.<Italic>.Value = Data
                Case "PointChart:Settings:ChartTitle:Underline"
                    PointChartSettingsList.<PointChart>.<Settings>.<ChartTitle>.<Underline>.Value = Data
                Case "PointChart:Settings:ChartTitle:Strikeout"
                    PointChartSettingsList.<PointChart>.<Settings>.<ChartTitle>.<Strikeout>.Value = Data
                Case "PointChart:Settings:ChartTitle:Alignment"
                    PointChartSettingsList.<PointChart>.<Settings>.<ChartTitle>.<Alignment>.Value = Data

                Case "PointChart:Settings:XAxis:TitleText"
                    PointChartSettingsList.<PointChart>.<Settings>.<XAxis>.<TitleText>.Value = Data
                Case "PointChart:Settings:XAxis:TitleFontName"
                    PointChartSettingsList.<PointChart>.<Settings>.<XAxis>.<TitleFontName>.Value = Data
                Case "PointChart:Settings:XAxis:TitleColor"
                    PointChartSettingsList.<PointChart>.<Settings>.<XAxis>.<TitleColor>.Value = Data
                Case "PointChart:Settings:XAxis:TitleSize"
                    PointChartSettingsList.<PointChart>.<Settings>.<XAxis>.<TitleSize>.Value = Data
                Case "PointChart:Settings:XAxis:TitleBold"
                    PointChartSettingsList.<PointChart>.<Settings>.<XAxis>.<TitleBold>.Value = Data
                Case "PointChart:Settings:XAxis:TitleItalic"
                    PointChartSettingsList.<PointChart>.<Settings>.<XAxis>.<TitleItalic>.Value = Data
                Case "PointChart:Settings:XAxis:TitleUnderline"
                    PointChartSettingsList.<PointChart>.<Settings>.<XAxis>.<TitleUnderline>.Value = Data
                Case "PointChart:Settings:XAxis:TitleStrikeout"
                    PointChartSettingsList.<PointChart>.<Settings>.<XAxis>.<TitleStrikeout>.Value = Data
                Case "PointChart:Settings:XAxis:TitleAlignment"
                    PointChartSettingsList.<PointChart>.<Settings>.<XAxis>.<TitleAlignment>.Value = Data
                Case "PointChart:Settings:XAxis:AutoMinimum"
                    PointChartSettingsList.<PointChart>.<Settings>.<XAxis>.<AutoMinimum>.Value = Data
                Case "PointChart:Settings:XAxis:Minimum"
                    PointChartSettingsList.<PointChart>.<Settings>.<XAxis>.<Minimum>.Value = Data
                Case "PointChart:Settings:XAxis:AutoMaximum"
                    PointChartSettingsList.<PointChart>.<Settings>.<XAxis>.<AutoMaximum>.Value = Data
                Case "PointChart:Settings:XAxis:Maximum"
                    PointChartSettingsList.<PointChart>.<Settings>.<XAxis>.<Maximum>.Value = Data
                Case "PointChart:Settings:XAxis:AutoInterval"
                    PointChartSettingsList.<PointChart>.<Settings>.<XAxis>.<AutoInterval>.Value = Data
                Case "PointChart:Settings:XAxis:Interval"
                    PointChartSettingsList.<PointChart>.<Settings>.<XAxis>.<Interval>.Value = Data
                Case "PointChart:Settings:XAxis:AutoMajorGridInterval"
                    PointChartSettingsList.<PointChart>.<Settings>.<XAxis>.<AutoMajorGridInterval>.Value = Data
                Case "PointChart:Settings:XAxis:MajorGridInterval"
                    PointChartSettingsList.<PointChart>.<Settings>.<XAxis>.<MajorGridInterval>.Value = Data

                Case "PointChart:Settings:YAxis:TitleText"
                    PointChartSettingsList.<PointChart>.<Settings>.<YAxis>.<TitleText>.Value = Data
                Case "PointChart:Settings:YAxis:TitleFontName"
                    PointChartSettingsList.<PointChart>.<Settings>.<YAxis>.<TitleFontName>.Value = Data
                Case "PointChart:Settings:YAxis:TitleColor"
                    PointChartSettingsList.<PointChart>.<Settings>.<YAxis>.<TitleColor>.Value = Data
                Case "PointChart:Settings:YAxis:TitleSize"
                    PointChartSettingsList.<PointChart>.<Settings>.<YAxis>.<TitleSize>.Value = Data
                Case "PointChart:Settings:YAxis:TitleBold"
                    PointChartSettingsList.<PointChart>.<Settings>.<YAxis>.<TitleBold>.Value = Data
                Case "PointChart:Settings:YAxis:TitleItalic"
                    PointChartSettingsList.<PointChart>.<Settings>.<YAxis>.<TitleItalic>.Value = Data
                Case "PointChart:Settings:YAxis:TitleUnderline"
                    PointChartSettingsList.<PointChart>.<Settings>.<YAxis>.<TitleUnderline>.Value = Data
                Case "PointChart:Settings:YAxis:TitleStrikeout"
                    PointChartSettingsList.<PointChart>.<Settings>.<YAxis>.<TitleStrikeout>.Value = Data
                Case "PointChart:Settings:YAxis:TitleAlignment"
                    PointChartSettingsList.<PointChart>.<Settings>.<YAxis>.<TitleAlignment>.Value = Data
                Case "PointChart:Settings:YAxis:AutoMinimum"
                    PointChartSettingsList.<PointChart>.<Settings>.<YAxis>.<AutoMinimum>.Value = Data
                Case "PointChart:Settings:YAxis:Minimum"
                    PointChartSettingsList.<PointChart>.<Settings>.<YAxis>.<Minimum>.Value = Data
                Case "PointChart:Settings:YAxis:AutoMaximum"
                    PointChartSettingsList.<PointChart>.<Settings>.<YAxis>.<AutoMaximum>.Value = Data
                Case "PointChart:Settings:YAxis:Maximum"
                    PointChartSettingsList.<PointChart>.<Settings>.<YAxis>.<Maximum>.Value = Data
                Case "PointChart:Settings:YAxis:AutoInterval"
                    PointChartSettingsList.<PointChart>.<Settings>.<YAxis>.<AutoInterval>.Value = Data
                Case "PointChart:Settings:YAxis:Interval"
                    PointChartSettingsList.<PointChart>.<Settings>.<YAxis>.<Interval>.Value = Data
                Case "PointChart:Settings:YAxis:AutoMajorGridInterval"
                    PointChartSettingsList.<PointChart>.<Settings>.<YAxis>.<AutoMajorGridInterval>.Value = Data
                Case "PointChart:Settings:YAxis:MajorGridInterval"
                    PointChartSettingsList.<PointChart>.<Settings>.<YAxis>.<MajorGridInterval>.Value = Data

           'Get Share Information ====================================================

           'Get GICS List
                Case "GetGicsList"
                    If Data = "OK" Then
                        GetGicsList()
                    End If

          'Get company list in GICS group. Data contains the GICS code.
                Case "GetGicsCompanyList"
                    GetGicsCompanyList(Data)

           'Get company name. Data contains the ASX code.
                Case "GetCompanyName"
                    GetCompanyName(Data)



           'Startup Command Arguments ================================================
                Case "ProjectName"
                    If Project.OpenProject(Data) = True Then
                        ProjectSelected = True 'Project has been opened OK.
                    Else
                        ProjectSelected = False 'Project could not be opened.
                    End If

                Case "ProjectID"
                    Message.AddWarning("Add code to handle ProjectID parameter at StartUp!" & vbCrLf)
                'Note the AppNet will usually select a project using ProjectPath.

                Case "ProjectPath"
                    If Project.OpenProjectPath(Data) = True Then
                        ProjectSelected = True 'Project has been opened OK.

                    Else
                        ProjectSelected = False 'Project could not be opened.
                    End If

                Case "ConnectionName"
                    StartupConnectionName = Data
            '--------------------------------------------------------------------------

            'Application Information  =================================================
            'returned by client.GetAdvlNetworkAppInfoAsync()
                'Case "MessageServiceAppInfo:Name"
                ''The name of the Message Service Application. (Not used.)
                Case "AdvlNetworkAppInfo:Name"
                'The name of the Andorville™ Network Application. (Not used.)

                'Case "MessageServiceAppInfo:ExePath"
                '    'The executable file path of the Message Service Application.
                '    MsgServiceExePath = Data
                Case "AdvlNetworkAppInfo:ExePath"
                    'The executable file path of the Andorville™ Network Application.
                    AdvlNetworkExePath = Data

                'Case "MessageServiceAppInfo:Path"
                '    'The path of the Message Service Application (ComNet). (This is where an Application.Lock file will be found while ComNet is running.)
                '    MsgServiceAppPath = Data
                Case "AdvlNetworkAppInfo:Path"
                    'The path of the Andorville™ Network Application (ComNet). (This is where an Application.Lock file will be found while ComNet is running.)
                    AdvlNetworkAppPath = Data

            '---------------------------------------------------------------------------

            'Show Share Price Table ====================================================
                Case "ShowSharePriceTable:Name"
                    SharePricesFormNo = OpenOrAppendSharePricesView(Data) 'Open the Calculations View named Data. If not found, create a new view named Data.

                Case "ShowSharePriceTable:Command"
                    Select Case Data
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
                        SharePricesFormList(SharePricesFormNo).Query = Data
                    Else
                        SharePricesFormList(SharePricesFormNo).Query = Data
                    End If

            '---------------------------------------------------------------------------


            'Show Share Calculations Table =============================================
                Case "ShowCalculationsTable:Name"
                    CalculationsFormNo = OpenOrAppendCalculationsView(Data) 'Open the Calculations View named Data. If not found, create a new view named Data.

                Case "ShowCalculationsTable:Command"
                    Select Case Data
                        Case "Apply"
                            If CalculationsFormNo = -1 Then
                                Message.AddWarning("The Calculations Form Number is not known." & vbCrLf)
                            Else
                                CalculationsFormList(CalculationsFormNo).ApplyQuery
                            End If
                        Case "OpenNewForm"
                            Debug.Print("Starting: CalculationsFormNo = AppendCalculationsView()")
                            CalculationsFormNo = AppendCalculationsView()

                    End Select

                Case "ShowCalculationsTable:Query"
                    Debug.Print("ShowCalculationsTable:Query: CalculationsFormNo: " & CalculationsFormNo)
                    If CalculationsFormNo = -1 Then
                        'Message.AddWarning("The Calculations Form Number is not known." & vbCrLf)
                        CalculationsFormNo = AppendCalculationsView()
                        CalculationsFormList(CalculationsFormNo).Query = Data
                    Else
                        CalculationsFormList(CalculationsFormNo).Query = Data
                    End If

            '---------------------------------------------------------------------------


            'Show Share Price Chart ====================================================
                Case "ShowSharePriceChart:Query"
                    txtSPChartQuery.Text = Data 'Specify the Query used to extract the data to chart.
                    UpdateChartSharePricesTab()

                Case "ShowSharePriceChart:SeriesName"
                    txtSeriesName.Text = Data 'Set the Series Name -  the name of the series of points being charted.

                Case "ShowSharePriceChart:ChartTitle"
                    txtChartTitle.Text = Data 'Set the Chart Title.

                Case "ShowSharePriceChart:Command"
                    Select Case Data
                        Case "Apply"
                            DisplayStockChart()
                        Case Else
                            Message.AddWarning("Unknown ShowSharePriceChart command: " & Data & vbCrLf)
                    End Select

            '---------------------------------------------------------------------------


            'Update Project List =======================================================
                Case "Update:ProjectList:Project:Name"
                    ProjListNo = Proj.List.Count
                    Proj.List.Add(New ProjSummary)
                    Proj.List(ProjListNo).Name = Data
                    'Message.Add("Proj.List.Count = " & Proj.List.Count & vbCrLf)
                Case "Update:ProjectList:Project:Description"
                    Proj.List(ProjListNo).Description = Data
                Case "Update:ProjectList:Project:ProjectNetworkName"
                    Proj.List(ProjListNo).ProNetName = Data
                Case "Update:ProjectList:Project:ID"
                    Proj.List(ProjListNo).ID = Data
                Case "Update:ProjectList:Project:Type"
                    Select Case Data
                        Case "Archive"
                            Proj.List(ProjListNo).Type = ADVL_Utilities_Library_1.Project.Types.Archive
                        Case "Directory"
                            Proj.List(ProjListNo).Type = ADVL_Utilities_Library_1.Project.Types.Directory
                        Case "Hybrid"
                            Proj.List(ProjListNo).Type = ADVL_Utilities_Library_1.Project.Types.Hybrid
                        Case "None"
                            Proj.List(ProjListNo).Type = ADVL_Utilities_Library_1.Project.Types.None
                        Case Else
                            Message.AddWarning("Unknown project type: " & Data & vbCrLf)
                    End Select
                Case "Update:ProjectList:Project:Path"
                    Proj.List(ProjListNo).Path = Data
                Case "Update:ProjectList:Project:ApplicationName"
                    Proj.List(ProjListNo).ApplicationName = Data
                Case "Update:ProjectList:Project:ParentProjectName"
                    Proj.List(ProjListNo).ParentProjectName = Data
                Case "Update:ProjectList:Project:ParentProjectID"
                    Proj.List(ProjListNo).ParentProjectID = Data

            '---------------------------------------------------------------------------


               'Message Window Instructions  ==============================================
                Case "MessageWindow:Left"
                    If IsNothing(Message.MessageForm) Then
                        Message.ApplicationName = ApplicationInfo.Name
                        Message.SettingsLocn = Project.SettingsLocn
                        Message.Show()
                    End If
                    Message.MessageForm.Left = Data
                Case "MessageWindow:Top"
                    If IsNothing(Message.MessageForm) Then
                        Message.ApplicationName = ApplicationInfo.Name
                        Message.SettingsLocn = Project.SettingsLocn
                        Message.Show()
                    End If
                    Message.MessageForm.Top = Data
                Case "MessageWindow:Width"
                    If IsNothing(Message.MessageForm) Then
                        Message.ApplicationName = ApplicationInfo.Name
                        Message.SettingsLocn = Project.SettingsLocn
                        Message.Show()
                    End If
                    Message.MessageForm.Width = Data
                Case "MessageWindow:Height"
                    If IsNothing(Message.MessageForm) Then
                        Message.ApplicationName = ApplicationInfo.Name
                        Message.SettingsLocn = Project.SettingsLocn
                        Message.Show()
                    End If
                    Message.MessageForm.Height = Data
                Case "MessageWindow:Command"
                    Select Case Data
                        Case "BringToFront"
                            If IsNothing(Message.MessageForm) Then
                                Message.ApplicationName = ApplicationInfo.Name
                                Message.SettingsLocn = Project.SettingsLocn
                                Message.Show()
                            End If
                            'Message.MessageForm.BringToFront()
                            Message.MessageForm.Activate()
                            Message.MessageForm.TopMost = True
                            Message.MessageForm.TopMost = False
                        Case "SaveSettings"
                            Message.MessageForm.SaveFormSettings()
                    End Select

            '---------------------------------------------------------------------------

                'Command to bring the Application window to the front:
                Case "ApplicationWindow:Command"
                    Select Case Data
                        Case "BringToFront"
                            Me.Activate()
                            Me.TopMost = True
                            Me.TopMost = False
                    End Select


                Case "EndOfSequence"
                    'End of Information Sequence reached.
                    'Add Status OK element at the end of the sequence:
                    Dim statusOK As New XElement("Status", "OK") 'Add Status OK element at the end of the sequence
                    xlocns(xlocns.Count - 1).Add(statusOK)
                    'xlocns(0).Add(statusOK) 'TESTING - Check if this puts the instruction at the main location.

                    ''DEPRECATED:
                    ''Add the final OnCompletion instruction:
                    ''Dim onCompletion As New XElement("OnCompletion", CompletionInstruction) '
                    'Dim onCompletion As New XElement("OnCompletion", XCompletionInstruction) '
                    'xlocns(xlocns.Count - 1).Add(onCompletion)
                    ''CompletionInstruction = "Stop" 'Reset the Completion Instruction
                    'XCompletionInstruction = "Stop" 'Reset the Completion Instruction
                    ''RunningXMsg = False 'Added 12Jan2020 (DOESNT WORK CORRECTLY.)

                    ''DEPRECATED:
                    'Select Case CompletionInstruction
                    '    Case ""
                    '        'No instructions.
                    '    Case "UpdateStockChartProjects"
                    '        UpdateStockChartProjects()
                    '        Message.Add("Stock Chart projects updated!" & vbCrLf)
                    '    Case "UpdateChartProjLists"
                    '        UpdateChartProjLists()
                    '    Case Else
                    '        Message.AddWarning("Unknown Completion Instruction: " & CompletionInstruction & vbCrLf)
                    'End Select

                    Select Case EndInstruction
                        Case "Stop"
                            'No instructions.
                        Case "UpdateStockChartProjects"
                            UpdateStockChartProjects()
                            Message.Add("Stock Chart projects updated!" & vbCrLf)
                        Case "UpdateChartProjLists"
                            UpdateChartProjLists()
                        Case Else
                            Message.AddWarning("Unknown End Instruction: " & EndInstruction & vbCrLf)
                    End Select
                    EndInstruction = "Stop"

                    ''REPLACED BY:
                    'Select Case EndInstructionData
                    '    Case "Stop"
                    '        'Do nothing

                    '        'Process other EndInstructions here.
                    '        If EndInstructionLocn = "" Then
                    '            'Process EndInstructionInfo here.

                    '        Else
                    '            'Send the EndInstructionInfo to the EndInstructionLocn:
                    '            InstructionParams.Data = EndInstructionInfo
                    '            InstructionParams.Locn = EndInstructionLocn
                    '            If bgwRunInstruction.IsBusy Then
                    '                Message.AddWarning("Send Run Instruction backgroundworker is busy." & vbCrLf)
                    '            Else
                    '                bgwRunInstruction.RunWorkerAsync(InstructionParams)
                    '            End If

                    '        End If

                    '    Case Else
                    '        Message.AddWarning("Unknown End Instruction: " & EndInstructionInfo & vbCrLf)
                    'End Select

                    'EndInstructionInfo = "Stop" 'Reset the End Instruction Data
                    'EndInstructionLocn = ""     'Reset the End Instruction Locn

                    ''TESTING: - OK
                    ''Add the final EndInstruction:
                    'Dim xEndInstruction As New XElement("EndInstruction", EndInstruction)
                    'xlocns(xlocns.Count - 1).Add(xEndInstruction)
                    'EndInstruction = "Stop" 'Reset the End Instruction

                    ''Final Version:
                    ''Add the final EndInstruction:
                    'Dim xEndInstruction As New XElement("EndInstruction", OnCompletionInstruction)
                    'xlocns(xlocns.Count - 1).Add(xEndInstruction)
                    ''xlocns(0).Add(xEndInstruction) 'TESTING - Check if this puts the instruction at the main location.
                    'OnCompletionInstruction = "Stop" 'Reset the OnCompletion Instruction

                    'Add the final EndInstruction:
                    If OnCompletionInstruction = "Stop" Then
                        'Final EndInstruction is not required.
                    Else
                        Dim xEndInstruction As New XElement("EndInstruction", OnCompletionInstruction)
                        xlocns(xlocns.Count - 1).Add(xEndInstruction)
                        OnCompletionInstruction = "Stop" 'Reset the OnCompletion Instruction
                    End If

                Case Else
                    Message.AddWarning("Unknown location: " & Locn & vbCrLf)
                    Message.AddWarning("            data: " & Data & vbCrLf & vbCrLf)
            End Select
        End If
    End Sub

    'Private Sub ClearStockChartDefaults()
    Private Sub ClearStockChartSettingsList()
        'Clear the settings in the StockChartDefaults XDocument.

        'StockChartDefaults = <?xml version="1.0" encoding="utf-8"?>
        StockChartSettingsList = <?xml version="1.0" encoding="utf-8"?>
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

    'Private Sub ClearPointChartDefaults()
    Private Sub ClearPointChartSettingsList()
        'Clear the settings in the PointChartDefaults XDocument.

        'PointChartDefaults = <?xml version="1.0" encoding="utf-8"?>
        PointChartSettingsList = <?xml version="1.0" encoding="utf-8"?>
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

        If System.IO.File.Exists(SharePriceDbPath) Then
            'Share Price Database file exists.
        Else
            'Share Price Database file does not exist!
            Message.AddWarning("The share price database was not found: " & SharePriceDbPath & vbCrLf)
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

        If System.IO.File.Exists(SharePriceDbPath) Then
            'Share Price Database file exists.
        Else
            'Share Price Database file does not exist!
            Message.AddWarning("The share price database was not found: " & SharePriceDbPath & vbCrLf)
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

        If System.IO.File.Exists(SharePriceDbPath) Then
            'Share Price Database file exists.
        Else
            'Share Price Database file does not exist!
            Message.AddWarning("The share price database was not found: " & SharePriceDbPath & vbCrLf)
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

    'REPLACED 23Jul19 - by bgwSendMessage()
    'Private Sub SendMessage()
    '    'Code used to send a message after a timer delay.
    '    'The message destination is stored in MessageDest
    '    'The message text is stored in MessageText
    '    Timer1.Interval = 100 '100ms delay
    '    Timer1.Enabled = True 'Start the timer.
    'End Sub

    'REPLACED 23Jul19 - by bgwSendMessage()
    'Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick

    '    If IsNothing(client) Then
    '        Message.AddWarning("No client connection available!" & vbCrLf)
    '    Else
    '        If client.State = ServiceModel.CommunicationState.Faulted Then
    '            Message.AddWarning("client state is faulted. Message not sent!" & vbCrLf)
    '        Else
    '            Try
    '                'client.SendMessage(ClientAppNetName, ClientConnName, MessageText) 'Added 2Feb19
    '                client.SendMessage(ClientProNetName, ClientConnName, MessageText)
    '                MessageText = "" 'Clear the message after it has been sent.
    '                ClientAppName = "" 'Clear the Client Application Name after the message has been sent.
    '                ClientConnName = "" 'Clear the Client Application Name after the message has been sent.
    '                xlocns.Clear()
    '            Catch ex As Exception
    '                Message.AddWarning("Error sending message: " & ex.Message & vbCrLf)
    '            End Try
    '        End If
    '    End If

    '    'Stop timer:
    '    Timer1.Enabled = False
    'End Sub

#End Region 'Process XMessages --------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Settings Tab" '==============================================================================================================================================================

    Private Sub btnFindSharePriceDatabase_Click(sender As Object, e As EventArgs) Handles btnFindSharePriceDatabase.Click
        'Find a Share Price database:

        If SharePriceDbPath = "" Then
            'OpenFileDialog1.InitialDirectory = System.Environment.SpecialFolder.MyDocuments
            If Project.Type = ADVL_Utilities_Library_1.Project.Types.Archive Then
                OpenFileDialog1.InitialDirectory = IO.Path.GetDirectoryName(Project.Path)
            Else
                OpenFileDialog1.InitialDirectory = Project.Path
            End If
            OpenFileDialog1.Filter = "Database |*.accdb; *.mdb"
            OpenFileDialog1.FileName = ""
        Else
            Dim fInfo As New System.IO.FileInfo(SharePriceDbPath)
            OpenFileDialog1.InitialDirectory = fInfo.DirectoryName
            OpenFileDialog1.Filter = "Database |*.accdb; *.mdb"
            OpenFileDialog1.FileName = fInfo.Name
        End If

        If OpenFileDialog1.ShowDialog() = vbOK Then
            'SharePriceDbPath = OpenFileDialog1.FileName
            dgvSPDatabase.Rows.Add()
            Dim RowCount = dgvSPDatabase.Rows.Count
            dgvSPDatabase.Rows(RowCount - 1).Cells(4).Value = OpenFileDialog1.FileName

            Dim DirectoryPath As String = System.IO.Path.GetDirectoryName(OpenFileDialog1.FileName)
            If DirectoryPath = Project.Path Then
                dgvSPDatabase.Rows(RowCount - 1).Cells(3).Value = "Project"
            ElseIf DirectoryPath = Project.SettingsLocn.Path Then
                dgvSPDatabase.Rows(RowCount - 1).Cells(3).Value = "Settings"
            ElseIf DirectoryPath = Project.DataLocn.Path Then
                dgvSPDatabase.Rows(RowCount - 1).Cells(3).Value = "Data"
            ElseIf DirectoryPath = Project.SystemLocn.Path Then
                dgvSPDatabase.Rows(RowCount - 1).Cells(3).Value = "System"
            Else
                dgvSPDatabase.Rows(RowCount - 1).Cells(3).Value = "External"
            End If

            Dim DbFileName As String = System.IO.Path.GetFileName(OpenFileDialog1.FileName)
            dgvSPDatabase.Rows(RowCount - 1).Cells(2).Value = DbFileName

            ''Set the corresponding Project Parameter:
            'Project.AddParameter("SharePriceDatabasePath", SharePriceDbPath, "The path of the Share Price database.")
            'Project.SaveParameters()
        End If
    End Sub

    Private Sub btnFindFinancialsDatabase_Click(sender As Object, e As EventArgs) Handles btnFindFinancialsDatabase.Click
        'Find a Historical Financials database:

        If FinancialsDbPath = "" Then
            'OpenFileDialog1.InitialDirectory = System.Environment.SpecialFolder.MyDocuments
            If Project.Type = ADVL_Utilities_Library_1.Project.Types.Archive Then
                OpenFileDialog1.InitialDirectory = IO.Path.GetDirectoryName(Project.Path)
            Else
                OpenFileDialog1.InitialDirectory = Project.Path
            End If
            OpenFileDialog1.Filter = "Database |*.accdb; *.mdb"
            OpenFileDialog1.FileName = ""
        Else
            Dim fInfo As New System.IO.FileInfo(FinancialsDbPath)
            OpenFileDialog1.InitialDirectory = fInfo.DirectoryName
            OpenFileDialog1.Filter = "Database |*.accdb; *.mdb"
            OpenFileDialog1.FileName = fInfo.Name
        End If

        If OpenFileDialog1.ShowDialog() = vbOK Then
            'FinancialsDbPath = OpenFileDialog1.FileName

            dgvFinDatabase.Rows.Add()
            Dim RowCount = dgvFinDatabase.Rows.Count
            dgvFinDatabase.Rows(RowCount - 1).Cells(4).Value = OpenFileDialog1.FileName

            Dim DirectoryPath As String = System.IO.Path.GetDirectoryName(OpenFileDialog1.FileName)
            If DirectoryPath = Project.Path Then
                dgvFinDatabase.Rows(RowCount - 1).Cells(3).Value = "Project"
            ElseIf DirectoryPath = Project.SettingsLocn.Path Then
                dgvFinDatabase.Rows(RowCount - 1).Cells(3).Value = "Settings"
            ElseIf DirectoryPath = Project.DataLocn.Path Then
                dgvFinDatabase.Rows(RowCount - 1).Cells(3).Value = "Data"
            ElseIf DirectoryPath = Project.SystemLocn.Path Then
                dgvFinDatabase.Rows(RowCount - 1).Cells(3).Value = "System"
            Else
                dgvFinDatabase.Rows(RowCount - 1).Cells(3).Value = "External"
            End If

            Dim DbFileName As String = System.IO.Path.GetFileName(OpenFileDialog1.FileName)
            dgvFinDatabase.Rows(RowCount - 1).Cells(2).Value = DbFileName


            ''Set the corresponding Project Parameter:
            'Project.AddParameter("FinancialsDatabasePath", FinancialsDbPath, "The path of the Historical Financials database.")
            'Project.SaveParameters()
        End If
    End Sub

    Private Sub btnFindCalcsDatabase_Click(sender As Object, e As EventArgs) Handles btnFindCalcsDatabase.Click
        'Find a Calculations database:

        If CalculationsDbPath = "" Then
            'OpenFileDialog1.InitialDirectory = System.Environment.SpecialFolder.MyDocuments
            If Project.Type = ADVL_Utilities_Library_1.Project.Types.Archive Then
                OpenFileDialog1.InitialDirectory = IO.Path.GetDirectoryName(Project.Path)
            Else
                OpenFileDialog1.InitialDirectory = Project.Path
            End If
            OpenFileDialog1.Filter = "Database |*.accdb; *.mdb"
            OpenFileDialog1.FileName = ""
        Else
            Dim fInfo As New System.IO.FileInfo(CalculationsDbPath)
            OpenFileDialog1.InitialDirectory = fInfo.DirectoryName
            OpenFileDialog1.Filter = "Database |*.accdb; *.mdb"
            OpenFileDialog1.FileName = fInfo.Name
        End If

        If OpenFileDialog1.ShowDialog() = vbOK Then
            'CalculationsDbPath = OpenFileDialog1.FileName
            dgvCalcDatabase.Rows.Add()
            Dim RowCount = dgvCalcDatabase.Rows.Count
            dgvCalcDatabase.Rows(RowCount - 1).Cells(4).Value = OpenFileDialog1.FileName

            Dim DirectoryPath As String = System.IO.Path.GetDirectoryName(OpenFileDialog1.FileName)
            If DirectoryPath = Project.Path Then
                dgvCalcDatabase.Rows(RowCount - 1).Cells(3).Value = "Project"
            ElseIf DirectoryPath = Project.SettingsLocn.Path Then
                dgvCalcDatabase.Rows(RowCount - 1).Cells(3).Value = "Settings"
            ElseIf DirectoryPath = Project.DataLocn.Path Then
                dgvCalcDatabase.Rows(RowCount - 1).Cells(3).Value = "Data"
            ElseIf DirectoryPath = Project.SystemLocn.Path Then
                dgvCalcDatabase.Rows(RowCount - 1).Cells(3).Value = "System"
            Else
                dgvCalcDatabase.Rows(RowCount - 1).Cells(3).Value = "External"
            End If

            Dim DbFileName As String = System.IO.Path.GetFileName(OpenFileDialog1.FileName)
            dgvCalcDatabase.Rows(RowCount - 1).Cells(2).Value = DbFileName


            ''Set the corresponding Project Parameter:
            'Project.AddParameter("CalculationsDatabasePath", CalculationsDbPath, "The path of the Calculations database.")
            'Project.SaveParameters()
        End If
    End Sub

    Private Sub btnFindNewsDirectory_Click(sender As Object, e As EventArgs) Handles btnFindNewsDirectory.Click
        'Find a News directory:

    End Sub

#End Region 'Settings Tab -------------------------------------------------------------------------------------------------------------------------------------------------------------

#Region " View Data Tab" '=============================================================================================================================================================

#Region " View Share Prices Sub Tab" '=================================================================================================================================================

    'Public Sub UpdateSharePricesDataDescr(ByVal IndexNo As Integer, ByVal Description As String)
    Public Sub UpdateSharePricesDataName(ByVal IndexNo As Integer, ByVal Name As String)
        'Set the Share Prices data name in lstSharePrices list box.
        '  IndexNo is the index number of the item in the list.

        Dim ListCount As Integer = lstSharePrices.Items.Count

        If IndexNo >= ListCount Then
            'Pad out entries in lstSharePrices list box:
            Dim I As Integer
            For I = ListCount To IndexNo
                lstSharePrices.Items.Add("")
            Next
        End If
        lstSharePrices.Items(IndexNo) = Name
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
            'SharePricesFormList(1).DataSummary = "New Share Price Data View"
            SharePricesFormList(1).DataName = "New Share Price Data View"
            SharePricesFormList(1).Version = "Version 1"
        Else
            If SelectedIndex >= 0 Then
                If SelectedIndex = NViews - 1 Then 'Last item selected
                    'Add the new View to the end of the list.
                    lstSharePrices.Items.Add("")
                    Dim NewSettings As New DataViewSettings
                    SharePricesSettings.List.Add(NewSettings)
                    OpenSharePricesFormNo(NViews)
                    'SharePricesFormList(NViews).DataSummary = "New Share Price Data View"
                    SharePricesFormList(NViews).DataName = "New Share Price Data View"
                    SharePricesFormList(NViews).Version = "Version 1"
                Else
                    lstSharePrices.Items.Insert(SelectedIndex + 1, "")
                    'Insert a new Settings entry in SharePricesSettings:
                    Dim NewSettings As New DataViewSettings
                    SharePricesSettings.List.Insert(SelectedIndex + 1, NewSettings)
                    OpenSharePricesFormNo(SelectedIndex + 1)
                    'SharePricesFormList(SelectedIndex + 1).DataSummary = "New Share Price Data View"
                    SharePricesFormList(SelectedIndex + 1).DataName = "New Share Price Data View"
                    SharePricesFormList(SelectedIndex + 1).Version = "Version 1"
                End If
            Else
                'No item selected
                'Add the new View to the end of the list.
                lstSharePrices.Items.Add("")
                Dim NewSettings As New DataViewSettings
                SharePricesSettings.List.Add(NewSettings)
                OpenSharePricesFormNo(NViews + 1)
                'SharePricesFormList(NViews + 1).DataSummary = "New Share Price Data View"
                SharePricesFormList(NViews + 1).DataName = "New Share Price Data View"
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
        'SharePricesFormList(NViews + 1).DataSummary = "New Share Price Data View"
        SharePricesFormList(NViews + 1).DataName = "New Share Price Data View"
        SharePricesFormList(NViews + 1).Version = "Version 1"
        Return NViews + 1
    End Function

    Private Function AppendCalculationsView() As Integer
        'Append a Calculations view to the list.
        'Return the Form Number of the Data View

        Dim NViews As Integer = lstCalculations.Items.Count 'The number of views in the Calculations list.
        Debug.Print("NViews: " & NViews)
        'lstCalculations.Items.Add("") 'Add an entry with a blank name. - Blank entries can be removed when the application is closed.
        'lstCalculations.Items.Add("Temp") 'Add an entry with a Temp name. - Temp entries can be removed when the application is closed.
        lstCalculations.Items.Add("New Calculations Data View") 'Add an entry.
        Dim NewSettings As New DataViewSettings
        CalculationsSettings.List.Add(NewSettings)
        Debug.Print("Finished: CalculationsSettings.List.Add(NewSettings)")
        OpenCalculationsFormNo(NViews + 1)
        Debug.Print("Finished: OpenCalculationsFormNo(NViews + 1)")
        'CalculationsFormList(NViews + 1).DataSummary = "New Calculations Data View"
        CalculationsFormList(NViews + 1).DataName = "New Calculations Data View"
        CalculationsFormList(NViews + 1).Version = "Version 1"
        Return NViews + 1
    End Function

    'Private Function AppendSharePricesView() As Integer
    '    'Append a SharePrices view to the list.
    '    'Return the Form Number of the Data View

    '    Dim NViews As Integer = lstSharePrices.Items.Count 'The number of views in the SharePrices list.
    '    'Debug.Print("NViews: " & NViews)
    '    lstSharePrices.Items.Add("New SharePrices Data View") 'Add an entry.
    '    Dim NewSettings As New DataViewSettings
    '    SharePricesSettings.List.Add(NewSettings)
    '    'Debug.Print("Finished: CalculationsSettings.List.Add(NewSettings)")
    '    OpenSharePricesFormNo(NViews + 1)
    '    'Debug.Print("Finished: OpenCalculationsFormNo(NViews + 1)")
    '    SharePricesFormList(NViews + 1).DataName = "New SharePrices Data View"
    '    SharePricesFormList(NViews + 1).Version = "Version 1"
    '    Return NViews + 1
    'End Function



    Private Function AppendCalculationsView(ByVal ViewName As String) As Integer
        'Append a Calculations view with the name ViewName to the list.
        'Return the Form Number of the Data View

        Dim NViews As Integer = lstCalculations.Items.Count 'The number of views in the Calculations list.

        lstCalculations.Items.Add(ViewName) 'Add an entry with the name ViewName.
        Dim NewSettings As New DataViewSettings
        CalculationsSettings.List.Add(NewSettings)
        OpenCalculationsFormNo(NViews + 1)
        CalculationsFormList(NViews + 1).DataName = ViewName
        CalculationsFormList(NViews + 1).Version = "Version 1"
        Return NViews + 1

    End Function

    Private Function AppendSharePricesView(ByVal ViewName As String) As Integer
        'Append a SharePrices view with the name ViewName to the list.
        'Return the Form Number of the Data View

        Dim NViews As Integer = lstSharePrices.Items.Count 'The number of views in the SharePrices list.

        lstSharePrices.Items.Add(ViewName) 'Add an entry with the name ViewName.
        Dim NewSettings As New DataViewSettings
        SharePricesSettings.List.Add(NewSettings)
        OpenSharePricesFormNo(NViews + 1)
        SharePricesFormList(NViews + 1).DataName = ViewName
        SharePricesFormList(NViews + 1).Version = "Version 1"
        Return NViews + 1
    End Function

    Private Function OpenOrAppendCalculationsView(ByVal ViewName As String) As Integer
        'Open the Calculations view named ViewName.
        'If not found, create a new view named ViewName.
        'Return the form number of the view.

        Dim NViews As Integer = lstCalculations.Items.Count 'The number of views in the Calculations list.
        Dim I As Integer
        Dim FoundViewIndex As Integer = -1

        'Search the Calculations Views List for a View with name ViewName:
        For I = 0 To NViews - 1
            If lstCalculations.Items(I).ToString = ViewName Then
                FoundViewIndex = I
                Exit For
            End If
        Next

        If FoundViewIndex = -1 Then 'Create a new Calculations View:
            Return AppendCalculationsView(ViewName)
        Else 'Open the existing Calculations View named ViewName:
            OpenCalculationsFormNo(FoundViewIndex)
            Return FoundViewIndex
        End If

    End Function

    Private Function OpenOrAppendSharePricesView(ByVal ViewName As String) As Integer
        'Open the SharePrices view named ViewName.
        'If not found, create a new view named ViewName.
        'Return the form number of the view.

        Dim NViews As Integer = lstSharePrices.Items.Count 'The number of views in the SharePrices list.
        Dim I As Integer
        Dim FoundViewIndex As Integer = -1

        'Search the SharePrices Views List for a View with name ViewName:
        For I = 0 To NViews - 1
            If lstSharePrices.Items(I).ToString = ViewName Then
                FoundViewIndex = I
                Exit For
            End If
        Next

        If FoundViewIndex = -1 Then 'Create a new SharePrices View:
            Return AppendSharePricesView(ViewName)
        Else 'Open the existing SharePrices View named ViewName:
            OpenSharePricesFormNo(FoundViewIndex)
            Return FoundViewIndex
        End If
    End Function

    Private Sub btnDeleteViewSP_Click(sender As Object, e As EventArgs) Handles btnDeleteViewSP.Click
        'Delete selected view

        Dim SelectedIndex As Integer = lstSharePrices.SelectedIndex 'The index of the selected view.

        If SelectedIndex = -1 Then
            Message.AddWarning("No item selected." & vbCrLf)
            Exit Sub
        End If

        Dim NViews As Integer = lstSharePrices.Items.Count 'The number of views in the Share Prices list.

        If SharePricesFormList.Count < SelectedIndex + 1 Then 'The Share Price data view is not being displayed.
            lstSharePrices.Items.RemoveAt(SelectedIndex) 'Remove the entry from the list displayed on the form.
            SharePricesSettings.List.RemoveAt(SelectedIndex)  'Delete the entry in SharePricesSettings
            SharePricesSettings.LastEditDate = Now
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
            SharePricesSettings.LastEditDate = Now
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
                'Zip.SelectFileForm.ApplicationName = Project.ApplicationName
                Zip.SelectFileForm.ApplicationName = Project.Application.Name
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

    'Public Sub UpdateFinancialsDataDescr(ByVal IndexNo As Integer, ByVal Description As String)
    Public Sub UpdateFinancialsDataName(ByVal IndexNo As Integer, ByVal Name As String)
        'Set the Financials data name in lstFinancials list box.
        '  IndexNo is the index number of the item in the list.

        Dim ListCount As Integer = lstFinancials.Items.Count

        If IndexNo >= ListCount Then
            'Pad out entries in lstFinancials list box:
            Dim I As Integer
            For I = ListCount To IndexNo
                lstFinancials.Items.Add("")
            Next
        End If
        lstFinancials.Items(IndexNo) = Name
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
            'FinancialsFormList(1).DataSummary = "New Financials Data View"
            FinancialsFormList(1).DataName = "New Financials Data View"
            FinancialsFormList(1).Version = "Version 1"
        Else
            If SelectedIndex >= 0 Then
                If SelectedIndex = NViews - 1 Then 'Last item selected
                    'Add the new View to the end of the list.
                    lstFinancials.Items.Add("")
                    Dim NewSettings As New DataViewSettings
                    FinancialsSettings.List.Add(NewSettings)
                    OpenFinancialsFormNo(NViews)
                    'FinancialsFormList(NViews).DataSummary = "New Financials Data View"
                    FinancialsFormList(NViews).DataName = "New Financials Data View"
                    FinancialsFormList(NViews).Version = "Version 1"
                Else
                    lstFinancials.Items.Insert(SelectedIndex + 1, "")
                    'Insert a new Settings entry in FinancialSettings:
                    Dim NewSettings As New DataViewSettings
                    FinancialsSettings.List.Insert(SelectedIndex + 1, NewSettings)
                    OpenFinancialsFormNo(SelectedIndex + 1)
                    'FinancialsFormList(SelectedIndex + 1).DataSummary = "New Financials Data View"
                    FinancialsFormList(SelectedIndex + 1).DataName = "New Financials Data View"
                    FinancialsFormList(SelectedIndex + 1).Version = "Version 1"
                End If
            Else
                'No item selected
                'Add the new View to the end of the list.
                lstFinancials.Items.Add("")
                Dim NewSettings As New DataViewSettings
                FinancialsSettings.List.Add(NewSettings)
                OpenFinancialsFormNo(NViews + 1)
                'FinancialsFormList(NViews + 1).DataSummary = "New Financials Data View"
                FinancialsFormList(NViews + 1).DataName = "New Financials Data View"
                FinancialsFormList(NViews + 1).Version = "Version 1"
            End If
        End If
    End Sub



    Private Sub btnDeleteViewFin_Click(sender As Object, e As EventArgs) Handles btnDeleteViewFin.Click
        'Delete selected view

        Dim SelectedIndex As Integer = lstFinancials.SelectedIndex 'The index of the selected view.

        If SelectedIndex = -1 Then
            Message.AddWarning("No item selected." & vbCrLf)
            Exit Sub
        End If

        Dim NViews As Integer = lstFinancials.Items.Count 'The number of views in the Financials list.

        If FinancialsFormList.Count < SelectedIndex + 1 Then 'The SFinancials data view is not being displayed.
            lstFinancials.Items.RemoveAt(SelectedIndex) 'Remove the entry from the list displayed on the form.
            FinancialsSettings.List.RemoveAt(SelectedIndex) 'Delete the entry in FinancialsSettings
            FinancialsSettings.LastEditDate = Now
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
            FinancialsSettings.LastEditDate = Now
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
                'Zip.SelectFileForm.ApplicationName = Project.ApplicationName
                Zip.SelectFileForm.ApplicationName = Project.Application.Name
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

    'Public Sub UpdateCalculationsDataDescr(ByVal IndexNo As Integer, ByVal Description As String)
    Public Sub UpdateCalculationsDataName(ByVal IndexNo As Integer, ByVal Name As String)
        'Set the Calculations data name in lstCalculations list box.
        '  IndexNo is the index number of the item in the list.

        Dim ListCount As Integer = lstCalculations.Items.Count

        If IndexNo >= ListCount Then
            'Pad out entries in lstCalculations list box:
            Dim I As Integer
            For I = ListCount To IndexNo
                lstCalculations.Items.Add("")
            Next
        End If
        lstCalculations.Items(IndexNo) = Name
    End Sub

    Private Sub btnDeleteViewCalcs_Click(sender As Object, e As EventArgs) Handles btnDeleteViewCalcs.Click
        'Delete selected view

        Dim SelectedIndex As Integer = lstCalculations.SelectedIndex 'The index of the selected view.

        If SelectedIndex = -1 Then
            Message.AddWarning("No item selected." & vbCrLf)
            Exit Sub
        End If

        Dim NViews As Integer = lstCalculations.Items.Count 'The number of views in the Calculations list.

        'If CalculationsDataViewList.Count < SelectedIndex + 1 Then
        If CalculationsFormList.Count < SelectedIndex + 1 Then 'The Calculations data view is not being displayed.
            lstCalculations.Items.RemoveAt(SelectedIndex) 'Remove the entry from the list displayed on the form.
            CalculationsSettings.List.RemoveAt(SelectedIndex) 'te the entry in CalculationsSettings
            CalculationsSettings.LastEditDate = Now
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
            CalculationsSettings.LastEditDate = Now
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
            'CalculationsFormList(1).DataSummary = "New Calculations Data View"
            CalculationsFormList(1).DataName = "New Calculations Data View"
            CalculationsFormList(1).Version = "Version 1"
        Else
            If SelectedIndex >= 0 Then
                If SelectedIndex = NViews - 1 Then 'Last item selected
                    'Add the new View to the end of the list.
                    lstCalculations.Items.Add("")
                    Dim NewSettings As New DataViewSettings
                    CalculationsSettings.List.Add(NewSettings)
                    OpenCalculationsFormNo(NViews)
                    'CalculationsFormList(NViews).DataSummary = "New Calculations Data View"
                    CalculationsFormList(NViews).DataName = "New Calculations Data View"
                    CalculationsFormList(NViews).Version = "Version 1"
                Else
                    'Insert a new Settings entry in FinancialSettings:
                    lstCalculations.Items.Insert(SelectedIndex + 1, "")
                    Dim NewSettings As New DataViewSettings
                    CalculationsSettings.List.Insert(SelectedIndex + 1, NewSettings)
                    OpenCalculationsFormNo(SelectedIndex + 1)
                    'CalculationsFormList(SelectedIndex + 1).DataSummary = "New Calculations Data View"
                    CalculationsFormList(SelectedIndex + 1).DataName = "New Calculations Data View"
                    CalculationsFormList(SelectedIndex + 1).Version = "Version 1"
                End If
            Else
                'No item selected
                'Add the new View to the end of the list.
                lstCalculations.Items.Add("")
                Dim NewSettings As New DataViewSettings
                CalculationsSettings.List.Add(NewSettings)
                OpenCalculationsFormNo(NViews + 1)
                'CalculationsFormList(NViews + 1).DataSummary = "New Calculations Data View"
                CalculationsFormList(NViews + 1).DataName = "New Calculations Data View"
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
            CalculationsSettings.LastEditDate = Now
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
                'Zip.SelectFileForm.ApplicationName = Project.ApplicationName
                Zip.SelectFileForm.ApplicationName = Project.Application.Name
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
        'Find a Copy Data Settings file.

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
                'Zip.SelectFileForm.ApplicationName = Project.ApplicationName
                Zip.SelectFileForm.ApplicationName = Project.Application.Name
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
                If SharePriceDbPath = "" Then
                    Message.AddWarning("A share price database has not been selected!" & vbCrLf)
                    Exit Sub
                End If
                If System.IO.File.Exists(SharePriceDbPath) Then
                    'Share Price Database file exists.
                Else
                    'Share Price Database file does not exist!
                    Message.AddWarning("The share price database was not found: " & SharePriceDbPath & vbCrLf)
                    Exit Sub
                End If
                Query = txtCopyDataInputQuery.Text
                connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source =" & SharePriceDbPath
                myConnection.ConnectionString = connString
                myConnection.Open()
                da = New OleDb.OleDbDataAdapter(Query, myConnection)
                da.MissingSchemaAction = MissingSchemaAction.AddWithKey
                da.Fill(dsInput, "myData")
                myConnection.Close()
            Case "Financials"
                If FinancialsDbPath = "" Then
                    Message.AddWarning("A financials database has not been selected!" & vbCrLf)
                    Exit Sub
                End If
                If System.IO.File.Exists(FinancialsDbPath) Then
                    'Financials Database file exists.
                Else
                    'Financials Database file does not exist!
                    Message.AddWarning("The financials database was not found: " & FinancialsDbPath & vbCrLf)
                    Exit Sub
                End If
                Query = txtCopyDataInputQuery.Text
                connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source =" & FinancialsDbPath
                myConnection.ConnectionString = connString
                myConnection.Open()
                da = New OleDb.OleDbDataAdapter(Query, myConnection)
                da.MissingSchemaAction = MissingSchemaAction.AddWithKey
                da.Fill(dsInput, "myData")
                myConnection.Close()
            Case "Calculations"
                If CalculationsDbPath = "" Then
                    Message.AddWarning("A calculations database has not been selected!" & vbCrLf)
                    Exit Sub
                End If
                If System.IO.File.Exists(CalculationsDbPath) Then
                    'Calculations Database file exists.
                Else
                    'Calculations Database file does not exist!
                    Message.AddWarning("The calculations database was not found: " & CalculationsDbPath & vbCrLf)
                    Exit Sub
                End If
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
                If SharePriceDbPath = "" Then
                    Message.AddWarning("A share price database has not been selected!" & vbCrLf)
                    Exit Sub
                End If
                If System.IO.File.Exists(SharePriceDbPath) Then
                    'Share Price Database file exists.
                Else
                    'Share Price Database file does not exist!
                    Message.AddWarning("The share price database was not found: " & SharePriceDbPath & vbCrLf)
                    Exit Sub
                End If
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
                If FinancialsDbPath = "" Then
                    Message.AddWarning("A financials database has not been selected!" & vbCrLf)
                    Exit Sub
                End If
                If System.IO.File.Exists(FinancialsDbPath) Then
                    'Financials Database file exists.
                Else
                    'Financials Database file does not exist!
                    Message.AddWarning("The financials database was not found: " & FinancialsDbPath & vbCrLf)
                    Exit Sub
                End If
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
                If CalculationsDbPath = "" Then
                    Message.AddWarning("A calculations database has not been selected!" & vbCrLf)
                    Exit Sub
                End If
                If System.IO.File.Exists(CalculationsDbPath) Then
                    'Calculations Database file exists.
                Else
                    'Calculations Database file does not exist!
                    Message.AddWarning("The calculations database was not found: " & CalculationsDbPath & vbCrLf)
                    Exit Sub
                End If
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
                'Zip.SelectFileForm.ApplicationName = Project.ApplicationName
                Zip.SelectFileForm.ApplicationName = Project.Application.Name
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
                If SharePriceDbPath = "" Then
                    Message.AddWarning("A share price database has not been selected!" & vbCrLf)
                    Exit Sub
                End If
                If System.IO.File.Exists(SharePriceDbPath) Then
                    'Share Price Database file exists.
                Else
                    'Share Price Database file does not exist!
                    Message.AddWarning("The share price database was not found: " & SharePriceDbPath & vbCrLf)
                    Exit Sub
                End If
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
                If FinancialsDbPath = "" Then
                    Message.AddWarning("A financials database has not been selected!" & vbCrLf)
                    Exit Sub
                End If
                If System.IO.File.Exists(FinancialsDbPath) Then
                    'Financials Database file exists.
                Else
                    'Financials Database file does not exist!
                    Message.AddWarning("The financials database was not found: " & FinancialsDbPath & vbCrLf)
                    Exit Sub
                End If
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
                If CalculationsDbPath = "" Then
                    Message.AddWarning("A calculations database has not been selected!" & vbCrLf)
                    Exit Sub
                End If
                If System.IO.File.Exists(CalculationsDbPath) Then
                    'Calculations Database file exists.
                Else
                    'Calculations Database file does not exist!
                    Message.AddWarning("The calculations database was not found: " & CalculationsDbPath & vbCrLf)
                    Exit Sub
                End If
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
                If SharePriceDbPath = "" Then
                    Message.AddWarning("A share price database has not been selected!" & vbCrLf)
                    Exit Sub
                End If
                If System.IO.File.Exists(SharePriceDbPath) Then
                    'Share Price Database file exists.
                Else
                    'Share Price Database file does not exist!
                    Message.AddWarning("The share price database was not found: " & SharePriceDbPath & vbCrLf)
                    Exit Sub
                End If
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
                If FinancialsDbPath = "" Then
                    Message.AddWarning("A financials database has not been selected!" & vbCrLf)
                    Exit Sub
                End If
                If System.IO.File.Exists(FinancialsDbPath) Then
                    'Financials Database file exists.
                Else
                    'Financials Database file does not exist!
                    Message.AddWarning("The financials database was not found: " & FinancialsDbPath & vbCrLf)
                    Exit Sub
                End If
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
                If CalculationsDbPath = "" Then
                    Message.AddWarning("A calculations database has not been selected!" & vbCrLf)
                    Exit Sub
                End If
                If System.IO.File.Exists(CalculationsDbPath) Then
                    'Calculations Database file exists.
                Else
                    'Calculations Database file does not exist!
                    Message.AddWarning("The calculations database was not found: " & CalculationsDbPath & vbCrLf)
                    Exit Sub
                End If
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
                'Zip.SelectFileForm.ApplicationName = Project.ApplicationName
                Zip.SelectFileForm.ApplicationName = Project.Application.Name
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
                If SharePriceDbPath = "" Then
                    Message.AddWarning("A share price database has not been selected!" & vbCrLf)
                    Exit Sub
                End If
                If System.IO.File.Exists(SharePriceDbPath) Then
                    'Share Price Database file exists.
                Else
                    'Share Price Database file does not exist!
                    Message.AddWarning("The share price database was not found: " & SharePriceDbPath & vbCrLf)
                    Exit Sub
                End If
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
                If FinancialsDbPath = "" Then
                    Message.AddWarning("A financials database has not been selected!" & vbCrLf)
                    Exit Sub
                End If
                If System.IO.File.Exists(FinancialsDbPath) Then
                    'Financials Database file exists.
                Else
                    'Financials Database file does not exist!
                    Message.AddWarning("The financials database was not found: " & FinancialsDbPath & vbCrLf)
                    Exit Sub
                End If
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
                If CalculationsDbPath = "" Then
                    Message.AddWarning("A calculations database has not been selected!" & vbCrLf)
                    Exit Sub
                End If
                If System.IO.File.Exists(CalculationsDbPath) Then
                    'Calculations Database file exists.
                Else
                    'Calculations Database file does not exist!
                    Message.AddWarning("The calculations database was not found: " & CalculationsDbPath & vbCrLf)
                    Exit Sub
                End If
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
                'Zip.SelectFileForm.ApplicationName = Project.ApplicationName
                Zip.SelectFileForm.ApplicationName = Project.Application.Name
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
                If SharePriceDbPath = "" Then
                    Message.AddWarning("A share price database has not been selected!" & vbCrLf)
                    Exit Sub
                End If
                If System.IO.File.Exists(SharePriceDbPath) Then
                    'Share Price Database file exists.
                Else
                    'Share Price Database file does not exist!
                    Message.AddWarning("The share price database was not found: " & SharePriceDbPath & vbCrLf)
                    Exit Sub
                End If
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
                If FinancialsDbPath = "" Then
                    Message.AddWarning("A financials database has not been selected!" & vbCrLf)
                    Exit Sub
                End If
                If System.IO.File.Exists(FinancialsDbPath) Then
                    'Financials Database file exists.
                Else
                    'Financials Database file does not exist!
                    Message.AddWarning("The financials database was not found: " & FinancialsDbPath & vbCrLf)
                    Exit Sub
                End If
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
                If CalculationsDbPath = "" Then
                    Message.AddWarning("A calculations database has not been selected!" & vbCrLf)
                    Exit Sub
                End If
                If System.IO.File.Exists(CalculationsDbPath) Then
                    'Calculations Database file exists.
                Else
                    'Calculations Database file does not exist!
                    Message.AddWarning("The calculations database was not found: " & CalculationsDbPath & vbCrLf)
                    Exit Sub
                End If
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
                'Zip.SelectFileForm.ApplicationName = Project.ApplicationName
                Zip.SelectFileForm.ApplicationName = Project.Application.Name
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
                If SharePriceDbPath = "" Then
                    Message.AddWarning("A share price database has not been selected!" & vbCrLf)
                    Exit Sub
                End If
                If System.IO.File.Exists(SharePriceDbPath) Then
                    'Share Price Database file exists.
                Else
                    'Share Price Database file does not exist!
                    Message.AddWarning("The share price database was not found: " & SharePriceDbPath & vbCrLf)
                    Exit Sub
                End If
                Query = txtDateSelInputQuery.Text
                connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source =" & SharePriceDbPath
                myConnection.ConnectionString = connString
                myConnection.Open()
                da = New OleDb.OleDbDataAdapter(Query, myConnection)
                da.MissingSchemaAction = MissingSchemaAction.AddWithKey
                da.Fill(dsInput, "myData")
                myConnection.Close()
            Case "Financials"
                If FinancialsDbPath = "" Then
                    Message.AddWarning("A financials database has not been selected!" & vbCrLf)
                    Exit Sub
                End If
                If System.IO.File.Exists(FinancialsDbPath) Then
                    'Financials Database file exists.
                Else
                    'Financials Database file does not exist!
                    Message.AddWarning("The financials database was not found: " & FinancialsDbPath & vbCrLf)
                    Exit Sub
                End If
                Query = txtDateSelInputQuery.Text
                connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source =" & FinancialsDbPath
                myConnection.ConnectionString = connString
                myConnection.Open()
                da = New OleDb.OleDbDataAdapter(Query, myConnection)
                da.MissingSchemaAction = MissingSchemaAction.AddWithKey
                da.Fill(dsInput, "myData")
                myConnection.Close()
            Case "Calculations"
                If CalculationsDbPath = "" Then
                    Message.AddWarning("A calculations database has not been selected!" & vbCrLf)
                    Exit Sub
                End If
                If System.IO.File.Exists(CalculationsDbPath) Then
                    'Calculations Database file exists.
                Else
                    'Calculations Database file does not exist!
                    Message.AddWarning("The calculations database was not found: " & CalculationsDbPath & vbCrLf)
                    Exit Sub
                End If
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
                If SharePriceDbPath = "" Then
                    Message.AddWarning("A share price database has not been selected!" & vbCrLf)
                    Exit Sub
                End If
                If System.IO.File.Exists(SharePriceDbPath) Then
                    'Share Price Database file exists.
                Else
                    'Share Price Database file does not exist!
                    Message.AddWarning("The share price database was not found: " & SharePriceDbPath & vbCrLf)
                    Exit Sub
                End If
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
                If FinancialsDbPath = "" Then
                    Message.AddWarning("A financials database has not been selected!" & vbCrLf)
                    Exit Sub
                End If
                If System.IO.File.Exists(FinancialsDbPath) Then
                    'Financials Database file exists.
                Else
                    'Financials Database file does not exist!
                    Message.AddWarning("The financials database was not found: " & FinancialsDbPath & vbCrLf)
                    Exit Sub
                End If
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
                If CalculationsDbPath = "" Then
                    Message.AddWarning("A calculations database has not been selected!" & vbCrLf)
                    Exit Sub
                End If
                If System.IO.File.Exists(CalculationsDbPath) Then
                    'Calculations Database file exists.
                Else
                    'Calculations Database file does not exist!
                    Message.AddWarning("The calculations database was not found: " & CalculationsDbPath & vbCrLf)
                    Exit Sub
                End If
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
                    If SharePriceDbPath = "" Then
                        Message.AddWarning("A share price database has not been selected!" & vbCrLf)
                        Exit Sub
                    End If
                    If System.IO.File.Exists(SharePriceDbPath) Then
                        'Share Price Database file exists.
                    Else
                        'Share Price Database file does not exist!
                        Message.AddWarning("The share price database was not found: " & SharePriceDbPath & vbCrLf)
                        Exit Sub
                    End If
                    connectionString = "provider=Microsoft.ACE.OLEDB.12.0;data source = " + SharePriceDbPath
                Case "Financials"
                    If FinancialsDbPath = "" Then
                        Message.AddWarning("A financials database has not been selected!" & vbCrLf)
                        Exit Sub
                    End If
                    If System.IO.File.Exists(FinancialsDbPath) Then
                        'Financials Database file exists.
                    Else
                        'Financials Database file does not exist!
                        Message.AddWarning("The financials database was not found: " & FinancialsDbPath & vbCrLf)
                        Exit Sub
                    End If
                    connectionString = "provider=Microsoft.ACE.OLEDB.12.0;data source = " + FinancialsDbPath
                Case "Calculations"
                    If CalculationsDbPath = "" Then
                        Message.AddWarning("A calculations database has not been selected!" & vbCrLf)
                        Exit Sub
                    End If
                    If System.IO.File.Exists(CalculationsDbPath) Then
                        'Calculations Database file exists.
                    Else
                        'Calculations Database file does not exist!
                        Message.AddWarning("The calculations database was not found: " & CalculationsDbPath & vbCrLf)
                        Exit Sub
                    End If
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
                    If SharePriceDbPath = "" Then
                        Message.AddWarning("A share price database has not been selected!" & vbCrLf)
                        Exit Sub
                    End If
                    If System.IO.File.Exists(SharePriceDbPath) Then
                        'Share Price Database file exists.
                    Else
                        'Share Price Database file does not exist!
                        Message.AddWarning("The share price database was not found: " & SharePriceDbPath & vbCrLf)
                        Exit Sub
                    End If
                    connectionString = "provider=Microsoft.ACE.OLEDB.12.0;data source = " + SharePriceDbPath
                Case "Financials"
                    If FinancialsDbPath = "" Then
                        Message.AddWarning("A financials database has not been selected!" & vbCrLf)
                        Exit Sub
                    End If
                    If System.IO.File.Exists(FinancialsDbPath) Then
                        'Financials Database file exists.
                    Else
                        'Financials Database file does not exist!
                        Message.AddWarning("The financials database was not found: " & FinancialsDbPath & vbCrLf)
                        Exit Sub
                    End If
                    connectionString = "provider=Microsoft.ACE.OLEDB.12.0;data source = " + FinancialsDbPath
                Case "Calculations"
                    If CalculationsDbPath = "" Then
                        Message.AddWarning("A calculations database has not been selected!" & vbCrLf)
                        Exit Sub
                    End If
                    If System.IO.File.Exists(CalculationsDbPath) Then
                        'Calculations Database file exists.
                    Else
                        'Calculations Database file does not exist!
                        Message.AddWarning("The calculations database was not found: " & CalculationsDbPath & vbCrLf)
                        Exit Sub
                    End If
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
                ElseIf System.IO.File.Exists(SharePriceDbPath) Then
                    connectionString = "provider=Microsoft.ACE.OLEDB.12.0;data source = " + SharePriceDbPath
                Else
                    Message.AddWarning("The share price database was not found: " & SharePriceDbPath & vbCrLf)
                    Exit Sub
                End If
            Case "Financials"
                If FinancialsDbPath = "" Then
                    Message.AddWarning("Calculations: Daily Prices: No input Financials database selected!" & vbCrLf)
                    Exit Sub
                ElseIf System.IO.File.Exists(FinancialsDbPath) Then
                    connectionString = "provider=Microsoft.ACE.OLEDB.12.0;data source = " + FinancialsDbPath
                Else
                    Message.AddWarning("The financials database was not found: " & FinancialsDbPath & vbCrLf)
                    Exit Sub
                End If
            Case "Calculations"
                If CalculationsDbPath = "" Then
                    Message.AddWarning("Calculations: Daily Prices: No input Calculations database selected!" & vbCrLf)
                    Exit Sub
                ElseIf System.IO.File.Exists(CalculationsDbPath) Then
                    connectionString = "provider=Microsoft.ACE.OLEDB.12.0;data source = " + CalculationsDbPath
                Else
                    Message.AddWarning("The calculations database was not found: " & CalculationsDbPath & vbCrLf)
                    Exit Sub
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
                ElseIf System.IO.File.Exists(SharePriceDbPath) Then
                    connectionString = "provider=Microsoft.ACE.OLEDB.12.0;data source = " + SharePriceDbPath
                Else
                    Message.AddWarning("The share price database was not found: " & SharePriceDbPath & vbCrLf)
                    Exit Sub
                End If
            Case "Financials"
                If FinancialsDbPath = "" Then
                    Message.AddWarning("Calculations: Daily Prices: No output Financials database selected!" & vbCrLf)
                    Exit Sub
                ElseIf System.IO.File.Exists(FinancialsDbPath) Then
                    connectionString = "provider=Microsoft.ACE.OLEDB.12.0;data source = " + FinancialsDbPath
                Else
                    Message.AddWarning("The financials database was not found: " & FinancialsDbPath & vbCrLf)
                    Exit Sub
                End If
            Case "Calculations"
                If CalculationsDbPath = "" Then
                    Message.AddWarning("Calculations: Daily Prices: No output Calculations database selected!" & vbCrLf)
                    Exit Sub
                ElseIf System.IO.File.Exists(CalculationsDbPath) Then
                    connectionString = "provider=Microsoft.ACE.OLEDB.12.0;data source = " + CalculationsDbPath
                Else
                    Message.AddWarning("The calculations database was not found: " & CalculationsDbPath & vbCrLf)
                    Exit Sub
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

    Private Sub SetupChartLinePlotTab()
        'Set up the CLine Chart tab under the Charts tab.

        'Set up database selection options:
        cmbLineChartDb.Items.Clear()
        cmbLineChartDb.Items.Add("Share Prices")
        cmbLineChartDb.Items.Add("Financials")
        cmbLineChartDb.Items.Add("Calculations")

        'Set up the Chart Title alignment options:
        cmbLineChartAlignment.Items.Clear()
        'Show the list of ContentAlignment enumerations in the cmbAlignment combobox:
        For Each item In System.Enum.GetValues(GetType(ContentAlignment))
            cmbLineChartAlignment.Items.Add(item)
        Next
        cmbLineChartAlignment.SelectedIndex = 1 'Top Center
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

        CheckOpenProjectAtRelativePath("\Stock Chart", "ADVL_Stock_Chart_1")

        'Dim StartTime As Date = Now
        'Dim Duration As TimeSpan
        ''Wait up to 16 seconds for the connection ConnName to be established
        'While client.ConnectionExists(ProNetName, "ADVL_Stock_Chart_1") = False 'Wait until the required connection is made.
        '    System.Threading.Thread.Sleep(1000) 'Pause for 1000ms
        '    Duration = Now - StartTime
        '    If Duration.Seconds > 16 Then Exit While 'Stop waiting after 16 seconds.
        'End While

        'Wait up to 8 seconds for the Stock Chart project to connect:
        If WaitForConnection(ProNetName, "ADVL_Stock_Chart_1", 8000) = False Then
            Message.AddWarning("The Stock Chart project did not connect." & vbCrLf)
        End If

        DisplayStockChart()

    End Sub

    Private Sub btnOpenSPChart_Click(sender As Object, e As EventArgs) Handles btnOpenSPChart.Click
        'Open the Share Price Chart project:
        CheckOpenProjectAtRelativePath("\Stock Chart", "ADVL_Stock_Chart_1")
        'Wait up to 8 seconds for the Stock Chart project to connect:
        If WaitForConnection(ProNetName, "ADVL_Stock_Chart_1", 8000) = False Then
            Message.AddWarning("The Stock Chart project did not connect." & vbCrLf)
        End If
    End Sub

    Private Sub lstTables_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lstTables.SelectedIndexChanged
        'Show the list of table columns in lstFields (Point Chart)

        If lstTables.SelectedIndex = -1 Then
            Message.Add("No table selected." & vbCrLf)
        Else
            'Database access for MS Access:
            Dim connectionString As String 'Declare a connection string for MS Access - defines the database or server to be used.
            Dim conn As System.Data.OleDb.OleDbConnection 'Declare a connection for MS Access - used by the Data Adapter to connect to and disconnect from the database.
            Dim commandString As String 'Declare a command string - contains the query to be passed to the database.
            Dim ds As DataSet 'Declate a Dataset.
            Dim dt As DataTable

            'cmbCompanyCodeCol.Items.Clear()
            lstFields.Items.Clear()

            'Specify the connection string (Access 2007):
            connectionString = "provider=Microsoft.ACE.OLEDB.12.0;" +
            "data source = " + txtPointChartDbPath.Text

            'Connect to the Access database:
            conn = New System.Data.OleDb.OleDbConnection(connectionString)
            conn.Open()

            'Specify the commandString to query the database:
            'commandString = "SELECT Top 500 * FROM " + cmbChartDataTable.SelectedItem.ToString
            commandString = "SELECT Top 500 * FROM " + lstTables.SelectedItem.ToString
            Dim dataAdapter As New System.Data.OleDb.OleDbDataAdapter(commandString, conn)
            ds = New DataSet
            dataAdapter.Fill(ds, "SelTable") 'ds was defined earlier as a DataSet
            dt = ds.Tables("SelTable")

            Dim NFields As Integer
            NFields = dt.Columns.Count
            Dim I As Integer
            For I = 0 To NFields - 1
                'cmbCompanyCodeCol.Items.Add(dt.Columns(I).ColumnName.ToString)
                lstFields.Items.Add(dt.Columns(I).ColumnName.ToString)
            Next

            conn.Close()
        End If
    End Sub

    Private Sub lstLineTables_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lstLineTables.SelectedIndexChanged
        'Show the list of table columns in lstLineFields (Line Chart)

        If lstLineTables.SelectedIndex = -1 Then
            Message.Add("No table selected." & vbCrLf)
        Else
            'Database access for MS Access:
            Dim connectionString As String 'Declare a connection string for MS Access - defines the database or server to be used.
            Dim conn As System.Data.OleDb.OleDbConnection 'Declare a connection for MS Access - used by the Data Adapter to connect to and disconnect from the database.
            Dim commandString As String 'Declare a command string - contains the query to be passed to the database.
            Dim ds As DataSet 'Declate a Dataset.
            Dim dt As DataTable

            'cmbCompanyCodeCol.Items.Clear()
            lstLineFields.Items.Clear()

            'Specify the connection string (Access 2007):
            connectionString = "provider=Microsoft.ACE.OLEDB.12.0;" +
            "data source = " + txtLineChartDbPath.Text

            'Connect to the Access database:
            conn = New System.Data.OleDb.OleDbConnection(connectionString)
            conn.Open()

            'Specify the commandString to query the database:
            'commandString = "SELECT Top 500 * FROM " + cmbChartDataTable.SelectedItem.ToString
            commandString = "SELECT Top 500 * FROM " + lstLineTables.SelectedItem.ToString
            Dim dataAdapter As New System.Data.OleDb.OleDbDataAdapter(commandString, conn)
            ds = New DataSet
            dataAdapter.Fill(ds, "SelTable") 'ds was defined earlier as a DataSet
            dt = ds.Tables("SelTable")

            Dim NFields As Integer
            NFields = dt.Columns.Count
            Dim I As Integer
            For I = 0 To NFields - 1
                'cmbCompanyCodeCol.Items.Add(dt.Columns(I).ColumnName.ToString)
                lstLineFields.Items.Add(dt.Columns(I).ColumnName.ToString)
            Next

            conn.Close()
        End If


    End Sub

    'Private Function WaitForConnection(ByVal ProNetName As String, ByVal ConnName As String, ByVal MaxMilliSecs As Integer) As Boolean
    Public Function WaitForConnection(ByVal ProNetName As String, ByVal ConnName As String, ByVal MaxMilliSecs As Integer) As Boolean
        'Wait for the connection to be made for [ProNetName].ConnName
        'Return True if the connection was made within MaxMilliSecs time.
        'Return False if the connection was not made within MaxMilliSecs time.
        Dim StartTime As Date = Now
        Dim Duration As TimeSpan
        Dim Timeout As Boolean = False
        'Wait up to 16 seconds for the connection ConnName to be established
        'While client.ConnectionExists(ProNetName, "ADVL_Stock_Chart_1") = False 'Wait until the required connection is made.
        While client.ConnectionExists(ProNetName, ConnName) = False 'Wait until the required connection is made.
            System.Threading.Thread.Sleep(1000) 'Pause for 1000ms
            Duration = Now - StartTime
            'If Duration.Seconds > 16 Then
            'If Duration.Milliseconds > MaxMilliSecs Then
            If Duration.TotalMilliseconds > MaxMilliSecs Then
                Timeout = True
                Exit While 'Stop waiting after 16 seconds.
            End If
        End While
        If Timeout Then
            Return False 'The Timeout period elapsed before the connection was made.
        Else
            Return True 'The connection was made within the Timeout period.
        End If
    End Function

    'Public Sub DisplayStockChartUsingDefaults()
    Public Sub DisplayStockChartUsingSettingsList_Old()
        'Display Stock Chart.
        'Use the default parameters in StockChartDefaults
        'Send the instructions to the Chart application to display the stock chart.

        'If StockChartDefaults Is Nothing Then
        If StockChartSettingsList Is Nothing Then
            'Message.AddWarning("No Stock Chart default settings loaded." & vbCrLf)
            Message.AddWarning("No Stock Chart settings list loaded." & vbCrLf)
            'DisplayStockChartNoDefaults()
            DisplayStockChartNoSettingsList()
            Exit Sub
        End If

        'Create the xml instructions
        Dim decl As New XDeclaration("1.0", "utf-8", "yes")
        Dim doc As New XDocument(decl, Nothing) 'Create an XDocument to store the instructions.
        Dim xmessage As New XElement("XMsg") 'This indicates the start of the message in the XMessage class

        'Dim clientAppNetName As New XElement("ClientAppNetName", AppNetName)
        'xmessage.Add(clientAppNetName)
        Dim clientProNetName As New XElement("ClientProNetName", ProNetName)
        xmessage.Add(clientProNetName)

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

        'If StockChartDefaults.<StockChart>.<Settings>.<ChartTitle>.<FontName>.Value <> Nothing Then
        If StockChartSettingsList.<StockChart>.<Settings>.<ChartTitle>.<FontName>.Value <> Nothing Then
            'Dim chartTitleFontName As New XElement("FontName", StockChartDefaults.<StockChart>.<Settings>.<ChartTitle>.<FontName>.Value)
            Dim chartTitleFontName As New XElement("FontName", StockChartSettingsList.<StockChart>.<Settings>.<ChartTitle>.<FontName>.Value)
            chartTitle.Add(chartTitleFontName)
        Else
            Message.AddWarning("Default Chart Title Font Name settings not found." & vbCrLf)
            Dim chartTitleFontName As New XElement("FontName", txtChartTitle.Font.Name)
            chartTitle.Add(chartTitleFontName)
        End If

        If StockChartSettingsList.<StockChart>.<Settings>.<ChartTitle>.<Color>.Value <> Nothing Then
            Dim chartTitleColor As New XElement("Color", StockChartSettingsList.<StockChart>.<Settings>.<ChartTitle>.<Color>.Value)
            chartTitle.Add(chartTitleColor)
        Else
            Message.AddWarning("Default Chart Title Color settings not found." & vbCrLf)
            Dim chartTitleColor As New XElement("Color", txtChartTitle.ForeColor)
            chartTitle.Add(chartTitleColor)
        End If

        If StockChartSettingsList.<StockChart>.<Settings>.<ChartTitle>.<Size>.Value <> Nothing Then
            Dim chartTitleSize As New XElement("Size", StockChartSettingsList.<StockChart>.<Settings>.<ChartTitle>.<Size>.Value)
            chartTitle.Add(chartTitleSize)
        Else
            Message.AddWarning("Default Chart Title Size settings not found." & vbCrLf)
            Dim chartTitleSize As New XElement("Size", txtChartTitle.Font.Size)
            chartTitle.Add(chartTitleSize)
        End If

        If StockChartSettingsList.<StockChart>.<Settings>.<ChartTitle>.<Bold>.Value <> Nothing Then
            Dim chartTitleBold As New XElement("Bold", StockChartSettingsList.<StockChart>.<Settings>.<ChartTitle>.<Bold>.Value)
            chartTitle.Add(chartTitleBold)
        Else
            Message.AddWarning("Default Chart Title Bold settings not found." & vbCrLf)
            Dim chartTitleBold As New XElement("Bold", txtChartTitle.Font.Bold)
            chartTitle.Add(chartTitleBold)
        End If

        If StockChartSettingsList.<StockChart>.<Settings>.<ChartTitle>.<Italic>.Value <> Nothing Then
            Dim chartTitleItalic As New XElement("Italic", StockChartSettingsList.<StockChart>.<Settings>.<ChartTitle>.<Italic>.Value)
            chartTitle.Add(chartTitleItalic)
        Else
            Message.AddWarning("Default Chart Title Italic settings not found." & vbCrLf)
            Dim chartTitleItalic As New XElement("Italic", txtChartTitle.Font.Italic)
            chartTitle.Add(chartTitleItalic)
        End If

        If StockChartSettingsList.<StockChart>.<Settings>.<ChartTitle>.<Underline>.Value <> Nothing Then
            Dim chartTitleUnderline As New XElement("Underline", StockChartSettingsList.<StockChart>.<Settings>.<ChartTitle>.<Underline>.Value)
            chartTitle.Add(chartTitleUnderline)
        Else
            Message.AddWarning("Default Chart Title Underline settings not found." & vbCrLf)
            Dim chartTitleUnderline As New XElement("Underline", txtChartTitle.Font.Underline)
            chartTitle.Add(chartTitleUnderline)
        End If

        If StockChartSettingsList.<StockChart>.<Settings>.<ChartTitle>.<Strikeout>.Value <> Nothing Then
            Dim chartTitleStrikeout As New XElement("Strikeout", StockChartSettingsList.<StockChart>.<Settings>.<ChartTitle>.<Strikeout>.Value)
            chartTitle.Add(chartTitleStrikeout)
        Else
            Message.AddWarning("Default Chart Title Strikeout settings not found." & vbCrLf)
            Dim chartTitleStrikeout As New XElement("Strikeout", txtChartTitle.Font.Strikeout)
            chartTitle.Add(chartTitleStrikeout)
        End If

        If StockChartSettingsList.<StockChart>.<Settings>.<ChartTitle>.<Alignment>.Value <> Nothing Then
            Dim chartTitleAlignment As New XElement("Alignment", StockChartSettingsList.<StockChart>.<Settings>.<ChartTitle>.<Alignment>.Value)
            chartTitle.Add(chartTitleAlignment)
        Else
            Message.AddWarning("Default Chart Title Alignment settings not found." & vbCrLf)
            Dim chartTitleAlignment As New XElement("Alignment", cmbAlignment.SelectedItem.ToString)
            chartTitle.Add(chartTitleAlignment)
        End If

        chartSettings.Add(chartTitle)

        Dim xAxis As New XElement("XAxis")

        If StockChartSettingsList.<StockChart>.<Settings>.<XAxis>.<TitleText>.Value <> Nothing Then
            Dim xAxisTitleText As New XElement("TitleText", StockChartSettingsList.<StockChart>.<Settings>.<XAxis>.<TitleText>.Value)
            xAxis.Add(xAxisTitleText)
        End If

        If StockChartSettingsList.<StockChart>.<Settings>.<XAxis>.<TitleFontName>.Value <> Nothing Then
            Dim xAxisTitleFontName As New XElement("TitleFontName", StockChartSettingsList.<StockChart>.<Settings>.<XAxis>.<TitleFontName>.Value)
            xAxis.Add(xAxisTitleFontName)
        End If

        If StockChartSettingsList.<StockChart>.<Settings>.<XAxis>.<TitleColor>.Value <> Nothing Then
            Dim xAxisTitleColor As New XElement("TitleColor", StockChartSettingsList.<StockChart>.<Settings>.<XAxis>.<TitleColor>.Value)
            xAxis.Add(xAxisTitleColor)
        End If

        If StockChartSettingsList.<StockChart>.<Settings>.<XAxis>.<TitleSize>.Value <> Nothing Then
            Dim xAxisTitleSize As New XElement("TitleSize", StockChartSettingsList.<StockChart>.<Settings>.<XAxis>.<TitleSize>.Value)
            xAxis.Add(xAxisTitleSize)
        End If

        If StockChartSettingsList.<StockChart>.<Settings>.<XAxis>.<TitleBold>.Value <> Nothing Then
            Dim xAxisTitleBold As New XElement("TitleBold", StockChartSettingsList.<StockChart>.<Settings>.<XAxis>.<TitleBold>.Value)
            xAxis.Add(xAxisTitleBold)
        End If

        If StockChartSettingsList.<StockChart>.<Settings>.<XAxis>.<TitleItalic>.Value <> Nothing Then
            Dim xAxisTitleItalic As New XElement("TitleItalic", StockChartSettingsList.<StockChart>.<Settings>.<XAxis>.<TitleItalic>.Value)
            xAxis.Add(xAxisTitleItalic)
        End If

        If StockChartSettingsList.<StockChart>.<Settings>.<XAxis>.<TitleUnderline>.Value <> Nothing Then
            Dim xAxisTitleUnderline As New XElement("TitleUnderline", StockChartSettingsList.<StockChart>.<Settings>.<XAxis>.<TitleUnderline>.Value)
            xAxis.Add(xAxisTitleUnderline)
        End If

        If StockChartSettingsList.<StockChart>.<Settings>.<XAxis>.<TitleStrikeout>.Value <> Nothing Then
            Dim xAxisTitleStrikeout As New XElement("TitleStrikeout", StockChartSettingsList.<StockChart>.<Settings>.<XAxis>.<TitleStrikeout>.Value)
            xAxis.Add(xAxisTitleStrikeout)
        End If

        If StockChartSettingsList.<StockChart>.<Settings>.<XAxis>.<TitleAlignment>.Value <> Nothing Then
            Dim xAxisTitleAlignment As New XElement("TitleAlignment", StockChartSettingsList.<StockChart>.<Settings>.<XAxis>.<TitleAlignment>.Value)
            xAxis.Add(xAxisTitleAlignment)
        End If

        If StockChartSettingsList.<StockChart>.<Settings>.<XAxis>.<AutoMinimum>.Value <> Nothing Then
            Dim xAxisAutoMinimum As New XElement("AutoMinimum", StockChartSettingsList.<StockChart>.<Settings>.<XAxis>.<AutoMinimum>.Value)
            xAxis.Add(xAxisAutoMinimum)
        End If

        If StockChartSettingsList.<StockChart>.<Settings>.<XAxis>.<Minimum>.Value <> Nothing Then
            Dim xAxisMinimum As New XElement("Minimum", StockChartSettingsList.<StockChart>.<Settings>.<XAxis>.<Minimum>.Value)
            xAxis.Add(xAxisMinimum)
        End If

        If StockChartSettingsList.<StockChart>.<Settings>.<XAxis>.<AutoMaximum>.Value <> Nothing Then
            Dim xAxisAutoMaximum As New XElement("AutoMaximum", StockChartSettingsList.<StockChart>.<Settings>.<XAxis>.<AutoMaximum>.Value)
            xAxis.Add(xAxisAutoMaximum)
        End If

        If StockChartSettingsList.<StockChart>.<Settings>.<XAxis>.<Maximum>.Value <> Nothing Then
            Dim xAxisMaximum As New XElement("Maximum", StockChartSettingsList.<StockChart>.<Settings>.<XAxis>.<Maximum>.Value)
            xAxis.Add(xAxisMaximum)
        End If

        If StockChartSettingsList.<StockChart>.<Settings>.<XAxis>.<AutoInterval>.Value <> Nothing Then
            Dim xAxisAutoInterval As New XElement("AutoInterval", StockChartSettingsList.<StockChart>.<Settings>.<XAxis>.<AutoInterval>.Value)
            xAxis.Add(xAxisAutoInterval)
        End If

        If StockChartSettingsList.<StockChart>.<Settings>.<XAxis>.<MajorGridInterval>.Value <> Nothing Then
            Dim xAxisMajorGridInterval As New XElement("MajorGridInterval", StockChartSettingsList.<StockChart>.<Settings>.<XAxis>.<MajorGridInterval>.Value)
            xAxis.Add(xAxisMajorGridInterval)
        End If

        If StockChartSettingsList.<StockChart>.<Settings>.<XAxis>.<AutoMajorGridInterval>.Value <> Nothing Then
            Dim xAxisAutoMajorGridInterval As New XElement("AutoMajorGridInterval", StockChartSettingsList.<StockChart>.<Settings>.<XAxis>.<AutoMajorGridInterval>.Value)
            xAxis.Add(xAxisAutoMajorGridInterval)
        End If

        chartSettings.Add(xAxis)

        Dim yAxis As New XElement("YAxis")

        If StockChartSettingsList.<StockChart>.<Settings>.<YAxis>.<TitleText>.Value <> Nothing Then
            Dim yAxisTitleText As New XElement("TitleText", StockChartSettingsList.<StockChart>.<Settings>.<YAxis>.<TitleText>.Value)
            yAxis.Add(yAxisTitleText)
        End If

        If StockChartSettingsList.<StockChart>.<Settings>.<YAxis>.<TitleFontName>.Value <> Nothing Then
            Dim yAxisTitleFontName As New XElement("TitleFontName", StockChartSettingsList.<StockChart>.<Settings>.<YAxis>.<TitleFontName>.Value)
            yAxis.Add(yAxisTitleFontName)
        End If

        If StockChartSettingsList.<StockChart>.<Settings>.<YAxis>.<TitleColor>.Value <> Nothing Then
            Dim yAxisTitleColor As New XElement("TitleColor", StockChartSettingsList.<StockChart>.<Settings>.<YAxis>.<TitleColor>.Value)
            yAxis.Add(yAxisTitleColor)
        End If

        If StockChartSettingsList.<StockChart>.<Settings>.<YAxis>.<TitleSize>.Value <> Nothing Then
            Dim yAxisTitleSize As New XElement("TitleSize", StockChartSettingsList.<StockChart>.<Settings>.<YAxis>.<TitleSize>.Value)
            yAxis.Add(yAxisTitleSize)
        End If

        If StockChartSettingsList.<StockChart>.<Settings>.<YAxis>.<TitleBold>.Value <> Nothing Then
            Dim yAxisTitleBold As New XElement("TitleBold", StockChartSettingsList.<StockChart>.<Settings>.<YAxis>.<TitleBold>.Value)
            yAxis.Add(yAxisTitleBold)
        End If

        If StockChartSettingsList.<StockChart>.<Settings>.<YAxis>.<TitleItalic>.Value <> Nothing Then
            Dim yAxisTitleItalic As New XElement("TitleItalic", StockChartSettingsList.<StockChart>.<Settings>.<YAxis>.<TitleItalic>.Value)
            yAxis.Add(yAxisTitleItalic)
        End If

        If StockChartSettingsList.<StockChart>.<Settings>.<YAxis>.<TitleUnderline>.Value <> Nothing Then
            Dim yAxisTitleUnderline As New XElement("TitleUnderline", StockChartSettingsList.<StockChart>.<Settings>.<YAxis>.<TitleUnderline>.Value)
            yAxis.Add(yAxisTitleUnderline)
        End If

        If StockChartSettingsList.<StockChart>.<Settings>.<YAxis>.<TitleStrikeout>.Value <> Nothing Then
            Dim yAxisTitleStrikeout As New XElement("TitleStrikeout", StockChartSettingsList.<StockChart>.<Settings>.<YAxis>.<TitleStrikeout>.Value)
            yAxis.Add(yAxisTitleStrikeout)
        End If

        If StockChartSettingsList.<StockChart>.<Settings>.<YAxis>.<TitleAlignment>.Value <> Nothing Then
            Dim yAxisTitleAlignment As New XElement("TitleAlignment", StockChartSettingsList.<StockChart>.<Settings>.<YAxis>.<TitleAlignment>.Value)
            yAxis.Add(yAxisTitleAlignment)
        End If

        If StockChartSettingsList.<StockChart>.<Settings>.<YAxis>.<AutoMinimum>.Value <> Nothing Then
            Dim yAxisAutoMinimum As New XElement("AutoMinimum", StockChartSettingsList.<StockChart>.<Settings>.<YAxis>.<AutoMinimum>.Value)
            yAxis.Add(yAxisAutoMinimum)
        End If

        If StockChartSettingsList.<StockChart>.<Settings>.<YAxis>.<Minimum>.Value <> Nothing Then
            Dim yAxisMinimum As New XElement("Minimum", StockChartSettingsList.<StockChart>.<Settings>.<YAxis>.<Minimum>.Value)
            yAxis.Add(yAxisMinimum)
        End If

        If StockChartSettingsList.<StockChart>.<Settings>.<YAxis>.<AutoMaximum>.Value <> Nothing Then
            Dim yAxisAutoMaximum As New XElement("AutoMaximum", StockChartSettingsList.<StockChart>.<Settings>.<YAxis>.<AutoMaximum>.Value)
            yAxis.Add(yAxisAutoMaximum)
        End If

        If StockChartSettingsList.<StockChart>.<Settings>.<YAxis>.<Maximum>.Value <> Nothing Then
            Dim yAxisMaximum As New XElement("Maximum", StockChartSettingsList.<StockChart>.<Settings>.<YAxis>.<Maximum>.Value)
            yAxis.Add(yAxisMaximum)
        End If

        If StockChartSettingsList.<StockChart>.<Settings>.<YAxis>.<AutoInterval>.Value <> Nothing Then
            Dim yAxisAutoInterval As New XElement("AutoInterval", StockChartSettingsList.<StockChart>.<Settings>.<YAxis>.<AutoInterval>.Value)
            yAxis.Add(yAxisAutoInterval)
        End If

        If StockChartSettingsList.<StockChart>.<Settings>.<YAxis>.<MajorGridInterval>.Value <> Nothing Then
            Dim yAxisMajorGridInterval As New XElement("MajorGridInterval", StockChartSettingsList.<StockChart>.<Settings>.<YAxis>.<MajorGridInterval>.Value)
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

            'client.SendMessageAsync(AppNetName, "ADVL_Stock_Chart_1", doc.ToString) 'Added 3Feb19
            client.SendMessageAsync(ProNetName, "ADVL_Stock_Chart_1", doc.ToString)
            'Message.XAddText("Message sent to " & "ADVL_Stock_Chart_1" & ":" & vbCrLf, "XmlSentNotice")
            Message.XAddText("Message sent to [" & ProNetName & "]." & "ADVL_Stock_Chart_1" & ":" & vbCrLf, "XmlSentNotice")
            Message.XAddXml(doc.ToString)
            Message.XAddText(vbCrLf, "Message") 'Add extra line
        End If

    End Sub

    Public Sub DisplayLineChartUsingSettingsList()
        'Display the Line Chart using a Settings List.

        Message.Add("Displaying the Line Chart using the Settings List" & vbCrLf)

        'Send the instructions to the Chart application to display the line chart.

        'Check that required selections have been made:
        If cmbLineXValues.SelectedItem Is Nothing Then
            Message.AddWarning("Select a field for the X Values." & vbCrLf)
            Exit Sub
        End If

        'Build the XMessageBlock containing the Line Chart settings.
        'This will be send to the Line Chart application to create the chart display.

        Dim ChartSettingsList As XDocument = XDocument.Parse("<?xml version=""1.0"" encoding=""utf-8"" standalone=""yes""?>" & vbCrLf & XmlLineChartSettingsList.Text)

        Dim LineChartXMsgBlk = <?xml version="1.0" encoding="utf-8"?>
                               <XMsgBlk>
                                   <ClientLocn>DisplayChart</ClientLocn>
                                   <XInfo>
                                       <%= ChartSettingsList.<ChartSettings> %>
                                   </XInfo>
                               </XMsgBlk>

        'Update the Settings List with the current chart settings:

        'Update the Input Data settinga:
        '<InputDataType>Database</InputDataType> - Currently only the Database type is available.
        LineChartXMsgBlk.<XMsgBlk>.<XInfo>.<ChartSettings>.<InputDatabasePath>.Value = txtLineChartDbPath.Text
        LineChartXMsgBlk.<XMsgBlk>.<XInfo>.<ChartSettings>.<InputQuery>.Value = txtLineChartQuery.Text
        LineChartXMsgBlk.<XMsgBlk>.<XInfo>.<ChartSettings>.<InputDataDescr>.Value = txtLineSeriesName.Text

        'Add a warning if there is more than one entry in the SeriesInfoList:
        If LineChartXMsgBlk.<XMsgBlk>.<XInfo>.<ChartSettings>.<SeriesInfoList>.<SeriesInfo>.Count > 1 Then AddWarning("There is more than one entry in the Series Info List!" & vbCrLf)
        'Update the first entry in the SeriesInfoList:
        LineChartXMsgBlk.<XMsgBlk>.<XInfo>.<ChartSettings>.<SeriesInfoList>.<SeriesInfo>.<Name>.Value = txtLineSeriesName.Text
        LineChartXMsgBlk.<XMsgBlk>.<XInfo>.<ChartSettings>.<SeriesInfoList>.<SeriesInfo>.<XValuesFieldName>.Value = cmbLineXValues.SelectedItem.ToString
        LineChartXMsgBlk.<XMsgBlk>.<XInfo>.<ChartSettings>.<SeriesInfoList>.<SeriesInfo>.<YValuesFieldName>.Value = cmbLineYValues.SelectedItem.ToString

        'Leave the AreaInfoList unchanged.

        'Update the Chart title settings:
        LineChartXMsgBlk.<XMsgBlk>.<XInfo>.<ChartSettings>.<TitlesCollection>.<Title>.<Text>.Value = txtLineChartTitle.Text
        LineChartXMsgBlk.<XMsgBlk>.<XInfo>.<ChartSettings>.<TitlesCollection>.<Title>.<Alignment>.Value = cmbLineChartAlignment.SelectedItem.ToString
        If txtLineChartTitle.ForeColor.ToArgb.ToString = "0" Then 'This color value is not valid for a chart title.
            LineChartXMsgBlk.<XMsgBlk>.<XInfo>.<ChartSettings>.<TitlesCollection>.<Title>.<ForeColor>.Value = Color.Black.ToArgb.ToString
        Else
            LineChartXMsgBlk.<XMsgBlk>.<XInfo>.<ChartSettings>.<TitlesCollection>.<Title>.<ForeColor>.Value = txtLineChartTitle.ForeColor.ToArgb.ToString
        End If
        LineChartXMsgBlk.<XMsgBlk>.<XInfo>.<ChartSettings>.<TitlesCollection>.<Title>.<Font>.<Name>.Value = txtLineChartTitle.Font.Name
        LineChartXMsgBlk.<XMsgBlk>.<XInfo>.<ChartSettings>.<TitlesCollection>.<Title>.<Font>.<Size>.Value = txtLineChartTitle.Font.Size
        LineChartXMsgBlk.<XMsgBlk>.<XInfo>.<ChartSettings>.<TitlesCollection>.<Title>.<Font>.<Bold>.Value = txtLineChartTitle.Font.Bold
        LineChartXMsgBlk.<XMsgBlk>.<XInfo>.<ChartSettings>.<TitlesCollection>.<Title>.<Font>.<Italic>.Value = txtLineChartTitle.Font.Italic
        LineChartXMsgBlk.<XMsgBlk>.<XInfo>.<ChartSettings>.<TitlesCollection>.<Title>.<Font>.<Strikeout>.Value = txtLineChartTitle.Font.Strikeout
        LineChartXMsgBlk.<XMsgBlk>.<XInfo>.<ChartSettings>.<TitlesCollection>.<Title>.<Font>.<Underline>.Value = txtLineChartTitle.Font.Underline

        'Add a warning if there is more than one entry in the SeriesCollection:
        If LineChartXMsgBlk.<XMsgBlk>.<XInfo>.<ChartSettings>.<SeriesCollection>.<Series>.Count > 1 Then AddWarning("There is more than one entry in the Series Collection!" & vbCrLf)
        'Update the first entry in the SeriesCollection:
        LineChartXMsgBlk.<XMsgBlk>.<XInfo>.<ChartSettings>.<SeriesCollection>.<Series>.<Name>.Value = txtLineSeriesName.Text


        'Send the XMessageBlock to the Line Chart application:
        Message.XAddText("Message sent to [" & ProNetName & "]." & "ADVL_Line_Chart_1" & ":" & vbCrLf, "XmlSentNotice")
        Message.XAddXml(LineChartXMsgBlk.ToString)
        Message.XAddText(vbCrLf, "Normal") 'Add extra line

        SendMessageParams.ProjectNetworkName = ProNetName
        SendMessageParams.ConnectionName = "ADVL_Line_Chart_1"
        SendMessageParams.Message = LineChartXMsgBlk.ToString
        If bgwSendMessage.IsBusy Then
            Message.AddWarning("Send Message backgroundworker is busy." & vbCrLf)
        Else
            bgwSendMessage.RunWorkerAsync(SendMessageParams)
        End If

    End Sub

    'TO BE DELETED:
    Public Sub DisplayLineChartUsingDefaults()

        'Display Line Chart.
        'Use the default parameters in LineChartDefaults
        'Send the instructions to the Chart application to display the line chart.

        'If LineChartDefaults Is Nothing Then
        If LineChartSettingsList Is Nothing Then
            'Message.AddWarning("No Stock Chart default settings loaded." & vbCrLf)
            Message.AddWarning("No Line Chart settings list loaded." & vbCrLf)
            'DisplayStockChartNoDefaults()
            'DisplayLineChartNoDefaults() 
            DisplayLineChartNoSettingsList()
            Exit Sub
        End If

        'Create the xml instructions
        Dim decl As New XDeclaration("1.0", "utf-8", "yes")
        Dim doc As New XDocument(decl, Nothing) 'Create an XDocument to store the instructions.
        Dim xmessage As New XElement("XMsg") 'This indicates the start of the message in the XMessage class

        'Dim clientAppNetName As New XElement("ClientAppNetName", AppNetName)
        'xmessage.Add(clientAppNetName)
        Dim clientProNetName As New XElement("ClientProNetName", ProNetName)
        xmessage.Add(clientProNetName)

        Dim clientName As New XElement("ClientName", ApplicationInfo.Name) 'This tells the Chart server the name of the client making the request.
        xmessage.Add(clientName)

        Dim clientLocn As New XElement("ClientLocn", "LineChart")
        xmessage.Add(clientLocn)

        Dim chartSettings As New XElement("LineChartSettings")

        'THe ChartType is specified in the Series section.
        'Dim chartType As New XElement("ChartType", "Line")
        'chartSettings.Add(chartType)

        Dim commandClearChart As New XElement("Command", "ClearChart")
        chartSettings.Add(commandClearChart)

        Dim inputData As New XElement("InputData")
        Dim dataType As New XElement("Type", "Database")
        inputData.Add(dataType)

        Dim databasePath As New XElement("DatabasePath", txtLineChartDbPath.Text)
        inputData.Add(databasePath)

        Dim dataDescription As New XElement("DataDescription", txtLineSeriesName.Text)
        inputData.Add(dataDescription)

        Dim databaseQuery As New XElement("DatabaseQuery", txtLineChartQuery.Text)
        inputData.Add(databaseQuery)

        chartSettings.Add(inputData)

        Dim chartProperties As New XElement("ChartProperties")
        Dim seriesName As New XElement("SeriesName", txtLineSeriesName.Text)
        chartProperties.Add(seriesName)
        Dim xValuesFieldName As New XElement("XValuesFieldName", cmbXValues.SelectedItem.ToString)
        chartProperties.Add(xValuesFieldName)
        Dim yValuesHighFieldName As New XElement("YValuesFieldName", DataGridView1.Rows(0).Cells(1).Value)
        chartProperties.Add(yValuesHighFieldName)

        'Dim yValuesHighFieldName As New XElement("YValuesHighFieldName", DataGridView1.Rows(0).Cells(1).Value)
        'chartProperties.Add(yValuesHighFieldName)
        'Dim yValuesLowFieldName As New XElement("YValuesLowFieldName", DataGridView1.Rows(1).Cells(1).Value)
        'chartProperties.Add(yValuesLowFieldName)
        'Dim yValuesOpenFieldName As New XElement("YValuesOpenFieldName", DataGridView1.Rows(2).Cells(1).Value)
        'chartProperties.Add(yValuesOpenFieldName)
        'Dim yValuesCloseFieldName As New XElement("YValuesCloseFieldName", DataGridView1.Rows(3).Cells(1).Value)
        'chartProperties.Add(yValuesCloseFieldName)

        chartSettings.Add(chartProperties)

        Dim chartTitle As New XElement("ChartTitle")
        Dim chartTitleLabelName As New XElement("LabelName", "Label1")
        chartTitle.Add(chartTitleLabelName)
        Dim chartTitleText As New XElement("Text", txtChartTitle.Text)
        chartTitle.Add(chartTitleText)

        'If StockChartDefaults.<StockChart>.<Settings>.<ChartTitle>.<FontName>.Value <> Nothing Then
        If StockChartSettingsList.<StockChart>.<Settings>.<ChartTitle>.<FontName>.Value <> Nothing Then
            'Dim chartTitleFontName As New XElement("FontName", StockChartDefaults.<StockChart>.<Settings>.<ChartTitle>.<FontName>.Value)
            Dim chartTitleFontName As New XElement("FontName", StockChartSettingsList.<StockChart>.<Settings>.<ChartTitle>.<FontName>.Value)
            chartTitle.Add(chartTitleFontName)
        Else
            Message.AddWarning("Default Chart Title Font Name settings not found." & vbCrLf)
            Dim chartTitleFontName As New XElement("FontName", txtChartTitle.Font.Name)
            chartTitle.Add(chartTitleFontName)
        End If

        If StockChartSettingsList.<StockChart>.<Settings>.<ChartTitle>.<Color>.Value <> Nothing Then
            Dim chartTitleColor As New XElement("Color", StockChartSettingsList.<StockChart>.<Settings>.<ChartTitle>.<Color>.Value)
            chartTitle.Add(chartTitleColor)
        Else
            Message.AddWarning("Default Chart Title Color settings not found." & vbCrLf)
            Dim chartTitleColor As New XElement("Color", txtChartTitle.ForeColor)
            chartTitle.Add(chartTitleColor)
        End If

        If StockChartSettingsList.<StockChart>.<Settings>.<ChartTitle>.<Size>.Value <> Nothing Then
            Dim chartTitleSize As New XElement("Size", StockChartSettingsList.<StockChart>.<Settings>.<ChartTitle>.<Size>.Value)
            chartTitle.Add(chartTitleSize)
        Else
            Message.AddWarning("Default Chart Title Size settings not found." & vbCrLf)
            Dim chartTitleSize As New XElement("Size", txtChartTitle.Font.Size)
            chartTitle.Add(chartTitleSize)
        End If

        If StockChartSettingsList.<StockChart>.<Settings>.<ChartTitle>.<Bold>.Value <> Nothing Then
            Dim chartTitleBold As New XElement("Bold", StockChartSettingsList.<StockChart>.<Settings>.<ChartTitle>.<Bold>.Value)
            chartTitle.Add(chartTitleBold)
        Else
            Message.AddWarning("Default Chart Title Bold settings not found." & vbCrLf)
            Dim chartTitleBold As New XElement("Bold", txtChartTitle.Font.Bold)
            chartTitle.Add(chartTitleBold)
        End If

        If StockChartSettingsList.<StockChart>.<Settings>.<ChartTitle>.<Italic>.Value <> Nothing Then
            Dim chartTitleItalic As New XElement("Italic", StockChartSettingsList.<StockChart>.<Settings>.<ChartTitle>.<Italic>.Value)
            chartTitle.Add(chartTitleItalic)
        Else
            Message.AddWarning("Default Chart Title Italic settings not found." & vbCrLf)
            Dim chartTitleItalic As New XElement("Italic", txtChartTitle.Font.Italic)
            chartTitle.Add(chartTitleItalic)
        End If

        If StockChartSettingsList.<StockChart>.<Settings>.<ChartTitle>.<Underline>.Value <> Nothing Then
            Dim chartTitleUnderline As New XElement("Underline", StockChartSettingsList.<StockChart>.<Settings>.<ChartTitle>.<Underline>.Value)
            chartTitle.Add(chartTitleUnderline)
        Else
            Message.AddWarning("Default Chart Title Underline settings not found." & vbCrLf)
            Dim chartTitleUnderline As New XElement("Underline", txtChartTitle.Font.Underline)
            chartTitle.Add(chartTitleUnderline)
        End If

        If StockChartSettingsList.<StockChart>.<Settings>.<ChartTitle>.<Strikeout>.Value <> Nothing Then
            Dim chartTitleStrikeout As New XElement("Strikeout", StockChartSettingsList.<StockChart>.<Settings>.<ChartTitle>.<Strikeout>.Value)
            chartTitle.Add(chartTitleStrikeout)
        Else
            Message.AddWarning("Default Chart Title Strikeout settings not found." & vbCrLf)
            Dim chartTitleStrikeout As New XElement("Strikeout", txtChartTitle.Font.Strikeout)
            chartTitle.Add(chartTitleStrikeout)
        End If

        If StockChartSettingsList.<StockChart>.<Settings>.<ChartTitle>.<Alignment>.Value <> Nothing Then
            Dim chartTitleAlignment As New XElement("Alignment", StockChartSettingsList.<StockChart>.<Settings>.<ChartTitle>.<Alignment>.Value)
            chartTitle.Add(chartTitleAlignment)
        Else
            Message.AddWarning("Default Chart Title Alignment settings not found." & vbCrLf)
            Dim chartTitleAlignment As New XElement("Alignment", cmbAlignment.SelectedItem.ToString)
            chartTitle.Add(chartTitleAlignment)
        End If

        chartSettings.Add(chartTitle)

        Dim xAxis As New XElement("XAxis")

        If StockChartSettingsList.<StockChart>.<Settings>.<XAxis>.<TitleText>.Value <> Nothing Then
            Dim xAxisTitleText As New XElement("TitleText", StockChartSettingsList.<StockChart>.<Settings>.<XAxis>.<TitleText>.Value)
            xAxis.Add(xAxisTitleText)
        End If

        If StockChartSettingsList.<StockChart>.<Settings>.<XAxis>.<TitleFontName>.Value <> Nothing Then
            Dim xAxisTitleFontName As New XElement("TitleFontName", StockChartSettingsList.<StockChart>.<Settings>.<XAxis>.<TitleFontName>.Value)
            xAxis.Add(xAxisTitleFontName)
        End If

        If StockChartSettingsList.<StockChart>.<Settings>.<XAxis>.<TitleColor>.Value <> Nothing Then
            Dim xAxisTitleColor As New XElement("TitleColor", StockChartSettingsList.<StockChart>.<Settings>.<XAxis>.<TitleColor>.Value)
            xAxis.Add(xAxisTitleColor)
        End If

        If StockChartSettingsList.<StockChart>.<Settings>.<XAxis>.<TitleSize>.Value <> Nothing Then
            Dim xAxisTitleSize As New XElement("TitleSize", StockChartSettingsList.<StockChart>.<Settings>.<XAxis>.<TitleSize>.Value)
            xAxis.Add(xAxisTitleSize)
        End If

        If StockChartSettingsList.<StockChart>.<Settings>.<XAxis>.<TitleBold>.Value <> Nothing Then
            Dim xAxisTitleBold As New XElement("TitleBold", StockChartSettingsList.<StockChart>.<Settings>.<XAxis>.<TitleBold>.Value)
            xAxis.Add(xAxisTitleBold)
        End If

        If StockChartSettingsList.<StockChart>.<Settings>.<XAxis>.<TitleItalic>.Value <> Nothing Then
            Dim xAxisTitleItalic As New XElement("TitleItalic", StockChartSettingsList.<StockChart>.<Settings>.<XAxis>.<TitleItalic>.Value)
            xAxis.Add(xAxisTitleItalic)
        End If

        If StockChartSettingsList.<StockChart>.<Settings>.<XAxis>.<TitleUnderline>.Value <> Nothing Then
            Dim xAxisTitleUnderline As New XElement("TitleUnderline", StockChartSettingsList.<StockChart>.<Settings>.<XAxis>.<TitleUnderline>.Value)
            xAxis.Add(xAxisTitleUnderline)
        End If

        If StockChartSettingsList.<StockChart>.<Settings>.<XAxis>.<TitleStrikeout>.Value <> Nothing Then
            Dim xAxisTitleStrikeout As New XElement("TitleStrikeout", StockChartSettingsList.<StockChart>.<Settings>.<XAxis>.<TitleStrikeout>.Value)
            xAxis.Add(xAxisTitleStrikeout)
        End If

        If StockChartSettingsList.<StockChart>.<Settings>.<XAxis>.<TitleAlignment>.Value <> Nothing Then
            Dim xAxisTitleAlignment As New XElement("TitleAlignment", StockChartSettingsList.<StockChart>.<Settings>.<XAxis>.<TitleAlignment>.Value)
            xAxis.Add(xAxisTitleAlignment)
        End If

        If StockChartSettingsList.<StockChart>.<Settings>.<XAxis>.<AutoMinimum>.Value <> Nothing Then
            Dim xAxisAutoMinimum As New XElement("AutoMinimum", StockChartSettingsList.<StockChart>.<Settings>.<XAxis>.<AutoMinimum>.Value)
            xAxis.Add(xAxisAutoMinimum)
        End If

        If StockChartSettingsList.<StockChart>.<Settings>.<XAxis>.<Minimum>.Value <> Nothing Then
            Dim xAxisMinimum As New XElement("Minimum", StockChartSettingsList.<StockChart>.<Settings>.<XAxis>.<Minimum>.Value)
            xAxis.Add(xAxisMinimum)
        End If

        If StockChartSettingsList.<StockChart>.<Settings>.<XAxis>.<AutoMaximum>.Value <> Nothing Then
            Dim xAxisAutoMaximum As New XElement("AutoMaximum", StockChartSettingsList.<StockChart>.<Settings>.<XAxis>.<AutoMaximum>.Value)
            xAxis.Add(xAxisAutoMaximum)
        End If

        If StockChartSettingsList.<StockChart>.<Settings>.<XAxis>.<Maximum>.Value <> Nothing Then
            Dim xAxisMaximum As New XElement("Maximum", StockChartSettingsList.<StockChart>.<Settings>.<XAxis>.<Maximum>.Value)
            xAxis.Add(xAxisMaximum)
        End If

        If StockChartSettingsList.<StockChart>.<Settings>.<XAxis>.<AutoInterval>.Value <> Nothing Then
            Dim xAxisAutoInterval As New XElement("AutoInterval", StockChartSettingsList.<StockChart>.<Settings>.<XAxis>.<AutoInterval>.Value)
            xAxis.Add(xAxisAutoInterval)
        End If

        If StockChartSettingsList.<StockChart>.<Settings>.<XAxis>.<MajorGridInterval>.Value <> Nothing Then
            Dim xAxisMajorGridInterval As New XElement("MajorGridInterval", StockChartSettingsList.<StockChart>.<Settings>.<XAxis>.<MajorGridInterval>.Value)
            xAxis.Add(xAxisMajorGridInterval)
        End If

        If StockChartSettingsList.<StockChart>.<Settings>.<XAxis>.<AutoMajorGridInterval>.Value <> Nothing Then
            Dim xAxisAutoMajorGridInterval As New XElement("AutoMajorGridInterval", StockChartSettingsList.<StockChart>.<Settings>.<XAxis>.<AutoMajorGridInterval>.Value)
            xAxis.Add(xAxisAutoMajorGridInterval)
        End If

        chartSettings.Add(xAxis)

        Dim yAxis As New XElement("YAxis")

        If StockChartSettingsList.<StockChart>.<Settings>.<YAxis>.<TitleText>.Value <> Nothing Then
            Dim yAxisTitleText As New XElement("TitleText", StockChartSettingsList.<StockChart>.<Settings>.<YAxis>.<TitleText>.Value)
            yAxis.Add(yAxisTitleText)
        End If

        If StockChartSettingsList.<StockChart>.<Settings>.<YAxis>.<TitleFontName>.Value <> Nothing Then
            Dim yAxisTitleFontName As New XElement("TitleFontName", StockChartSettingsList.<StockChart>.<Settings>.<YAxis>.<TitleFontName>.Value)
            yAxis.Add(yAxisTitleFontName)
        End If

        If StockChartSettingsList.<StockChart>.<Settings>.<YAxis>.<TitleColor>.Value <> Nothing Then
            Dim yAxisTitleColor As New XElement("TitleColor", StockChartSettingsList.<StockChart>.<Settings>.<YAxis>.<TitleColor>.Value)
            yAxis.Add(yAxisTitleColor)
        End If

        If StockChartSettingsList.<StockChart>.<Settings>.<YAxis>.<TitleSize>.Value <> Nothing Then
            Dim yAxisTitleSize As New XElement("TitleSize", StockChartSettingsList.<StockChart>.<Settings>.<YAxis>.<TitleSize>.Value)
            yAxis.Add(yAxisTitleSize)
        End If

        If StockChartSettingsList.<StockChart>.<Settings>.<YAxis>.<TitleBold>.Value <> Nothing Then
            Dim yAxisTitleBold As New XElement("TitleBold", StockChartSettingsList.<StockChart>.<Settings>.<YAxis>.<TitleBold>.Value)
            yAxis.Add(yAxisTitleBold)
        End If

        If StockChartSettingsList.<StockChart>.<Settings>.<YAxis>.<TitleItalic>.Value <> Nothing Then
            Dim yAxisTitleItalic As New XElement("TitleItalic", StockChartSettingsList.<StockChart>.<Settings>.<YAxis>.<TitleItalic>.Value)
            yAxis.Add(yAxisTitleItalic)
        End If

        If StockChartSettingsList.<StockChart>.<Settings>.<YAxis>.<TitleUnderline>.Value <> Nothing Then
            Dim yAxisTitleUnderline As New XElement("TitleUnderline", StockChartSettingsList.<StockChart>.<Settings>.<YAxis>.<TitleUnderline>.Value)
            yAxis.Add(yAxisTitleUnderline)
        End If

        If StockChartSettingsList.<StockChart>.<Settings>.<YAxis>.<TitleStrikeout>.Value <> Nothing Then
            Dim yAxisTitleStrikeout As New XElement("TitleStrikeout", StockChartSettingsList.<StockChart>.<Settings>.<YAxis>.<TitleStrikeout>.Value)
            yAxis.Add(yAxisTitleStrikeout)
        End If

        If StockChartSettingsList.<StockChart>.<Settings>.<YAxis>.<TitleAlignment>.Value <> Nothing Then
            Dim yAxisTitleAlignment As New XElement("TitleAlignment", StockChartSettingsList.<StockChart>.<Settings>.<YAxis>.<TitleAlignment>.Value)
            yAxis.Add(yAxisTitleAlignment)
        End If

        If StockChartSettingsList.<StockChart>.<Settings>.<YAxis>.<AutoMinimum>.Value <> Nothing Then
            Dim yAxisAutoMinimum As New XElement("AutoMinimum", StockChartSettingsList.<StockChart>.<Settings>.<YAxis>.<AutoMinimum>.Value)
            yAxis.Add(yAxisAutoMinimum)
        End If

        If StockChartSettingsList.<StockChart>.<Settings>.<YAxis>.<Minimum>.Value <> Nothing Then
            Dim yAxisMinimum As New XElement("Minimum", StockChartSettingsList.<StockChart>.<Settings>.<YAxis>.<Minimum>.Value)
            yAxis.Add(yAxisMinimum)
        End If

        If StockChartSettingsList.<StockChart>.<Settings>.<YAxis>.<AutoMaximum>.Value <> Nothing Then
            Dim yAxisAutoMaximum As New XElement("AutoMaximum", StockChartSettingsList.<StockChart>.<Settings>.<YAxis>.<AutoMaximum>.Value)
            yAxis.Add(yAxisAutoMaximum)
        End If

        If StockChartSettingsList.<StockChart>.<Settings>.<YAxis>.<Maximum>.Value <> Nothing Then
            Dim yAxisMaximum As New XElement("Maximum", StockChartSettingsList.<StockChart>.<Settings>.<YAxis>.<Maximum>.Value)
            yAxis.Add(yAxisMaximum)
        End If

        If StockChartSettingsList.<StockChart>.<Settings>.<YAxis>.<AutoInterval>.Value <> Nothing Then
            Dim yAxisAutoInterval As New XElement("AutoInterval", StockChartSettingsList.<StockChart>.<Settings>.<YAxis>.<AutoInterval>.Value)
            yAxis.Add(yAxisAutoInterval)
        End If

        If StockChartSettingsList.<StockChart>.<Settings>.<YAxis>.<MajorGridInterval>.Value <> Nothing Then
            Dim yAxisMajorGridInterval As New XElement("MajorGridInterval", StockChartSettingsList.<StockChart>.<Settings>.<YAxis>.<MajorGridInterval>.Value)
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

            'client.SendMessageAsync(AppNetName, "ADVL_Stock_Chart_1", doc.ToString) 'Added 3Feb19
            client.SendMessageAsync(ProNetName, "ADVL_Stock_Chart_1", doc.ToString)
            'Message.XAddText("Message sent to " & "ADVL_Stock_Chart_1" & ":" & vbCrLf, "XmlSentNotice")
            Message.XAddText("Message sent to [" & ProNetName & "]." & "ADVL_Stock_Chart_1" & ":" & vbCrLf, "XmlSentNotice")
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

        'Check if a Share Chart Project has been selected:
        If SelShareChartProjNo = -1 Then
            Message.AddWarning("A Share Chart project has not been selected." & vbCrLf)
            Exit Sub
        End If

        'Check if the selected Share Chart Project is in the current Project Network:
        Dim StockChartProNetName As String = ShareChartProj.List(SelShareChartProjNo).ProNetName
        If StockChartProNetName <> ProNetName Then
            Message.AddWarning("The selected Share Chart project is not in the current Project Network." & vbCrLf)
            Exit Sub
        End If

        'Check if The Stock Chart Project is connected:
        Dim StockChartProjectName As String = ShareChartProj.List(SelShareChartProjNo).Name
        Dim StockChartConnName As String = client.ConnNameFromProjName(StockChartProjectName, ProNetName, "ADVL_Stock_Chart_1")
        If StockChartConnName = "" Then
            'Connect the Stock Chart application to the Project Network:
            client.StartProjectWithName(StockChartProjectName, ProNetName, "ADVL_Stock_Chart_1", "ADVL_Stock_Chart_1")
        End If

        If chkUseStockChartSettingsList.Checked Then
            'DisplayStockChartUsingDefaults()
            DisplayStockChartUsingSettingsList()
        Else
            'DisplayStockChartNoDefaults()
            DisplayStockChartNoSettingsList()
        End If
    End Sub

    Private Sub DisplayStockChartNoSettingsList()
        'Display the Stock Chart without using a Settings List.

        Message.Add("Displaying the Stock Chart with no Settings List" & vbCrLf)

        'Send the instructions to the Chart application to display the stock chart.

        'Check that required selections have been made:
        If cmbXValues.SelectedItem Is Nothing Then
            Message.AddWarning("Select a field for the X Values." & vbCrLf)
            Exit Sub
        End If

        'Build the XMessageBlock containing the Stock Chart settings.
        'This will be send to the Stock Chart application to create the chart display.

        Dim StockChartXMsgBlk = <?xml version="1.0" encoding="utf-8"?>
                                <XMsgBlk>
                                    <ClientLocn>DisplayChart</ClientLocn>
                                    <XInfo>
                                        <ChartSettings>
                                            <!--Input Data:-->
                                            <InputDataType>Database</InputDataType>
                                            <InputDatabasePath><%= txtSPChartDbPath.Text %></InputDatabasePath>
                                            <InputQuery><%= txtSPChartQuery.Text %></InputQuery>
                                            <InputDataDescr><%= txtSeriesName.Text %></InputDataDescr>
                                            <SeriesInfoList>
                                                <SeriesInfo>
                                                    <Name><%= txtSeriesName.Text %></Name>
                                                    <XValuesFieldName><%= cmbXValues.SelectedItem.ToString %></XValuesFieldName>
                                                    <YValuesHighFieldName><%= DataGridView1.Rows(0).Cells(1).Value %></YValuesHighFieldName>
                                                    <YValuesLowFieldName><%= DataGridView1.Rows(1).Cells(1).Value %></YValuesLowFieldName>
                                                    <YValuesOpenFieldName><%= DataGridView1.Rows(2).Cells(1).Value %></YValuesOpenFieldName>
                                                    <YValuesCloseFieldName><%= DataGridView1.Rows(3).Cells(1).Value %></YValuesCloseFieldName>
                                                    <ChartArea>ChartArea1</ChartArea>
                                                </SeriesInfo>
                                            </SeriesInfoList>
                                            <AreaInfoList>
                                                <AreaInfo>
                                                    <Name>ChartArea1</Name>
                                                    <AutoXAxisMinimum>true</AutoXAxisMinimum>
                                                    <AutoXAxisMaximum>true</AutoXAxisMaximum>
                                                    <AutoXAxisMajorGridInterval>true</AutoXAxisMajorGridInterval>
                                                    <AutoX2AxisMinimum>true</AutoX2AxisMinimum>
                                                    <AutoX2AxisMaximum>true</AutoX2AxisMaximum>
                                                    <AutoX2AxisMajorGridInterval>true</AutoX2AxisMajorGridInterval>
                                                    <AutoYAxisMinimum>true</AutoYAxisMinimum>
                                                    <AutoYAxisMaximum>true</AutoYAxisMaximum>
                                                    <AutoYAxisMajorGridInterval>true</AutoYAxisMajorGridInterval>
                                                    <AutoY2AxisMinimum>true</AutoY2AxisMinimum>
                                                    <AutoY2AxisMaximum>true</AutoY2AxisMaximum>
                                                    <AutoY2AxisMajorGridInterval>true</AutoY2AxisMajorGridInterval>
                                                </AreaInfo>
                                            </AreaInfoList>
                                            <!--Chart Properties:-->
                                            <TitlesCollection>
                                                <Title>
                                                    <Name>Title1</Name>
                                                    <Text>Stock Chart</Text>
                                                    <TextOrientation>Auto</TextOrientation>
                                                    <Alignment>MiddleCenter</Alignment>
                                                    <ForeColor>-16777216</ForeColor>
                                                    <Font>
                                                        <Name>Microsoft Sans Serif</Name>
                                                        <Size>14.25</Size>
                                                        <Bold>true</Bold>
                                                        <Italic>false</Italic>
                                                        <Strikeout>false</Strikeout>
                                                        <Underline>false</Underline>
                                                    </Font>
                                                </Title>
                                            </TitlesCollection>
                                            <SeriesCollection>
                                                <Series>
                                                    <Name><%= txtSeriesName.Text %></Name>
                                                    <ChartType>Stock</ChartType>
                                                    <ChartArea>ChartArea1</ChartArea>
                                                    <Legend>Legend1</Legend>
                                                    <LabelValueType>Close</LabelValueType>
                                                    <MaxPixelPointWidth>2</MaxPixelPointWidth>
                                                    <MinPixelPointWidth>1</MinPixelPointWidth>
                                                    <OpenCloseStyle>Line</OpenCloseStyle>
                                                    <PixelPointDepth>2</PixelPointDepth>
                                                    <PixelPointGapDepth>2</PixelPointGapDepth>
                                                    <PixelPointWidth>2</PixelPointWidth>
                                                    <PointWidth>2</PointWidth>
                                                    <ShowOpenClose>Both</ShowOpenClose>
                                                    <AxisLabel/>
                                                    <XAxisType>Primary</XAxisType>
                                                    <XValueType>Date</XValueType>
                                                    <YAxisType>Primary</YAxisType>
                                                    <YValueType>Single</YValueType>
                                                    <Marker>
                                                        <BorderColor>-16777216</BorderColor>
                                                        <BorderWidth>1</BorderWidth>
                                                        <Color>-8355712</Color>
                                                        <Size>5</Size>
                                                        <Step>1</Step>
                                                        <Style>None</Style>
                                                    </Marker>
                                                    <Color>-16776961</Color>
                                                </Series>
                                            </SeriesCollection>
                                            <ChartAreasCollection>
                                                <ChartArea>
                                                    <Name>ChartArea1</Name>
                                                    <AxisX>
                                                        <Title>
                                                            <Text>Trade Date</Text>
                                                            <Alignment>Center</Alignment>
                                                            <ForeColor>-16777216</ForeColor>
                                                            <Font>
                                                                <Name>Microsoft Sans Serif</Name>
                                                                <Size>12</Size>
                                                                <Bold>true</Bold>
                                                                <Italic>false</Italic>
                                                                <Strikeout>false</Strikeout>
                                                                <Underline>false</Underline>
                                                            </Font>
                                                        </Title>
                                                        <LabelStyleFormat/>
                                                        <Minimum>40910</Minimum>
                                                        <Maximum>43743</Maximum>
                                                        <LineWidth>1</LineWidth>
                                                        <Interval>0</Interval>
                                                        <IntervalOffset>0</IntervalOffset>
                                                        <Crossing>NaN</Crossing>
                                                        <MajorGrid>
                                                            <Interval>6</Interval>
                                                            <IntervalOffset>NaN</IntervalOffset>
                                                        </MajorGrid>
                                                    </AxisX>
                                                    <AxisX2>
                                                        <Title>
                                                            <Text/>
                                                            <Alignment>Center</Alignment>
                                                            <ForeColor>-16777216</ForeColor>
                                                            <Font>
                                                                <Name>Microsoft Sans Serif</Name>
                                                                <Size>8</Size>
                                                                <Bold>false</Bold>
                                                                <Italic>false</Italic>
                                                                <Strikeout>false</Strikeout>
                                                                <Underline>false</Underline>
                                                            </Font>
                                                        </Title>
                                                        <LabelStyleFormat/>
                                                        <Minimum>40910</Minimum>
                                                        <Maximum>43743</Maximum>
                                                        <LineWidth>1</LineWidth>
                                                        <Interval>0</Interval>
                                                        <IntervalOffset>0</IntervalOffset>
                                                        <Crossing>NaN</Crossing>
                                                        <MajorGrid>
                                                            <Interval>2</Interval>
                                                            <IntervalOffset>NaN</IntervalOffset>
                                                        </MajorGrid>
                                                    </AxisX2>
                                                    <AxisY>
                                                        <Title>
                                                            <Text>Price</Text>
                                                            <Alignment>Center</Alignment>
                                                            <ForeColor>-16777216</ForeColor>
                                                            <Font>
                                                                <Name>Microsoft Sans Serif</Name>
                                                                <Size>12</Size>
                                                                <Bold>true</Bold>
                                                                <Italic>false</Italic>
                                                                <Strikeout>false</Strikeout>
                                                                <Underline>false</Underline>
                                                            </Font>
                                                        </Title>
                                                        <LabelStyleFormat/>
                                                        <Minimum>0</Minimum>
                                                        <Maximum>50</Maximum>
                                                        <LineWidth>1</LineWidth>
                                                        <Interval>0</Interval>
                                                        <IntervalOffset>0</IntervalOffset>
                                                        <Crossing>NaN</Crossing>
                                                        <MajorGrid>
                                                            <Interval>10</Interval>
                                                            <IntervalOffset>NaN</IntervalOffset>
                                                        </MajorGrid>
                                                    </AxisY>
                                                    <AxisY2>
                                                        <Title>
                                                            <Text/>
                                                            <Alignment>Center</Alignment>
                                                            <ForeColor>-16777216</ForeColor>
                                                            <Font>
                                                                <Name>Microsoft Sans Serif</Name>
                                                                <Size>8</Size>
                                                                <Bold>false</Bold>
                                                                <Italic>false</Italic>
                                                                <Strikeout>false</Strikeout>
                                                                <Underline>false</Underline>
                                                            </Font>
                                                        </Title>
                                                        <LabelStyleFormat/>
                                                        <Minimum>0</Minimum>
                                                        <Maximum>50</Maximum>
                                                        <LineWidth>1</LineWidth>
                                                        <Interval>0</Interval>
                                                        <IntervalOffset>0</IntervalOffset>
                                                        <Crossing>NaN</Crossing>
                                                        <MajorGrid>
                                                            <Interval>10</Interval>
                                                            <IntervalOffset>NaN</IntervalOffset>
                                                        </MajorGrid>
                                                    </AxisY2>
                                                </ChartArea>
                                            </ChartAreasCollection>
                                        </ChartSettings>
                                    </XInfo>
                                </XMsgBlk>

        'Send the XMessageBlock to the Stock Chart application:
        Message.XAddText("Message sent to [" & ProNetName & "]." & "ADVL_Stock_Chart_1" & ":" & vbCrLf, "XmlSentNotice")
        Message.XAddXml(StockChartXMsgBlk.ToString)
        Message.XAddText(vbCrLf, "Normal") 'Add extra line

        SendMessageParams.ProjectNetworkName = ProNetName
        SendMessageParams.ConnectionName = "ADVL_Stock_Chart_1"
        SendMessageParams.Message = StockChartXMsgBlk.ToString
        If bgwSendMessage.IsBusy Then
            Message.AddWarning("Send Message backgroundworker is busy." & vbCrLf)
        Else
            bgwSendMessage.RunWorkerAsync(SendMessageParams)
        End If

    End Sub

    Private Sub DisplayStockChartUsingSettingsList()
        'Display the Stock Chart using a Settings List.

        Message.Add("Displaying the Stock Chart using the Settings List" & vbCrLf)

        'Send the instructions to the Chart application to display the stock chart.

        'Check that required selections have been made:
        If cmbXValues.SelectedItem Is Nothing Then
            Message.AddWarning("Select a field for the X Values." & vbCrLf)
            Exit Sub
        End If

        'Build the XMessageBlock containing the Stock Chart settings.
        'This will be send to the Stock Chart application to create the chart display.

        Dim ChartSettingsList As XDocument = XDocument.Parse("<?xml version=""1.0"" encoding=""utf-8"" standalone=""yes""?>" & vbCrLf & XmlStockChartSettingsList.Text)

        Dim StockChartXMsgBlk = <?xml version="1.0" encoding="utf-8"?>
                                <XMsgBlk>
                                    <ClientLocn>DisplayChart</ClientLocn>
                                    <XInfo>
                                        <%= ChartSettingsList.<ChartSettings> %>
                                    </XInfo>
                                </XMsgBlk>

        'Update the Settings List with the current chart settings:

        'Update the Input Data settinga:
        '<InputDataType>Database</InputDataType> - Currently only the Database type is available.
        StockChartXMsgBlk.<XMsgBlk>.<XInfo>.<ChartSettings>.<InputDatabasePath>.Value = txtSPChartDbPath.Text
        StockChartXMsgBlk.<XMsgBlk>.<XInfo>.<ChartSettings>.<InputQuery>.Value = txtSPChartQuery.Text
        StockChartXMsgBlk.<XMsgBlk>.<XInfo>.<ChartSettings>.<InputDataDescr>.Value = txtSeriesName.Text

        'Add a warning if there is more than one entry in the SeriesInfoList:
        If StockChartXMsgBlk.<XMsgBlk>.<XInfo>.<ChartSettings>.<SeriesInfoList>.<SeriesInfo>.Count > 1 Then AddWarning("There is more than one entry in the Series Info List!" & vbCrLf)
        'Update the first entry in the SeriesInfoList:
        StockChartXMsgBlk.<XMsgBlk>.<XInfo>.<ChartSettings>.<SeriesInfoList>.<SeriesInfo>.<Name>.Value = txtSeriesName.Text
        StockChartXMsgBlk.<XMsgBlk>.<XInfo>.<ChartSettings>.<SeriesInfoList>.<SeriesInfo>.<XValuesFieldName>.Value = cmbXValues.SelectedItem.ToString
        StockChartXMsgBlk.<XMsgBlk>.<XInfo>.<ChartSettings>.<SeriesInfoList>.<SeriesInfo>.<YValuesHighFieldName>.Value = DataGridView1.Rows(0).Cells(1).Value
        StockChartXMsgBlk.<XMsgBlk>.<XInfo>.<ChartSettings>.<SeriesInfoList>.<SeriesInfo>.<YValuesLowFieldName>.Value = DataGridView1.Rows(1).Cells(1).Value
        StockChartXMsgBlk.<XMsgBlk>.<XInfo>.<ChartSettings>.<SeriesInfoList>.<SeriesInfo>.<YValuesOpenFieldName>.Value = DataGridView1.Rows(2).Cells(1).Value
        StockChartXMsgBlk.<XMsgBlk>.<XInfo>.<ChartSettings>.<SeriesInfoList>.<SeriesInfo>.<YValuesCloseFieldName>.Value = DataGridView1.Rows(3).Cells(1).Value

        'Leave the AreaInfoList unchanged.

        'Update the Chart title settings:
        StockChartXMsgBlk.<XMsgBlk>.<XInfo>.<ChartSettings>.<TitlesCollection>.<Title>.<Text>.Value = txtChartTitle.Text
        StockChartXMsgBlk.<XMsgBlk>.<XInfo>.<ChartSettings>.<TitlesCollection>.<Title>.<Alignment>.Value = cmbAlignment.SelectedItem.ToString
        If txtChartTitle.ForeColor.ToArgb.ToString = "0" Then 'This color value is not valid for a chart title.
            StockChartXMsgBlk.<XMsgBlk>.<XInfo>.<ChartSettings>.<TitlesCollection>.<Title>.<ForeColor>.Value = Color.Black.ToArgb.ToString
        Else
            StockChartXMsgBlk.<XMsgBlk>.<XInfo>.<ChartSettings>.<TitlesCollection>.<Title>.<ForeColor>.Value = txtChartTitle.ForeColor.ToArgb.ToString
        End If
        StockChartXMsgBlk.<XMsgBlk>.<XInfo>.<ChartSettings>.<TitlesCollection>.<Title>.<Font>.<Name>.Value = txtChartTitle.Font.Name
        StockChartXMsgBlk.<XMsgBlk>.<XInfo>.<ChartSettings>.<TitlesCollection>.<Title>.<Font>.<Size>.Value = txtChartTitle.Font.Size
        StockChartXMsgBlk.<XMsgBlk>.<XInfo>.<ChartSettings>.<TitlesCollection>.<Title>.<Font>.<Bold>.Value = txtChartTitle.Font.Bold
        StockChartXMsgBlk.<XMsgBlk>.<XInfo>.<ChartSettings>.<TitlesCollection>.<Title>.<Font>.<Italic>.Value = txtChartTitle.Font.Italic
        StockChartXMsgBlk.<XMsgBlk>.<XInfo>.<ChartSettings>.<TitlesCollection>.<Title>.<Font>.<Strikeout>.Value = txtChartTitle.Font.Strikeout
        StockChartXMsgBlk.<XMsgBlk>.<XInfo>.<ChartSettings>.<TitlesCollection>.<Title>.<Font>.<Underline>.Value = txtChartTitle.Font.Underline

        'Add a warning if there is more than one entry in the SeriesCollection:
        If StockChartXMsgBlk.<XMsgBlk>.<XInfo>.<ChartSettings>.<SeriesCollection>.<Series>.Count > 1 Then AddWarning("There is more than one entry in the Series Collection!" & vbCrLf)
        'Update the first entry in the SeriesCollection:
        StockChartXMsgBlk.<XMsgBlk>.<XInfo>.<ChartSettings>.<SeriesCollection>.<Series>.<Name>.Value = txtSeriesName.Text


        'Send the XMessageBlock to the Stock Chart application:
        Message.XAddText("Message sent to [" & ProNetName & "]." & "ADVL_Stock_Chart_1" & ":" & vbCrLf, "XmlSentNotice")
        Message.XAddXml(StockChartXMsgBlk.ToString)
        Message.XAddText(vbCrLf, "Normal") 'Add extra line

        SendMessageParams.ProjectNetworkName = ProNetName
        SendMessageParams.ConnectionName = "ADVL_Stock_Chart_1"
        SendMessageParams.Message = StockChartXMsgBlk.ToString
        If bgwSendMessage.IsBusy Then
            Message.AddWarning("Send Message backgroundworker is busy." & vbCrLf)
        Else
            bgwSendMessage.RunWorkerAsync(SendMessageParams)
        End If

    End Sub

    'Private Sub DisplayStockChartNoDefaults()
    Private Sub DisplayStockChartNoSettingsList_Old()
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

        'Dim clientAppNetName As New XElement("ClientAppNetName", AppNetName)
        'xmessage.Add(clientAppNetName)
        Dim clientProNetName As New XElement("ClientProNetName", ProNetName)
        xmessage.Add(clientProNetName)

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
            client.SendMessageAsync(ProNetName, "ADVL_Stock_Chart_1", doc.ToString)
            Message.XAddText("Message sent to [" & ProNetName & "]." & "ADVL_Stock_Chart_1" & ":" & vbCrLf, "XmlSentNotice")
            Message.XAddXml(doc.ToString)
            Message.XAddText(vbCrLf, "Message") 'Add extra line
        End If
    End Sub

    'Private Sub btnGetStockChartDefaults_Click(sender As Object, e As EventArgs) Handles btnGetStockChartSettings.Click
    Private Sub btnGetStockChartSettings_Click(sender As Object, e As EventArgs) Handles btnGetStockChartSettings.Click
        'Send a request to ADVL_Charts_1 for the current Stock Chart settings.

        'Create the xml instructions
        Dim decl As New XDeclaration("1.0", "utf-8", "yes")
        Dim doc As New XDocument(decl, Nothing) 'Create an XDocument to store the instructions.
        Dim xmessage As New XElement("XMsg") 'This indicates the start of the message in the XMessage class

        'ADDED 3Feb19:
        'Dim clientAppNetName As New XElement("ClientAppNetName", AppNetName)
        'xmessage.Add(clientAppNetName)
        Dim clientProNetName As New XElement("ClientProNetName", ProNetName)
        xmessage.Add(clientProNetName)

        Dim clientName As New XElement("ClientName", ApplicationInfo.Name) 'This tells the coordinate server the name of the client making the request.
        xmessage.Add(clientName)

        Dim clientConnName As New XElement("ClientConnectionName", ApplicationInfo.Name) 'This tells the coordinate server the name of the client making the request.
        xmessage.Add(clientConnName)

        'Dim clientLocn As New XElement("ClientLocn", "StockChart")
        'xmessage.Add(clientLocn)

        'Dim commandGetSettings As New XElement("Command", "GetStockChartSettings")
        'xmessage.Add(commandGetSettings)

        'New Code: Send chart settings using XDataMsg instead of XMsg.
        '  XDataMsg contains a block of data instead of a set of instructions.
        '  Format: 
        '  <XDataMsg>
        '    <ClientLocn>StockChart</ClientLocn>
        '    <XData>
        '      The block of XML data to be send to the Client Location. (The Client Location will contain code to 
        '    </XData>
        '  </XDataMsg>

        Dim getChartDefaultsCommand As New XElement("GetChartSettings", "StockChart")
        xmessage.Add(getChartDefaultsCommand)

        doc.Add(xmessage)

        If IsNothing(client) Then
            Message.AddWarning("No client connection available!" & vbCrLf)
            Beep()
        Else
            client.SendMessageAsync(ProNetName, "ADVL_Stock_Chart_1", doc.ToString)
            Message.XAddText("Message sent to [" & ProNetName & "]." & "ADVL_Stock_Chart_1" & ":" & vbCrLf, "XmlSentNotice")
            Message.XAddXml(doc.ToString)
            Message.XAddText(vbCrLf, "Normal") 'Add extra line

        End If
    End Sub

    'Private Function ApplySettingUpdates(ByRef DefaultSettings As XDocument, ByRef UpdateSettings As XDocument) As XDocument
    Private Function ApplySettingUpdates(ByVal DefaultSettings As XDocument, ByRef UpdateSettings As XDocument) As XDocument
        'Apply the Update Settings to the Default Settings.
        'Return the updated settings XDocument.

        'DefaultSettings is an XDocument containing a collection of default settings.
        'UpdateSettings is an XDocument containing the settings to be updated.



    End Function

    Private Sub DisplayLineChart()
        'Display Line Chart.

        'Check if connected to ComNet:
        If ConnectedToComNet = False Then
            ConnectToComNet()
        End If

        'Check if a Line Chart Project has been selected:
        If SelLineChartProjNo = -1 Then
            Message.AddWarning("A Line Chart project has not been selected." & vbCrLf)
            Exit Sub
        End If

        'Check if the selected Line Chart Project is in the current Project Network:
        Dim LineChartProNetName As String = LineChartProj.List(SelLineChartProjNo).ProNetName
        If LineChartProNetName <> ProNetName Then
            Message.AddWarning("The selected Line Chart project is not in the current Project Network." & vbCrLf)
            Exit Sub
        End If

        'Check if The Line Chart Project is connected:
        Dim LineChartProjectName As String = LineChartProj.List(SelLineChartProjNo).Name
        Dim LineChartConnName As String = client.ConnNameFromProjName(LineChartProjectName, ProNetName, "ADVL_Line_Chart_1")
        If LineChartConnName = "" Then
            'Connect the Line Chart project to the Project Network:
            client.StartProjectWithName(LineChartProjectName, ProNetName, "ADVL_Line_Chart_1", "ADVL_Line_Chart_1")
        End If

        If chkUseLineChartDefaults.Checked Then
            'DisplayLineChartUsingDefaults()
            DisplayLineChartUsingSettingsList()
        Else
            'DisplayLineChartNoDefaults()
            DisplayLineChartNoSettingsList()
        End If
    End Sub

    'Private Sub DisplayLineChartNoDefaults()
    Private Sub DisplayLineChartNoSettingsList()
        'Display the Line chart without using the chart settings list.

        'Check that required selections have been made:
        If cmbLineXValues.SelectedItem Is Nothing Then
            Message.AddWarning("Select a field for the X Values." & vbCrLf)
            Exit Sub
        End If

        'Build the XMessageBlock containing the Point Chart settings.
        'This will be send to the Point Chart application to create the chart display.

        Dim LineChartXMsgBlk = <?xml version="1.0" encoding="utf-8"?>
                               <XMsgBlk>
                                   <ClientLocn>DisplayChart</ClientLocn>
                                   <XInfo>
                                       <ChartSettings>
                                           <!--Input Data:-->
                                           <InputDataType>Database</InputDataType>
                                           <InputDatabasePath><%= txtLineChartDbPath.Text %></InputDatabasePath>
                                           <InputQuery><%= txtLineChartQuery.Text %></InputQuery>
                                           <InputDataDescr><%= txtLineSeriesName.Text %></InputDataDescr>
                                           <SeriesInfoList>
                                               <SeriesInfo>
                                                   <Name><%= txtLineSeriesName.Text %></Name>
                                                   <XValuesFieldName><%= cmbLineXValues.SelectedItem.ToString %></XValuesFieldName>
                                                   <YValuesFieldName><%= cmbLineYValues.SelectedItem.ToString %></YValuesFieldName>
                                                   <ChartArea>ChartArea1</ChartArea>
                                               </SeriesInfo>
                                           </SeriesInfoList>
                                           <AreaInfoList>
                                               <AreaInfo>
                                                   <Name>ChartArea1</Name>
                                                   <AutoXAxisMinimum>true</AutoXAxisMinimum>
                                                   <AutoXAxisMaximum>true</AutoXAxisMaximum>
                                                   <AutoXAxisMajorGridInterval>true</AutoXAxisMajorGridInterval>
                                                   <AutoX2AxisMinimum>true</AutoX2AxisMinimum>
                                                   <AutoX2AxisMaximum>true</AutoX2AxisMaximum>
                                                   <AutoX2AxisMajorGridInterval>true</AutoX2AxisMajorGridInterval>
                                                   <AutoYAxisMinimum>true</AutoYAxisMinimum>
                                                   <AutoYAxisMaximum>true</AutoYAxisMaximum>
                                                   <AutoYAxisMajorGridInterval>true</AutoYAxisMajorGridInterval>
                                                   <AutoY2AxisMinimum>true</AutoY2AxisMinimum>
                                                   <AutoY2AxisMaximum>true</AutoY2AxisMaximum>
                                                   <AutoY2AxisMajorGridInterval>true</AutoY2AxisMajorGridInterval>
                                               </AreaInfo>
                                           </AreaInfoList>
                                           <!--Chart Properties:-->
                                           <TitlesCollection>
                                               <Title>
                                                   <Name>Title1</Name>
                                                   <Text><%= txtLineChartTitle.Text %></Text>
                                                   <TextOrientation>Auto</TextOrientation>
                                                   <Alignment><%= cmbLineChartAlignment.SelectedItem.ToString %></Alignment>
                                                   <ForeColor>-16777216</ForeColor>
                                                   <Font>
                                                       <Name>Microsoft Sans Serif</Name>
                                                       <Size>14.25</Size>
                                                       <Bold>true</Bold>
                                                       <Italic>false</Italic>
                                                       <Strikeout>false</Strikeout>
                                                       <Underline>false</Underline>
                                                   </Font>
                                               </Title>
                                           </TitlesCollection>
                                           <SeriesCollection>
                                               <Series>
                                                   <Name><%= txtLineSeriesName.Text %></Name>
                                                   <ChartType>Line</ChartType>
                                                   <ChartArea>ChartArea1</ChartArea>
                                                   <Legend>Legend1</Legend>
                                                   <EmptyPointValue>Average</EmptyPointValue>
                                                   <LabelStyle>Auto</LabelStyle>
                                                   <PixelPointDepth>1</PixelPointDepth>
                                                   <PixelPointGapDepth>1</PixelPointGapDepth>
                                                   <ShowMarkerLines/>
                                                   <AxisLabel/>
                                                   <XAxisType>Primary</XAxisType>
                                                   <XValueType>Single</XValueType>
                                                   <YAxisType>Primary</YAxisType>
                                                   <YValueType>Single</YValueType>
                                                   <Marker>
                                                       <BorderColor>-16777216</BorderColor>
                                                       <BorderWidth>1</BorderWidth>
                                                       <Color>-8355712</Color>
                                                       <Size>5</Size>
                                                       <Step>1</Step>
                                                       <Style>None</Style>
                                                   </Marker>
                                                   <Color>-16776961</Color>
                                               </Series>
                                           </SeriesCollection>
                                           <ChartAreasCollection>
                                               <ChartArea>
                                                   <Name>ChartArea1</Name>
                                                   <AxisX>
                                                       <Title>
                                                           <Text><%= cmbLineXValues.SelectedItem.ToString %></Text>
                                                           <Alignment>Center</Alignment>
                                                           <ForeColor>-16777216</ForeColor>
                                                           <Font>
                                                               <Name>Microsoft Sans Serif</Name>
                                                               <Size>12</Size>
                                                               <Bold>true</Bold>
                                                               <Italic>false</Italic>
                                                               <Strikeout>false</Strikeout>
                                                               <Underline>false</Underline>
                                                           </Font>
                                                       </Title>
                                                       <LabelStyleFormat/>
                                                       <Minimum>40910</Minimum>
                                                       <Maximum>43743</Maximum>
                                                       <LineWidth>1</LineWidth>
                                                       <Interval>0</Interval>
                                                       <IntervalOffset>0</IntervalOffset>
                                                       <Crossing>NaN</Crossing>
                                                       <MajorGrid>
                                                           <Interval>6</Interval>
                                                           <IntervalOffset>NaN</IntervalOffset>
                                                       </MajorGrid>
                                                   </AxisX>
                                                   <AxisX2>
                                                       <Title>
                                                           <Text/>
                                                           <Alignment>Center</Alignment>
                                                           <ForeColor>-16777216</ForeColor>
                                                           <Font>
                                                               <Name>Microsoft Sans Serif</Name>
                                                               <Size>8</Size>
                                                               <Bold>false</Bold>
                                                               <Italic>false</Italic>
                                                               <Strikeout>false</Strikeout>
                                                               <Underline>false</Underline>
                                                           </Font>
                                                       </Title>
                                                       <LabelStyleFormat/>
                                                       <Minimum>40910</Minimum>
                                                       <Maximum>43743</Maximum>
                                                       <LineWidth>1</LineWidth>
                                                       <Interval>0</Interval>
                                                       <IntervalOffset>0</IntervalOffset>
                                                       <Crossing>NaN</Crossing>
                                                       <MajorGrid>
                                                           <Interval>2</Interval>
                                                           <IntervalOffset>NaN</IntervalOffset>
                                                       </MajorGrid>
                                                   </AxisX2>
                                                   <AxisY>
                                                       <Title>
                                                           <Text><%= cmbLineYValues.SelectedItem.ToString %></Text>
                                                           <Alignment>Center</Alignment>
                                                           <ForeColor>-16777216</ForeColor>
                                                           <Font>
                                                               <Name>Microsoft Sans Serif</Name>
                                                               <Size>12</Size>
                                                               <Bold>true</Bold>
                                                               <Italic>false</Italic>
                                                               <Strikeout>false</Strikeout>
                                                               <Underline>false</Underline>
                                                           </Font>
                                                       </Title>
                                                       <LabelStyleFormat/>
                                                       <Minimum>0</Minimum>
                                                       <Maximum>50</Maximum>
                                                       <LineWidth>1</LineWidth>
                                                       <Interval>0</Interval>
                                                       <IntervalOffset>0</IntervalOffset>
                                                       <Crossing>NaN</Crossing>
                                                       <MajorGrid>
                                                           <Interval>10</Interval>
                                                           <IntervalOffset>NaN</IntervalOffset>
                                                       </MajorGrid>
                                                   </AxisY>
                                                   <AxisY2>
                                                       <Title>
                                                           <Text/>
                                                           <Alignment>Center</Alignment>
                                                           <ForeColor>-16777216</ForeColor>
                                                           <Font>
                                                               <Name>Microsoft Sans Serif</Name>
                                                               <Size>8</Size>
                                                               <Bold>false</Bold>
                                                               <Italic>false</Italic>
                                                               <Strikeout>false</Strikeout>
                                                               <Underline>false</Underline>
                                                           </Font>
                                                       </Title>
                                                       <LabelStyleFormat/>
                                                       <Minimum>0</Minimum>
                                                       <Maximum>50</Maximum>
                                                       <LineWidth>1</LineWidth>
                                                       <Interval>0</Interval>
                                                       <IntervalOffset>0</IntervalOffset>
                                                       <Crossing>NaN</Crossing>
                                                       <MajorGrid>
                                                           <Interval>10</Interval>
                                                           <IntervalOffset>NaN</IntervalOffset>
                                                       </MajorGrid>
                                                   </AxisY2>
                                               </ChartArea>
                                           </ChartAreasCollection>
                                       </ChartSettings>
                                   </XInfo>
                               </XMsgBlk>


        '   <ChartType>Point</ChartType>

        'Send the XMessageBlock to the Line Chart application:
        Message.XAddText("Message sent to [" & ProNetName & "]." & "ADVL_Line_Chart_1" & ":" & vbCrLf, "XmlSentNotice")
        Message.XAddXml(LineChartXMsgBlk.ToString)
        Message.XAddText(vbCrLf, "Normal") 'Add extra line

        SendMessageParams.ProjectNetworkName = ProNetName
        SendMessageParams.ConnectionName = "ADVL_Line_Chart_1"
        SendMessageParams.Message = LineChartXMsgBlk.ToString
        If bgwSendMessage.IsBusy Then
            Message.AddWarning("Send Message backgroundworker is busy." & vbCrLf)
        Else
            bgwSendMessage.RunWorkerAsync(SendMessageParams)
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

    Private Sub UpdateChartLineTab()
        'Update the field selection options on the Chart Line tab.

        If txtLineChartDbPath.Text = "" Then
            Message.AddWarning("Charts: Line: No database has been selected." & vbCrLf)
            Exit Sub
        End If

        If txtLineChartQuery.Text = "" Then
            Message.AddWarning("No query has been specified." & vbCrLf)
        End If

        Dim connString As String
        connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source = " & txtLineChartDbPath.Text

        Dim myConnection As OleDb.OleDbConnection = New OleDb.OleDbConnection
        myConnection.ConnectionString = connString
        myConnection.Open()

        Dim Query As String = txtLineChartQuery.Text & " AND 1 = 2" 'This is used to get all the fields in the query. " AND 1 = 2" ensures no data rows are retrieved.

        Dim da As OleDb.OleDbDataAdapter
        da = New OleDb.OleDbDataAdapter(Query, myConnection)
        da.MissingSchemaAction = MissingSchemaAction.AddWithKey

        Dim ds As DataSet = New DataSet

        Try
            cmbLineXValues.Items.Clear()
            cmbLineYValues.Items.Clear()
            da.Fill(ds, "myData")

            If ds.Tables(0).Columns.Count > 0 Then
                Dim I As Integer 'Loop index
                Dim Name As String
                For I = 1 To ds.Tables(0).Columns.Count
                    Name = ds.Tables(0).Columns(I - 1).ColumnName
                    cmbLineXValues.Items.Add(Name)
                    cmbLineYValues.Items.Add(Name)
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

    Private Sub cmbLineChartDb_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbLineChartDb.SelectedIndexChanged
        'The selected Line Chart database has been changed.
        Select Case cmbLineChartDb.SelectedItem.ToString
            Case "Share Prices"
                txtLineChartDbPath.Text = SharePriceDbPath
            Case "Financials"
                txtLineChartDbPath.Text = FinancialsDbPath
            Case "Calculations"
                txtLineChartDbPath.Text = CalculationsDbPath
        End Select
        FillLineChartDataTableList()
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

    Private Sub FillLineChartDataTableList()
        'Fill the lstLineTables listbox with the available tables in the selected database.

        'Database access for MS Access:
        Dim connectionString As String 'Declare a connection string for MS Access - defines the database or server to be used.
        Dim conn As System.Data.OleDb.OleDbConnection 'Declare a connection for MS Access - used by the Data Adapter to connect to and disconnect from the database.
        Dim dt As DataTable

        lstLineTables.Items.Clear()
        lstLineFields.Items.Clear()

        'Specify the connection string:
        'Access 2003
        'connectionString = "provider=Microsoft.Jet.OLEDB.4.0;" + _
        '"data source = " + txtDatabase.Text

        If txtLineChartDbPath.Text = "" Then
            Exit Sub
        End If

        'Access 2007:
        'connectionString = "provider=Microsoft.ACE.OLEDB.12.0;" +
        '"data source = " + txtDatabase.Text
        connectionString = "provider=Microsoft.ACE.OLEDB.12.0;" +
        "data source = " + txtLineChartDbPath.Text 'DatabasePath

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
                lstLineTables.Items.Add(dt.Rows(I).Item(2).ToString)
            Next I

            conn.Close()
        Catch ex As Exception
            Message.Add("Error opening database: " & txtLineChartDbPath.Text & vbCrLf)
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
        Dim RegExString4 As String = "(?<= <!--)([A-Za-z\d \.,_:]+)(?=-->)"
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

    Private Sub btnChartTitleColor_Click(sender As Object, e As EventArgs) Handles btnChartTitleColor.Click
        ColorDialog1.Color = txtChartTitle.ForeColor
        ColorDialog1.ShowDialog()
        txtChartTitle.ForeColor = ColorDialog1.Color
    End Sub

    Private Sub btnShowFontInfo_Click(sender As Object, e As EventArgs) Handles btnShowFontInfo.Click
        'Show the Font information for txtChartTitle
        'txtChartTitle
        Message.Add("Font information for txtChartTitle:" & vbCrLf)
        Message.Add("  txtChartTitle.Text:" & txtChartTitle.Text & vbCrLf)
        Message.Add("  txtChartTitle.ForeColor.ToArgb.ToString :" & txtChartTitle.ForeColor.ToArgb.ToString & vbCrLf)
        Message.Add("  Color.Black.ToArgb.ToString :" & Color.Black.ToArgb.ToString & vbCrLf)
        Message.Add("  txtChartTitle.Font.Name :" & txtChartTitle.Font.Name & vbCrLf)
        Message.Add("  txtChartTitle.Font.Size :" & txtChartTitle.Font.Size & vbCrLf)
        Message.Add("  txtChartTitle.Font.Bold :" & txtChartTitle.Font.Bold & vbCrLf)
        Message.Add("  txtChartTitle.Font.Italic :" & txtChartTitle.Font.Italic & vbCrLf)
        Message.Add("  txtChartTitle.Font.Strikeout :" & txtChartTitle.Font.Strikeout & vbCrLf)
        Message.Add("  txtChartTitle.Font.Underline :" & txtChartTitle.Font.Underline & vbCrLf)

    End Sub

    'Private Sub btnSaveStockDefaults_Click(sender As Object, e As EventArgs) Handles btnSaveStockChartSettings.Click
    Private Sub btnSaveStockChartSettings_Click(sender As Object, e As EventArgs) Handles btnSaveStockChartSettings.Click
        'Save the Stock Chart Default Settings in a file:

        Try
            'Dim xmlSeq As System.Xml.Linq.XDocument = XDocument.Parse(rtbStockChartDefaults.Text)
            Dim xmlSeq As System.Xml.Linq.XDocument = XDocument.Parse(XmlStockChartSettingsList.Text)

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

    'Private Sub btnOpenStockChartDefaults_Click(sender As Object, e As EventArgs) Handles btnOpenStockChartSettings.Click
    Private Sub btnOpenStockChartSettings_Click(sender As Object, e As EventArgs) Handles btnOpenStockChartSettings.Click
        'Open a Stock Chart Default Settings file:

        Dim SelectedFileName As String = ""

        SelectedFileName = Project.SelectDataFile("Share Price Chart Defaults", "SPChartDefaults")
        Message.Add("Selected Stock Chart Default Settings: " & SelectedFileName & vbCrLf)

        txtStockChartSettings.Text = SelectedFileName

        'Project.ReadXmlData(SelectedFileName, StockChartDefaults)
        Project.ReadXmlData(SelectedFileName, StockChartSettingsList)

        'If StockChartDefaults Is Nothing Then
        If StockChartSettingsList Is Nothing Then
            Exit Sub
        End If

        'rtbStockChartDefaults.Text = StockChartDefaults.ToString
        'FormatXmlText(rtbStockChartDefaults)
        'XmlStockChartDefaults.Rtf = XmlStockChartDefaults.XmlToRtf(StockChartDefaults.ToString, True)
        'XmlStockChartSettingsList.Rtf = XmlStockChartSettingsList.XmlToRtf(StockChartDefaults.ToString, False)
        XmlStockChartSettingsList.Rtf = XmlStockChartSettingsList.XmlToRtf(StockChartSettingsList.ToString, False)

    End Sub

    Private Sub DesignPointChartQuery_Apply(myQuery As String) Handles DesignPointChartQuery.Apply
        txtPointChartQuery.Text = myQuery
        UpdateChartCrossPlotsTab()
    End Sub

    Private Sub DesignLineChartQuery_Apply(myQuery As String) Handles DesignLineChartQuery.Apply
        txtLineChartQuery.Text = myQuery
        UpdateChartLineTab()
    End Sub

    Private Sub btnPointChartTitleFont_Click(sender As Object, e As EventArgs) Handles btnPointChartTitleFont.Click
        'Edit chart title font
        FontDialog1.Font = txtPointChartTitle.Font
        FontDialog1.ShowDialog()
        txtPointChartTitle.Font = FontDialog1.Font

    End Sub

    Private Sub btnDisplayPointChart_Click(sender As Object, e As EventArgs) Handles btnDisplayPointChart.Click

        CheckOpenProjectAtRelativePath("\Point Chart", "ADVL_Point_Chart_1") 'Open the Point Chart project if it is not already connected.

        'Wait up to 8 seconds for the Point Chart project to connect:
        If WaitForConnection(ProNetName, "ADVL_Point_Chart_1", 8000) = False Then
            Message.AddWarning("The Point Chart project did not connect." & vbCrLf)
        End If

        DisplayPointChart()
    End Sub

    Private Sub btnOpenPointChart_Click(sender As Object, e As EventArgs) Handles btnOpenPointChart.Click
        'Open the Point Chart project:
        CheckOpenProjectAtRelativePath("\Point Chart", "ADVL_Point_Chart_1") 'Open the Point Chart project if it is not already connected.
        'Wait up to 8 seconds for the Point Chart project to connect:
        If WaitForConnection(ProNetName, "ADVL_Point_Chart_1", 8000) = False Then
            Message.AddWarning("The Point Chart project did not connect." & vbCrLf)
        End If
    End Sub

    Private Sub DisplayPointChart()
        'Display Cross Plot Chart (Point Chart).

        'Check if connected to ComNet:
        If ConnectedToComNet = False Then
            ConnectToComNet()
        End If

        'Check if a Point Chart Project has been selected:
        If SelPointChartProjNo = -1 Then
            Message.AddWarning("A Point Chart project has not been selected." & vbCrLf)
            Exit Sub
        End If

        'Check if the selected Point Chart Project is in the current Project Network:
        Dim PointChartProNetName As String = ShareChartProj.List(SelPointChartProjNo).ProNetName
        If PointChartProNetName <> ProNetName Then
            Message.AddWarning("The selected Point Chart project is not in the current Project Network." & vbCrLf)
            Exit Sub
        End If

        'Check if the Point Chart Project is connected:
        Dim PointChartProjectName As String = PointChartProj.List(SelPointChartProjNo).Name
        Dim PointChartConnName As String = client.ConnNameFromProjName(PointChartProjectName, ProNetName, "ADVL_Point_Chart_1")
        If PointChartConnName = "" Then
            'Connect the Stock Chart application to the Project Network:
            client.StartProjectWithName(PointChartProjectName, ProNetName, "ADVL_Point_Chart_1", "ADVL_Point_Chart_1")
        End If

        If chkUsePointChartDefaults.Checked Then
            'DisplayPointChartUsingDefaults()
            DisplayPointChartUsingSettingsList()
        Else
            'DisplayPointChartNoDefaults()
            DisplayPointChartNoSettingsList()
        End If
    End Sub

    Private Sub DisplayPointChartUsingSettingsList()
        'Display the Point Chart using a Settings List.

        Message.Add("Displaying the Point Chart using the Settings List" & vbCrLf)

        'Send the instructions to the Chart application to display the stock chart.

        'Check that required selections have been made:
        If cmbPointXValues.SelectedItem Is Nothing Then
            Message.AddWarning("Select a field for the X Values." & vbCrLf)
            Exit Sub
        End If

        'Build the XMessageBlock containing the Stock Chart settings.
        'This will be send to the Stock Chart application to create the chart display.

        Dim ChartSettingsList As XDocument = XDocument.Parse("<?xml version=""1.0"" encoding=""utf-8"" standalone=""yes""?>" & vbCrLf & XmlPointChartSettingsList.Text)

        Dim PointChartXMsgBlk = <?xml version="1.0" encoding="utf-8"?>
                                <XMsgBlk>
                                    <ClientLocn>DisplayChart</ClientLocn>
                                    <XInfo>
                                        <%= ChartSettingsList.<ChartSettings> %>
                                    </XInfo>
                                </XMsgBlk>

        'Update the Settings List with the current chart settings:

        'Update the Input Data settinga:
        '<InputDataType>Database</InputDataType> - Currently only the Database type is available.
        PointChartXMsgBlk.<XMsgBlk>.<XInfo>.<ChartSettings>.<InputDatabasePath>.Value = txtPointChartDbPath.Text
        PointChartXMsgBlk.<XMsgBlk>.<XInfo>.<ChartSettings>.<InputQuery>.Value = txtPointChartQuery.Text
        PointChartXMsgBlk.<XMsgBlk>.<XInfo>.<ChartSettings>.<InputDataDescr>.Value = txtPointSeriesName.Text

        'Add a warning if there is more than one entry in the SeriesInfoList:
        If PointChartXMsgBlk.<XMsgBlk>.<XInfo>.<ChartSettings>.<SeriesInfoList>.<SeriesInfo>.Count > 1 Then AddWarning("There is more than one entry in the Series Info List!" & vbCrLf)
        'Update the first entry in the SeriesInfoList:
        PointChartXMsgBlk.<XMsgBlk>.<XInfo>.<ChartSettings>.<SeriesInfoList>.<SeriesInfo>.<Name>.Value = txtPointSeriesName.Text
        PointChartXMsgBlk.<XMsgBlk>.<XInfo>.<ChartSettings>.<SeriesInfoList>.<SeriesInfo>.<XValuesFieldName>.Value = cmbPointXValues.SelectedItem.ToString
        PointChartXMsgBlk.<XMsgBlk>.<XInfo>.<ChartSettings>.<SeriesInfoList>.<SeriesInfo>.<YValuesFieldName>.Value = cmbPointYValues.SelectedItem.ToString

        'StockChartXMsgBlk.<XMsgBlk>.<XInfo>.<ChartSettings>.<SeriesInfoList>.<SeriesInfo>.<YValuesHighFieldName>.Value = DataGridView1.Rows(0).Cells(1).Value
        'StockChartXMsgBlk.<XMsgBlk>.<XInfo>.<ChartSettings>.<SeriesInfoList>.<SeriesInfo>.<YValuesLowFieldName>.Value = DataGridView1.Rows(1).Cells(1).Value
        'StockChartXMsgBlk.<XMsgBlk>.<XInfo>.<ChartSettings>.<SeriesInfoList>.<SeriesInfo>.<YValuesOpenFieldName>.Value = DataGridView1.Rows(2).Cells(1).Value
        'StockChartXMsgBlk.<XMsgBlk>.<XInfo>.<ChartSettings>.<SeriesInfoList>.<SeriesInfo>.<YValuesCloseFieldName>.Value = DataGridView1.Rows(3).Cells(1).Value

        'Leave the AreaInfoList unchanged.

        'Update the Chart title settings:
        PointChartXMsgBlk.<XMsgBlk>.<XInfo>.<ChartSettings>.<TitlesCollection>.<Title>.<Text>.Value = txtPointChartTitle.Text
        PointChartXMsgBlk.<XMsgBlk>.<XInfo>.<ChartSettings>.<TitlesCollection>.<Title>.<Alignment>.Value = cmbPointChartAlignment.SelectedItem.ToString
        If txtPointChartTitle.ForeColor.ToArgb.ToString = "0" Then 'This color value is not valid for a chart title.
            PointChartXMsgBlk.<XMsgBlk>.<XInfo>.<ChartSettings>.<TitlesCollection>.<Title>.<ForeColor>.Value = Color.Black.ToArgb.ToString
        Else
            PointChartXMsgBlk.<XMsgBlk>.<XInfo>.<ChartSettings>.<TitlesCollection>.<Title>.<ForeColor>.Value = txtPointChartTitle.ForeColor.ToArgb.ToString
        End If
        PointChartXMsgBlk.<XMsgBlk>.<XInfo>.<ChartSettings>.<TitlesCollection>.<Title>.<Font>.<Name>.Value = txtPointChartTitle.Font.Name
        PointChartXMsgBlk.<XMsgBlk>.<XInfo>.<ChartSettings>.<TitlesCollection>.<Title>.<Font>.<Size>.Value = txtPointChartTitle.Font.Size
        PointChartXMsgBlk.<XMsgBlk>.<XInfo>.<ChartSettings>.<TitlesCollection>.<Title>.<Font>.<Bold>.Value = txtPointChartTitle.Font.Bold
        PointChartXMsgBlk.<XMsgBlk>.<XInfo>.<ChartSettings>.<TitlesCollection>.<Title>.<Font>.<Italic>.Value = txtPointChartTitle.Font.Italic
        PointChartXMsgBlk.<XMsgBlk>.<XInfo>.<ChartSettings>.<TitlesCollection>.<Title>.<Font>.<Strikeout>.Value = txtPointChartTitle.Font.Strikeout
        PointChartXMsgBlk.<XMsgBlk>.<XInfo>.<ChartSettings>.<TitlesCollection>.<Title>.<Font>.<Underline>.Value = txtPointChartTitle.Font.Underline

        'Add a warning if there is more than one entry in the SeriesCollection:
        If PointChartXMsgBlk.<XMsgBlk>.<XInfo>.<ChartSettings>.<SeriesCollection>.<Series>.Count > 1 Then AddWarning("There is more than one entry in the Series Collection!" & vbCrLf)
        'Update the first entry in the SeriesCollection:
        PointChartXMsgBlk.<XMsgBlk>.<XInfo>.<ChartSettings>.<SeriesCollection>.<Series>.<Name>.Value = txtPointSeriesName.Text


        'Send the XMessageBlock to the Stock Chart application:
        Message.XAddText("Message sent to [" & ProNetName & "]." & "ADVL_Point_Chart_1" & ":" & vbCrLf, "XmlSentNotice")
        Message.XAddXml(PointChartXMsgBlk.ToString)
        Message.XAddText(vbCrLf, "Normal") 'Add extra line

        SendMessageParams.ProjectNetworkName = ProNetName
        SendMessageParams.ConnectionName = "ADVL_Point_Chart_1"
        SendMessageParams.Message = PointChartXMsgBlk.ToString
        If bgwSendMessage.IsBusy Then
            Message.AddWarning("Send Message backgroundworker is busy." & vbCrLf)
        Else
            bgwSendMessage.RunWorkerAsync(SendMessageParams)
        End If

    End Sub

    Private Sub DisplayPointChartNoSettingsList()
        'Display the Point Chart without using a Settings List.

        'Check that required selections have been made:
        If cmbPointXValues.SelectedItem Is Nothing Then
            Message.AddWarning("Select a field for the X Values." & vbCrLf)
            Exit Sub
        End If

        'Build the XMessageBlock containing the Point Chart settings.
        'This will be send to the Point Chart application to create the chart display.

        Dim PointChartXMsgBlk = <?xml version="1.0" encoding="utf-8"?>
                                <XMsgBlk>
                                    <ClientLocn>DisplayChart</ClientLocn>
                                    <XInfo>
                                        <ChartSettings>
                                            <!--Input Data:-->
                                            <InputDataType>Database</InputDataType>
                                            <InputDatabasePath><%= txtPointChartDbPath.Text %></InputDatabasePath>
                                            <InputQuery><%= txtPointChartQuery.Text %></InputQuery>
                                            <InputDataDescr><%= txtPointSeriesName.Text %></InputDataDescr>
                                            <SeriesInfoList>
                                                <SeriesInfo>
                                                    <Name><%= txtPointSeriesName.Text %></Name>
                                                    <XValuesFieldName><%= cmbPointXValues.SelectedItem.ToString %></XValuesFieldName>
                                                    <YValuesFieldName><%= cmbPointYValues.SelectedItem.ToString %></YValuesFieldName>
                                                    <ChartArea>ChartArea1</ChartArea>
                                                </SeriesInfo>
                                            </SeriesInfoList>
                                            <AreaInfoList>
                                                <AreaInfo>
                                                    <Name>ChartArea1</Name>
                                                    <AutoXAxisMinimum>true</AutoXAxisMinimum>
                                                    <AutoXAxisMaximum>true</AutoXAxisMaximum>
                                                    <AutoXAxisMajorGridInterval>true</AutoXAxisMajorGridInterval>
                                                    <AutoX2AxisMinimum>true</AutoX2AxisMinimum>
                                                    <AutoX2AxisMaximum>true</AutoX2AxisMaximum>
                                                    <AutoX2AxisMajorGridInterval>true</AutoX2AxisMajorGridInterval>
                                                    <AutoYAxisMinimum>true</AutoYAxisMinimum>
                                                    <AutoYAxisMaximum>true</AutoYAxisMaximum>
                                                    <AutoYAxisMajorGridInterval>true</AutoYAxisMajorGridInterval>
                                                    <AutoY2AxisMinimum>true</AutoY2AxisMinimum>
                                                    <AutoY2AxisMaximum>true</AutoY2AxisMaximum>
                                                    <AutoY2AxisMajorGridInterval>true</AutoY2AxisMajorGridInterval>
                                                </AreaInfo>
                                            </AreaInfoList>
                                            <!--Chart Properties:-->
                                            <TitlesCollection>
                                                <Title>
                                                    <Name>Title1</Name>
                                                    <Text><%= txtPointChartTitle.Text %></Text>
                                                    <TextOrientation>Auto</TextOrientation>
                                                    <Alignment><%= cmbPointChartAlignment.SelectedItem.ToString %></Alignment>
                                                    <ForeColor>-16777216</ForeColor>
                                                    <Font>
                                                        <Name>Microsoft Sans Serif</Name>
                                                        <Size>14.25</Size>
                                                        <Bold>true</Bold>
                                                        <Italic>false</Italic>
                                                        <Strikeout>false</Strikeout>
                                                        <Underline>false</Underline>
                                                    </Font>
                                                </Title>
                                            </TitlesCollection>
                                            <SeriesCollection>
                                                <Series>
                                                    <Name><%= txtPointSeriesName.Text %></Name>
                                                    <ChartType>Point</ChartType>
                                                    <ChartArea>ChartArea1</ChartArea>
                                                    <Legend>Legend1</Legend>
                                                    <EmptyPointValue>Average</EmptyPointValue>
                                                    <LabelStyle>Auto</LabelStyle>
                                                    <PixelPointDepth>1</PixelPointDepth>
                                                    <PixelPointGapDepth>1</PixelPointGapDepth>
                                                    <ShowMarkerLines/>
                                                    <AxisLabel/>
                                                    <XAxisType>Primary</XAxisType>
                                                    <XValueType>Single</XValueType>
                                                    <YAxisType>Primary</YAxisType>
                                                    <YValueType>Single</YValueType>
                                                    <Marker>
                                                        <BorderColor>-16777216</BorderColor>
                                                        <BorderWidth>1</BorderWidth>
                                                        <Color>-8355712</Color>
                                                        <Size>5</Size>
                                                        <Step>1</Step>
                                                        <Style>None</Style>
                                                    </Marker>
                                                    <Color>-16776961</Color>
                                                </Series>
                                            </SeriesCollection>
                                            <ChartAreasCollection>
                                                <ChartArea>
                                                    <Name>ChartArea1</Name>
                                                    <AxisX>
                                                        <Title>
                                                            <Text><%= cmbPointXValues.SelectedItem.ToString %></Text>
                                                            <Alignment>Center</Alignment>
                                                            <ForeColor>-16777216</ForeColor>
                                                            <Font>
                                                                <Name>Microsoft Sans Serif</Name>
                                                                <Size>12</Size>
                                                                <Bold>true</Bold>
                                                                <Italic>false</Italic>
                                                                <Strikeout>false</Strikeout>
                                                                <Underline>false</Underline>
                                                            </Font>
                                                        </Title>
                                                        <LabelStyleFormat/>
                                                        <Minimum>40910</Minimum>
                                                        <Maximum>43743</Maximum>
                                                        <LineWidth>1</LineWidth>
                                                        <Interval>0</Interval>
                                                        <IntervalOffset>0</IntervalOffset>
                                                        <Crossing>NaN</Crossing>
                                                        <MajorGrid>
                                                            <Interval>6</Interval>
                                                            <IntervalOffset>NaN</IntervalOffset>
                                                        </MajorGrid>
                                                    </AxisX>
                                                    <AxisX2>
                                                        <Title>
                                                            <Text/>
                                                            <Alignment>Center</Alignment>
                                                            <ForeColor>-16777216</ForeColor>
                                                            <Font>
                                                                <Name>Microsoft Sans Serif</Name>
                                                                <Size>8</Size>
                                                                <Bold>false</Bold>
                                                                <Italic>false</Italic>
                                                                <Strikeout>false</Strikeout>
                                                                <Underline>false</Underline>
                                                            </Font>
                                                        </Title>
                                                        <LabelStyleFormat/>
                                                        <Minimum>40910</Minimum>
                                                        <Maximum>43743</Maximum>
                                                        <LineWidth>1</LineWidth>
                                                        <Interval>0</Interval>
                                                        <IntervalOffset>0</IntervalOffset>
                                                        <Crossing>NaN</Crossing>
                                                        <MajorGrid>
                                                            <Interval>2</Interval>
                                                            <IntervalOffset>NaN</IntervalOffset>
                                                        </MajorGrid>
                                                    </AxisX2>
                                                    <AxisY>
                                                        <Title>
                                                            <Text><%= cmbPointYValues.SelectedItem.ToString %></Text>
                                                            <Alignment>Center</Alignment>
                                                            <ForeColor>-16777216</ForeColor>
                                                            <Font>
                                                                <Name>Microsoft Sans Serif</Name>
                                                                <Size>12</Size>
                                                                <Bold>true</Bold>
                                                                <Italic>false</Italic>
                                                                <Strikeout>false</Strikeout>
                                                                <Underline>false</Underline>
                                                            </Font>
                                                        </Title>
                                                        <LabelStyleFormat/>
                                                        <Minimum>0</Minimum>
                                                        <Maximum>50</Maximum>
                                                        <LineWidth>1</LineWidth>
                                                        <Interval>0</Interval>
                                                        <IntervalOffset>0</IntervalOffset>
                                                        <Crossing>NaN</Crossing>
                                                        <MajorGrid>
                                                            <Interval>10</Interval>
                                                            <IntervalOffset>NaN</IntervalOffset>
                                                        </MajorGrid>
                                                    </AxisY>
                                                    <AxisY2>
                                                        <Title>
                                                            <Text/>
                                                            <Alignment>Center</Alignment>
                                                            <ForeColor>-16777216</ForeColor>
                                                            <Font>
                                                                <Name>Microsoft Sans Serif</Name>
                                                                <Size>8</Size>
                                                                <Bold>false</Bold>
                                                                <Italic>false</Italic>
                                                                <Strikeout>false</Strikeout>
                                                                <Underline>false</Underline>
                                                            </Font>
                                                        </Title>
                                                        <LabelStyleFormat/>
                                                        <Minimum>0</Minimum>
                                                        <Maximum>50</Maximum>
                                                        <LineWidth>1</LineWidth>
                                                        <Interval>0</Interval>
                                                        <IntervalOffset>0</IntervalOffset>
                                                        <Crossing>NaN</Crossing>
                                                        <MajorGrid>
                                                            <Interval>10</Interval>
                                                            <IntervalOffset>NaN</IntervalOffset>
                                                        </MajorGrid>
                                                    </AxisY2>
                                                </ChartArea>
                                            </ChartAreasCollection>
                                        </ChartSettings>
                                    </XInfo>
                                </XMsgBlk>

        'Send the XMessageBlock to the Point Chart application:
        Message.XAddText("Message sent to [" & ProNetName & "]." & "ADVL_Point_Chart_1" & ":" & vbCrLf, "XmlSentNotice")
        Message.XAddXml(PointChartXMsgBlk.ToString)
        Message.XAddText(vbCrLf, "Normal") 'Add extra line

        SendMessageParams.ProjectNetworkName = ProNetName
        SendMessageParams.ConnectionName = "ADVL_Point_Chart_1"
        SendMessageParams.Message = PointChartXMsgBlk.ToString
        If bgwSendMessage.IsBusy Then
            Message.AddWarning("Send Message backgroundworker is busy." & vbCrLf)
        Else
            bgwSendMessage.RunWorkerAsync(SendMessageParams)
        End If

    End Sub

    'Remove the following method:
    Private Sub DisplayPointChartUsingDefaults()
        'Display Point Chart.
        'Use the default parameters in PointChartDefaults
        'Send the instructions to the Chart application to display the point chart (crossplot).

        'If PointChartDefaults Is Nothing Then
        If PointChartSettingsList Is Nothing Then
            Message.AddWarning("No Point Chart default settings loaded." & vbCrLf)
            DisplayPointChartNoDefaults()
            Exit Sub
        End If

        'Create the xml instructions
        Dim decl As New XDeclaration("1.0", "utf-8", "yes")
        Dim doc As New XDocument(decl, Nothing) 'Create an XDocument to store the instructions.
        Dim xmessage As New XElement("XMsg") 'This indicates the start of the message in the XMessage class

        'Dim clientAppNetName As New XElement("ClientAppNetName", AppNetName)
        'xmessage.Add(clientAppNetName)
        Dim clientProNetName As New XElement("ClientProNetName", ProNetName)
        xmessage.Add(clientProNetName)

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

        'If PointChartDefaults.<PointChart>.<Settings>.<ChartTitle>.<FontName>.Value <> Nothing Then
        If PointChartSettingsList.<PointChart>.<Settings>.<ChartTitle>.<FontName>.Value <> Nothing Then
            'Dim chartTitleFontName As New XElement("FontName", PointChartDefaults.<PointChart>.<Settings>.<ChartTitle>.<FontName>.Value)
            Dim chartTitleFontName As New XElement("FontName", PointChartSettingsList.<PointChart>.<Settings>.<ChartTitle>.<FontName>.Value)
            chartTitle.Add(chartTitleFontName)
        Else
            Message.AddWarning("Default Chart Title Font Name settings not found." & vbCrLf)
            Dim chartTitleFontName As New XElement("FontName", txtPointChartTitle.Font.Name)
            chartTitle.Add(chartTitleFontName)
        End If

        If PointChartSettingsList.<PointChart>.<Settings>.<ChartTitle>.<Color>.Value <> Nothing Then
            Dim chartTitleColor As New XElement("Color", PointChartSettingsList.<PointChart>.<Settings>.<ChartTitle>.<Color>.Value)
            chartTitle.Add(chartTitleColor)
        Else
            Message.AddWarning("Default Chart Title Color settings not found." & vbCrLf)
            Dim chartTitleColor As New XElement("Color", txtChartTitle.ForeColor)
            chartTitle.Add(chartTitleColor)
        End If

        If PointChartSettingsList.<PointChart>.<Settings>.<ChartTitle>.<Size>.Value <> Nothing Then
            Dim chartTitleSize As New XElement("Size", PointChartSettingsList.<PointChart>.<Settings>.<ChartTitle>.<Size>.Value)
            chartTitle.Add(chartTitleSize)
        Else
            Message.AddWarning("Default Chart Title Size settings not found." & vbCrLf)
            Dim chartTitleSize As New XElement("Size", txtChartTitle.Font.Size)
            chartTitle.Add(chartTitleSize)
        End If

        If PointChartSettingsList.<PointChart>.<Settings>.<ChartTitle>.<Bold>.Value <> Nothing Then
            Dim chartTitleBold As New XElement("Bold", PointChartSettingsList.<PointChart>.<Settings>.<ChartTitle>.<Bold>.Value)
            chartTitle.Add(chartTitleBold)
        Else
            Message.AddWarning("Default Chart Title Bold settings not found." & vbCrLf)
            Dim chartTitleBold As New XElement("Bold", txtChartTitle.Font.Bold)
            chartTitle.Add(chartTitleBold)
        End If

        If PointChartSettingsList.<PointChart>.<Settings>.<ChartTitle>.<Italic>.Value <> Nothing Then
            Dim chartTitleItalic As New XElement("Italic", PointChartSettingsList.<PointChart>.<Settings>.<ChartTitle>.<Italic>.Value)
            chartTitle.Add(chartTitleItalic)
        Else
            Message.AddWarning("Default Chart Title Italic settings not found." & vbCrLf)
            Dim chartTitleItalic As New XElement("Italic", txtChartTitle.Font.Italic)
            chartTitle.Add(chartTitleItalic)
        End If

        If PointChartSettingsList.<PointChart>.<Settings>.<ChartTitle>.<Underline>.Value <> Nothing Then
            Dim chartTitleUnderline As New XElement("Underline", PointChartSettingsList.<PointChart>.<Settings>.<ChartTitle>.<Underline>.Value)
            chartTitle.Add(chartTitleUnderline)
        Else
            Message.AddWarning("Default Chart Title Underline settings not found." & vbCrLf)
            Dim chartTitleUnderline As New XElement("Underline", txtChartTitle.Font.Underline)
            chartTitle.Add(chartTitleUnderline)
        End If

        If PointChartSettingsList.<PointChart>.<Settings>.<ChartTitle>.<Strikeout>.Value <> Nothing Then
            Dim chartTitleStrikeout As New XElement("Strikeout", PointChartSettingsList.<PointChart>.<Settings>.<ChartTitle>.<Strikeout>.Value)
            chartTitle.Add(chartTitleStrikeout)
        Else
            Message.AddWarning("Default Chart Title Strikeout settings not found." & vbCrLf)
            Dim chartTitleStrikeout As New XElement("Strikeout", txtChartTitle.Font.Strikeout)
            chartTitle.Add(chartTitleStrikeout)
        End If

        If PointChartSettingsList.<PointChart>.<Settings>.<ChartTitle>.<Alignment>.Value <> Nothing Then
            Dim chartTitleAlignment As New XElement("Alignment", PointChartSettingsList.<PointChart>.<Settings>.<ChartTitle>.<Alignment>.Value)
            chartTitle.Add(chartTitleAlignment)
        Else
            Message.AddWarning("Default Chart Title Alignment settings not found." & vbCrLf)
            Dim chartTitleAlignment As New XElement("Alignment", cmbAlignment.SelectedItem.ToString)
            chartTitle.Add(chartTitleAlignment)
        End If

        chartSettings.Add(chartTitle)

        Dim xAxis As New XElement("XAxis")

        If PointChartSettingsList.<PointChart>.<Settings>.<XAxis>.<TitleText>.Value <> Nothing Then
            Dim xAxisTitleText As New XElement("TitleText", PointChartSettingsList.<PointChart>.<Settings>.<XAxis>.<TitleText>.Value)
            xAxis.Add(xAxisTitleText)
        End If

        If PointChartSettingsList.<PointChart>.<Settings>.<XAxis>.<TitleFontName>.Value <> Nothing Then
            Dim xAxisTitleFontName As New XElement("TitleFontName", PointChartSettingsList.<PointChart>.<Settings>.<XAxis>.<TitleFontName>.Value)
            xAxis.Add(xAxisTitleFontName)
        End If

        If PointChartSettingsList.<PointChart>.<Settings>.<XAxis>.<TitleColor>.Value <> Nothing Then
            Dim xAxisTitleColor As New XElement("TitleColor", PointChartSettingsList.<PointChart>.<Settings>.<XAxis>.<TitleColor>.Value)
            xAxis.Add(xAxisTitleColor)
        End If

        If PointChartSettingsList.<PointChart>.<Settings>.<XAxis>.<TitleSize>.Value <> Nothing Then
            Dim xAxisTitleSize As New XElement("TitleSize", PointChartSettingsList.<PointChart>.<Settings>.<XAxis>.<TitleSize>.Value)
            xAxis.Add(xAxisTitleSize)
        End If

        If PointChartSettingsList.<PointChart>.<Settings>.<XAxis>.<TitleBold>.Value <> Nothing Then
            Dim xAxisTitleBold As New XElement("TitleBold", PointChartSettingsList.<PointChart>.<Settings>.<XAxis>.<TitleBold>.Value)
            xAxis.Add(xAxisTitleBold)
        End If

        If PointChartSettingsList.<PointChart>.<Settings>.<XAxis>.<TitleItalic>.Value <> Nothing Then
            Dim xAxisTitleItalic As New XElement("TitleItalic", PointChartSettingsList.<PointChart>.<Settings>.<XAxis>.<TitleItalic>.Value)
            xAxis.Add(xAxisTitleItalic)
        End If

        If PointChartSettingsList.<PointChart>.<Settings>.<XAxis>.<TitleUnderline>.Value <> Nothing Then
            Dim xAxisTitleUnderline As New XElement("TitleUnderline", PointChartSettingsList.<PointChart>.<Settings>.<XAxis>.<TitleUnderline>.Value)
            xAxis.Add(xAxisTitleUnderline)
        End If

        If PointChartSettingsList.<PointChart>.<Settings>.<XAxis>.<TitleStrikeout>.Value <> Nothing Then
            Dim xAxisTitleStrikeout As New XElement("TitleStrikeout", PointChartSettingsList.<PointChart>.<Settings>.<XAxis>.<TitleStrikeout>.Value)
            xAxis.Add(xAxisTitleStrikeout)
        End If

        If PointChartSettingsList.<PointChart>.<Settings>.<XAxis>.<TitleAlignment>.Value <> Nothing Then
            Dim xAxisTitleAlignment As New XElement("TitleAlignment", PointChartSettingsList.<PointChart>.<Settings>.<XAxis>.<TitleAlignment>.Value)
            xAxis.Add(xAxisTitleAlignment)
        End If

        If PointChartSettingsList.<PointChart>.<Settings>.<XAxis>.<AutoMinimum>.Value <> Nothing Then
            Dim xAxisAutoMinimum As New XElement("AutoMinimum", PointChartSettingsList.<PointChart>.<Settings>.<XAxis>.<AutoMinimum>.Value)
            xAxis.Add(xAxisAutoMinimum)
        End If

        If PointChartSettingsList.<PointChart>.<Settings>.<XAxis>.<Minimum>.Value <> Nothing Then
            Dim xAxisMinimum As New XElement("Minimum", PointChartSettingsList.<PointChart>.<Settings>.<XAxis>.<Minimum>.Value)
            xAxis.Add(xAxisMinimum)
        End If

        If PointChartSettingsList.<PointChart>.<Settings>.<XAxis>.<AutoMaximum>.Value <> Nothing Then
            Dim xAxisAutoMaximum As New XElement("AutoMaximum", PointChartSettingsList.<PointChart>.<Settings>.<XAxis>.<AutoMaximum>.Value)
            xAxis.Add(xAxisAutoMaximum)
        End If

        If PointChartSettingsList.<PointChart>.<Settings>.<XAxis>.<Maximum>.Value <> Nothing Then
            Dim xAxisMaximum As New XElement("Maximum", PointChartSettingsList.<PointChart>.<Settings>.<XAxis>.<Maximum>.Value)
            xAxis.Add(xAxisMaximum)
        End If

        If PointChartSettingsList.<PointChart>.<Settings>.<XAxis>.<AutoInterval>.Value <> Nothing Then
            Dim xAxisAutoInterval As New XElement("AutoInterval", PointChartSettingsList.<PointChart>.<Settings>.<XAxis>.<AutoInterval>.Value)
            xAxis.Add(xAxisAutoInterval)
        End If

        If PointChartSettingsList.<PointChart>.<Settings>.<XAxis>.<MajorGridInterval>.Value <> Nothing Then
            Dim xAxisMajorGridInterval As New XElement("MajorGridInterval", PointChartSettingsList.<PointChart>.<Settings>.<XAxis>.<MajorGridInterval>.Value)
            xAxis.Add(xAxisMajorGridInterval)
        End If

        If PointChartSettingsList.<PointChart>.<Settings>.<XAxis>.<AutoMajorGridInterval>.Value <> Nothing Then
            Dim xAxisAutoMajorGridInterval As New XElement("AutoMajorGridInterval", PointChartSettingsList.<PointChart>.<Settings>.<XAxis>.<AutoMajorGridInterval>.Value)
            xAxis.Add(xAxisAutoMajorGridInterval)
        End If

        chartSettings.Add(xAxis)


        Dim yAxis As New XElement("YAxis")

        If PointChartSettingsList.<PointChart>.<Settings>.<YAxis>.<TitleText>.Value <> Nothing Then
            Dim yAxisTitleText As New XElement("TitleText", PointChartSettingsList.<PointChart>.<Settings>.<YAxis>.<TitleText>.Value)
            yAxis.Add(yAxisTitleText)
        End If

        If PointChartSettingsList.<PointChart>.<Settings>.<YAxis>.<TitleFontName>.Value <> Nothing Then
            Dim yAxisTitleFontName As New XElement("TitleFontName", PointChartSettingsList.<PointChart>.<Settings>.<YAxis>.<TitleFontName>.Value)
            yAxis.Add(yAxisTitleFontName)
        End If

        If PointChartSettingsList.<PointChart>.<Settings>.<YAxis>.<TitleColor>.Value <> Nothing Then
            Dim yAxisTitleColor As New XElement("TitleColor", PointChartSettingsList.<PointChart>.<Settings>.<YAxis>.<TitleColor>.Value)
            yAxis.Add(yAxisTitleColor)
        End If

        If PointChartSettingsList.<PointChart>.<Settings>.<YAxis>.<TitleSize>.Value <> Nothing Then
            Dim yAxisTitleSize As New XElement("TitleSize", PointChartSettingsList.<PointChart>.<Settings>.<YAxis>.<TitleSize>.Value)
            yAxis.Add(yAxisTitleSize)
        End If

        If PointChartSettingsList.<PointChart>.<Settings>.<YAxis>.<TitleBold>.Value <> Nothing Then
            Dim yAxisTitleBold As New XElement("TitleBold", PointChartSettingsList.<PointChart>.<Settings>.<YAxis>.<TitleBold>.Value)
            yAxis.Add(yAxisTitleBold)
        End If

        If PointChartSettingsList.<PointChart>.<Settings>.<YAxis>.<TitleItalic>.Value <> Nothing Then
            Dim yAxisTitleItalic As New XElement("TitleItalic", PointChartSettingsList.<PointChart>.<Settings>.<YAxis>.<TitleItalic>.Value)
            yAxis.Add(yAxisTitleItalic)
        End If

        If PointChartSettingsList.<PointChart>.<Settings>.<YAxis>.<TitleUnderline>.Value <> Nothing Then
            Dim yAxisTitleUnderline As New XElement("TitleUnderline", PointChartSettingsList.<PointChart>.<Settings>.<YAxis>.<TitleUnderline>.Value)
            yAxis.Add(yAxisTitleUnderline)
        End If

        If PointChartSettingsList.<PointChart>.<Settings>.<YAxis>.<TitleStrikeout>.Value <> Nothing Then
            Dim yAxisTitleStrikeout As New XElement("TitleStrikeout", PointChartSettingsList.<PointChart>.<Settings>.<YAxis>.<TitleStrikeout>.Value)
            yAxis.Add(yAxisTitleStrikeout)
        End If

        If PointChartSettingsList.<PointChart>.<Settings>.<YAxis>.<TitleAlignment>.Value <> Nothing Then
            Dim yAxisTitleAlignment As New XElement("TitleAlignment", PointChartSettingsList.<PointChart>.<Settings>.<YAxis>.<TitleAlignment>.Value)
            yAxis.Add(yAxisTitleAlignment)
        End If

        If PointChartSettingsList.<PointChart>.<Settings>.<YAxis>.<AutoMinimum>.Value <> Nothing Then
            Dim yAxisAutoMinimum As New XElement("AutoMinimum", PointChartSettingsList.<PointChart>.<Settings>.<YAxis>.<AutoMinimum>.Value)
            yAxis.Add(yAxisAutoMinimum)
        End If

        If PointChartSettingsList.<PointChart>.<Settings>.<YAxis>.<Minimum>.Value <> Nothing Then
            Dim yAxisMinimum As New XElement("Minimum", PointChartSettingsList.<PointChart>.<Settings>.<YAxis>.<Minimum>.Value)
            yAxis.Add(yAxisMinimum)
        End If

        If PointChartSettingsList.<PointChart>.<Settings>.<YAxis>.<AutoMaximum>.Value <> Nothing Then
            Dim yAxisAutoMaximum As New XElement("AutoMaximum", PointChartSettingsList.<PointChart>.<Settings>.<YAxis>.<AutoMaximum>.Value)
            yAxis.Add(yAxisAutoMaximum)
        End If

        If PointChartSettingsList.<PointChart>.<Settings>.<YAxis>.<Maximum>.Value <> Nothing Then
            Dim yAxisMaximum As New XElement("Maximum", PointChartSettingsList.<PointChart>.<Settings>.<YAxis>.<Maximum>.Value)
            yAxis.Add(yAxisMaximum)
        End If

        If PointChartSettingsList.<PointChart>.<Settings>.<YAxis>.<AutoInterval>.Value <> Nothing Then
            Dim yAxisAutoInterval As New XElement("AutoInterval", PointChartSettingsList.<PointChart>.<Settings>.<YAxis>.<AutoInterval>.Value)
            yAxis.Add(yAxisAutoInterval)
        End If

        If PointChartSettingsList.<PointChart>.<Settings>.<YAxis>.<MajorGridInterval>.Value <> Nothing Then
            Dim yAxisMajorGridInterval As New XElement("MajorGridInterval", PointChartSettingsList.<PointChart>.<Settings>.<YAxis>.<MajorGridInterval>.Value)
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
            'client.SendMessageAsync(AppNetName, "ADVL_Chart_1", doc.ToString)
            'client.SendMessageAsync(ProNetName, "ADVL_Chart_1", doc.ToString)
            client.SendMessageAsync(ProNetName, "ADVL_Point_Chart_1", doc.ToString)
            'Message.XAddText("Message sent to " & "ADVL_Chart_1" & ":" & vbCrLf, "XmlSentNotice")
            Message.XAddText("Message sent to [" & ProNetName & "]." & "ADVL_Point_Chart_1" & ":" & vbCrLf, "XmlSentNotice")
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
            'Dim clientAppNetName As New XElement("ClientAppNetName", AppNetName)
            'xmessage.Add(clientAppNetName)
            Dim clientProNetName As New XElement("ClientProNetName", ProNetName)
            xmessage.Add(clientProNetName)

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
            'client.SendMessageAsync(AppNetName, "ADVL_Chart_1", doc.ToString)
            'client.SendMessageAsync(ProNetName, "ADVL_Chart_1", doc.ToString)
            client.SendMessageAsync(ProNetName, "ADVL_Point_Chart_1", doc.ToString)
            'Message.XAddText("Message sent to " & "ADVL_Chart_1" & ":" & vbCrLf, "XmlSentNotice")
            Message.XAddText("Message sent to [" & ProNetName & "]." & "ADVL_Point_Chart_1" & ":" & vbCrLf, "XmlSentNotice")
            Message.XAddXml(doc.ToString)
            Message.XAddText(vbCrLf, "Message") 'Add extra line
        End If
    End Sub

    'Private Sub btnGetPointChartDefaults_Click(sender As Object, e As EventArgs) Handles btnGetPointChartSettings.Click
    Private Sub btnGetPointChartSettings_Click(sender As Object, e As EventArgs) Handles btnGetPointChartSettings.Click
        'Send a request to ADVL_Charts_1 for the current Point Chart settings.

        'Create the xml instructions
        Dim decl As New XDeclaration("1.0", "utf-8", "yes")
        Dim doc As New XDocument(decl, Nothing) 'Create an XDocument to store the instructions.
        Dim xmessage As New XElement("XMsg") 'This indicates the start of the message in the XMessage class

        'Dim clientAppNetName As New XElement("ClientAppNetName", AppNetName)
        'xmessage.Add(clientAppNetName)
        Dim clientProNetName As New XElement("ClientProNetName", ProNetName)
        xmessage.Add(clientProNetName)

        Dim clientName As New XElement("ClientName", ApplicationInfo.Name) 'This tells the coordinate server the name of the client making the request.
        xmessage.Add(clientName)

        Dim clientConnName As New XElement("ClientConnectionName", ApplicationInfo.Name) 'This tells the coordinate server the name of the client making the request.
        xmessage.Add(clientConnName)


        'Dim clientLocn As New XElement("ClientLocn", "PointChart")
        'xmessage.Add(clientLocn)

        'Dim commandGetSettings As New XElement("Command", "GetPointChartSettings")
        'xmessage.Add(commandGetSettings)

        Dim getChartDefaultsCommand As New XElement("GetChartSettings", "PointChart")
        xmessage.Add(getChartDefaultsCommand)

        doc.Add(xmessage)

        If IsNothing(client) Then
            Message.AddWarning("No client connection available!" & vbCrLf)
            Beep()
        Else
            'client.SendMessageAsync(AppNetName, "ADVL_Chart_1", doc.ToString)
            'client.SendMessageAsync(ProNetName, "ADVL_Chart_1", doc.ToString)
            client.SendMessageAsync(ProNetName, "ADVL_Point_Chart_1", doc.ToString)
            'Message.XAddText("Message sent to " & "ADVL_Chart_1" & ":" & vbCrLf, "XmlSentNotice")
            'Message.XAddText("Message sent to [" & ProNetName & "]." & "ADVL_Chart_1" & ":" & vbCrLf, "XmlSentNotice")
            Message.XAddText("Message sent to [" & ProNetName & "]." & "ADVL_Point_Chart_1" & ":" & vbCrLf, "XmlSentNotice")
            Message.XAddXml(doc.ToString)
            Message.XAddText(vbCrLf, "Message") 'Add extra line
        End If
    End Sub

    Private Sub btnGetLineChartDefaults_Click(sender As Object, e As EventArgs) Handles btnGetLineChartDefaults.Click
        'Send a request to ADVL_Line_Chart_1 for the current Line Chart settings.

        'Create the xml instructions
        Dim decl As New XDeclaration("1.0", "utf-8", "yes")
        Dim doc As New XDocument(decl, Nothing) 'Create an XDocument to store the instructions.
        Dim xmessage As New XElement("XMsg") 'This indicates the start of the message in the XMessage class

        Dim clientProNetName As New XElement("ClientProNetName", ProNetName)
        xmessage.Add(clientProNetName)

        Dim clientName As New XElement("ClientName", ApplicationInfo.Name) 'This tells the Chart server the name of the client making the request.
        xmessage.Add(clientName)

        Dim clientConnName As New XElement("ClientConnectionName", ApplicationInfo.Name) 'This tells the Chart server the name of the client making the request.
        xmessage.Add(clientConnName)

        Dim getChartDefaultsCommand As New XElement("GetChartSettings", "LineChart")
        xmessage.Add(getChartDefaultsCommand)

        doc.Add(xmessage)

        If IsNothing(client) Then
            Message.AddWarning("No client connection available!" & vbCrLf)
            Beep()
        Else
            client.SendMessageAsync(ProNetName, "ADVL_Line_Chart_1", doc.ToString)
            Message.XAddText("Message sent to [" & ProNetName & "]." & "ADVL_Line_Chart_1" & ":" & vbCrLf, "XmlSentNotice")
            Message.XAddXml(doc.ToString)
            Message.XAddText(vbCrLf, "Message") 'Add extra line
        End If
    End Sub

    'Private Sub btnSavePointDefaults_Click(sender As Object, e As EventArgs) Handles btnSavePointSettings.Click
    Private Sub btnSavePointChartSettings_Click(sender As Object, e As EventArgs) Handles btnSavePointChartSettings.Click
        'Save the Point Chart Default Settings in a file:

        Try
            'Dim xmlSeq As System.Xml.Linq.XDocument = XDocument.Parse(rtbPointChartDefaults.Text)
            Dim xmlSeq As System.Xml.Linq.XDocument = XDocument.Parse(XmlPointChartSettingsList.Text)


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

    'Private Sub btnOpenPointChartDefaults_Click(sender As Object, e As EventArgs) Handles btnOpenPointChartSettings.Click
    Private Sub btnOpenPointChartSettings_Click(sender As Object, e As EventArgs) Handles btnOpenPointChartSettings.Click
        'Open a Point Chart Default Settings file:

        Dim SelectedFileName As String = ""

        SelectedFileName = Project.SelectDataFile("Cross Plot Chart Defaults", "PointChartDefaults")
        Message.Add("Selected Cross Plot Chart Default Settings: " & SelectedFileName & vbCrLf)

        txtPointChartSettings.Text = SelectedFileName

        'Project.ReadXmlData(SelectedFileName, PointChartDefaults)
        Project.ReadXmlData(SelectedFileName, PointChartSettingsList)

        'If PointChartDefaults Is Nothing Then
        If PointChartSettingsList Is Nothing Then
            Exit Sub
        End If

        'rtbPointChartDefaults.Text = PointChartDefaults.ToString
        'rtbPointChartDefaults.Text = PointChartSettingsList.ToString
        'FormatXmlText(rtbPointChartDefaults)
        XmlPointChartSettingsList.Rtf = XmlPointChartSettingsList.XmlToRtf(PointChartSettingsList.ToString, False)
    End Sub

    Private Sub btnDisplayLineChart_Click(sender As Object, e As EventArgs) Handles btnDisplayLineChart.Click
        CheckOpenProjectAtRelativePath("\Line Chart", "ADVL_Line_Chart_1") 'Open the Line Chart project if it is not already connected.

        'Wait up to 8 seconds for the Line Chart project to connect:
        If WaitForConnection(ProNetName, "ADVL_Line_Chart_1", 8000) = False Then
            Message.AddWarning("The Line Chart project did not connect." & vbCrLf)
        End If

        DisplayLineChart()

    End Sub


    Private Sub btnOpenLineChart_Click(sender As Object, e As EventArgs) Handles btnOpenLineChart.Click
        'Open the Line Chart project:
        CheckOpenProjectAtRelativePath("\Line Chart", "ADVL_Line_Chart_1") 'Open the Line Chart project if it is not already connected.
        'Wait up to 8 seconds for the Line Chart project to connect:
        If WaitForConnection(ProNetName, "ADVL_Line_Chart_1", 8000) = False Then
            Message.AddWarning("The Line Chart project did not connect." & vbCrLf)
        End If
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
        If System.IO.File.Exists(DatabasePath) Then
            'Database file exists.
        Else
            'FDatabase file does not exist!
            Message.AddWarning("The database was not found: " & DatabasePath & vbCrLf)
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
        If System.IO.File.Exists(DatabasePath) Then
            'Database file exists.
        Else
            'FDatabase file does not exist!
            Message.AddWarning("The database was not found: " & DatabasePath & vbCrLf)
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
        If System.IO.File.Exists(DatabasePath) Then
            'Database file exists.
        Else
            'FDatabase file does not exist!
            Message.AddWarning("The database was not found: " & DatabasePath & vbCrLf)
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
        If System.IO.File.Exists(DatabasePath) Then
            'Database file exists.
        Else
            'FDatabase file does not exist!
            Message.AddWarning("The database was not found: " & DatabasePath & vbCrLf)
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
        If System.IO.File.Exists(DatabasePath) Then
            'Database file exists.
        Else
            'FDatabase file does not exist!
            Message.AddWarning("The database was not found: " & DatabasePath & vbCrLf)
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

    'Function used by JavaScript code section to check if a table is present in a specified database:
    Function TableExists(ByVal DatabasePath As String, ByVal TableName As String) As Boolean
        'Return True if TableName is found in Database at DatabasePath.

        If DatabasePath = "" Then
            Message.AddWarning("No database specified!" & vbCrLf)
            Return False
            Exit Function
        End If

        If TableName = "" Then
            Message.AddWarning("No table spacified!" & vbCrLf)
            Return False
            Exit Function
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


        Dim restrictions As String() = New String() {Nothing, Nothing, Nothing, "TABLE"} 'This restriction removes system tables
        Dim dt As DataTable
        dt = conn.GetSchema("Tables", restrictions)

        'Search the tables in the database for TableName:
        'Dim dr As DataRow
        Dim I As Integer 'Loop index
        Dim MaxI As Integer
        Dim TableFound As Boolean = False

        MaxI = dt.Rows.Count
        For I = 0 To MaxI - 1
            'dr = dt.Rows(0)
            'cmbSelectTable.Items.Add(dt.Rows(I).Item(2).ToString)
            If dt.Rows(I).Item(2).ToString = TableName Then
                TableFound = True
                Exit For
            End If
        Next I

        Return TableFound

        conn.Close()


    End Function


#End Region 'Database Tables Sub Tab ----------------------------------------------------------------------------------------------------------------------------------------------------------


#End Region 'Utilities Tab --------------------------------------------------------------------------------------------------------------------------------------------------------------------

#Region " Run XSequence Code" '================================================================================================================================================================

    Private Sub XSeq_ErrorMsg(ErrMsg As String) Handles XSeq.ErrorMsg
        Message.Add(ErrMsg & vbCrLf)
    End Sub

    Private Sub XSeq_Instruction(Data As String, Locn As String) Handles XSeq.Instruction
        'Execute each instruction that is produced by running the XSequence file.

        If IsDBNull(Data) Then
            Data = ""
        End If

        ''Intercept and instructions with the prefix "WebPage_"
        'If Locn.StartsWith("WebPage_") Then 'Send the Data, Location data to the correct Web Page:
        '    If Locn.Contains(":") Then
        '        Dim EndOfWebPageNoString As Integer = Locn.IndexOf(":")
        '        Dim PageNoLen As Integer = EndOfWebPageNoString - 8
        '        Dim WebPageNoString As String = Locn.Substring(8, PageNoLen)
        '        Dim WebPageNo As Integer = CInt(WebPageNoString)
        '        Dim WebPageData As String = Data
        '        Dim WebPageLocn As String = Locn.Substring(EndOfWebPageNoString + 1)

        '        WebPageFormList(WebPageNo).XMsgInstruction(WebPageData, WebPageLocn)
        '    Else
        '        Message.AddWarning("XMessage instruction location is not complete: " & Locn & vbCrLf)
        '    End If
        'Else

        'Intercept instructions with the prefix "WebPage_"
        If Locn.StartsWith("WebPage_") Then 'Send the Data, Location data to the correct Web Page:
            'Message.Add("Web Page Location: " & Locn & vbCrLf)
            If Locn.Contains(":") Then
                Dim EndOfWebPageNoString As Integer = Locn.IndexOf(":")
                If Locn.Contains("-") Then
                    Dim HyphenLocn As Integer = Locn.IndexOf("-")
                    If HyphenLocn < EndOfWebPageNoString Then 'Web Page Location contains a sub-location in the web page - WebPage_1-SubLocn:Locn - SubLocn:Locn will be sent to Web page 1
                        EndOfWebPageNoString = HyphenLocn
                    End If
                End If
                Dim PageNoLen As Integer = EndOfWebPageNoString - 8
                Dim WebPageNoString As String = Locn.Substring(8, PageNoLen)
                Dim WebPageNo As Integer = CInt(WebPageNoString)
                Dim WebPageData As String = Data
                Dim WebPageLocn As String = Locn.Substring(EndOfWebPageNoString + 1)

                'Message.Add("WebPageData = " & WebPageData & "  WebPageLocn = " & WebPageLocn & vbCrLf)

                WebPageFormList(WebPageNo).XMsgInstruction(WebPageData, WebPageLocn)
            Else
                'Message.AddWarning("XSequence: " & vbCrLf)
                'Message.AddWarning("XMessage instruction location is not complete: " & Locn & vbCrLf)
                Message.AddWarning("XSequence instruction location is not complete: " & Locn & vbCrLf)
            End If
        Else


            Select Case Locn

             'Restore Web Page Settings: -------------------------------------------------
                Case "Settings:Form:Name"
                    FormName = Data

                Case "Settings:Form:Item:Name"
                    ItemName = Data

                Case "Settings:Form:Item:Value"
                    RestoreSetting(FormName, ItemName, Data)

                Case "Settings:Form:SelectId"
                    SelectId = Data

                Case "Settings:Form:OptionText"
                    RestoreOption(SelectId, Data)
            'END Restore Web Page Settings: ---------------------------------------------


            'Parameter Code: -----------------------------------------------------------------------------------------------------------------------------
                Case "Parameter:Name"
                    XSeq.NewParameter.Name = Data
                Case "Parameter:Description"
                    XSeq.NewParameter.Description = Data
                Case "Parameter:Value"
                    XSeq.NewParameter.Value = Data
                Case "Parameter:Command"
                    Select Case Data
                        Case "Add"
                            XSeq.AddParameter()
                        Case Else
                            Message.AddWarning("Unknown Parameter:Command Information Value: " & Data & vbCrLf)
                    End Select
            'Copy Data Code: -----------------------------------------------------------------------------------------------------------------------------
                Case "CopyData:InputDatabase"
                    cmbCopyDataInputDb.SelectedIndex = cmbCopyDataInputDb.FindStringExact(Data)
                    'Set up dgvCopyData:
                    dgvCopyData.AllowUserToAddRows = False 'This removes the last edit row that allows the user to add new rows.
                    dgvCopyData.Rows.Clear()
                Case "CopyData:InputDatabasePath"
                'This information is not used.
                'Consider writing a warning message if the current database path is different from the one in this Processing Sequence.
                Case "CopyData:InputQuery"
                    txtCopyDataInputQuery.Text = Data
                Case "CopyData:InputQuery:ReadParameter"
                    If XSeq.Parameter.ContainsKey(Data) Then
                        txtCopyDataInputQuery.Text = XSeq.Parameter(Data).Value
                    Else
                        Message.AddWarning("Instruction error - CopyData:InputQuery:ReadParameter - The following parameter was not found: " & Data & vbCrLf)
                    End If
                Case "CopyData:OutputDatabase"
                    cmbCopyDataOutputDb.SelectedIndex = cmbCopyDataOutputDb.FindStringExact(Data)
                Case "CopyData:OutputDatabasePath"
                'This information is not used.
                'Consider writing a warning message if the current database path is different from the one in this Processing Sequence.
                Case "CopyData:OutputQuery"
                    txtCopyDataOutputQuery.Text = Data
                Case "CopyData:OutputQuery:ReadParameter"
                    'If XSeq.Parameter(Data).
                    If XSeq.Parameter.ContainsKey(Data) Then
                        txtCopyDataOutputQuery.Text = XSeq.Parameter(Data).Value
                    Else
                        Message.AddWarning("Instruction error - CopyData:OutputQuery:ReadParameter - The following parameter was not found: " & Data & vbCrLf)
                    End If
                Case "CopyData:CopyList:CopyColumn:From"
                    dgvCopyData.Rows.Add() 'Add a new blank row.
                    dgvCopyData.Rows(dgvCopyData.Rows.Count - 1).Cells(0).Value = Data 'Add the From column name to the last row.
                Case "CopyData:CopyList:CopyColumn:To"
                    dgvCopyData.Rows(dgvCopyData.Rows.Count - 1).Cells(1).Value = Data 'Add the To column name to the last row.
                Case "CopyData:Command"
                    Select Case Data
                        Case "Apply"
                            ApplyCopyData()
                        Case Else
                            Message.AddWarning("Unknown CopyData:Command Information Value: " & Data & vbCrLf)
                    End Select


            'Select Data Code: ---------------------------------------------------------------------------------------------------------------------------
                Case "SelectData:InputDatabase"
                    cmbSelectDataInputDb.SelectedIndex = cmbSelectDataInputDb.FindStringExact(Data)
                    'Set up dgvCopyData:
                    dgvSelectData.AllowUserToAddRows = False 'This removes the last edit row that allows the user to add new rows.
                    dgvSelectData.Rows.Clear()
                    dgvSelectConstraints.AllowUserToAddRows = False 'This removes the last edit row that allows the user to add new rows.
                    dgvSelectConstraints.Rows.Clear()
                Case "SelectData:InputDatabasePath"
                'This information is not used.
                'Consider writing a warning message if the current database path is different from the one in this Processing Sequence.
                Case "SelectData:InputQuery"
                    txtSelectDataInputQuery.Text = Data
                Case "SelectData:InputQuery:ReadParameter"
                    If XSeq.Parameter.ContainsKey(Data) Then
                        txtSelectDataInputQuery.Text = XSeq.Parameter(Data).Value
                    Else
                        Message.AddWarning("Instruction error - SelectData:InputQuery:ReadParameter - The following parameter was not found: " & Data & vbCrLf)
                    End If
                Case "SelectData:OutputDatabase"
                    cmbSelectDataOutputDb.SelectedIndex = cmbSelectDataOutputDb.FindStringExact(Data)
                Case "SelectData:OutputDatabasePath"
                'This information is not used.
                'Consider writing a warning message if the current database path is different from the one in this Processing Sequence.
                Case "SelectData:OutputQuery"
                    txtSelectDataOutputQuery.Text = Data
                Case "SelectData:OutputQuery:ReadParameter"
                    If XSeq.Parameter.ContainsKey(Data) Then
                        txtSelectDataOutputQuery.Text = XSeq.Parameter(Data).Value
                    Else
                        Message.AddWarning("Instruction error - SelectData:OutputQuery:ReadParameter - The following parameter was not found: " & Data & vbCrLf)
                    End If
                Case "SelectData:SelectConstraintList:Constraint:WhereInputColumn"
                    dgvSelectConstraints.Rows.Add() 'Add a new blank row.
                    dgvSelectConstraints.Rows(dgvSelectConstraints.Rows.Count - 1).Cells(0).Value = Data 'Add the WhereInputColumn name to the last row.
                Case "SelectData:SelectConstraintList:Constraint:EqualsOutputColumn"
                    dgvSelectConstraints.Rows(dgvSelectConstraints.Rows.Count - 1).Cells(1).Value = Data 'Add the EqualsOutputColumn name to the last row.
                Case "SelectData:SelectDataList:CopyColumn:From"
                    dgvSelectData.Rows.Add() 'Add a new blank row.
                    dgvSelectData.Rows(dgvSelectData.Rows.Count - 1).Cells(0).Value = Data 'Add the From column name to the last row.
                Case "SelectData:SelectDataList:CopyColumn:To"
                    dgvSelectData.Rows(dgvSelectData.Rows.Count - 1).Cells(1).Value = Data 'Add the To column name to the last row.
                Case "SelectData:Command"
                    Select Case Data
                        Case "Apply"
                            ApplySelectData()
                        Case Else
                            Message.AddWarning("Unknown SelectData:Command Information Value: " & Data & vbCrLf)
                    End Select


            'Simple Calculations Code: -------------------------------------------------------------------------------------------------------------------
                Case "SimpleCalculations:SelectedDatabase"
                    cmbSimpleCalcDb.SelectedIndex = cmbSimpleCalcDb.FindStringExact(Data)
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
                    txtSimpleCalcsQuery.Text = Data
                Case "SimpleCalculations:DataQuery:ReadParameter"
                    If XSeq.Parameter.ContainsKey(Data) Then
                        txtSimpleCalcsQuery.Text = XSeq.Parameter(Data).Value
                    Else
                        Message.AddWarning("Instruction error - SimpleCalculations:DataQuery:ReadParameter - The following parameter was not found: " & Data & vbCrLf)
                    End If
                Case "SimpleCalculations:ParameterList:Parameter:Name"
                    dgvSimpleCalcsParameterList.Rows.Add() 'Add a new blank row.
                    dgvSimpleCalcsParameterList.Rows(dgvSimpleCalcsParameterList.Rows.Count - 1).Cells(0).Value = Data 'Add the Parameter Name to the last row.
                Case "SimpleCalculations:ParameterList:Parameter:Type"
                    dgvSimpleCalcsParameterList.Rows(dgvSimpleCalcsParameterList.Rows.Count - 1).Cells(1).Value = Data 'Add the Parameter Type to the last row.
                Case "SimpleCalculations:ParameterList:Parameter:Value"
                    dgvSimpleCalcsParameterList.Rows(dgvSimpleCalcsParameterList.Rows.Count - 1).Cells(2).Value = Data 'Add the Parameter Value to the last row.
                Case "SimpleCalculations:ParameterList:Parameter:Description"
                    dgvSimpleCalcsParameterList.Rows(dgvSimpleCalcsParameterList.Rows.Count - 1).Cells(3).Value = Data 'Add the Parameter Description to the last row.
                Case "SimpleCalculations:InputDataList:InputData:Parameter"
                    dgvSimpleCalcsInputData.Rows.Add() 'Add a new blank row.
                    dgvSimpleCalcsInputData.Rows(dgvSimpleCalcsInputData.Rows.Count - 1).Cells(0).Value = Data 'Add the Input Parameter Name to the last row.
                Case "SimpleCalculations:InputDataList:InputData:Column"
                    dgvSimpleCalcsInputData.Rows(dgvSimpleCalcsInputData.Rows.Count - 1).Cells(1).Value = Data 'Add the Input Column Name to the last row.
                Case "SimpleCalculations:CalculationList:Calculation:Input1"
                    dgvSimpleCalcsCalculations.Rows.Add() 'Add a new blank row.
                    dgvSimpleCalcsCalculations.Rows(dgvSimpleCalcsCalculations.Rows.Count - 1).Cells(0).Value = Data 'Add the Input1 Parameter Name to the last row.
                Case "SimpleCalculations:CalculationList:Calculation:Input2"
                    dgvSimpleCalcsCalculations.Rows(dgvSimpleCalcsCalculations.Rows.Count - 1).Cells(1).Value = Data 'Add the Input2 Parameter Name to the last row.
                Case "SimpleCalculations:CalculationList:Calculation:Operation"
                    dgvSimpleCalcsCalculations.Rows(dgvSimpleCalcsCalculations.Rows.Count - 1).Cells(2).Value = Data 'Add the Operation to the last row.
                Case "SimpleCalculations:CalculationList:Calculation:Output"
                    dgvSimpleCalcsCalculations.Rows(dgvSimpleCalcsCalculations.Rows.Count - 1).Cells(3).Value = Data 'Add the Output Parameter Name to the last row.
                Case "SimpleCalculations:OutputDataList:OutputData:Parameter"
                    dgvSimpleCalcsOutputData.Rows.Add() 'Add a new blank row.
                    dgvSimpleCalcsOutputData.Rows(dgvSimpleCalcsOutputData.Rows.Count - 1).Cells(0).Value = Data 'Add the Output Parameter Name to the last row.
                Case "SimpleCalculations:OutputDataList:OutputData:Column"
                    dgvSimpleCalcsOutputData.Rows(dgvSimpleCalcsOutputData.Rows.Count - 1).Cells(1).Value = Data 'Add the Output Column Name to the last row.
                Case "SimpleCalculations:Command"
                    Select Case Data
                        Case "Apply"
                            ApplySimpleCalcs()
                        Case Else
                            Message.AddWarning("Unknown SimpleCalculations:Command Information Value: " & Data & vbCrLf)
                    End Select


            'Date Calculations Code: ---------------------------------------------------------------------------------------------------------------------
                Case "DateCalculations:SelectedDatabase"
                    cmbDateCalcDb.SelectedIndex = cmbDateCalcDb.FindStringExact(Data)
                Case "DateCalculations:SelectedDatabasePath"
                  'This information is not used.
                'Consider writing a warning message if the current database path is different from the one in this Processing Sequence.
                Case "DateCalculations:DataQuery"
                    txtDateCalcsQuery.Text = Data
                Case "DateCalculations:DataQuery:ReadParameter"
                    If XSeq.Parameter.ContainsKey(Data) Then
                        txtDateCalcsQuery.Text = XSeq.Parameter(Data).Value
                    Else
                        Message.AddWarning("Instruction error - DataCalculations:DataQuery:ReadParameter - The following parameter was not found: " & Data & vbCrLf)
                    End If
                Case "DateCalculations:CalculationType"
                    cmbDateCalcType.SelectedIndex = cmbDateCalcType.FindStringExact(Data)
                Case "DateCalculations:FixedDate"
                    txtFixedDate.Text = Data
                Case "DateCalculations:FixedDate:ReadParameter"
                    If XSeq.Parameter.ContainsKey(Data) Then
                        txtFixedDate.Text = XSeq.Parameter(Data).Value
                    Else
                        Message.AddWarning("Instruction error - DateCalculations:FixedDate:ReadParameter - The following parameter was not found: " & Data & vbCrLf)
                    End If
                Case "DateCalculations:DateFormatString"
                    txtDateFormatString.Text = Data
                Case "DateCalculations:MonthColumn"
                    cmbDateCalcParam1.Text = Data
                Case "DateCalculations:YearColumn"
                    cmbDateCalcParam2.Text = Data
                Case "DateCalculations:StartDateColumn"
                    cmbDateCalcParam1.Text = Data
                Case "DateCalculations:NDaysColumn"
                    cmbDateCalcParam2.Text = Data
                Case "DateCalculations:OutputDateColumn"
                    cmbDateCalcOutput.Text = Data
                Case "DateCalculations:Command"
                    Select Case Data
                        Case "Apply"
                            ApplyDateCalcs()
                        Case Else
                            Message.AddWarning("Unknown CopyData:Command Information Value: " & Data & vbCrLf)
                    End Select


            'Date Select Code: ---------------------------------------------------------------------------------------------------------------------------
                Case "DateSelect:InputDatabase"
                    cmbDateSelectInputDb.SelectedIndex = cmbDateSelectInputDb.FindStringExact(Data)
                    'Set up DataGridViews:
                    dgvDateSelectData.AllowUserToAddRows = False 'This removes the last edit row that allows the user to add new rows.
                    dgvDateSelectData.Rows.Clear()
                    dgvDateSelectConstraints.AllowUserToAddRows = False 'This removes the last edit row that allows the user to add new rows.
                    dgvDateSelectConstraints.Rows.Clear()
                Case "DateSelect:InputDatabasePath"
                'This information is not used.
                'Consider writing a warning message if the current database path is different from the one in this Processing Sequence.
                Case "DateSelect:InputQuery"
                    txtDateSelInputQuery.Text = Data
                Case "DateSelect:InputQuery:ReadParameter"
                    If XSeq.Parameter.ContainsKey(Data) Then
                        txtDateSelInputQuery.Text = XSeq.Parameter(Data).Value
                    Else
                        Message.AddWarning("Instruction error - DateSelect:InputQuery:ReadParameter - The following parameter was not found: " & Data & vbCrLf)
                    End If
                Case "DateSelect:OutputDatabase"
                    cmbDateSelectOutputDb.SelectedIndex = cmbDateSelectOutputDb.FindStringExact(Data)
                Case "DateSelect:OutputDatabasePath"
                'This information is not used.
                'Consider writing a warning message if the current database path is different from the one in this Processing Sequence.
                Case "DateSelect:OutputQuery"
                    txtDateSelOutputQuery.Text = Data
                Case "DateSelect:OutputQuery:ReadParameter"
                    If XSeq.Parameter.ContainsKey(Data) Then
                        txtDateSelOutputQuery.Text = XSeq.Parameter(Data).Value
                    Else
                        Message.AddWarning("Instruction error - DateSelect:OutputQuery:ReadParameter - The following parameter was not found: " & Data & vbCrLf)
                    End If
                Case "DateSelect:DateSelectionType"
                    cmbDateSelectionType.Text = Data
                Case "DateSelect:InputDateColumn"
                    cmbDateSelInputDateCol.Text = Data
                Case "DateSelect:OutputDateColumn"
                    cmbDateSelOutputDateCol.Text = Data
                Case "DateSelect:SelectConstraintList:Constraint:WhereInputColumn"
                    dgvDateSelectConstraints.Rows.Add() 'Add a new blank row.
                    dgvDateSelectConstraints.Rows(dgvDateSelectConstraints.Rows.Count - 1).Cells(0).Value = Data 'Add the WhereInputColumn name to the last row.
                Case "DateSelect:SelectConstraintList:Constraint:EqualsOutputColumn"
                    dgvDateSelectConstraints.Rows(dgvDateSelectConstraints.Rows.Count - 1).Cells(1).Value = Data 'Add the EqualsOutputColumn name to the last row.
                Case "DateSelect:SelectDataList:CopyColumn:From"
                    dgvDateSelectData.Rows.Add() 'Add a new blank row.
                    dgvDateSelectData.Rows(dgvDateSelectData.Rows.Count - 1).Cells(0).Value = Data 'Add the From column name to the last row.
                Case "DateSelect:SelectDataList:CopyColumn:To"
                    dgvDateSelectData.Rows(dgvDateSelectData.Rows.Count - 1).Cells(1).Value = Data 'Add the To column name to the last row.
                Case "DateSelect:Command"
                    Select Case Data
                        Case "Apply"
                            ApplyDateSelections()
                        Case Else
                            Message.AddWarning("Unknown DateSelect:Command Information Value: " & Data & vbCrLf)
                    End Select

            'Utilities - Database Tables Code: -----------------------------------------------------------------------------------------------------------
                Case "DeleteTable:Database"
                    cmbUtilTablesDatabase.SelectedIndex = cmbUtilTablesDatabase.FindStringExact(Data)
                Case "DeleteTable:DatabasePath"
                 'Not used. Database path determined from Database selection.
                Case "DeleteTable:TableToDelete"
                    cmbSelectTable.SelectedIndex = cmbSelectTable.FindStringExact(Data)
                Case "DeleteTable:TableToDelete:ReadParameter"
                    If XSeq.Parameter.ContainsKey(Data) Then
                        cmbSelectTable.SelectedIndex = cmbSelectTable.FindStringExact(XSeq.Parameter(Data).Value)
                    Else
                        Message.AddWarning("Instruction error - DeleteTable:TableToDelete:ReadParameter - The following parameter was not found: " & Data & vbCrLf)
                    End If
                Case "DeleteTable:Command"
                    Select Case Data
                        Case "Apply"
                            DeleteTable()
                        Case Else
                            Message.AddWarning("Unknown DeleteTable:Command Information Value: " & Data & vbCrLf)
                    End Select

                Case "DeleteRecords:Database"
                    cmbUtilTablesDatabase.SelectedIndex = cmbUtilTablesDatabase.FindStringExact(Data)
                Case "DeleteRecords:DatabasePath"
                 'Not used. Database path determined from Database selection.
                Case "DeleteRecords:Table"
                    cmbSelectTable.SelectedIndex = cmbSelectTable.FindStringExact(Data)
                Case "DeleteRecords:Table:ReadParameter"
                    If XSeq.Parameter.ContainsKey(Data) Then
                        cmbSelectTable.SelectedIndex = cmbSelectTable.FindStringExact(XSeq.Parameter(Data).Value)
                    Else
                        Message.AddWarning("Instruction error - DeleteRecords:Table:ReadParameter - The following parameter was not found: " & Data & vbCrLf)
                    End If

                Case "DeleteRecords:Command"
                    Select Case Data
                        Case "Apply"
                            DeleteRecords()
                        Case Else
                            Message.AddWarning("Unknown DeleteRecords:Command Information Value: " & Data & vbCrLf)
                    End Select

                Case "CopyTableColumns:Database"
                    cmbUtilTablesDatabase.SelectedIndex = cmbUtilTablesDatabase.FindStringExact(Data)
                Case "CopyTableColumns:DatabasePath"
                'Not used. Database path determined from Database selection.
                Case "CopyTableColumns:Table"
                    cmbSelectTable.SelectedIndex = cmbSelectTable.FindStringExact(Data)
                Case "CopyTableColumns:Table:ReadParameter"
                    If XSeq.Parameter.ContainsKey(Data) Then
                        cmbSelectTable.SelectedIndex = cmbSelectTable.FindStringExact(XSeq.Parameter(Data).Value)
                    Else
                        Message.AddWarning("Instruction error - CopyTableColumns:Table:ReadParameter - The following parameter was not found: " & Data & vbCrLf)
                    End If

                Case "CopyTableColumns:NewTable"
                    txtNewTableName.Text = Data
                Case "CopyTableColumns:NewTable:ReadParameter"
                    If XSeq.Parameter.ContainsKey(Data) Then
                        txtNewTableName.Text = XSeq.Parameter(Data).Value
                    Else
                        Message.AddWarning("Instruction error - CopyTableColumns:Table:ReadParameter - The following parameter was not found: " & Data & vbCrLf)
                    End If

                Case "CopyTableColumns:Command"
                    Select Case Data
                        Case "Apply"
                            CopyTableColumns()
                        Case Else
                            Message.AddWarning("Unknown CopyTableColumns:Command Information Value: " & Data & vbCrLf)
                    End Select

                Case "CopyTable:Database"
                    cmbUtilTablesDatabase.SelectedIndex = cmbUtilTablesDatabase.FindStringExact(Data)
                Case "CopyTable:DatabasePath"
                 'Not used. Database path determined from Database selection.
                Case "CopyTable:Table"
                    cmbSelectTable.SelectedIndex = cmbSelectTable.FindStringExact(Data)
                Case "CopyTable:Table:ReadParameter"
                    If XSeq.Parameter.ContainsKey(Data) Then
                        cmbSelectTable.SelectedIndex = cmbSelectTable.FindStringExact(XSeq.Parameter(Data).Value)
                    Else
                        Message.AddWarning("Instruction error - CopyTable:Table:ReadParameter - The following parameter was not found: " & Data & vbCrLf)
                    End If

                Case "CopyTable:NewTable"
                    txtNewTableName.Text = Data
                Case "CopyTable:NewTable:ReadParameter"
                    If XSeq.Parameter.ContainsKey(Data) Then
                        txtNewTableName.Text = XSeq.Parameter(Data).Value
                    Else
                        Message.AddWarning("Instruction error - CopyTable:NewTable:ReadParameter - The following parameter was not found: " & Data & vbCrLf)
                    End If

                Case "CopyTable:Command"
                    Select Case Data
                        Case "Apply"
                            CopyTable()
                        Case Else
                            Message.AddWarning("Unknown CopyTable:Command Information Value: " & Data & vbCrLf)
                    End Select


            'End of Sequence Code: -----------------------------------------------------------------------------------------------------------------------
                Case "EndOfSequence"
                    XSeq.Parameter.Clear() 'Clear the Parameter dictionary.
                    Message.Add("Processing sequence has completed." & vbCrLf)

                Case Else
                    Message.AddWarning("Unknown Information Location: " & Locn & vbCrLf)
            End Select
        End If
    End Sub


    Private Sub XMsgLocal_Instruction(Data As String, Locn As String) Handles XMsgLocal.Instruction
        'Process an XMessage instruction locally.

        If IsDBNull(Data) Then
            Data = ""
        End If

        ''Intercept and instructions with the prefix "WebPage_"
        'If Locn.StartsWith("WebPage_") Then 'Send the Data, Location data to the correct Web Page:
        '    If Locn.Contains(":") Then
        '        Dim EndOfWebPageNoString As Integer = Locn.IndexOf(":")
        '        Dim PageNoLen As Integer = EndOfWebPageNoString - 8
        '        Dim WebPageNoString As String = Locn.Substring(8, PageNoLen)
        '        Dim WebPageNo As Integer = CInt(WebPageNoString)
        '        Dim WebPageData As String = Data
        '        Dim WebPageLocn As String = Locn.Substring(EndOfWebPageNoString + 1)

        '        WebPageFormList(WebPageNo).XMsgInstruction(WebPageData, WebPageLocn)
        '    Else
        '        Message.AddWarning("XMessage instruction location is not complete: " & Locn & vbCrLf)
        '    End If
        'Else


        'Intercept instructions with the prefix "WebPage_"
        If Locn.StartsWith("WebPage_") Then 'Send the Data, Location data to the correct Web Page:
            'Message.Add("Web Page Location: " & Locn & vbCrLf)
            If Locn.Contains(":") Then
                Dim EndOfWebPageNoString As Integer = Locn.IndexOf(":")
                If Locn.Contains("-") Then
                    Dim HyphenLocn As Integer = Locn.IndexOf("-")
                    If HyphenLocn < EndOfWebPageNoString Then 'Web Page Location contains a sub-location in the web page - WebPage_1-SubLocn:Locn - SubLocn:Locn will be sent to Web page 1
                        EndOfWebPageNoString = HyphenLocn
                    End If
                End If
                Dim PageNoLen As Integer = EndOfWebPageNoString - 8
                Dim WebPageNoString As String = Locn.Substring(8, PageNoLen)
                Dim WebPageNo As Integer = CInt(WebPageNoString)
                Dim WebPageData As String = Data
                Dim WebPageLocn As String = Locn.Substring(EndOfWebPageNoString + 1)

                'Message.Add("WebPageData = " & WebPageData & "  WebPageLocn = " & WebPageLocn & vbCrLf)

                WebPageFormList(WebPageNo).XMsgInstruction(WebPageData, WebPageLocn)
            Else
                Message.AddWarning("XMessage instruction location is not complete: " & Locn & vbCrLf)
            End If
        Else

            Select Case Locn
                Case "ClientName"
                    ClientAppName = Data 'The name of the Client requesting service.

                       'UPDATE:
                Case "OnCompletion"
                    OnCompletionInstruction = Data

                Case "Main"
                 'Blank message - do nothing.

                ' 'LEGACY CODE:
                'Case "Main:OnCompletion"
                '    Select Case "Stop"
                '        'Stop on completion of the instruction sequence.
                '    End Select

                Case "Main:EndInstruction"
                    Select Case Data
                        Case "Stop"
                            'Stop at the end of the instruction sequence.

                            'Add other cases here:
                    End Select

                Case "Main:Status"
                    Select Case Data
                        Case "OK"
                            'Main instructions completed OK
                    End Select

                Case "EndOfSequence"
                    'End of Information Sequence reached.

                Case Else
                    Message.AddWarning("Local XMessage: " & Locn & vbCrLf)
                    Message.AddWarning("Unknown location: " & Locn & vbCrLf)
                    Message.AddWarning("            info: " & Data & vbCrLf & vbCrLf)
            End Select
        End If
    End Sub

#End Region 'Run XSequence Code ---------------------------------------------------------------------------------------------------------------------------------------------------------------

    Private Sub ApplicationInfo_UpdateExePath() Handles ApplicationInfo.UpdateExePath
        'Update the Executable Path.
        ApplicationInfo.ExecutablePath = Application.ExecutablePath
    End Sub

    Private Sub Zip_FileSelected(FileName As String) Handles Zip.FileSelected

    End Sub


    Private Sub chkConnect_LostFocus(sender As Object, e As EventArgs) Handles chkConnect.LostFocus
        If chkConnect.Checked Then
            Project.ConnectOnOpen = True
        Else
            Project.ConnectOnOpen = False
        End If
        Project.SaveProjectInfoFile()

    End Sub

    'Private Sub Timer3_Tick(sender As Object, e As EventArgs) Handles Timer3.Tick
    '    'Keet the connection awake with each tick:

    '    If ConnectedToComNet = True Then
    '        Try
    '            If client.IsAlive() Then
    '                Message.Add(Format(Now, "HH:mm:ss") & " Connection OK." & vbCrLf)
    '                'Timer3.Interval = TimeSpan.FromMinutes(55).TotalMilliseconds '55 minute interval
    '                Timer3.Interval = TimeSpan.FromHours(1).TotalMilliseconds '1 hour interval
    '            Else
    '                Message.Add(Format(Now, "HH:mm:ss") & " Connection Fault." & vbCrLf)
    '                'Timer3.Interval = TimeSpan.FromMinutes(55).TotalMilliseconds '55 minute interval
    '                Timer3.Interval = TimeSpan.FromHours(1).TotalMilliseconds '1 hour interval
    '            End If
    '        Catch ex As Exception
    '            Message.AddWarning(ex.Message & vbCrLf)
    '            'Set interval to five minutes - try again in five minutes:
    '            Timer3.Interval = TimeSpan.FromMinutes(5).TotalMilliseconds '5 minute interval
    '        End Try
    '    Else
    '        Message.Add(Format(Now, "HH:mm:ss") & " Not connected." & vbCrLf)
    '    End If

    'End Sub

    'Public Sub CalculateDailyStatistics(ByVal TableName As String, ByVal DateValue As Date)
    Public Sub CalculateDailyStatistics(ByVal TableName As String, ByVal DateString As String)
        'Calculate the Daily share trading Statistics for the date in DateValue.
        'The results are stored in the Calculation database in the table named TableName.

        Dim DateValue As Date = CDate(DateString)

        Dim DbDateString As String = "#" & Format(DateValue, "MM-dd-yyyy") & "#"

        If SharePriceDbPath = "" Then
            Message.AddWarning("A share price database has not been selected!" & vbCrLf)
            Exit Sub
        End If
        If System.IO.File.Exists(SharePriceDbPath) Then
            'Share Price Database file exists.
        Else
            'Share Price Database file does not exist!
            Message.AddWarning("The share price database was not found: " & SharePriceDbPath & vbCrLf)
            Exit Sub
        End If

        If CalculationsDbPath = "" Then
            Message.AddWarning("A Calculations database has not been selected!" & vbCrLf)
            Exit Sub
        End If
        If System.IO.File.Exists(CalculationsDbPath) Then
            'Calculations Database file exists.
        Else
            'Calculations Database file does not exist!
            Message.AddWarning("The calculations database was not found: " & CalculationsDbPath & vbCrLf)
            Exit Sub
        End If

        Dim connString As String
        Dim myConnection As OleDb.OleDbConnection = New OleDb.OleDbConnection

        Dim Query As String = "Select * From ASX_Share_Prices Where Trade_Date = " & DbDateString

        Dim ds As DataSet = New DataSet
        Dim da As OleDb.OleDbDataAdapter

        connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source =" & SharePriceDbPath 'DatabasePath
        myConnection.ConnectionString = connString
        myConnection.Open()

        da = New OleDb.OleDbDataAdapter(Query, myConnection)
        da.MissingSchemaAction = MissingSchemaAction.AddWithKey

        ds.Clear()

        Try
            da.Fill(ds, "myData")
            Dim NRows As Integer = ds.Tables(0).Rows.Count
            Dim I As Integer

            Dim DayType As String = ""
            Dim NCompaniesTraded As Integer = 0
            Dim OpenValueTraded As Double = 0
            Dim HighValueTraded As Double = 0
            Dim LowValueTraded As Double = 0
            Dim CloseValueTraded As Double = 0
            Dim VolumeTraded As Integer = 0

            For I = 0 To NRows - 1
                'If ds.Tables(0).Rows(I).Field("Volume") = 0 Then
                If ds.Tables(0).Rows(I).Item("Volume") = 0 Then
                    'No shares in this company have been traded.
                Else
                    'Shares in this company have been traded.
                    NCompaniesTraded += 1
                    VolumeTraded = ds.Tables(0).Rows(I).Item("Volume")
                    OpenValueTraded = OpenValueTraded + VolumeTraded * ds.Tables(0).Rows(I).Item("Open_Price")
                    HighValueTraded = HighValueTraded + VolumeTraded * ds.Tables(0).Rows(I).Item("High_Price")
                    LowValueTraded = LowValueTraded + VolumeTraded * ds.Tables(0).Rows(I).Item("Low_Price")
                    CloseValueTraded = CloseValueTraded + VolumeTraded * ds.Tables(0).Rows(I).Item("Close_Price")
                End If
            Next

            myConnection.Close()
            ds.Clear()

            'Write the data to the Daily Statistics table:
            connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source =" & CalculationsDbPath 'DatabasePath
            myConnection.ConnectionString = connString
            myConnection.Open()

            'Dim CommandString As String = "Insert Into [Daily Statistics] (Trade_Date, Day_Type, Companies_Traded, Open_Value_Traded, High_Value_Traded, Low_Value_Traded, Close_Value_Traded) Values (" & DateString & ", Trade_Day, " & NCompaniesTraded.ToString & ", " & OpenValueTraded.ToString & ", " & HighValueTraded.ToString & ", " & LowValueTraded.ToString & ", " & CloseValueTraded.ToString & ")"
            'Dim CommandString As String = "INSERT INTO [Daily Statistics] (Trade_Date, Day_Type, Companies_Traded, Open_Value_Traded, High_Value_Traded, Low_Value_Traded, Close_Value_Traded) VALUES (" & DateString & ", Trade_Day, " & NCompaniesTraded.ToString & ", " & OpenValueTraded.ToString & ", " & HighValueTraded.ToString & ", " & LowValueTraded.ToString & ", " & CloseValueTraded.ToString & ")"
            'Dim CommandString As String = "INSERT INTO [Daily Statistics] (Trade_Date, Day_Type, Companies_Traded, Open_Value_Traded, High_Value_Traded, Low_Value_Traded, Close_Value_Traded) VALUES ('" & DateString & "', 'Trade_Day', " & NCompaniesTraded.ToString & ", " & OpenValueTraded.ToString & ", " & HighValueTraded.ToString & ", " & LowValueTraded.ToString & ", " & CloseValueTraded.ToString & ")"
            Dim CommandString As String = "INSERT INTO [Daily Statistics] (Trade_Date, Day_Type, Companies_Traded, Open_Value_Traded, High_Value_Traded, Low_Value_Traded, Close_Value_Traded) VALUES (" & DbDateString & ", 'Trade_Day', " & NCompaniesTraded.ToString & ", " & OpenValueTraded.ToString & ", " & HighValueTraded.ToString & ", " & LowValueTraded.ToString & ", " & CloseValueTraded.ToString & ")"
            Message.Add("CommandString = " & CommandString & vbCrLf)

            Dim cmd As New System.Data.OleDb.OleDbCommand
            da.InsertCommand = cmd
            da.InsertCommand.Connection = myConnection

            da.InsertCommand.CommandText = CommandString
            da.InsertCommand.ExecuteNonQuery()


            'Close the database:
            da.InsertCommand.Connection.Close()

        Catch ex As Exception
            Message.AddWarning("Error writing to database: " & ex.Message & vbCrLf)
        End Try

    End Sub

    'Public Sub DailyStatsBetweenDates(ByVal TableName As String, ByVal FromDateValue As Date, ByVal ToDateValue As Date)
    Public Sub DailyStatsBetweenDates(ByVal TableName As String, ByVal FromDateString As String, ByVal ToDateString As String)
        'Calculate the Daily share trading Statistics between the dates in FromDateValue and ToDateValue.
        'The results are stored in the Calculation database in the table named TableName.

        Dim FromDateValue As Date = CDate(FromDateString)
        Dim ToDateValue As Date = CDate(ToDateString)

        Dim DateValue As Date = FromDateValue
        Dim DateString As String

        If SharePriceDbPath = "" Then
            Message.AddWarning("A share price database has not been selected!" & vbCrLf)
            Exit Sub
        End If

        If System.IO.File.Exists(SharePriceDbPath) Then
            'Share Price Database file exists.
        Else
            'Share Price Database file does not exist!
            Message.AddWarning("The share price database was not found: " & SharePriceDbPath & vbCrLf)
            Exit Sub
        End If

        If CalculationsDbPath = "" Then
            Message.AddWarning("A Calculations database has not been selected!" & vbCrLf)
            Exit Sub
        End If

        If System.IO.File.Exists(CalculationsDbPath) Then
            'Calculations Database file exists.
        Else
            'Calculations Database file does not exist!
            Message.AddWarning("The calculations database was not found: " & CalculationsDbPath & vbCrLf)
            Exit Sub
        End If

        'Connection to Share Price database:
        Dim SharePriceConnString As String
        Dim mySharePriceConnection As OleDb.OleDbConnection = New OleDb.OleDbConnection
        Dim SharePriceQuery As String

        Try
            SharePriceConnString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source =" & SharePriceDbPath 'DatabasePath
            mySharePriceConnection.ConnectionString = SharePriceConnString
            mySharePriceConnection.Open()
        Catch ex As Exception
            Message.AddWarning("Error connecting to Share Price database: " & ex.Message & vbCrLf)
            Exit Sub
        End Try

        Dim ds As DataSet = New DataSet
        Dim da As OleDb.OleDbDataAdapter

        'Connection to Calculations database:
        Dim CalcsConnString As String
        Dim myCalcsConnection As OleDb.OleDbConnection = New OleDb.OleDbConnection
        'Dim QueryCalcs As String
        Dim InsertCommandString As String

        Try
            CalcsConnString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source =" & CalculationsDbPath 'DatabasePath
            myCalcsConnection.ConnectionString = CalcsConnString
            myCalcsConnection.Open()
        Catch ex As Exception
            Message.AddWarning("Error connecting to Calculations database: " & ex.Message & vbCrLf)
            Exit Sub
        End Try

        Dim NRows As Integer
        Dim I As Integer

        Dim DayType As String = ""
        Dim NCompaniesTraded As Integer = 0
        Dim OpenValueTraded As Double = 0
        Dim HighValueTraded As Double = 0
        Dim LowValueTraded As Double = 0
        Dim CloseValueTraded As Double = 0
        Dim VolumeTraded As Integer = 0

        Message.Add("Processing data from date: " & "#" & Format(FromDateValue, "MM-dd-yyyy") & "#" & "  to date: " & "#" & Format(ToDateValue, "MM-dd-yyyy") & "#" & vbCrLf)
        Message.Add("Start time: " & Format(Now, "HH:mm:ss") & vbCrLf)
        Message.Add("The following dates are displayed with the format: #Month-Day-Year#." & vbCrLf)

        While Date.Compare(DateValue, ToDateValue) <= 0 '<0 changed to <= 0 to Include the To Date in the calculations.
            'Process the share trades for one day:
        DateString = "#" & Format(DateValue, "MM-dd-yyyy") & "#"
            Message.Add("Processing date: " & DateString & vbCrLf)

            SharePriceQuery = "Select * From ASX_Share_Prices Where Trade_Date = " & DateString

            da = New OleDb.OleDbDataAdapter(SharePriceQuery, mySharePriceConnection)
            da.MissingSchemaAction = MissingSchemaAction.AddWithKey

            ds.Clear()

            NCompaniesTraded = 0
            OpenValueTraded = 0
            HighValueTraded = 0
            LowValueTraded = 0
            CloseValueTraded = 0

            Try
                da.Fill(ds, "myData")
                NRows = ds.Tables(0).Rows.Count

                For I = 0 To NRows - 1
                    If ds.Tables(0).Rows(I).Item("Volume") = 0 Then
                        'No shares in this company have been traded.
                    Else
                        'Shares in this company have been traded.
                        NCompaniesTraded += 1
                        VolumeTraded = ds.Tables(0).Rows(I).Item("Volume")
                        OpenValueTraded = OpenValueTraded + VolumeTraded * ds.Tables(0).Rows(I).Item("Open_Price")
                        HighValueTraded = HighValueTraded + VolumeTraded * ds.Tables(0).Rows(I).Item("High_Price")
                        LowValueTraded = LowValueTraded + VolumeTraded * ds.Tables(0).Rows(I).Item("Low_Price")
                        CloseValueTraded = CloseValueTraded + VolumeTraded * ds.Tables(0).Rows(I).Item("Close_Price")
                    End If
                Next

                ds.Clear()

                If DateValue.DayOfWeek = DayOfWeek.Sunday Then
                    DayType = "WeekEnd"
                ElseIf DateValue.DayOfWeek = DayOfWeek.Saturday Then
                    DayType = "WeekEnd"
                ElseIf NCompaniesTraded = 0 Then
                    DayType = "No_Trades"
                Else
                    DayType = "Trade_Day"
                End If


                'InsertCommandString = "INSERT INTO [Daily Statistics] (Trade_Date, Day_Type, Companies_Traded, Open_Value_Traded, High_Value_Traded, Low_Value_Traded, Close_Value_Traded) VALUES (" & DateString & ", 'Trade_Day', " & NCompaniesTraded.ToString & ", " & OpenValueTraded.ToString & ", " & HighValueTraded.ToString & ", " & LowValueTraded.ToString & ", " & CloseValueTraded.ToString & ")"
                'InsertCommandString = "INSERT INTO [" & TableName & "] (Trade_Date, Day_Type, Companies_Traded, Open_Value_Traded, High_Value_Traded, Low_Value_Traded, Close_Value_Traded) VALUES (" & DateString & ", '" & DayType & "', " & NCompaniesTraded.ToString & ", " & OpenValueTraded.ToString & ", " & HighValueTraded.ToString & ", " & LowValueTraded.ToString & ", " & CloseValueTraded.ToString & ")"
                'InsertCommandString = "INSERT INTO [" & TableName & "] (Trade_Date, Day_Type, Day_Of_Week, Companies_Traded, Open_Value_Traded, High_Value_Traded, Low_Value_Traded, Close_Value_Traded) VALUES (" & DateString & ", '" & DayType & ", '" & DateValue.DayOfWeek & "', " & NCompaniesTraded.ToString & ", " & OpenValueTraded.ToString & ", " & HighValueTraded.ToString & ", " & LowValueTraded.ToString & ", " & CloseValueTraded.ToString & ")"
                'InsertCommandString = "INSERT INTO [" & TableName & "] (Trade_Date, Day_Type, Day_Of_Week, Companies_Traded, Open_Value_Traded, High_Value_Traded, Low_Value_Traded, Close_Value_Traded) VALUES (" & DateString & ", '" & DayType & "', " & DateValue.DayOfWeek & ", " & NCompaniesTraded.ToString & ", " & OpenValueTraded.ToString & ", " & HighValueTraded.ToString & ", " & LowValueTraded.ToString & ", " & CloseValueTraded.ToString & ")"
                InsertCommandString = "INSERT INTO [" & TableName & "] (Trade_Date, Day_Type, Day_Of_Week, Day_No, Companies_Traded, Open_Value_Traded, High_Value_Traded, Low_Value_Traded, Close_Value_Traded) VALUES (" & DateString & ", '" & DayType & "', " & DateValue.DayOfWeek & ", " & Int(DateValue.ToOADate) & ", " & NCompaniesTraded.ToString & ", " & OpenValueTraded.ToString & ", " & HighValueTraded.ToString & ", " & LowValueTraded.ToString & ", " & CloseValueTraded.ToString & ")"
                'Int(CDbl(DateValue.ToOADate)) 

                'NOTE: The following message was only shown for debugging:
                'Message.Add("CommandString = " & InsertCommandString & vbCrLf)

                Dim cmd As New System.Data.OleDb.OleDbCommand
                da.InsertCommand = cmd
                da.InsertCommand.Connection = myCalcsConnection

                da.InsertCommand.CommandText = InsertCommandString
                da.InsertCommand.ExecuteNonQuery()

                'Close the database:
                'da.InsertCommand.Connection.Close()

            Catch ex As Exception
                Message.AddWarning("Error writing to database: " & ex.Message & vbCrLf)
            End Try

            'Increment DateValue by one day:
            'DateValue.AddDays(1)
            DateValue = DateValue.AddDays(1)

            'This will allow the connection to the checked occasionally:
            Application.DoEvents() 'The will check if a CancelImport button has been pressed.
            System.Threading.Thread.Sleep(100) 'This allows time for the CancelImport property to be updated.

        End While

        mySharePriceConnection.Close()
        myCalcsConnection.Close()

        Message.Add("Daily statistics calculations complete." & vbCrLf)
        Message.Add("End time: " & Format(Now, "HH:mm:ss") & vbCrLf & vbCrLf)

    End Sub

    Public Sub DailyStatsDayNos(ByVal TableName As String)
        'Caluclate the Rel_Day_No and Trade_Day_No values in the Daily share trading Statistics table.

        If CalculationsDbPath = "" Then
            Message.AddWarning("A Calculations database has not been selected!" & vbCrLf)
            Exit Sub
        End If

        If System.IO.File.Exists(CalculationsDbPath) Then
            'Calculations Database file exists.
        Else
            'Calculations Database file does not exist!
            Message.AddWarning("The calculations database was not found: " & CalculationsDbPath & vbCrLf)
            Exit Sub
        End If

        'Connection to Calculations database:
        Dim CalcsConnString As String
        Dim myCalcsConnection As OleDb.OleDbConnection = New OleDb.OleDbConnection
        Dim QueryCalcs As String
        Dim InsertCommandString As String

        Try
            CalcsConnString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source =" & CalculationsDbPath 'DatabasePath
            myCalcsConnection.ConnectionString = CalcsConnString
            myCalcsConnection.Open()
        Catch ex As Exception
            Message.AddWarning("Error connecting to Calculations database: " & ex.Message & vbCrLf)
            Exit Sub
        End Try

        'QueryCalcs = "Select Trade_Date, Day_Type, Day_No, Rel_Day_No, Trade_Day_No From " & TableName & " Order By Trade_Date"
        QueryCalcs = "Select Trade_Date, Day_Type, Day_No, Rel_Day_No, Trade_Day_No, Week_No, Day_Of_Week, Month_No, Day_Of_Month, Year_No, Day_Of_Year From " & TableName & " Order By Trade_Date"

        Dim ds As DataSet = New DataSet
        Dim da As OleDb.OleDbDataAdapter

        da = New OleDb.OleDbDataAdapter(QueryCalcs, myCalcsConnection)
        da.MissingSchemaAction = MissingSchemaAction.AddWithKey

        ds.Clear()

        Message.Add("Calculating the Relative Day Number and Trade Day Number for each day in the Daily Statistics table." & vbCrLf)
        Message.Add("Start time: " & Format(Now, "HH:mm:ss") & vbCrLf)

        Try
            da.Fill(ds, "myData")
            Dim NRows As Integer = ds.Tables(0).Rows.Count
            Message.Add("Number of Daily Statistics days to process: " & NRows & vbCrLf)

            Dim LastRelativeDayNumber As Integer = 0
            Dim LastTradeDayNumber As Integer = 0

            Dim DateValue As Date

            'These are used to generate the Week_No:
            Dim CI As New System.Globalization.CultureInfo("en-US")
            Dim Cal As System.Globalization.Calendar = CI.Calendar
            Dim CWR As System.Globalization.CalendarWeekRule = CI.DateTimeFormat.CalendarWeekRule
            Dim FirstDOW As DayOfWeek = CI.DateTimeFormat.FirstDayOfWeek


            'Columns already calculated: Trade_Date, Day_Type, Day_No, 
            'Columns to calculate:       Rel_Day_No, Trade_Day_No, Week_No, Day_Of_Week, Month_No, Day_Of_Month, Year_No, Day_Of_Year

            Dim I As Integer
            For I = 0 To NRows - 1
                If I Mod 100 = 0 Then
                    Message.Add("Processing row: " & I & vbCrLf)
                End If

                LastRelativeDayNumber += 1
                ds.Tables(0).Rows(I).Item("Rel_Day_No") = LastRelativeDayNumber

                If ds.Tables(0).Rows(I).Item("Day_Type") = "Trade_Day" Then
                    LastTradeDayNumber += 1
                    ds.Tables(0).Rows(I).Item("Trade_Day_No") = LastTradeDayNumber
                Else
                    ds.Tables(0).Rows(I).Item("Trade_Day_No") = DBNull.Value
                End If

                DateValue = ds.Tables(0).Rows(I).Item("Trade_Date")

                ds.Tables(0).Rows(I).Item("Week_No") = Cal.GetWeekOfYear(DateValue, CWR, FirstDOW)

                ds.Tables(0).Rows(I).Item("Day_Of_Week") = DateValue.DayOfWeek

                ds.Tables(0).Rows(I).Item("Month_No") = DateValue.Month

                ds.Tables(0).Rows(I).Item("Day_Of_Month") = DateValue.Day

                ds.Tables(0).Rows(I).Item("Year_No") = DateValue.Year

                ds.Tables(0).Rows(I).Item("Day_Of_Year") = DateValue.DayOfYear
            Next

            'The following line is required to prevent the error: Update requires a valid UpdateCommand when passed DataRow collection with modified rows.
            Dim cb = New OleDb.OleDbCommandBuilder(da)

            da.Update(ds.Tables(0))
            ds.Tables(0).AcceptChanges()

            Message.Add("Daily statistics Day Number calculations complete." & vbCrLf)
            Message.Add("End time: " & Format(Now, "HH:mm:ss") & vbCrLf & vbCrLf)

        Catch ex As Exception
            Message.AddWarning("Error generating day numbers: " & ex.Message & vbCrLf)
        End Try

    End Sub

    'Public Sub CompanyDailyStatsBetweenDates(ByVal TableName As String, ByVal AsxCode As String, ByVal FromDateValue As Date, ByVal ToDateValue As Date)
    Public Sub CompanyDailyStatsBetweenDates(ByVal TableName As String, ByVal AsxCode As String, ByVal FromDateString As String, ByVal ToDateString As String)
        'Calculate the daily statistics for one company between two dates.
        'The results are written to the TableName table.

        Dim FromDateValue As Date = CDate(FromDateString)
        Dim ToDateValue As Date = CDate(ToDateString)

        If SharePriceDbPath = "" Then
            Message.AddWarning("A share price database has not been selected!" & vbCrLf)
            Exit Sub
        End If

        If System.IO.File.Exists(SharePriceDbPath) Then
            'Share Price Database file exists.
        Else
            'Share Price Database file does not exist!
            Message.AddWarning("The share price database was not found: " & SharePriceDbPath & vbCrLf)
            Exit Sub
        End If

        'Connection to Share Price database:
        Dim SharePriceConnString As String
        Dim mySharePriceConnection As OleDb.OleDbConnection = New OleDb.OleDbConnection
        Dim SharePriceQuery As String

        Try
            SharePriceConnString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source =" & SharePriceDbPath 'DatabasePath
            mySharePriceConnection.ConnectionString = SharePriceConnString
            mySharePriceConnection.Open()
        Catch ex As Exception
            Message.AddWarning("Error connecting to Share Price database: " & ex.Message & vbCrLf)
            Exit Sub
        End Try

        Dim ds As DataSet = New DataSet
        Dim da As OleDb.OleDbDataAdapter

        Dim FromDateStr As String = "#" & Format(FromDateValue, "MM-dd-yyyy") & "#"
        Dim ToDateStr As String = "#" & Format(ToDateValue, "MM-dd-yyyy") & "#"
        SharePriceQuery = "Select * From ASX_Share_Prices Where ASX_Code = '" & AsxCode & "' And Trade_Date Between " & FromDateStr & " And " & ToDateStr

        da = New OleDb.OleDbDataAdapter(SharePriceQuery, mySharePriceConnection)
        da.MissingSchemaAction = MissingSchemaAction.AddWithKey

        ds.Clear()

        Dim MinTradeValue As Single = 0
        Dim MaxTradeValue As Single = 0
        Dim MeanTradeValue As Single = 0
        Dim StdDevTradeValue As Single = 0

        Dim TradeVals As New List(Of Single)
        Dim CloseTradeVal As Single = 0

        Message.Add("Generating trading statistics for company: " & AsxCode & vbCrLf)

        Try
            da.Fill(ds, "myData")
            Dim NRows As Integer = ds.Tables(0).Rows.Count

            If NRows = 0 Then
                Message.AddWarning("There are no records to process between the specified dates." & vbCrLf)
                Exit Sub
            End If

            Message.Add("Processing " & NRows & " records between the two dates." & vbCrLf)

            Dim I As Integer
            For I = 0 To NRows - 1
                CloseTradeVal = ds.Tables(0).Rows(I).Item("Close_Price") * ds.Tables(0).Rows(I).Item("Volume")
                TradeVals.Add(CloseTradeVal)
            Next

            Dim Variance As Double = 0
            'Dim Average As Double = TradeVals.Average
            MeanTradeValue = TradeVals.Average

            For Each Val As Single In TradeVals
                Variance += (Val - MeanTradeValue) ^ 2
            Next

            Variance /= TradeVals.Count
            StdDevTradeValue = Math.Sqrt(Variance)

            MinTradeValue = TradeVals.Min
            MaxTradeValue = TradeVals.Max

            'Write the data to the table:
            Dim CommandString As String = "INSERT INTO [" & TableName & "] (ASX_Code, From_Date, To_Date, Minimum, Maximum, Mean, StdDev) VALUES ('" & AsxCode & "', " & FromDateStr & ", " & ToDateStr & ", " & MinTradeValue & ", " & MaxTradeValue & ", " & MeanTradeValue & ", " & StdDevTradeValue & ")"
            'Message.Add("CommandString = " & CommandString & vbCrLf)

            Dim cmd As New System.Data.OleDb.OleDbCommand
            da.InsertCommand = cmd
            da.InsertCommand.Connection = mySharePriceConnection

            da.InsertCommand.CommandText = CommandString
            da.InsertCommand.ExecuteNonQuery()


            'Close the database:
            da.InsertCommand.Connection.Close()


        Catch ex As Exception
            Message.AddWarning("Error writing to database: " & ex.Message & vbCrLf)
        End Try

    End Sub

    'Public Sub CalcSharePriceDistributions(ByVal TableName As String, ByVal AsxCode As String, ByVal FromDateString As String, ByVal ToDateString As String, ByVal WindowLen As Integer)
    'Public Sub CalcSharePriceDistributions(ByVal TableName As String, ByVal AsxCode As String, ByVal FromDateString As String, ByVal ToDateString As String, ByVal WindowLen As Integer, ByVal TradeDaysOnly As Boolean)
    Public Sub CalcSharePriceDistAllDays(ByVal TableName As String, ByVal AsxCode As String, ByVal FromDateString As String, ByVal ToDateString As String, ByVal WindowLen As Integer)
        'Calculate the closing share price distributions between the dates in FromDateString and ToDateString for the single company with the code AsxCode.
        'The mean and standard deviation of the closing prices are calculated over a woindow length of side WindowLen days.
        'The trade volume is also analysed.
        'The WindowLength includes all days, trade days and non-trade days. (Non-trade days include weekends, holidays and trading halt days.)

        Dim FromDateValue As Date = CDate(FromDateString)
        Dim ToDateValue As Date = CDate(ToDateString)
        Dim LastDateValue As Date = ToDateValue.AddDays(WindowLen)

        If SharePriceDbPath = "" Then
            Message.AddWarning("A share price database has not been selected!" & vbCrLf)
            Exit Sub
        End If

        If System.IO.File.Exists(SharePriceDbPath) Then
            'Share Price Database file exists.
        Else
            'Share Price Database file does not exist!
            Message.AddWarning("The share price database was not found: " & SharePriceDbPath & vbCrLf)
            Exit Sub
        End If

        'Connection to Share Price database:
        Dim SharePriceConnString As String
        Dim mySharePriceConnection As OleDb.OleDbConnection = New OleDb.OleDbConnection
        Dim SharePriceQuery As String

        Try
            SharePriceConnString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source =" & SharePriceDbPath 'DatabasePath
            mySharePriceConnection.ConnectionString = SharePriceConnString
            mySharePriceConnection.Open()
        Catch ex As Exception
            Message.AddWarning("Error connecting to Share Price database: " & ex.Message & vbCrLf)
            Exit Sub
        End Try

        Dim ds As DataSet = New DataSet
        Dim da As OleDb.OleDbDataAdapter

        Dim FromDateStr As String = "#" & Format(FromDateValue, "MM-dd-yyyy") & "#"
        Dim ToDateStr As String = "#" & Format(LastDateValue, "MM-dd-yyyy") & "#"
        SharePriceQuery = "Select * From ASX_Share_Prices Where ASX_Code = '" & AsxCode & "' And Trade_Date Between " & FromDateStr & " And " & ToDateStr & " Order By Trade_Date"

        da = New OleDb.OleDbDataAdapter(SharePriceQuery, mySharePriceConnection)
        da.MissingSchemaAction = MissingSchemaAction.AddWithKey

        ds.Clear()

        'Connection to Calculations database:
        Dim CalcsConnString As String
        Dim myCalcsConnection As OleDb.OleDbConnection = New OleDb.OleDbConnection
        Dim InsertCommandString As String

        Try
            CalcsConnString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source =" & CalculationsDbPath 'DatabasePath
            myCalcsConnection.ConnectionString = CalcsConnString
            myCalcsConnection.Open()
        Catch ex As Exception
            Message.AddWarning("Error connecting to Calculations database: " & ex.Message & vbCrLf)
            Exit Sub
        End Try



        Dim MeanCloseValue As Single = 0
        Dim StdDevCloseValue As Single = 0
        Dim MeanVolume As Single = 0
        Dim StdDevVolume As Single = 0

        Message.Add("Calculating closing share price distributions for company: " & AsxCode & vbCrLf)

        Try
            da.Fill(ds, "myData")
            Dim NRows As Integer = ds.Tables(0).Rows.Count

            If NRows = 0 Then
                Message.AddWarning("There are no records to process between the specified dates." & vbCrLf)
                Exit Sub
            End If

            'The dataset ds contains all the share trade records for the AsxCode company between the dates to be processed.
            'The data within each WindowLen window will be extracted and processed to determine the mean and standard deviation.
            'https://www.dotnetperls.com/datatable-select-vbnet
            'The window length can included non-trade days.

            Dim WindowStart As Date = FromDateValue
            Dim WindowStartStr As String
            Dim WindowEnd As Date
            Dim WindowEndStr As String

            Dim FirstDate As Date
            Dim LastDate As Date

            Dim I As Integer 'Loop index
            'Row Fields to generate:
            'ASXCode (= AsxCode)
            'StartDate (= WindowStart)
            'EndDate (= WindowEnd)
            'NDays (= WindowLen)
            Dim NTradeDays As Integer '(=NRows)
            Dim MeanClose As Single
            Dim StdDevClose As Single
            Dim MeanChange As Single
            Dim StdDevChange As Single
            Dim MeanChangePct As Single
            Dim StdDevChangePct As Single
            Dim MeanVol As Single
            Dim StdDevVol As Single

            Dim SumClose As Double
            Dim Change As Double
            Dim SumChange As Double
            Dim SumChangePct As Double
            Dim SumVol As Double

            Dim SumSqCloseDev As Double
            Dim SumSqChangeDev As Double
            Dim SumSqChangePctDev As Double
            Dim SumSqVolDev As Double

            Dim LastWindowFirstClose As Single = ds.Tables(0).Rows(0).Item("Close_Price") 'Used to calculate the first price change.
            Dim LastClose As Single  'Used to calculate price change. Updated for each window.

            While WindowStart <= ToDateValue
                WindowStartStr = "#" & Format(WindowStart, "MM-dd-yyyy") & "#"
                WindowEnd = WindowStart.AddDays(WindowLen - 1)
                WindowEndStr = "#" & Format(WindowEnd, "MM-dd-yyyy") & "#"

                Dim WindowData() As DataRow = ds.Tables(0).Select("Trade_Date Between " & WindowStartStr & " And " & WindowEndStr) 'Select data within the analysis window.
                NRows = WindowData.Count
                If NRows > 0 Then
                    FirstDate = WindowData(0).Item("Trade_Date") 'The Trade_Date of the first record in the window. (May be after the WindowStart.)
                    LastDate = WindowData(NRows - 1).Item("Trade_Date") 'The Trade_Date of the last record in the window. (May be before the WindowEnd.)

                    'Calculate the Means:
                    SumClose = 0
                    SumChange = 0
                    SumChangePct = 0
                    SumVol = 0
                    LastClose = LastWindowFirstClose
                    For I = 0 To NRows - 1
                        SumClose += WindowData(I).Item("Close_Price")
                        Change = WindowData(I).Item("Close_Price") - LastClose
                        SumChange += Change
                        SumChangePct += Change / LastClose * 100
                        SumVol += WindowData(I).Item("Volume")
                        LastClose = WindowData(I).Item("Close_Price")
                    Next

                    MeanClose = SumClose / NRows
                    MeanChange = SumChange / NRows
                    MeanChangePct = SumChangePct / NRows
                    MeanVol = SumVol / NRows

                    'Calculate the standard deviations:
                    SumSqCloseDev = 0
                    SumSqChangeDev = 0
                    SumSqChangePctDev = 0
                    SumSqVolDev = 0
                    LastClose = LastWindowFirstClose
                    For I = 0 To NRows - 1
                        SumSqCloseDev += (WindowData(I).Item("Close_Price") - MeanClose) ^ 2
                        Change = WindowData(I).Item("Close_Price") - LastClose
                        SumSqChangeDev += (Change - MeanChange) ^ 2
                        SumSqChangePctDev += ((Change / LastClose * 100) - MeanChangePct) ^ 2
                        SumSqVolDev += (WindowData(I).Item("Volume") - MeanVol) ^ 2
                        LastClose = WindowData(I).Item("Close_Price")
                    Next
                    StdDevClose = Math.Sqrt(SumSqCloseDev / NRows)
                    StdDevChange = Math.Sqrt(SumSqChangeDev / NRows)
                    StdDevChangePct = Math.Sqrt(SumSqChangePctDev / NRows)
                    StdDevVol = Math.Sqrt(SumSqVolDev / NRows)

                    LastWindowFirstClose = ds.Tables(0).Rows(0).Item("Close_Price")

                    'Write the calculated values to the table:
                    InsertCommandString = "INSERT INTO [" & TableName & "] (ASX_Code, Start_Date, End_Date, NDays, NTradeDays, Mean_Close, StdDev_Close, Mean_Change, StdDev_Change, Mean_Change_Pct, StdDev_Change_Pct, Mean_Vol, StdDev_Vol) VALUES (" & AsxCode & ", " & WindowStart & ", " & WindowEnd & ", " & WindowLen & ", " & NRows & ", " & MeanClose & ", " & StdDevClose & ", " & MeanChange & ", " & StdDevChange & ", " & MeanChangePct & ", " & StdDevChangePct & ", " & MeanVol & ", " & StdDevVol & ")"

                    Message.Add("InsertCommandString = " & InsertCommandString & vbCrLf)

                    Dim cmd As New System.Data.OleDb.OleDbCommand
                    da.InsertCommand = cmd
                    da.InsertCommand.Connection = myCalcsConnection

                    da.InsertCommand.CommandText = InsertCommandString
                    da.InsertCommand.ExecuteNonQuery()

                Else
                    'No share trades in the analysis window.
                End If

                WindowStart = WindowStart.AddDays(1) 'Increment the window start date.
            End While

            mySharePriceConnection.Close()
            myCalcsConnection.Close()

            Message.Add("Closing share price distributions calculations complete." & vbCrLf)
            Message.Add("End time: " & Format(Now, "HH:mm:ss") & vbCrLf & vbCrLf)

        Catch ex As Exception
            Message.AddWarning("Error: " & ex.Message & vbCrLf)
        End Try



    End Sub

    Public Sub CalcSharePriceDistTradeDays(ByVal TableName As String, ByVal AsxCode As String, ByVal FromDateString As String, ByVal ToDateString As String, ByVal WindowLen As Integer)
        'Calculate the closing share price distributions between the dates in FromDateString and ToDateString for the single company with the code AsxCode.
        'The mean and standard deviation of the closing prices are calculated over a woindow length of side WindowLen days.
        'The trade volume is also analysed.
        'The WindowLength includes only trade days. Non-trade days are excluded from the statistics. (Non-trade days include weekends, holidays and trading halt days.)

        Dim FromDateValue As Date = CDate(FromDateString)
        Dim ToDateValue As Date = CDate(ToDateString)
        'Dim LastDateValue As Date = ToDateValue.AddDays(WindowLen)

        If SharePriceDbPath = "" Then
            Message.AddWarning("A share price database has not been selected!" & vbCrLf)
            Exit Sub
        End If

        If System.IO.File.Exists(SharePriceDbPath) Then
            'Share Price Database file exists.
        Else
            'Share Price Database file does not exist!
            Message.AddWarning("The share price database was not found: " & SharePriceDbPath & vbCrLf)
            Exit Sub
        End If

        'Connection to Share Price database:
        Dim SharePriceConnString As String
        Dim mySharePriceConnection As OleDb.OleDbConnection = New OleDb.OleDbConnection
        Dim SharePriceQuery As String

        Try
            SharePriceConnString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source =" & SharePriceDbPath 'DatabasePath
            mySharePriceConnection.ConnectionString = SharePriceConnString
            mySharePriceConnection.Open()
        Catch ex As Exception
            Message.AddWarning("Error connecting to Share Price database: " & ex.Message & vbCrLf)
            Exit Sub
        End Try

        Dim ds As DataSet = New DataSet
        Dim da As OleDb.OleDbDataAdapter

        Dim FromDateStr As String = "#" & Format(FromDateValue, "MM-dd-yyyy") & "#"
        'Dim ToDateStr As String = "#" & Format(LastDateValue, "MM-dd-yyyy") & "#"
        Dim ToDateStr As String = "#" & Format(ToDateValue, "MM-dd-yyyy") & "#"
        SharePriceQuery = "Select * From ASX_Share_Prices Where ASX_Code = '" & AsxCode & "' And Trade_Date Between " & FromDateStr & " And " & ToDateStr & " Order By Trade_Date"

        Message.Add("Query = " & SharePriceQuery & vbCrLf)

        da = New OleDb.OleDbDataAdapter(SharePriceQuery, mySharePriceConnection)
        da.MissingSchemaAction = MissingSchemaAction.AddWithKey

        ds.Clear()

        'Connection to Calculations database:
        Dim CalcsConnString As String
        Dim myCalcsConnection As OleDb.OleDbConnection = New OleDb.OleDbConnection
        Dim InsertCommandString As String

        Try
            CalcsConnString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source =" & CalculationsDbPath 'DatabasePath
            myCalcsConnection.ConnectionString = CalcsConnString
            myCalcsConnection.Open()
        Catch ex As Exception
            Message.AddWarning("Error connecting to Calculations database: " & ex.Message & vbCrLf)
            Exit Sub
        End Try

        'Dim MeanCloseValue As Single = 0
        'Dim StdDevCloseValue As Single = 0
        'Dim MeanVolume As Single = 0
        'Dim StdDevVolume As Single = 0

        Message.Add("Calculating closing share price distributions (trade days only) for company: " & AsxCode & vbCrLf)

        Try
            da.Fill(ds, "myData")
            Dim NRows As Integer = ds.Tables(0).Rows.Count
            Message.Add("Number of share price records loaded: " & NRows & vbCrLf)

            If NRows = 0 Then
                Message.AddWarning("There are no records to process between the specified dates." & vbCrLf)
                Exit Sub
            End If

            'The dataset ds contains all the share trade records for the AsxCode company between the dates to be processed.
            'The data within each WindowLen window will be extracted and processed to determine the mean and standard deviation.
            'The window length only includes trade days contained in dataset ds.

            Dim WindowStart As Date = FromDateValue
            Dim WindowStartStr As String
            Dim WindowEnd As Date
            Dim WindowEndStr As String
            Dim NextTradeDate As Date 'The next trade date after the end of the window.
            Dim NextTradeDateStr As String 'The next trade date string after the end of the window

            'Dim FirstDate As Date
            'Dim LastDate As Date

            Dim I As Integer 'Loop index
            Dim J As Integer 'Loop index
            'Row Fields to generate:
            'ASXCode (= AsxCode)
            'StartDate (= WindowStart)
            'EndDate (= WindowEnd)
            'NDays (= WindowLen)
            Dim NDays As Integer
            Dim TS As TimeSpan
            'Dim NTradeDays As Integer '(=NRows)
            Dim MeanClose As Single
            Dim StdDevClose As Single
            Dim MeanChange As Single
            Dim StdDevChange As Single
            Dim MeanChangePct As Single
            Dim StdDevChangePct As Single
            Dim MeanVol As Single
            Dim StdDevVol As Single

            Dim SumClose As Double
            Dim Change As Double
            Dim SumChange As Double
            Dim SumChangePct As Double
            Dim SumVol As Double

            Dim SumSqCloseDev As Double
            Dim SumSqChangeDev As Double
            Dim SumSqChangePctDev As Double
            Dim SumSqVolDev As Double

            Dim LastWindowFirstClose As Single = ds.Tables(0).Rows(0).Item("Close_Price") 'Used to calculate the first price change.
            Dim LastClose As Single  'Used to calculate price change. Updated for each window.

            'Loop Index example:
            'Item No:  1 2 3 4 5 6 7 8 9 - NRows = 9
            'Index No: 0 1 2 3 4 5 6 7 8
            '

            'For I = 0 To NRows - WindowLen 'Process each trade record in ds
            For I = 0 To NRows - WindowLen - 1 'Process each trade record in ds
                'Calculate the mean values over the analysis window:
                SumClose = 0
                SumChange = 0
                SumChangePct = 0
                SumVol = 0
                LastClose = LastWindowFirstClose
                For J = 0 To WindowLen - 1
                    SumClose += ds.Tables(0).Rows(I + J).Item("Close_Price")
                    Change = ds.Tables(0).Rows(I + J).Item("Close_Price") - LastClose
                    SumChange += Change
                    SumChangePct += Change / LastClose * 100
                    SumVol += ds.Tables(0).Rows(I + J).Item("Volume")
                    LastClose = ds.Tables(0).Rows(I + J).Item("Close_Price")
                Next
                MeanClose = SumClose / WindowLen
                MeanChange = SumChange / WindowLen
                MeanChangePct = SumChangePct / WindowLen
                MeanVol = SumVol / WindowLen

                'Message.Add("LastClose =" & LastClose & vbCrLf)
                'Message.Add("MeanClose =" & MeanClose & vbCrLf)

                'Calculate the standard deviations over the analysis window:
                SumSqCloseDev = 0
                SumSqChangeDev = 0
                SumSqChangePctDev = 0
                SumSqVolDev = 0
                LastClose = LastWindowFirstClose
                For J = 0 To WindowLen - 1
                    SumSqCloseDev += (ds.Tables(0).Rows(I + J).Item("Close_Price") - MeanClose) ^ 2
                    Change = ds.Tables(0).Rows(I + J).Item("Close_Price") - LastClose
                    SumSqChangeDev += (Change - MeanChange) ^ 2
                    SumSqChangePctDev += ((Change / LastClose * 100) - MeanChangePct) ^ 2
                    SumSqVolDev += (ds.Tables(0).Rows(I + J).Item("Volume") - MeanVol) ^ 2
                    LastClose = ds.Tables(0).Rows(I + J).Item("Close_Price")
                Next
                StdDevClose = Math.Sqrt(SumSqCloseDev / WindowLen)
                StdDevChange = Math.Sqrt(SumSqChangeDev / WindowLen)
                StdDevChangePct = Math.Sqrt(SumSqChangePctDev / WindowLen)
                StdDevVol = Math.Sqrt(SumSqVolDev / WindowLen)

                LastWindowFirstClose = ds.Tables(0).Rows(I).Item("Close_Price")

                'Write the calculated values to the table:
                WindowStart = ds.Tables(0).Rows(I).Item("Trade_Date")
                WindowEnd = ds.Tables(0).Rows(I + WindowLen - 1).Item("Trade_Date")
                NextTradeDate = ds.Tables(0).Rows(I + WindowLen).Item("Trade_Date")
                TS = WindowEnd.Subtract(WindowStart)
                NDays = TS.Days
                WindowStartStr = "#" & Format(WindowStart, "MM-dd-yyyy") & "#"
                WindowEndStr = "#" & Format(WindowEnd, "MM-dd-yyyy") & "#"
                NextTradeDateStr = "#" & Format(NextTradeDate, "MM-dd-yyyy") & "#"
                'InsertCommandString = "INSERT INTO [" & TableName & "] (ASX_Code, Start_Date, End_Date, NDays, NTradeDays, Mean_Close, StdDev_Close, Mean_Change, StdDev_Change, Mean_Change_Pct, StdDev_Change_Pct, Mean_Vol, StdDev_Vol) VALUES (" & AsxCode & ", " & WindowStart & ", " & WindowEnd & ", " & NDays & ", " & WindowLen & ", " & MeanClose & ", " & StdDevClose & ", " & MeanChange & ", " & StdDevChange & ", " & MeanChangePct & ", " & StdDevChangePct & ", " & MeanVol & ", " & StdDevVol & ")"
                'InsertCommandString = "INSERT INTO [" & TableName & "] (ASX_Code, Start_Date, End_Date, NDays, NTradeDays, Mean_Close, StdDev_Close, Mean_Change, StdDev_Change, Mean_Change_Pct, StdDev_Change_Pct, Mean_Vol, StdDev_Vol) VALUES ('" & AsxCode & "', " & WindowStartStr & ", " & WindowEndStr & ", " & NDays & ", " & WindowLen & ", " & MeanClose & ", " & StdDevClose & ", " & MeanChange & ", " & StdDevChange & ", " & MeanChangePct & ", " & StdDevChangePct & ", " & MeanVol & ", " & StdDevVol & ")"
                InsertCommandString = "INSERT INTO [" & TableName & "] (ASX_Code, Start_Date, End_Date, Next_Trade_Date, NDays, NTradeDays, Mean_Close, StdDev_Close, Mean_Change, StdDev_Change, Mean_Change_Pct, StdDev_Change_Pct, Mean_Vol, StdDev_Vol) VALUES ('" & AsxCode & "', " & WindowStartStr & ", " & WindowEndStr & ", " & NextTradeDateStr & ", " & NDays & ", " & WindowLen & ", " & MeanClose & ", " & StdDevClose & ", " & MeanChange & ", " & StdDevChange & ", " & MeanChangePct & ", " & StdDevChangePct & ", " & MeanVol & ", " & StdDevVol & ")"

                'Message.Add("InsertCommandString = " & InsertCommandString & vbCrLf)

                Dim cmd As New System.Data.OleDb.OleDbCommand
                da.InsertCommand = cmd
                da.InsertCommand.Connection = myCalcsConnection

                da.InsertCommand.CommandText = InsertCommandString
                da.InsertCommand.ExecuteNonQuery()


            Next

            mySharePriceConnection.Close()
            myCalcsConnection.Close()

            Message.Add("Closing share price distributions calculations complete." & vbCrLf)
            Message.Add("End time: " & Format(Now, "HH:mm:ss") & vbCrLf & vbCrLf)

        Catch ex As Exception
            Message.AddWarning("Error: " & ex.Message & vbCrLf)
        End Try

    End Sub


    'Public Sub AllCompanyDailyStatsBetweenDates(ByVal TableName As String, ByVal FromDateValue As Date, ByVal ToDateValue As Date)
    Public Sub AllCompanyDailyStatsBetweenDates(ByVal TableName As String, ByVal FromDateString As String, ByVal ToDateString As String)
        'Calculate the daily statistics for all companies trading between two dates.
        'The results are written to the TableName table.

        Dim FromDateValue As Date = CDate(FromDateString)
        Dim ToDateValue As Date = CDate(ToDateString)

        If SharePriceDbPath = "" Then
            Message.AddWarning("A share price database has not been selected!" & vbCrLf)
            Exit Sub
        End If

        If System.IO.File.Exists(SharePriceDbPath) Then
            'Share Price Database file exists.
        Else
            'Share Price Database file does not exist!
            Message.AddWarning("The share price database was not found: " & SharePriceDbPath & vbCrLf)
            Exit Sub
        End If

        'Connection to Share Price database:
        Dim SharePriceConnString As String
        Dim mySharePriceConnection As OleDb.OleDbConnection = New OleDb.OleDbConnection
        Dim SharePriceQuery As String

        Try
            SharePriceConnString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source =" & SharePriceDbPath 'DatabasePath
            mySharePriceConnection.ConnectionString = SharePriceConnString
            mySharePriceConnection.Open()
        Catch ex As Exception
            Message.AddWarning("Error connecting to Share Price database: " & ex.Message & vbCrLf)
            Exit Sub
        End Try

        Dim ds As DataSet = New DataSet
        Dim da As OleDb.OleDbDataAdapter

        Dim FromDateStr As String = "#" & Format(FromDateValue, "MM-dd-yyyy") & "#"
        Dim ToDateStr As String = "#" & Format(ToDateValue, "MM-dd-yyyy") & "#"
        SharePriceQuery = "Select Distinct ASX_Code From ASX_Share_Prices Where Trade_Date Between " & FromDateStr & " And " & ToDateStr

        da = New OleDb.OleDbDataAdapter(SharePriceQuery, mySharePriceConnection)
        da.MissingSchemaAction = MissingSchemaAction.AddWithKey

        ds.Clear()
        Message.Add("Generating daily trding statistics for all companies between " & Format(FromDateValue, "dd MMMM yyyy") & " and " & Format(ToDateValue, "dd MMMM yyyy") & vbCrLf)

        Try
            da.Fill(ds, "myData")
            Dim NRows As Integer = ds.Tables(0).Rows.Count

            If NRows = 0 Then
                Message.AddWarning("There are no companies to process between the specified dates." & vbCrLf)
                Exit Sub
            End If

            Message.Add("Processing " & NRows & " records between the two dates." & vbCrLf)

            Dim I As Integer
            'Dim AsxCode As String
            For I = 0 To NRows - 1
                'CloseTradeVal = ds.Tables(0).Rows(I).Item("Close_Price") * ds.Tables(0).Rows(I).Item("Volume")
                'TradeVals.Add(CloseTradeVal)
                'AsxCode = ds.Tables(0).Rows(I).Item("ASX_Code")

                CompanyDailyStatsBetweenDates(TableName, ds.Tables(0).Rows(I).Item("ASX_Code"), FromDateValue, ToDateValue)
            Next

        Catch ex As Exception
            Message.AddWarning("Error processing all company statistics between dates: " & ex.Message & vbCrLf)
        End Try

    End Sub


    'Public Sub GetDividendPaymentInfo(ByVal AsxCode As String)
    Public Sub DividendReboundAnalysis(ByVal AsxCode As String, ByVal MinBuyDelay As Integer, ByVal MaxBuyDelay As Integer, ByVal MinSellDelay As Integer, ByVal MaxSellDelay As Integer, ByVal BuyBrokPct As Single, ByVal SellBrokPct As Single)
        'Analyse the share price rebound after a dividend is paid.
        'Get dividend payment information for the company with code AsxCode.
        'The dividend info is stored in the table Dividend_Payment_Info in the Calculations database.

        'Step 1 - Get the Close_Price for the company:

        'Connection to Share Price database:
        Dim SharePriceConnString As String
        Dim mySharePriceConnection As OleDb.OleDbConnection = New OleDb.OleDbConnection
        Dim SharePriceQuery As String


        Try
            SharePriceConnString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source =" & SharePriceDbPath 'DatabasePath
            mySharePriceConnection.ConnectionString = SharePriceConnString
            mySharePriceConnection.Open()
        Catch ex As Exception
            Message.AddWarning("Error connecting to Share Price database: " & ex.Message & vbCrLf)
            Exit Sub
        End Try

        Dim ds As DataSet = New DataSet
        Dim da As OleDb.OleDbDataAdapter

        SharePriceQuery = "Select Trade_Date, Close_Price From ASX_Share_Prices Where ASX_Code = '" & AsxCode & "' Order by Trade_Date"

        da = New OleDb.OleDbDataAdapter(SharePriceQuery, mySharePriceConnection)
        da.MissingSchemaAction = MissingSchemaAction.AddWithKey

        ds.Clear()

        'Open the Calculations databse to write the DividendPaymentInfo rows:
        Dim CalcsConnString As String
        Dim myCalcsConnection As OleDb.OleDbConnection = New OleDb.OleDbConnection
        Dim InsertCommand As String
        Dim CalcQuery As String 'Contains the query used to select data from the Calculations database.

        Try
            CalcsConnString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source =" & CalculationsDbPath 'DatabasePath
            myCalcsConnection.ConnectionString = CalcsConnString
            myCalcsConnection.Open()
        Catch ex As Exception
            Message.AddWarning("Error connecting to Calculations database: " & ex.Message & vbCrLf)
            Exit Sub
        End Try

        Dim daCalc As New OleDb.OleDbDataAdapter

        Try
            da.Fill(ds, "ClosePrices")
            Dim NRows As Integer = ds.Tables(0).Rows.Count

            Message.Add(vbCrLf & NRows & " closing prices read for ASX code = " & AsxCode & vbCrLf)

            'Add the DayNo and TradeDayNo columns to the ClosePrices table: --------------------------------------
            Dim dc1 As DataColumn = New DataColumn("DayNo", System.Type.GetType("System.Int32"))
            Dim dc2 As DataColumn = New DataColumn("TradeDayNo", System.Type.GetType("System.Int32"))

            ds.Tables("ClosePrices").Columns.Add(dc1)
            ds.Tables("ClosePrices").Columns.Add(dc2)

            'Generate the DayNo and TradeDayNo values: -----------------------------------------------------------
            Dim TradeDate As Date
            Dim I As Integer
            For I = 0 To NRows - 1
                TradeDate = ds.Tables("ClosePrices").Rows.Item(I)("Trade_Date")
                ds.Tables("ClosePrices").Rows.Item(I)("DayNo") = Int(TradeDate.ToOADate)
                ds.Tables("ClosePrices").Rows.Item(I)("TradeDayNo") = I
                If I Mod 20 = 0 Then
                    Message.Add("DayNo = " & Int(TradeDate.ToOADate) & " TradeDayNo = " & I & " Close Price = " & ds.Tables("ClosePrices").Rows.Item(I)("Close_Price") & vbCrLf)
                End If
            Next

            Message.Add(vbCrLf & "DayNo and TradeDayNo values added to ClosePrices table. " & vbCrLf & vbCrLf)

            'Open Dividends records for AsxCode:
            SharePriceQuery = "Select * From Dividends Where ASX_Code = '" & AsxCode & "' Order By Ex_Div"

            da.SelectCommand = New OleDb.OleDbCommand(SharePriceQuery, mySharePriceConnection)
            'da.MissingSchemaAction = MissingSchemaAction.AddWithKey
            'da.SelectCommand.Connection = mySharePriceConnection
            'mySharePriceConnection.Open()
            da.Fill(ds, "Dividends") 'ASX_Code, Ex_Div, Amount, Rec_Date, Payable, Franking, Credit

            Message.Add(ds.Tables("Dividends").Rows.Count & " dividends loaded for company: " & AsxCode & vbCrLf)

            'Compile Dividend Payment Info ---------------------------------------------------------------
            'Dividend Payment Info columns:
            '  ASX_Code
            '  ExDivDate
            '  ExDivDayNo TO ADD
            '  ExDivTradeDay
            '  PreDivPrice
            '  ExDivPrice
            '  DivAmount
            '  FrCredit
            '  DivPct
            '  GrossDiv
            '  GrossDivPct
            '  DropAmt
            '  DropPct
            Dim DPI_AsxCode As String 'The ASX Code to be written to the Dividend Payment Information table
            Dim DPI_ExDivDate As Date 'The ex-dividend date
            Dim DPI_ExDivDayNo As Integer 'The ex-dividend day number
            Dim DPI_ExDivTradeDay As Integer 'The ex-dividend trade day number
            Dim DPI_PreDivPrice As Single 'The price of the share pre-dividend
            Dim DPI_ExDivPrice As Single 'The price of the share post-dividend
            Dim DPI_DivAmount As Single 'The amount of the dividend per share.
            Dim DPI_FrCredit As Single 'The franking credit per share
            Dim DPI_DivPct As Single 'The dividend as a percent of the pre-div share price
            Dim DPI_GrossDiv As Single 'The Dividend amount plus the franking credit
            Dim DPI_GrossDivPct As Single 'The gross dividend as a percent of the pre-div share price 
            Dim DPI_DropAmt As Single 'The amount the share price dropped ex-dividend
            Dim DPI_DropPct As Single 'The amount the share price dropped ex-dividend as a percent of the pre-div price.

            Dim ExDivDateStr As String 'Date string for queries

            'Process each row in the Dividends table
            For I = 0 To ds.Tables("Dividends").Rows.Count
                DPI_AsxCode = AsxCode
                DPI_ExDivDate = ds.Tables("Dividends").Rows(I)("Ex_Div")
                ExDivDateStr = "#" & Format(DPI_ExDivDate, "MM-dd-yyyy") & "#"
                DPI_DivAmount = ds.Tables("Dividends").Rows(I)("Amount")
                DPI_FrCredit = ds.Tables("Dividends").Rows(I)("Credit")
                DPI_GrossDiv = DPI_DivAmount + DPI_FrCredit 'The gross dividend = dividend amount + franking credit.

                Dim FoundRow() As DataRow = ds.Tables("ClosePrices").Select("Trade_Date = " & ExDivDateStr) 'Trade_Date, Close_Price, DayNo, TradeDayNo
                If FoundRow.Count = 0 Then
                    Message.AddWarning("No ClosePrices record found on date = " & ExDivDateStr & vbCrLf)
                    'The following values cannot be calculated - default to zero:
                    DPI_ExDivDayNo = 0
                    DPI_ExDivTradeDay = 0
                    DPI_ExDivPrice = 0
                    DPI_DropAmt = 0
                    DPI_DivPct = 0
                ElseIf FoundRow.Count = 1 Then
                    DPI_ExDivDayNo = FoundRow(0)("DayNo")
                    DPI_ExDivTradeDay = FoundRow(0)("TradeDayNo")
                    DPI_ExDivPrice = FoundRow(0)("Close_Price")

                    'Now get the pre-dividend closing price:
                    '  Get the ClosePrices record one trading day before the ex-dividend date:
                    Dim FoundRow2() As DataRow = ds.Tables("ClosePrices").Select("TradeDayNo = " & DPI_ExDivTradeDay - 1) 'Trade_Date, Close_Price, DayNo, TradeDayNo
                    If FoundRow2.Count = 0 Then
                        Message.AddWarning("No ClosePrices record found on TradeDayNo = " & DPI_ExDivTradeDay - 1 & vbCrLf)
                        'These values cannot be calculated - default to zero:
                        DPI_DropAmt = 0
                        DPI_DivPct = 0
                        DPI_GrossDivPct = 0
                        DPI_DropAmt = 0
                        DPI_DropPct = 0
                    ElseIf FoundRow2.Count = 1 Then
                        DPI_PreDivPrice = FoundRow2(0)("Close_Price")
                        DPI_DropAmt = DPI_PreDivPrice - DPI_ExDivPrice
                        DPI_DivPct = DPI_DivAmount / DPI_PreDivPrice * 100
                        DPI_GrossDivPct = DPI_GrossDiv / DPI_PreDivPrice * 100
                        DPI_DropAmt = DPI_PreDivPrice - DPI_ExDivPrice
                        DPI_DropPct = DPI_DropAmt / DPI_PreDivPrice * 100
                    Else
                        Message.AddWarning(FoundRow2.Count & " ClosePrices records found on TradeDayNo = " & DPI_ExDivTradeDay - 1 & vbCrLf)
                        'These values cannot be calculated - default to zero:
                        DPI_DropAmt = 0
                        DPI_DivPct = 0
                        DPI_GrossDivPct = 0
                        DPI_DropAmt = 0
                        DPI_DropPct = 0
                    End If
                Else
                    Message.AddWarning(FoundRow.Count & " ClosePrices records found on date = " & ExDivDateStr & vbCrLf)
                    'The following values cannot be calculated - default to zero:
                    DPI_ExDivDayNo = 0
                    DPI_ExDivTradeDay = 0
                    DPI_ExDivPrice = 0
                    DPI_DropAmt = 0
                    DPI_DivPct = 0
                End If

                'Write the row to the DividendPaymentInfo table: ---------------------------------------------------
                If DPI_ExDivDayNo = 0 Then
                    'Valid data not generated - do not write to table
                Else
                    Try
                        InsertCommand = "INSERT INTO Dividend_Payment_Info (ASX_Code, ExDivDate, ExDivDayNo, ExDivTradeDay, PreDivPrice, ExDivPrice, DivAmount, FrCredit, DivPct, GrossDiv, GrossDivPct, DropAmt, DropPct)" _
                        & " VALUES ('" & DPI_AsxCode & "', " & ExDivDateStr & ", " & DPI_ExDivDayNo & ", " & DPI_ExDivTradeDay & ", " & DPI_PreDivPrice & ", " & DPI_ExDivPrice & ", " & DPI_DivAmount & ", " & DPI_FrCredit _
                        & ", " & DPI_DivPct & ", " & DPI_GrossDiv & ", " & DPI_GrossDivPct & ", " & DPI_DropAmt & ", " & DPI_DropPct & ")"

                        daCalc.InsertCommand = New OleDb.OleDbCommand(InsertCommand, myCalcsConnection)
                        daCalc.InsertCommand.ExecuteNonQuery()

                    Catch ex As Exception
                        Message.AddWarning("Error writing to Dividend_Payment_Info table: " & ex.Message & vbCrLf)
                        Message.Add("Insert command:" & vbCrLf & InsertCommand & vbCrLf)
                    End Try
                End If
            Next

            ''Close the databases:
            'daCalc.InsertCommand.Connection.Close()

        Catch ex As Exception
            Message.AddWarning("Error processing dividend payment information: " & ex.Message & vbCrLf)
            'NOTE: If error is: No value given for one or more required parameters.
            '      Then check for a spelling error in the query.
        End Try

        'Compile Dividend Rebound Info ---------------------------------------------------------------
        'Dividend Payment Info columns:
        '  ASX_Code
        '  ExDivDate
        '  ExDivDayNo 
        '  ExDivTradeDay
        '  BuyDelayDays
        '  BuyDelayTradeDays
        '  BuyDate
        '  BuyDayNo
        '  BuyTradeDay
        '  BuyClosePrice
        '  HoldDays
        '  HoldTradeDays
        '  SellDate
        '  SellDayNo
        '  SellTradeDay
        '  SellClosePrice
        '  TradeFactor
        '  ProfitPct
        '  ProfitAnnPct
        '  BrokPct
        '  NetProfitPct
        '  NetProfitAnnPct

        Dim DRI_AsxCode As String 'The ASX Code to be written to the Dividend Rebounf Information table
        Dim DRI_ExDivDate As Date
        Dim DRI_ExDivDayNo As Integer
        Dim DRI_ExDivTradeDay As Integer
        Dim DRI_BuyDelayDays As Integer
        Dim DRI_BuyDelayTradeDays As Integer
        Dim DRI_BuyDate As Date
        Dim DRI_BuyDayNo As Integer
        Dim DRI_BuyTradeDay As Integer
        Dim DRI_BuyClosePrice As Single
        Dim DRI_HoldDays As Integer
        Dim DRI_HoldTradeDays As Integer
        Dim DRI_SellDate As Date
        Dim DRI_SellDayNo As Integer
        Dim DRI_SellTradeDay As Integer
        Dim DRI_SellClosePrice As Single
        Dim DRI_TradeFactor As Single
        Dim DRI_ProfitPct As Single
        Dim DRI_ProfitAnnPct As Single
        Dim DRI_BrokPct As Single = BuyBrokPct + SellBrokPct
        Dim DRI_NetProfitPct As Single
        Dim DRI_NetProfitAnnPct As Single

        Dim DRIExDivDateStr As String 'Date string for writing to table
        Dim BuyDateStr As String 'Date string for writing to table
        Dim SellDateStr As String 'Date string for writing to table

        DRI_AsxCode = AsxCode

        'Dataset ds contains the tables: ClosePrices (used here) and Dividends (not used here)
        'daCalc is the data adaptor connected to the Calculations database.

        'Read Dividend_Payment_Info into DivInfo in daCalc:
        CalcQuery = "Select * From Dividend_Payment_Info Where ASX_Code = '" & DRI_AsxCode & "' Order By ExDivDate"
        daCalc.SelectCommand = New OleDb.OleDbCommand(CalcQuery, myCalcsConnection)
        daCalc.Fill(ds, "DividendInfo") 'ASX_Code, ExDivDate, ExDivDayNo, ExDivTradeDay, PreDivPrice, ExDivPrice, DivAmount, FrCredit, DivPct, GrossDiv, GrossDivPct, DropAmt, DropPct

        Message.Add(ds.Tables("DividendInfo").Rows.Count & " dividend payment information loaded for company: " & AsxCode & vbCrLf)

        Dim DivNo As Integer
        Dim BuyDelay As Integer
        Dim HoldDays As Integer

        'Process each dividend payment:
        For DivNo = 0 To ds.Tables("DividendInfo").Rows.Count - 1
            DRI_ExDivDate = ds.Tables("DividendInfo").Rows.Item(DivNo)("ExDivDate")
            DRIExDivDateStr = "#" & Format(DRI_ExDivDate, "MM-dd-yyyy") & "#"
            DRI_ExDivDayNo = ds.Tables("DividendInfo").Rows.Item(DivNo)("ExDivDayNo")
            'DRI_ExDivTradeDay = ds.Tables("DividendInfo").Rows.Item(DivNo)("ExTradeDay")
            DRI_ExDivTradeDay = ds.Tables("DividendInfo").Rows.Item(DivNo)("ExDivTradeDay")
            'Process each buy delay:
            For BuyDelay = MinBuyDelay To MaxBuyDelay
                DRI_BuyDelayTradeDays = BuyDelay
                DRI_BuyTradeDay = DRI_ExDivTradeDay + BuyDelay
                'Read the ClosingPrice record on the Buy Day:
                Dim FoundRow3() As DataRow = ds.Tables("ClosePrices").Select("TradeDayNo = " & DRI_BuyTradeDay) 'Trade_Date, Close_Price, DayNo, TradeDayNo
                If FoundRow3.Count = 0 Then
                    Message.AddWarning("No ClosePrices record found on buy TradeDayNo = " & DRI_BuyTradeDay & vbCrLf)
                    'These values cannot be calculated - default to zero or ExDivDate:
                    DRI_BuyDate = DRI_ExDivDate
                    DRI_BuyDayNo = 0
                    DRI_BuyClosePrice = 0
                ElseIf FoundRow3.Count = 1 Then
                    DRI_BuyDate = FoundRow3(0)("Trade_Date")
                    BuyDateStr = "#" & Format(DRI_BuyDate, "MM-dd-yyyy") & "#"
                    DRI_BuyDayNo = FoundRow3(0)("DayNo")
                    DRI_BuyDelayDays = DRI_BuyDayNo - DRI_ExDivDayNo
                    DRI_BuyClosePrice = FoundRow3(0)("Close_Price")
                    'Process each sell delay:
                    For HoldDays = MinSellDelay To MaxSellDelay
                        DRI_HoldTradeDays = HoldDays
                        DRI_SellTradeDay = DRI_BuyTradeDay + HoldDays
                        Dim FoundRow4() As DataRow = ds.Tables("ClosePrices").Select("TradeDayNo = " & DRI_SellTradeDay) 'Trade_Date, Close_Price, DayNo, TradeDayNo
                        If FoundRow4.Count = 0 Then
                            Message.AddWarning("No ClosePrices record found on sell TradeDayNo = " & DRI_SellTradeDay & vbCrLf)

                        ElseIf FoundRow4.Count = 1 Then
                            DRI_SellDate = FoundRow4(0)("Trade_Date")
                            SellDateStr = "#" & Format(DRI_SellDate, "MM-dd-yyyy") & "#"
                            DRI_SellDayNo = FoundRow4(0)("DayNo")
                            DRI_HoldDays = DRI_SellDayNo - DRI_BuyDayNo
                            DRI_SellClosePrice = FoundRow4(0)("Close_Price")
                            DRI_TradeFactor = DRI_SellClosePrice / DRI_BuyClosePrice
                            DRI_ProfitPct = (DRI_SellClosePrice - DRI_BuyClosePrice) / DRI_BuyClosePrice * 100
                            DRI_ProfitAnnPct = DRI_ProfitPct * 365.25 / (DRI_SellDayNo - DRI_BuyDayNo)
                            DRI_NetProfitPct = DRI_ProfitPct - DRI_BrokPct
                            DRI_NetProfitAnnPct = DRI_NetProfitPct * 365.25 / (DRI_SellDayNo - DRI_BuyDayNo)

                            'Write the data to the table:
                            Try
                                InsertCommand = "INSERT INTO Dividend_Rebound_Info (ASX_Code, ExDivDate, ExDivDayNo, ExDivTradeDay, BuyDelayDays, BuyDelayTradeDays, BuyDate, BuyDayNo, BuyTradeDay, BuyClosePrice, HoldDays, HoldTradeDays, SellDate, SellDayNo, SellTradeDay, SellClosePrice, TradeFactor, ProfitPct, ProfitAnnPct, BrokPct, NetProfitPct, NetProfitAnnPct)" _
                                & " VALUES ('" & DRI_AsxCode & "', " & DRIExDivDateStr & ", " & DRI_ExDivDayNo & ", " & DRI_ExDivTradeDay & ", " & DRI_BuyDelayDays & ", " & DRI_BuyDelayTradeDays & ", " & BuyDateStr & ", " & DRI_BuyDayNo _
                                & ", " & DRI_BuyTradeDay & ", " & DRI_BuyClosePrice & ", " & DRI_HoldDays & ", " & DRI_HoldTradeDays & ", " & SellDateStr & ", " & DRI_SellDayNo & ", " & DRI_SellTradeDay & ", " & DRI_SellClosePrice & ", " & DRI_TradeFactor & ", " & DRI_ProfitPct & ", " & DRI_ProfitAnnPct & ", " & DRI_BrokPct & ", " & DRI_NetProfitPct & ", " & DRI_NetProfitAnnPct & ")"

                                daCalc.InsertCommand = New OleDb.OleDbCommand(InsertCommand, myCalcsConnection)
                                daCalc.InsertCommand.ExecuteNonQuery()

                            Catch ex As Exception
                                Message.AddWarning("Error writing to Dividend_Payment_Info table: " & ex.Message & vbCrLf)
                                Message.Add("DivNo = " & DivNo & "  BuyDelay = " & BuyDelay & "  HoldDays = " & HoldDays & vbCrLf)
                                Message.Add("Insert command:" & vbCrLf & InsertCommand & vbCrLf)
                            End Try
                        Else
                            Message.AddWarning(FoundRow4.Count & " ClosePrices records found on sell TradeDayNo = " & DRI_SellTradeDay & vbCrLf)

                        End If
                    Next
                Else
                    Message.AddWarning(FoundRow3.Count & " ClosePrices records found on buy TradeDayNo = " & DRI_BuyTradeDay & vbCrLf)
                    'These values cannot be calculated - default to zero or ExDivDate:
                    DRI_BuyDate = DRI_ExDivDate
                    DRI_BuyDayNo = 0
                    DRI_BuyClosePrice = 0
                End If
            Next
        Next

    End Sub

    Public Sub CalcDividendPaymentInfo(ByVal AsxCode As String)
        'Calculate the dividend payment information for the company with the code AsxCode.

        'Get the Close_Prices for the company:

        'Connection to Share Price database:
        Dim SharePriceConnString As String
        Dim mySharePriceConnection As OleDb.OleDbConnection = New OleDb.OleDbConnection
        Dim SharePriceQuery As String


        Try
            SharePriceConnString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source =" & SharePriceDbPath 'DatabasePath
            mySharePriceConnection.ConnectionString = SharePriceConnString
            mySharePriceConnection.Open()
        Catch ex As Exception
            Message.AddWarning("Error connecting to Share Price database: " & ex.Message & vbCrLf)
            Exit Sub
        End Try

        Dim ds As DataSet = New DataSet
        Dim da As OleDb.OleDbDataAdapter

        SharePriceQuery = "Select Trade_Date, Close_Price From ASX_Share_Prices Where ASX_Code = '" & AsxCode & "' Order by Trade_Date"

        da = New OleDb.OleDbDataAdapter(SharePriceQuery, mySharePriceConnection)
        da.MissingSchemaAction = MissingSchemaAction.AddWithKey

        ds.Clear()

        'Open the Calculations database to write the DividendPaymentInfo rows:
        Dim CalcsConnString As String
        Dim myCalcsConnection As OleDb.OleDbConnection = New OleDb.OleDbConnection
        Dim InsertCommand As String
        Dim CalcQuery As String 'Contains the query used to select data from the Calculations database.

        Try
            CalcsConnString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source =" & CalculationsDbPath 'DatabasePath
            myCalcsConnection.ConnectionString = CalcsConnString
            myCalcsConnection.Open()
        Catch ex As Exception
            Message.AddWarning("Error connecting to Calculations database: " & ex.Message & vbCrLf)
            Exit Sub
        End Try
        Dim daCalc As New OleDb.OleDbDataAdapter

        Try
            da.Fill(ds, "ClosePrices")
            Dim NRows As Integer = ds.Tables(0).Rows.Count

            Message.Add(vbCrLf & NRows & " closing prices read for ASX code = " & AsxCode & vbCrLf)

            'Add the DayNo and TradeDayNo columns to the ClosePrices table: --------------------------------------
            Dim dc1 As DataColumn = New DataColumn("DayNo", System.Type.GetType("System.Int32"))
            Dim dc2 As DataColumn = New DataColumn("TradeDayNo", System.Type.GetType("System.Int32"))

            ds.Tables("ClosePrices").Columns.Add(dc1)
            ds.Tables("ClosePrices").Columns.Add(dc2)

            'Generate the DayNo and TradeDayNo values: -----------------------------------------------------------
            Dim TradeDate As Date
            Dim I As Integer
            For I = 0 To NRows - 1
                TradeDate = ds.Tables("ClosePrices").Rows.Item(I)("Trade_Date")
                ds.Tables("ClosePrices").Rows.Item(I)("DayNo") = Int(TradeDate.ToOADate)
                ds.Tables("ClosePrices").Rows.Item(I)("TradeDayNo") = I
                If I Mod 20 = 0 Then
                    Message.Add("DayNo = " & Int(TradeDate.ToOADate) & " TradeDayNo = " & I & " Close Price = " & ds.Tables("ClosePrices").Rows.Item(I)("Close_Price") & vbCrLf)
                End If
            Next

            Message.Add(vbCrLf & "DayNo and TradeDayNo values added to ClosePrices table. " & vbCrLf & vbCrLf)

            'Open Dividends records for AsxCode:
            SharePriceQuery = "Select * From Dividends Where ASX_Code = '" & AsxCode & "' Order By Ex_Div"
            da.SelectCommand = New OleDb.OleDbCommand(SharePriceQuery, mySharePriceConnection)
            da.Fill(ds, "Dividends") 'ASX_Code, Ex_Div, Amount, Rec_Date, Payable, Franking, Credit

            Message.Add(ds.Tables("Dividends").Rows.Count & " dividends loaded for company: " & AsxCode & vbCrLf)

            'Compile Dividend Payment Info ---------------------------------------------------------------
            'Dividend Payment Info columns:
            Dim DPI_AsxCode As String 'The ASX Code to be written to the Dividend Payment Information table
            Dim DPI_ExDivDate As Date 'The ex-dividend date
            Dim DPI_ExDivDayNo As Integer 'The ex-dividend day number
            Dim DPI_ExDivTradeDay As Integer 'The ex-dividend trade day number
            Dim DPI_PreDivPrice As Single 'The price of the share pre-dividend
            Dim DPI_ExDivPrice As Single 'The price of the share post-dividend
            Dim DPI_DivAmount As Single 'The amount of the dividend per share.
            Dim DPI_FrCredit As Single 'The franking credit per share
            Dim DPI_DivPct As Single 'The dividend as a percent of the pre-div share price
            Dim DPI_GrossDiv As Single 'The Dividend amount plus the franking credit
            Dim DPI_GrossDivPct As Single 'The gross dividend as a percent of the pre-div share price 
            Dim DPI_DropAmt As Single 'The amount the share price dropped ex-dividend
            Dim DPI_DropPct As Single 'The amount the share price dropped ex-dividend as a percent of the pre-div price.

            Dim ExDivDateStr As String 'Date string for queries

            'Process each row in the Dividends table
            For I = 0 To ds.Tables("Dividends").Rows.Count
                DPI_AsxCode = AsxCode
                DPI_ExDivDate = ds.Tables("Dividends").Rows(I)("Ex_Div")
                ExDivDateStr = "#" & Format(DPI_ExDivDate, "MM-dd-yyyy") & "#"
                DPI_DivAmount = ds.Tables("Dividends").Rows(I)("Amount")
                DPI_FrCredit = ds.Tables("Dividends").Rows(I)("Credit")
                DPI_GrossDiv = DPI_DivAmount + DPI_FrCredit 'The gross dividend = dividend amount + franking credit.

                Dim FoundRow() As DataRow = ds.Tables("ClosePrices").Select("Trade_Date = " & ExDivDateStr) 'Trade_Date, Close_Price, DayNo, TradeDayNo
                If FoundRow.Count = 0 Then
                    Message.AddWarning("No ClosePrices record found on date = " & ExDivDateStr & vbCrLf)
                    'The following values cannot be calculated - default to zero:
                    DPI_ExDivDayNo = 0
                    DPI_ExDivTradeDay = 0
                    DPI_ExDivPrice = 0
                    DPI_DropAmt = 0
                    DPI_DivPct = 0
                ElseIf FoundRow.Count = 1 Then
                    DPI_ExDivDayNo = FoundRow(0)("DayNo")
                    DPI_ExDivTradeDay = FoundRow(0)("TradeDayNo")
                    DPI_ExDivPrice = FoundRow(0)("Close_Price")

                    'Now get the pre-dividend closing price:
                    '  Get the ClosePrices record one trading day before the ex-dividend date:
                    Dim FoundRow2() As DataRow = ds.Tables("ClosePrices").Select("TradeDayNo = " & DPI_ExDivTradeDay - 1) 'Trade_Date, Close_Price, DayNo, TradeDayNo
                    If FoundRow2.Count = 0 Then
                        Message.AddWarning("No ClosePrices record found on TradeDayNo = " & DPI_ExDivTradeDay - 1 & vbCrLf)
                        'These values cannot be calculated - default to zero:
                        DPI_DropAmt = 0
                        DPI_DivPct = 0
                        DPI_GrossDivPct = 0
                        DPI_DropAmt = 0
                        DPI_DropPct = 0
                    ElseIf FoundRow2.Count = 1 Then
                        DPI_PreDivPrice = FoundRow2(0)("Close_Price")
                        DPI_DropAmt = DPI_PreDivPrice - DPI_ExDivPrice
                        DPI_DivPct = DPI_DivAmount / DPI_PreDivPrice * 100
                        DPI_GrossDivPct = DPI_GrossDiv / DPI_PreDivPrice * 100
                        DPI_DropAmt = DPI_PreDivPrice - DPI_ExDivPrice
                        DPI_DropPct = DPI_DropAmt / DPI_PreDivPrice * 100
                    Else
                        Message.AddWarning(FoundRow2.Count & " ClosePrices records found on TradeDayNo = " & DPI_ExDivTradeDay - 1 & vbCrLf)
                        'These values cannot be calculated - default to zero:
                        DPI_DropAmt = 0
                        DPI_DivPct = 0
                        DPI_GrossDivPct = 0
                        DPI_DropAmt = 0
                        DPI_DropPct = 0
                    End If
                Else
                    Message.AddWarning(FoundRow.Count & " ClosePrices records found on date = " & ExDivDateStr & vbCrLf)
                    'The following values cannot be calculated - default to zero:
                    DPI_ExDivDayNo = 0
                    DPI_ExDivTradeDay = 0
                    DPI_ExDivPrice = 0
                    DPI_DropAmt = 0
                    DPI_DivPct = 0
                End If

                'Write the row to the DividendPaymentInfo table: ---------------------------------------------------
                If DPI_ExDivDayNo = 0 Then
                    'Valid data not generated - do not write to table
                Else
                    Try
                        InsertCommand = "INSERT INTO Dividend_Payment_Info (ASX_Code, ExDivDate, ExDivDayNo, ExDivTradeDay, PreDivPrice, ExDivPrice, DivAmount, FrCredit, DivPct, GrossDiv, GrossDivPct, DropAmt, DropPct)" _
                        & " VALUES ('" & DPI_AsxCode & "', " & ExDivDateStr & ", " & DPI_ExDivDayNo & ", " & DPI_ExDivTradeDay & ", " & DPI_PreDivPrice & ", " & DPI_ExDivPrice & ", " & DPI_DivAmount & ", " & DPI_FrCredit _
                        & ", " & DPI_DivPct & ", " & DPI_GrossDiv & ", " & DPI_GrossDivPct & ", " & DPI_DropAmt & ", " & DPI_DropPct & ")"

                        daCalc.InsertCommand = New OleDb.OleDbCommand(InsertCommand, myCalcsConnection)
                        daCalc.InsertCommand.ExecuteNonQuery()

                    Catch ex As Exception
                        Message.AddWarning("Error writing to Dividend_Payment_Info table: " & ex.Message & vbCrLf)
                        Message.Add("Insert command:" & vbCrLf & InsertCommand & vbCrLf)
                    End Try
                End If
            Next

            'Close the databases:
            daCalc.InsertCommand.Connection.Close()

        Catch ex As Exception
            Message.AddWarning("Error processing dividend payment information: " & ex.Message & vbCrLf)
            'NOTE: If error is: No value given for one or more required parameters.
            '      Then check for a spelling error in the query.
        End Try

    End Sub

    'Public Sub EventResponseAnalysis(ByVal AsxCode As String, ByVal EventQuery As String, ByVal MinBuyDelay As Integer, ByVal MaxBuyDelay As Integer, ByVal MinSellDelay As Integer, ByVal MaxSellDelay As Integer, ByVal BuyBrokPct As Single, ByVal SellBrokPct As Single)
    Public Sub EventResponseAnalysis(ByVal AsxCode As String, ByVal EventType As String, ByVal EventID As String, ByVal EventInfoTable As String, ByVal EventDateCol As String, ByVal EventDayNoCol As String, ByVal EventTradeDayNoCol As String, ByVal MinBuyDelay As Integer, ByVal MaxBuyDelay As Integer, ByVal MinSellDelay As Integer, ByVal MaxSellDelay As Integer, ByVal BuyBrokPct As Single, ByVal SellBrokPct As Single)
        'Generate the share price response of the AsxCode company to a set of events.

        'The Event parameters are listed in the EventTable.
        'The EventType is the type of event being analysed. (Dividend, Rights Issue, Momentum, Gap, Calendar, Pattern, ...) (Each event type has its own table of parameters.)
        'The EventID is an ID string used to identify different events of the same type.
        'The EventInfoTable is the name of the table containing the Event information.
        'The EventDateCol is the name of the column containing the event dates.
        'The EventDayNoCol is the name of the column containing the sequential day number of each event.
        'The EventTradingDayNoCol is the name of the column containing the sequential trading day number of each event.

        'The share trading profit is calculated for each buy date and sell date.
        'The MinBuyDelay is the minimum delay in trading days after the event to buy shares in the AsxCode company.
        'The MaxBuyDelay is the maximum delay in trading days to buy shares.
        'The MinSellDelay is the minimum holding time in trading days before selling the shares.
        'The MaxSellDelay is the maximum holding time in trading days before selling the shares.

        'The BuyBrokPct is the brokerage paid when buying the shares as a percentage of the price.
        'The SellBrokPct is the brokerage paid when selling the shares as a percentage of the price.

        'All combinations of buy delay and holding time are used to calculate the share trading profit following the event.
        'The share trading profit records are saved in the Event_Response_Info table in the Calculations database. NOTE: Add the Event_Type field to the table.

        'Get the Close_Prices for the company:
        'Connection to Share Price database:
        Dim SharePriceConnString As String
        Dim mySharePriceConnection As OleDb.OleDbConnection = New OleDb.OleDbConnection
        Dim SharePriceQuery As String

        Try
            SharePriceConnString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source =" & SharePriceDbPath 'DatabasePath
            mySharePriceConnection.ConnectionString = SharePriceConnString
            mySharePriceConnection.Open()
        Catch ex As Exception
            Message.AddWarning("Error connecting to Share Price database: " & ex.Message & vbCrLf)
            Exit Sub
        End Try

        Dim ds As DataSet = New DataSet
        Dim da As OleDb.OleDbDataAdapter

        SharePriceQuery = "Select Trade_Date, Close_Price From ASX_Share_Prices Where ASX_Code = '" & AsxCode & "' Order by Trade_Date"

        da = New OleDb.OleDbDataAdapter(SharePriceQuery, mySharePriceConnection)
        da.MissingSchemaAction = MissingSchemaAction.AddWithKey

        Try
            da.Fill(ds, "ClosePrices")
            Dim NRows As Integer = ds.Tables(0).Rows.Count
            Message.Add(vbCrLf & NRows & " closing prices read for ASX code = " & AsxCode & vbCrLf)

            'Add the DayNo and TradeDayNo columns to the ClosePrices table: --------------------------------------
            Dim dc1 As DataColumn = New DataColumn("DayNo", System.Type.GetType("System.Int32"))
            Dim dc2 As DataColumn = New DataColumn("TradeDayNo", System.Type.GetType("System.Int32"))

            ds.Tables("ClosePrices").Columns.Add(dc1)
            ds.Tables("ClosePrices").Columns.Add(dc2)

            'Generate the DayNo and TradeDayNo values: -----------------------------------------------------------
            Dim TradeDate As Date
            Dim I As Integer
            For I = 0 To NRows - 1
                TradeDate = ds.Tables("ClosePrices").Rows.Item(I)("Trade_Date")
                ds.Tables("ClosePrices").Rows.Item(I)("DayNo") = Int(TradeDate.ToOADate)
                ds.Tables("ClosePrices").Rows.Item(I)("TradeDayNo") = I
                If I Mod 20 = 0 Then
                    Message.Add("DayNo = " & Int(TradeDate.ToOADate) & " TradeDayNo = " & I & " Close Price = " & ds.Tables("ClosePrices").Rows.Item(I)("Close_Price") & vbCrLf)
                End If
            Next

            Message.Add(vbCrLf & "DayNo and TradeDayNo values added to ClosePrices table. " & vbCrLf & vbCrLf)

        Catch ex As Exception
            Message.AddWarning("Error reading Close_Prices for company: " & AsxCode & "  " & ex.Message & vbCrLf)
            Exit Sub
        End Try

        'Compile Dividend Rebound Info ---------------------------------------------------------------
        'Dividend Payment Info columns:
        '  ASX_Code
        '  ExDivDate
        '  ExDivDayNo 
        '  ExDivTradeDay
        '  BuyDelayDays
        '  BuyDelayTradeDays
        '  BuyDate
        '  BuyDayNo
        '  BuyTradeDay
        '  BuyClosePrice
        '  HoldDays
        '  HoldTradeDays
        '  SellDate
        '  SellDayNo
        '  SellTradeDay
        '  SellClosePrice
        '  TradeFactor
        '  ProfitPct
        '  ProfitAnnPct
        '  BrokPct
        '  NetProfitPct
        '  NetProfitAnnPct

        'DRI was the abbreviation for Dividend Response Info
        'Change this to ERA: Event Response Analysis
        Dim ERA_AsxCode As String 'The ASX Code to be written to the Dividend Rebounf Information table
        'Dim ERA_ExDivDate As Date        'CHANGE TO EventDate
        Dim ERA_EventDate As Date
        'Dim ERA_ExDivDayNo As Integer    'CHANGE TO EventDayNo
        Dim ERA_EventDayNo As Integer
        'Dim ERA_ExDivTradeDay As Integer 'CHANGE TO EventTradeDayNo
        Dim ERA_EventTradeDayNo As Integer
        Dim ERA_BuyDelayDays As Integer
        Dim ERA_BuyDelayTradeDays As Integer
        Dim ERA_BuyDate As Date
        Dim ERA_BuyDayNo As Integer
        Dim ERA_BuyTradeDay As Integer
        Dim ERA_BuyClosePrice As Single
        Dim ERA_HoldDays As Integer
        Dim ERA_HoldTradeDays As Integer
        Dim ERA_SellDate As Date
        Dim ERA_SellDayNo As Integer
        Dim ERA_SellTradeDay As Integer
        Dim ERA_SellClosePrice As Single
        Dim ERA_TradeFactor As Single
        Dim ERA_ProfitPct As Single
        Dim ERA_ProfitAnnPct As Single
        Dim ERA_BrokPct As Single = BuyBrokPct + SellBrokPct
        Dim ERA_NetProfitPct As Single
        Dim ERA_NetProfitAnnPct As Single

        'Dim ERAExDivDateStr As String 'Date string for writing to table
        Dim ERAEventDateStr As String 'Date string for writing to table
        Dim BuyDateStr As String 'Date string for writing to table
        Dim SellDateStr As String 'Date string for writing to table

        ERA_AsxCode = AsxCode

        'Dataset ds contains the tables: ClosePrices (used here) and Dividends (not used here)
        'daCalc is the data adaptor connected to the Calculations database.

        'Open the Calculations database to write the DividendPaymentInfo rows:
        Dim CalcsConnString As String
        Dim myCalcsConnection As OleDb.OleDbConnection = New OleDb.OleDbConnection
        Dim InsertCommand As String
        Dim CalcQuery As String 'Contains the query used to select data from the Calculations database.

        Try
            CalcsConnString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source =" & CalculationsDbPath 'DatabasePath
            myCalcsConnection.ConnectionString = CalcsConnString
            myCalcsConnection.Open()
        Catch ex As Exception
            Message.AddWarning("Error connecting to Calculations database: " & ex.Message & vbCrLf)
            Exit Sub
        End Try
        Dim daCalc As New OleDb.OleDbDataAdapter
        'Dim ds As DataSet = New DataSet

        ''Read Dividend_Payment_Info into DivInfo in daCalc:

        'Read Event Information into EventInfo in daCalc: (*** NOTE update queries when other Event Types and EventIDs are implemented. ***)
        'CalcQuery = "Select * From Dividend_Payment_Info Where ASX_Code = '" & ERA_AsxCode & "' Order By ExDivDate"
        CalcQuery = "Select * From " & EventInfoTable & "Where ASX_Code = '" & ERA_AsxCode & "' Order By " & EventDateCol
        daCalc.SelectCommand = New OleDb.OleDbCommand(CalcQuery, myCalcsConnection)
        'daCalc.Fill(ds, "DividendInfo") 'ASX_Code, ExDivDate, ExDivDayNo, ExDivTradeDay, PreDivPrice, ExDivPrice, DivAmount, FrCredit, DivPct, GrossDiv, GrossDivPct, DropAmt, DropPct
        daCalc.Fill(ds, "EventInfo") 'ASX_Code, ExDivDate, ExDivDayNo, ExDivTradeDay, PreDivPrice, ExDivPrice, DivAmount, FrCredit, DivPct, GrossDiv, GrossDivPct, DropAmt, DropPct

        'Message.Add(ds.Tables("DividendInfo").Rows.Count & " dividend payment information loaded for company: " & AsxCode & vbCrLf)
        Message.Add(ds.Tables("EventInfo").Rows.Count & " event information records loaded for company: " & AsxCode & vbCrLf)

        'Dim DivNo As Integer
        Dim EventNo As Integer
        Dim BuyDelay As Integer
        Dim HoldDays As Integer

        'Process each dividend payment:
        'For DivNo = 0 To ds.Tables("DividendInfo").Rows.Count - 1
        For EventNo = 0 To ds.Tables("EventInfo").Rows.Count - 1
            'ERA_ExDivDate = ds.Tables("DividendInfo").Rows.Item(DivNo)("ExDivDate")
            ERA_EventDate = ds.Tables("EventInfo").Rows.Item(EventNo)(EventDateCol)
            ERAEventDateStr = "#" & Format(ERA_EventDate, "MM-dd-yyyy") & "#"
            'ERA_ExDivDayNo = ds.Tables("DividendInfo").Rows.Item(DivNo)("ExDivDayNo")
            ERA_EventDayNo = ds.Tables("EventInfo").Rows.Item(EventNo)(EventDayNoCol)
            'DRI_ExDivTradeDay = ds.Tables("DividendInfo").Rows.Item(DivNo)("ExTradeDay")
            'ERA_ExDivTradeDay = ds.Tables("DividendInfo").Rows.Item(DivNo)("ExDivTradeDay")
            ERA_EventTradeDayNo = ds.Tables("EventInfo").Rows.Item(EventNo)(EventTradeDayNoCol)
            'Process each buy delay:
            For BuyDelay = MinBuyDelay To MaxBuyDelay
                ERA_BuyDelayTradeDays = BuyDelay
                ERA_BuyTradeDay = ERA_EventTradeDayNo + BuyDelay
                'Read the ClosingPrice record on the Buy Day:
                Dim FoundRow3() As DataRow = ds.Tables("ClosePrices").Select("TradeDayNo = " & ERA_BuyTradeDay) 'Trade_Date, Close_Price, DayNo, TradeDayNo
                If FoundRow3.Count = 0 Then
                    Message.AddWarning("No ClosePrices record found on buy TradeDayNo = " & ERA_BuyTradeDay & vbCrLf)
                    'These values cannot be calculated - default to zero or ExDivDate:
                    'ERA_BuyDate = ERA_ExDivDate
                    ERA_BuyDate = ERA_EventDate
                    ERA_BuyDayNo = 0
                    ERA_BuyClosePrice = 0
                ElseIf FoundRow3.Count = 1 Then
                    ERA_BuyDate = FoundRow3(0)("Trade_Date")
                    BuyDateStr = "#" & Format(ERA_BuyDate, "MM-dd-yyyy") & "#"
                    ERA_BuyDayNo = FoundRow3(0)("DayNo")
                    'ERA_BuyDelayDays = ERA_BuyDayNo - ERA_ExDivDayNo
                    ERA_BuyDelayDays = ERA_BuyDayNo - ERA_EventDayNo
                    ERA_BuyClosePrice = FoundRow3(0)("Close_Price")
                    'Process each sell delay:
                    For HoldDays = MinSellDelay To MaxSellDelay
                        ERA_HoldTradeDays = HoldDays
                        ERA_SellTradeDay = ERA_BuyTradeDay + HoldDays
                        Dim FoundRow4() As DataRow = ds.Tables("ClosePrices").Select("TradeDayNo = " & ERA_SellTradeDay) 'Trade_Date, Close_Price, DayNo, TradeDayNo
                        If FoundRow4.Count = 0 Then
                            Message.AddWarning("No ClosePrices record found on sell TradeDayNo = " & ERA_SellTradeDay & vbCrLf)

                        ElseIf FoundRow4.Count = 1 Then
                            ERA_SellDate = FoundRow4(0)("Trade_Date")
                            SellDateStr = "#" & Format(ERA_SellDate, "MM-dd-yyyy") & "#"
                            ERA_SellDayNo = FoundRow4(0)("DayNo")
                            ERA_HoldDays = ERA_SellDayNo - ERA_BuyDayNo
                            ERA_SellClosePrice = FoundRow4(0)("Close_Price")
                            ERA_TradeFactor = ERA_SellClosePrice / ERA_BuyClosePrice
                            ERA_ProfitPct = (ERA_SellClosePrice - ERA_BuyClosePrice) / ERA_BuyClosePrice * 100
                            ERA_ProfitAnnPct = ERA_ProfitPct * 365.25 / (ERA_SellDayNo - ERA_BuyDayNo)
                            ERA_NetProfitPct = ERA_ProfitPct - ERA_BrokPct
                            ERA_NetProfitAnnPct = ERA_NetProfitPct * 365.25 / (ERA_SellDayNo - ERA_BuyDayNo)

                            'Write the data to the table:
                            Try
                                'ADD EventType and EventID ******************************************
                                'InsertCommand = "INSERT INTO Dividend_Rebound_Info (ASX_Code, ExDivDate, ExDivDayNo, ExDivTradeDay, BuyDelayDays, BuyDelayTradeDays, BuyDate, BuyDayNo, BuyTradeDay, BuyClosePrice, HoldDays, HoldTradeDays, SellDate, SellDayNo, SellTradeDay, SellClosePrice, TradeFactor, ProfitPct, ProfitAnnPct, BrokPct, NetProfitPct, NetProfitAnnPct)" _
                                InsertCommand = "INSERT INTO Event_Response_Info (ASX_Code, ExDivDate, ExDivDayNo, ExDivTradeDay, BuyDelayDays, BuyDelayTradeDays, BuyDate, BuyDayNo, BuyTradeDay, BuyClosePrice, HoldDays, HoldTradeDays, SellDate, SellDayNo, SellTradeDay, SellClosePrice, TradeFactor, ProfitPct, ProfitAnnPct, BrokPct, NetProfitPct, NetProfitAnnPct)" _
                                & " VALUES ('" & ERA_AsxCode & "', " & ERAEventDateStr & ", " & ERA_EventDayNo & ", " & ERA_EventTradeDayNo & ", " & ERA_BuyDelayDays & ", " & ERA_BuyDelayTradeDays & ", " & BuyDateStr & ", " & ERA_BuyDayNo _
                                & ", " & ERA_BuyTradeDay & ", " & ERA_BuyClosePrice & ", " & ERA_HoldDays & ", " & ERA_HoldTradeDays & ", " & SellDateStr & ", " & ERA_SellDayNo & ", " & ERA_SellTradeDay & ", " & ERA_SellClosePrice & ", " & ERA_TradeFactor & ", " & ERA_ProfitPct & ", " & ERA_ProfitAnnPct & ", " & ERA_BrokPct & ", " & ERA_NetProfitPct & ", " & ERA_NetProfitAnnPct & ")"
                                '  & " VALUES ('" & ERA_AsxCode & "', " & ERAExDivDateStr & ", " & ERA_ExDivDayNo & ", " & ERA_ExDivTradeDay & ", " & ERA_BuyDelayDays & ", " & ERA_BuyDelayTradeDays & ", " & BuyDateStr & ", " & ERA_BuyDayNo _

                                daCalc.InsertCommand = New OleDb.OleDbCommand(InsertCommand, myCalcsConnection)
                                daCalc.InsertCommand.ExecuteNonQuery()

                            Catch ex As Exception
                                Message.AddWarning("Error writing to Dividend_Payment_Info table: " & ex.Message & vbCrLf)
                                'Message.Add("DivNo = " & DivNo & "  BuyDelay = " & BuyDelay & "  HoldDays = " & HoldDays & vbCrLf)
                                Message.Add("EventNo = " & EventNo & "  BuyDelay = " & BuyDelay & "  HoldDays = " & HoldDays & vbCrLf)
                                Message.Add("Insert command:" & vbCrLf & InsertCommand & vbCrLf)
                            End Try
                        Else
                            Message.AddWarning(FoundRow4.Count & " ClosePrices records found on sell TradeDayNo = " & ERA_SellTradeDay & vbCrLf)

                        End If
                    Next
                Else
                    Message.AddWarning(FoundRow3.Count & " ClosePrices records found on buy TradeDayNo = " & ERA_BuyTradeDay & vbCrLf)
                    'These values cannot be calculated - default to zero or ExDivDate:
                    'ERA_BuyDate = ERA_ExDivDate
                    ERA_BuyDate = ERA_EventDate
                    ERA_BuyDayNo = 0
                    ERA_BuyClosePrice = 0
                End If
            Next
        Next


    End Sub

    Function GetSharePriceValue(ByVal AsxCode As String, ByVal DateString As String, ByVal PriceType As String) As Single
        'Return a share price value at the specified date.

        'Dim DateValue As Date = CDate(DateString)
        Dim DateValue As Date

        Try
            DateValue = CDate(DateString)
        Catch ex As Exception
            Message.AddWarning("Date error in GetSharePriceValue: " & ex.Message & vbCrLf)
            Return 0
            Exit Function
        End Try

        If SharePriceDbPath = "" Then
            Message.AddWarning("A share price database has not been selected!" & vbCrLf)
            Return -1
            Exit Function
        End If

        If System.IO.File.Exists(SharePriceDbPath) Then
            'Share Price Database file exists.
        Else
            'Share Price Database file does not exist!
            Message.AddWarning("The share price database was not found: " & SharePriceDbPath & vbCrLf)
            Return -1
            Exit Function
        End If

        'Connection to Share Price database:
        Dim SharePriceConnString As String
        Dim mySharePriceConnection As OleDb.OleDbConnection = New OleDb.OleDbConnection
        Dim SharePriceQuery As String

        Try
            SharePriceConnString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source =" & SharePriceDbPath 'DatabasePath
            mySharePriceConnection.ConnectionString = SharePriceConnString
            mySharePriceConnection.Open()
        Catch ex As Exception
            Message.AddWarning("Error connecting to Share Price database: " & ex.Message & vbCrLf)
            Return -1
            Exit Function
        End Try

        Dim ds As DataSet = New DataSet
        Dim da As OleDb.OleDbDataAdapter

        Dim DateStr As String = "#" & Format(DateValue, "MM-dd-yyyy") & "#"

        SharePriceQuery = "Select * From ASX_Share_Prices Where ASX_Code = '" & AsxCode & "' And Trade_Date = " & DateStr

        Message.Add("Share price query is: " & SharePriceQuery & vbCrLf)

        da = New OleDb.OleDbDataAdapter(SharePriceQuery, mySharePriceConnection)
        da.MissingSchemaAction = MissingSchemaAction.AddWithKey

        ds.Clear()

        Try
            da.Fill(ds, "myData")
            Dim NRows As Integer = ds.Tables(0).Rows.Count
            If NRows = 0 Then
                Message.AddWarning("There are no records for company: " & AsxCode & " on the date: " & DateString & vbCrLf)
                Return -1
            ElseIf NRows = 1 Then
                Return ds.Tables(0).Rows(0).Item(PriceType)
            Else
                Message.AddWarning("More than one record was returned: " & NRows & " records. The first record value is returned." & vbCrLf)
                Return ds.Tables(0).Rows(0).Item(PriceType)
            End If
        Catch ex As Exception
            Message.AddWarning("Error getting share price value: " & ex.Message & vbCrLf)
        End Try
    End Function

    'Public Sub ShowFirstSharePriceValue(ByVal AsxCode As String, ByVal DateString As String, ByVal NDays As Integer, ByRef DateValue As Date, ByVal DateFormName As String, ByVal DateItemName As String, ByVal PriceType As String, ByRef PriceValue As Single, ByVal PriceFormName As String, ByVal PriceItemName As String, ByRef Status As String, ByVal StatusFormName As String, ByVal StatusItemName As String)
    Public Sub ShowFirstSharePriceValue(ByVal AsxCode As String, ByVal DateString As String, ByVal NDays As Integer, ByVal PriceType As String, ByVal PriceFormName As String, ByVal PriceItemName As String, ByVal DateFormName As String, ByVal DateItemName As String, ByVal StatusFormName As String, ByVal StatusItemName As String)

        'Get the first available Share price in the time window defined by the DateString and NDays.
        'Return the price date and status in the PriceValue, DateValue and Status variables.
        'The price, date and status values are then sent to the specified web page locations (PriceFormName, PriceItemName etc).

        Dim PriceValue As Single 'Stores the value of the first share price found in the date window
        Dim DateValue As Date    'Stores the date of the first share price found in the date window
        Dim Status As String     'Stores the status of the search for the first share price

        GetFirstSharePriceValue(AsxCode, DateString, NDays, DateValue, PriceType, PriceValue, Status)

        If Status = "OK" Then
            RestoreSetting(PriceFormName, PriceItemName, PriceValue)                         'Show the first share price found on the web page
            RestoreSetting(DateFormName, DateItemName, Format(DateValue, "dd MMMM yyyy"))    'Show the date of the first share price found on the web page.
            RestoreSetting(StatusFormName, StatusItemName, Status)                           'Show the Status string.
        Else 'Error finding the first share price value!
            RestoreSetting(PriceFormName, PriceItemName, "")       'Show a blank first share price found on the web page
            RestoreSetting(DateFormName, DateItemName, "")         'Show a blank date of the first share price found on the web page.
            RestoreSetting(StatusFormName, StatusItemName, Status) 'Show the Status string.
        End If

    End Sub

    Public Sub GetFirstSharePriceValue(ByVal AsxCode As String, ByVal DateString As String, ByVal NDays As Integer, ByRef DateValue As Date, ByVal PriceType As String, ByRef PriceValue As Single, ByRef Status As String)
        'Get the first available Share price in the time window defined by the DateString and NDays.
        'Return the price, date and status in the PriceValue, DateValue and Status variables.

        Try
            DateValue = CDate(DateString)
        Catch ex As Exception
            Message.AddWarning("Date error in GetSharePriceValue: " & ex.Message & vbCrLf)
            Status = "Error: Date String"
            Exit Sub
        End Try

        If SharePriceDbPath = "" Then
            Message.AddWarning("A share price database has not been selected!" & vbCrLf)
            Status = "Error: No Database"
            Exit Sub
        End If

        If System.IO.File.Exists(SharePriceDbPath) Then
            'Share Price Database file exists.
        Else
            'Share Price Database file does not exist!
            Message.AddWarning("The share price database was not found: " & SharePriceDbPath & vbCrLf)
            Status = "Error: Database Not Found"
            Exit Sub
        End If

        'Connection to Share Price database:
        Dim SharePriceConnString As String
        Dim mySharePriceConnection As OleDb.OleDbConnection = New OleDb.OleDbConnection
        Dim SharePriceQuery As String

        Try
            SharePriceConnString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source =" & SharePriceDbPath 'DatabasePath
            mySharePriceConnection.ConnectionString = SharePriceConnString
            mySharePriceConnection.Open()
        Catch ex As Exception
            Message.AddWarning("Error connecting to Share Price database: " & ex.Message & vbCrLf)
            Status = "Error: Database Connection"
            Exit Sub
        End Try

        Dim ds As DataSet = New DataSet
        Dim da As OleDb.OleDbDataAdapter

        Dim StartDateStr As String = "#" & Format(DateValue, "MM-dd-yyyy") & "#"
        Dim EndDateValue As Date = DateValue.AddDays(NDays)
        Dim EndDateStr As String = "#" & Format(EndDateValue, "MM-dd-yyyy") & "#"

        'SharePriceQuery = "Select * From ASX_Share_Prices Where ASX_Code = '" & AsxCode & "' And Trade_Date = " & StartDateStr
        SharePriceQuery = "Select * From ASX_Share_Prices Where ASX_Code = '" & AsxCode & "' And Trade_Date Between " & StartDateStr & " And " & EndDateStr & " Order By Trade_Date"

        'Message.Add("Share price query is: " & SharePriceQuery & vbCrLf)

        da = New OleDb.OleDbDataAdapter(SharePriceQuery, mySharePriceConnection)
        da.MissingSchemaAction = MissingSchemaAction.AddWithKey

        ds.Clear()

        Try
            da.Fill(ds, "myData")
            Dim NRows As Integer = ds.Tables(0).Rows.Count
            If NRows = 0 Then
                'Message.AddWarning("There are no records for company: " & AsxCode & " on the date: " & DateString & vbCrLf)
                Message.AddWarning("There are no records for company: " & AsxCode & " between dates: " & Format(DateValue, "dd MMMM yyyy") & " and " & Format(EndDateValue, "dd MMMM yyyy") & vbCrLf)
                PriceValue = 0
                DateValue = CDate("1 January 1900")
                Status = "No Records"
            ElseIf NRows = 1 Then
                'Return ds.Tables(0).Rows(0).Item(PriceType)
                PriceValue = ds.Tables(0).Rows(0).Item(PriceType)
                DateValue = ds.Tables(0).Rows(0).Item("Trade_Date")
                Status = "OK"
            Else
                'Message.AddWarning("More than one record was returned: " & NRows & " records. The first record value is returned." & vbCrLf)
                'Return ds.Tables(0).Rows(0).Item(PriceType)
                PriceValue = ds.Tables(0).Rows(0).Item(PriceType)
                DateValue = ds.Tables(0).Rows(0).Item("Trade_Date")
                Status = "OK"
            End If
        Catch ex As Exception
            Message.AddWarning("Error getting share price value: " & ex.Message & vbCrLf)
        End Try

    End Sub

    'Public Sub ExpRegressionFit(ByVal AsxCode As String, ByVal FromDateString As String, ByVal ToDateString As String, ByVal PriceType As String)
    '    'Use the Regression application to fit an exponential curve to a series of stock prices.
    '    xxx

    'End Sub

    'Public Function ExpRegressionForecastValue(ByVal DateString As String) As Single
    '    'Get a exponential regression forecast at the specified date.

    'End Function

    'Public Function ExpRegressionForecastError() As Single
    '    'Get the exponential regression forecast error (standard deviation).

    'End Function


    'Public Sub CloseAppAtConnection(ByVal AppNetName As String, ByVal ConnectionName As String)
    'Public Sub CloseAppAtConnection(ByVal ProNetName As String, ByVal ConnectionName As String)





    Private Sub Timer4_Tick(sender As Object, e As EventArgs) Handles Timer4.Tick
        'Exit the application:
        Me.btnExit.PerformClick() 'Press the Exit button
    End Sub

    Private Sub ToolStripMenuItem1_EditWorkflowTabPage_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem1_EditWorkflowTabPage.Click
        'Edit the Workflows Web Page:

        If WorkflowFileName = "" Then
            Message.AddWarning("No page to edit." & vbCrLf)
        Else
            Dim FormNo As Integer = OpenNewHtmlDisplayPage()
            HtmlDisplayFormList(FormNo).FileName = WorkflowFileName
            HtmlDisplayFormList(FormNo).OpenDocument
        End If

    End Sub

    Private Sub ToolStripMenuItem1_ShowStartPageInWorkflowTab_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem1_ShowStartPageInWorkflowTab.Click
        'Show the Start Page in the Workflows Tab:
        OpenStartPage()

    End Sub

    Private Sub bgwComCheck_DoWork(sender As Object, e As DoWorkEventArgs) Handles bgwComCheck.DoWork
        'The communications chack thread.
        While ConnectedToComNet
            Try
                If client.IsAlive() Then
                    'Message.Add(Format(Now, "HH:mm:ss") & " Connection OK." & vbCrLf) 'This produces the error: Cross thread operation not valid.
                    bgwComCheck.ReportProgress(1, Format(Now, "HH:mm:ss") & " Connection OK." & vbCrLf)
                Else
                    'Message.Add(Format(Now, "HH:mm:ss") & " Connection Fault." & vbCrLf) 'This produces the error: Cross thread operation not valid.
                    bgwComCheck.ReportProgress(1, Format(Now, "HH:mm:ss") & " Connection Fault.")
                End If
            Catch ex As Exception
                bgwComCheck.ReportProgress(1, "Error in bgeComCheck_DoWork!" & vbCrLf)
                bgwComCheck.ReportProgress(1, ex.Message & vbCrLf)
            End Try

            'System.Threading.Thread.Sleep(60000) 'Sleep time in milliseconds (60 seconds) - For testing only.
            'System.Threading.Thread.Sleep(3600000) 'Sleep time in milliseconds (60 minutes)
            System.Threading.Thread.Sleep(1800000) 'Sleep time in milliseconds (30 minutes)
        End While
    End Sub

    Private Sub bgwComCheck_ProgressChanged(sender As Object, e As ProgressChangedEventArgs) Handles bgwComCheck.ProgressChanged
        Message.Add(e.UserState.ToString) 'Show the ComCheck message 
    End Sub

    Private Sub bgwSendMessage_DoWork(sender As Object, e As DoWorkEventArgs) Handles bgwSendMessage.DoWork
        'Send a message on a separate thread:
        Try
            If IsNothing(client) Then
                bgwSendMessage.ReportProgress(1, "No Connection available. Message not sent!")
            Else
                If client.State = ServiceModel.CommunicationState.Faulted Then
                    bgwSendMessage.ReportProgress(1, "Connection state is faulted. Message not sent!")
                Else
                    Dim SendMessageParams As clsSendMessageParams = e.Argument
                    client.SendMessage(SendMessageParams.ProjectNetworkName, SendMessageParams.ConnectionName, SendMessageParams.Message)
                End If
            End If
        Catch ex As Exception
            bgwSendMessage.ReportProgress(1, ex.Message)
        End Try
    End Sub

    Private Sub bgwSendMessage_ProgressChanged(sender As Object, e As ProgressChangedEventArgs) Handles bgwSendMessage.ProgressChanged
        'Display an error message:
        Message.AddWarning("Send Message error: " & e.UserState.ToString & vbCrLf) 'Show the bgwSendMessage message 
    End Sub

    Private Sub bgwSendMessageAlt_DoWork(sender As Object, e As DoWorkEventArgs) Handles bgwSendMessageAlt.DoWork
        'Alternative SendMessage background worker - used to send a message while instructions are being processed. 
        'Send a message on a separate thread
        Try
            If IsNothing(client) Then
                bgwSendMessageAlt.ReportProgress(1, "No Connection available. Message not sent!")
            Else
                If client.State = ServiceModel.CommunicationState.Faulted Then
                    bgwSendMessageAlt.ReportProgress(1, "Connection state is faulted. Message not sent!")
                Else
                    Dim SendMessageParamsAlt As clsSendMessageParams = e.Argument
                    client.SendMessage(SendMessageParamsAlt.ProjectNetworkName, SendMessageParamsAlt.ConnectionName, SendMessageParamsAlt.Message)
                End If
            End If
        Catch ex As Exception
            bgwSendMessageAlt.ReportProgress(1, ex.Message)
        End Try
    End Sub

    Private Sub bgwSendMessageAlt_ProgressChanged(sender As Object, e As ProgressChangedEventArgs) Handles bgwSendMessageAlt.ProgressChanged
        'Display an error message:
        Message.AddWarning("Send Message error: " & e.UserState.ToString & vbCrLf) 'Show the bgwSendMessageAlt message 
    End Sub

    Private Sub bgwRunInstruction_DoWork(sender As Object, e As DoWorkEventArgs) Handles bgwRunInstruction.DoWork
        'Run a single instruction.
        Try
            Dim Instruction As clsInstructionParams = e.Argument
            XMsg_Instruction(Instruction.Info, Instruction.Locn)
        Catch ex As Exception
            bgwRunInstruction.ReportProgress(1, ex.Message)
        End Try
    End Sub

    Private Sub bgwRunInstruction_ProgressChanged(sender As Object, e As ProgressChangedEventArgs) Handles bgwRunInstruction.ProgressChanged
        'Display an error message:
        Message.AddWarning("Run Instruction error: " & e.UserState.ToString & vbCrLf) 'Show the bgwRunInstruction message 
    End Sub

    Private Sub Message_ShowXMessagesChanged(Show As Boolean) Handles Message.ShowXMessagesChanged
        ShowXMessages = Show
    End Sub

    Private Sub Message_ShowSysMessagesChanged(Show As Boolean) Handles Message.ShowSysMessagesChanged
        ShowSysMessages = Show
    End Sub

    Private Sub btnUpdateStockChartProject_Click(sender As Object, e As EventArgs) Handles btnUpdateStockChartProjects.Click
        'Find a Stock Chart Project:
        UpdateProjectList()
    End Sub

    Private Sub UpdateProjectList()
        'Update the list of projects stored in Proj:

        If ConnectedToComNet Then
            'Clear the current list of project:
            Message.Add("Updating project list." & vbCrLf)
            Proj.List.Clear()

            'CompletionInstruction = "Test"
            'CompletionInstruction = "UpdateStockChartProjects"
            'CompletionInstruction = "UpdateChartProjLists"
            EndInstruction = "UpdateChartProjLists"
            client.GetProjectListAsync("Update")
            'client.GetProjectListAsync("") 'Use blank Client Location to omit the location element from the Project list


            'THE FOLLOWING CODE DOESNT WORK PROPERLY: (This thread continues running instead of pausing for the Project List XMessage to be processed.)
            ''Application.DoEvents()

            'Dim I As Integer

            ''First wait for the Project List XML file to start processing:
            'For I = 1 To 20
            '    If RunningXMsg Then
            '        Message.Add("Waiting to start processing list: RunningXMsg = True" & vbCrLf)
            '        Exit For
            '    Else
            '        Message.Add("Waiting to start processing list: RunningXMsg = False" & vbCrLf)
            '        Application.DoEvents()
            '        System.Threading.Thread.Sleep(100) 'Pause for 100ms
            '    End If
            'Next

            ''Now wait for the processing to end:
            'For I = 1 To 20
            '    If RunningXMsg Then
            '        Message.Add("Waiting to end processing list: RunningXMsg = True" & vbCrLf)
            '        Application.DoEvents()
            '        System.Threading.Thread.Sleep(100) 'Pause for 100ms
            '    Else
            '        Message.Add("Waiting to end processing list: RunningXMsg = False" & vbCrLf)
            '        Exit For
            '    End If
            'Next

            'USE ONCOMPLETION CODE TO RUN UpdateStockChartProjects!
            'Message.Add("Project List import finished. Proj.List.Count = " & Proj.List.Count & vbCrLf)
            'UpdateStockChartProjects()
        Else
            Message.AddWarning("Not connected to the Network." & vbCrLf)
        End If
    End Sub

    Private Sub UpdateStockChartProjects()
        'Update the list of Stock Chart project in cmbStockChartProjects
        cmbStockChartProjects.Items.Clear()
        For Each item In Proj.List
            If item.ApplicationName = "ADVL_Stock_Chart_1" Then
                If item.ProNetName = ProNetName Then
                    'The Stock Chart Project is in the same Project Network as this Project:
                    cmbStockChartProjects.Items.Add(item.Name) 'Projects in the same Project Network will be left aligned.
                Else
                    cmbStockChartProjects.Items.Add("  " & item.Name) 'Projects that are not in the same Project Network will be offet to the right.
                End If
            End If
        Next
    End Sub

    Private Sub WriteProjectList()
        'Write the Project List in Proj.List() to the Project_List_ADVL_2.xml file in the Data Directory.

        Dim ProjectListXDoc = <?xml version="1.0" encoding="utf-8"?>
                              <!---->
                              <!--Project List File-->
                              <ProjectList>
                                  <FormatCode>ADVL_2</FormatCode>
                                  <%= From item In Proj.List
                                      Select
                                          <Project>
                                              <Name><%= item.Name %></Name>
                                              <ProNetName><%= item.ProNetName %></ProNetName>
                                              <ID><%= item.ID %></ID>
                                              <Type><%= item.Type %></Type>
                                              <Path><%= item.Path %></Path>
                                              <Description><%= item.Description %></Description>
                                              <ApplicationName><%= item.ApplicationName %></ApplicationName>
                                              <ParentProjectName><%= item.ParentProjectName %></ParentProjectName>
                                              <ParentProjectID><%= item.ParentProjectID %></ParentProjectID>
                                          </Project>
                                  %>
                              </ProjectList>

        'ProjectListXDoc.Save(ApplicationInfo.ApplicationDir & "\Global_Project_List_ADVL_2.xml")
        Project.SaveXmlData("Project_List_ADVL_2.xml", ProjectListXDoc)
    End Sub

    Private Sub ReadProjectList()
        'Read the Project List in Project_List_ADVL_2.xml file in the Data Directory and store the data in Proj.List().

        'If System.IO.File.Exists(ApplicationInfo.ApplicationDir & "\Global_Project_List_ADVL_2.xml") Then
        If Project.DataFileExists("Project_List_ADVL_2.xml") Then
            Dim ProjListXDoc As System.Xml.Linq.XDocument
            'ProjListXDoc = XDocument.Load(ApplicationInfo.ApplicationDir & "\Global_Project_List_ADVL_2.xml")
            Project.ReadXmlData("Project_List_ADVL_2.xml", ProjListXDoc)

            Dim Projects = From item In ProjListXDoc.<ProjectList>.<Project>

            Proj.List.Clear()

            For Each item In Projects
                Dim NewProj As New ProjSummary
                NewProj.Name = item.<Name>.Value

                If item.<ProNetName>.Value Is Nothing Then
                    'Check if the old AppNetName is used:
                    If item.<AppNetName>.Value Is Nothing Then
                        NewProj.ProNetName = ""
                    Else 'Use the old AppNetName value as the ProNetName:
                        NewProj.ProNetName = item.<AppNetName>.Value
                    End If
                Else
                    NewProj.ProNetName = item.<ProNetName>.Value
                End If

                NewProj.ID = item.<ID>.Value
                Select Case item.<Type>.Value
                    Case "None"
                        NewProj.Type = ADVL_Utilities_Library_1.Project.Types.None
                    Case "Directory"
                        NewProj.Type = ADVL_Utilities_Library_1.Project.Types.Directory
                    Case "Archive"
                        NewProj.Type = ADVL_Utilities_Library_1.Project.Types.Archive
                    Case "Hybrid"
                        NewProj.Type = ADVL_Utilities_Library_1.Project.Types.Hybrid
                    Case Else
                        Message.AddWarning("Unknown project type: " & item.<Type>.Value & vbCrLf)
                End Select
                NewProj.Path = item.<Path>.Value
                NewProj.Description = item.<Description>.Value
                NewProj.ApplicationName = item.<ApplicationName>.Value
                If item.<HostProjectName>.Value <> Nothing Then NewProj.ParentProjectName = item.<HostProjectName>.Value 'Legacy version - in case <HostProjectName> is used.
                If item.<ParentProjectName>.Value <> Nothing Then NewProj.ParentProjectName = item.<ParentProjectName>.Value 'Updated version.
                If item.<HostProjectID>.Value <> Nothing Then NewProj.ParentProjectID = item.<HostProjectID>.Value 'Legacy version - in case <HostProjectID> is used.
                If item.<ParentProjectID>.Value <> Nothing Then NewProj.ParentProjectID = item.<ParentProjectID>.Value 'Updated version.
                Proj.List.Add(NewProj)
            Next
            'UpdateProjectGrid()
        End If
    End Sub

    Private Sub UpdateShareChartProjList()
        'Update the list of Share Chart projects:

    End Sub

    Private Sub UpdatePointChartProjList()
        'Update the list of Point Chart projects:


    End Sub

    Private Sub UpdateLineChartProjList()
        'Update the list of LineChart projects:


    End Sub

    Private Sub UpdateChartProjLists()
        'Update the lists of Chart projects:

        ShareChartProj.List.Clear()
        PointChartProj.List.Clear()
        LineChartProj.List.Clear()

        cmbStockChartProjects.Items.Clear()
        'cmbPointChartProjects.Items.Clear()
        'cmbLineChartProjects.Items.Clear()

        Dim I As Integer
        For I = 0 To Proj.List.Count - 1
            Select Case Proj.List(I).ApplicationName
                Case "ADVL_Stock_Chart_1"
                    ShareChartProj.List.Add(Proj.List(I))
                    If Proj.List(I).ProNetName = ProNetName Then
                        cmbStockChartProjects.Items.Add(Proj.List(I).Name) 'Projects in the same Project Network will be left aligned.
                    Else
                        cmbStockChartProjects.Items.Add("  " & Proj.List(I).Name) 'Projects that are not in the same Project Network will be offet to the right.
                    End If
                Case "ADVL_Point_Chart_1"
                    PointChartProj.List.Add(Proj.List(I))
                    If Proj.List(I).ProNetName = ProNetName Then
                        cmbPointChartProjects.Items.Add(Proj.List(I).Name) 'Projects in the same Project Network will be left aligned.
                    Else
                        cmbPointChartProjects.Items.Add("  " & Proj.List(I).Name) 'Projects that are not in the same Project Network will be offet to the right.
                    End If
                Case "ADVL_Line_Chart_1"
                    LineChartProj.List.Add(Proj.List(I))
                    If Proj.List(I).ProNetName = ProNetName Then
                        cmbLineChartProjects.Items.Add(Proj.List(I).Name) 'Projects in the same Project Network will be left aligned.
                    Else
                        cmbLineChartProjects.Items.Add("  " & Proj.List(I).Name) 'Projects that are not in the same Project Network will be offet to the right.
                    End If
            End Select
        Next

        'Select Stock Chart on cmbStockChartProjects:
        If cmbStockChartProjects.Items.Count > 0 Then
            If cmbStockChartProjects.Items.Count > SelShareChartProjNo Then
                cmbStockChartProjects.SelectedIndex = SelShareChartProjNo
            Else
                cmbStockChartProjects.SelectedIndex = cmbStockChartProjects.Items.Count - 1
            End If
        Else
            cmbStockChartProjects.SelectedIndex = -1
        End If

        'Select Point Chart on cmbPointChartProjects:
        If cmbPointChartProjects.Items.Count > 0 Then
            If cmbPointChartProjects.Items.Count > SelPointChartProjNo Then
                cmbPointChartProjects.SelectedIndex = SelPointChartProjNo
            Else
                cmbPointChartProjects.SelectedIndex = cmbPointChartProjects.Items.Count - 1
            End If
        Else
            cmbPointChartProjects.SelectedIndex = -1
        End If

        'Select Line Chart on cmbLineChartProjects:
        If cmbLineChartProjects.Items.Count > 0 Then
            If cmbLineChartProjects.Items.Count > SelLineChartProjNo Then
                cmbLineChartProjects.SelectedIndex = SelLineChartProjNo
            Else
                cmbLineChartProjects.SelectedIndex = cmbLineChartProjects.Items.Count - 1
            End If
        Else
            cmbLineChartProjects.SelectedIndex = -1
        End If

    End Sub

    Private Sub Project_NewProjectCreated(ProjectPath As String) Handles Project.NewProjectCreated
        SendProjectInfo(ProjectPath) 'Send the path of the new project to the Network application. The new project will be added to the list of projects.
    End Sub



#End Region 'Form Methods ---------------------------------------------------------------------------------------------------------------------------------------------------------------------



#Region " Form Events - Events that can be triggered by this form." '==========================================================================================================================
#End Region 'Form Events ----------------------------------------------------------------------------------------------------------------------------------------------------------------------

#Region " Classes - Other classes used in this form." '========================================================================================================================================

    Public Class clsSendMessageParams
        'Parameters used when sending a message using the Message Service.
        Public ProjectNetworkName As String
        Public ConnectionName As String
        Public Message As String
    End Class

    Public Class clsInstructionParams
        'Parameters used when executing an instruction.
        Public Info As String 'The information in an instruction.
        Public Locn As String 'The location to send the information.
    End Class

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

    Private Sub btnShowVoices_Click(sender As Object, e As EventArgs) Handles btnShowVoices.Click
        'Show the list of installed voices:

        Dim mySynth As New System.Speech.Synthesis.SpeechSynthesizer

        Message.Add("List of voices installed in this computer:" & vbCrLf)

        For Each voice As Speech.Synthesis.InstalledVoice In mySynth.GetInstalledVoices
            Message.Add(voice.VoiceInfo.Description & vbCrLf)
        Next

    End Sub

    Private Sub XMsg_ErrorMsg(ErrMsg As String) Handles XMsg.ErrorMsg
        Message.AddWarning(ErrMsg & vbCrLf)
    End Sub




    Private Sub btnSelectSPDatabase_Click(sender As Object, e As EventArgs) Handles btnSelectSPDatabase.Click
        'Select the database selected in the dgvSPDatabase

        If dgvSPDatabase.SelectedRows.Count = 0 Then
            Message.AddWarning("No Share Price database has been selected." & vbCrLf)
        ElseIf dgvSPDatabase.SelectedRows.Count = 1 Then
            Dim SelRow As Integer = dgvSPDatabase.SelectedRows(0).Index
            SharePriceDbName = dgvSPDatabase.Rows(SelRow).Cells(0).Value
            SharePriceDbDescription = dgvSPDatabase.Rows(SelRow).Cells(1).Value
            SharePriceDbFileName = dgvSPDatabase.Rows(SelRow).Cells(2).Value
            SharePriceDbLocation = dgvSPDatabase.Rows(SelRow).Cells(3).Value
            SharePriceDbPath = dgvSPDatabase.Rows(SelRow).Cells(4).Value

            'Set the corresponding Project Parameter:
            Project.AddParameter("SharePriceDatabasePath", SharePriceDbPath, "The path of the Share Price database.")
            Project.SaveParameters()

        Else
            Message.AddWarning("More than one Share Price database has been selected." & vbCrLf)
        End If
    End Sub

    Private Sub btnSelectFinDb_Click(sender As Object, e As EventArgs) Handles btnSelectFinDb.Click
        'Select the database selected in the dgvFinDatabase

        If dgvFinDatabase.SelectedRows.Count = 0 Then
            Message.AddWarning("No Financials database has been selected." & vbCrLf)
        ElseIf dgvFinDatabase.SelectedRows.Count = 1 Then
            Dim SelRow As Integer = dgvFinDatabase.SelectedRows(0).Index
            FinancialsDbName = dgvFinDatabase.Rows(SelRow).Cells(0).Value
            FinancialsDbDescription = dgvFinDatabase.Rows(SelRow).Cells(1).Value
            FinancialsDbFileName = dgvFinDatabase.Rows(SelRow).Cells(2).Value
            FinancialsDbLocation = dgvFinDatabase.Rows(SelRow).Cells(3).Value
            FinancialsDbPath = dgvFinDatabase.Rows(SelRow).Cells(4).Value

            'Set the corresponding Project Parameter:
            Project.AddParameter("FinancialsDatabasePath", FinancialsDbPath, "The path of the Historical Financials database.")
            Project.SaveParameters()

        Else
            Message.AddWarning("More than one Financials database has been selected." & vbCrLf)
        End If
    End Sub

    Private Sub btnSelectCalcDb_Click(sender As Object, e As EventArgs) Handles btnSelectCalcDb.Click
        'Select the database selected in the dgvCalcDatabase

        If dgvCalcDatabase.SelectedRows.Count = 0 Then
            Message.AddWarning("No Calculations database has been selected." & vbCrLf)
        ElseIf dgvCalcDatabase.SelectedRows.Count = 1 Then
            Dim SelRow As Integer = dgvCalcDatabase.SelectedRows(0).Index
            CalculationsDbName = dgvCalcDatabase.Rows(SelRow).Cells(0).Value
            CalculationsDbDescription = dgvCalcDatabase.Rows(SelRow).Cells(1).Value
            CalculationsDbFileName = dgvCalcDatabase.Rows(SelRow).Cells(2).Value
            CalculationsDbLocation = dgvCalcDatabase.Rows(SelRow).Cells(3).Value
            CalculationsDbPath = dgvCalcDatabase.Rows(SelRow).Cells(4).Value

            'Set the corresponding Project Parameter:
            Project.AddParameter("CalculationsDatabasePath", CalculationsDbPath, "The path of the Calculations database.")
            Project.SaveParameters()
        Else
            Message.AddWarning("More than one Calculations database has been selected." & vbCrLf)
        End If
    End Sub

    Private Sub btnRemoveCalcDb_Click(sender As Object, e As EventArgs) Handles btnRemoveCalcDb.Click
        'Remove the Calculations database entry selected in dgvCalcDatabase

        If dgvCalcDatabase.SelectedRows.Count = 0 Then
            Message.AddWarning("No Calculations database has been selected." & vbCrLf)
        ElseIf dgvCalcDatabase.SelectedRows.Count = 1 Then
            Dim SelRow As Integer = dgvCalcDatabase.SelectedRows(0).Index
            dgvCalcDatabase.Rows.RemoveAt(SelRow)
        Else
            Message.AddWarning("More than one Calculations database has been selected." & vbCrLf)
        End If
    End Sub

    Private Sub btnRemoveFinDb_Click(sender As Object, e As EventArgs) Handles btnRemoveFinDb.Click
        'Remove the Financials database entry selected in dgvFinDatabase

        If dgvFinDatabase.SelectedRows.Count = 0 Then
            Message.AddWarning("No Financials database has been selected." & vbCrLf)
        ElseIf dgvFinDatabase.SelectedRows.Count = 1 Then
            Dim SelRow As Integer = dgvFinDatabase.SelectedRows(0).Index
            dgvFinDatabase.Rows.RemoveAt(SelRow)
        Else
            Message.AddWarning("More than one Financials database has been selected." & vbCrLf)
        End If
    End Sub

    Private Sub btnRemoveSPDatabase_Click(sender As Object, e As EventArgs) Handles btnRemoveSPDatabase.Click
        'Remove the Share Price database entry selected in dgvFinDatabase

        If dgvSPDatabase.SelectedRows.Count = 0 Then
            Message.AddWarning("No Share Price database has been selected." & vbCrLf)
        ElseIf dgvSPDatabase.SelectedRows.Count = 1 Then
            Dim SelRow As Integer = dgvSPDatabase.SelectedRows(0).Index
            dgvSPDatabase.Rows.RemoveAt(SelRow)
        Else
            Message.AddWarning("More than one Share Price database has been selected." & vbCrLf)
        End If
    End Sub

    Private Sub cmbStockChartProjects_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbStockChartProjects.SelectedIndexChanged
        SelShareChartProjNo = cmbStockChartProjects.SelectedIndex
    End Sub

    Private Sub cmbPointChartProjects_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbPointChartProjects.SelectedIndexChanged
        SelPointChartProjNo = cmbPointChartProjects.SelectedIndex
    End Sub

    Private Sub cmbLineChartProjects_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbLineChartProjects.SelectedIndexChanged
        SelLineChartProjNo = cmbLineChartProjects.SelectedIndex
    End Sub

    Private Sub btnStockChartProjectInfo_Click(sender As Object, e As EventArgs) Handles btnStockChartProjectInfo.Click
        'Display information about the selected Stock Chart Project:

        Message.Add("Stock Chart Project Information: ----------------------------" & vbCrLf)
        Message.Add("Project name: " & ShareChartProj.List(SelShareChartProjNo).Name & vbCrLf)
        Message.Add("Project description: " & ShareChartProj.List(SelShareChartProjNo).Description & vbCrLf)
        Message.Add("Project path: " & ShareChartProj.List(SelShareChartProjNo).Path & vbCrLf)
        Message.Add("Project type: " & ShareChartProj.List(SelShareChartProjNo).Type.ToString & vbCrLf)
        Message.Add("Project ID: " & ShareChartProj.List(SelShareChartProjNo).ID & vbCrLf)
        Message.Add("Application name: " & ShareChartProj.List(SelShareChartProjNo).ApplicationName & vbCrLf)
        Message.Add("Project Network name: " & ShareChartProj.List(SelShareChartProjNo).ProNetName & vbCrLf)
        Message.Add("Parent project name: " & ShareChartProj.List(SelShareChartProjNo).ParentProjectName & vbCrLf)
        Message.Add("Parent project ID: " & ShareChartProj.List(SelShareChartProjNo).ParentProjectID & vbCrLf)
        Message.Add("-------------------------------------------------------------" & vbCrLf & vbCrLf)
    End Sub



#End Region 'Classes --------------------------------------------------------------------------------------------------------------------------------------------------------------------------



End Class 'Main.

Public Class Proj
    'Class holds a list of projects.
    'This is used by the Proj object that contain a list of all projects.

    Public List As New List(Of ProjSummary) 'A list of projects

#Region "Application Methods" '--------------------------------------------------------------------------------------

    Public Function FindID(ByVal ProjID As String) As ProjSummary
        'Return the ProjSummary corresponding to the Project with ID ProjID

        Dim FoundID As ProjSummary

        FoundID = List.Find(Function(item As ProjSummary)
                                If IsNothing(item) Then
                                    '
                                Else
                                    Return item.ID = ProjID
                                End If
                            End Function)
        If IsNothing(FoundID) Then
            Return New ProjSummary 'Return blank record.
        Else
            Return FoundID
        End If
    End Function

    Public Function FindNameAndAppNet(ByVal Name As String, ByVal ProNetName As String) As ProjSummary
        'Return the ProjSummary corresponding to the Project with specified Name and ProNetName.

        Dim FoundProj As ProjSummary

        FoundProj = List.Find(Function(item As ProjSummary)
                                  'Return item.Name = Name And item.AppNetName = AppNetName
                                  Return item.Name = Name And item.ProNetName = ProNetName
                              End Function)
        If IsNothing(FoundProj) Then
            Return New ProjSummary
        Else
            Return FoundProj
        End If

    End Function



#End Region 'Application Methods ------------------------------------------------------------------------------------

End Class 'Proj
Public Class ProjSummary
    'Class holds summary information about a project.
    'This is used by the Proj class.

    Private _name As String = "" 'The name of the project.
    Property Name As String
        Get
            Return _name
        End Get
        Set(value As String)
            _name = value
        End Set
    End Property

    Private _proNetName As String = "" 'The name of the Project Network containing the project. 
    Property ProNetName As String
        Get
            Return _proNetName
        End Get
        Set(value As String)
            _proNetName = value
        End Set
    End Property

    Private _iD As String = "" 'The project ID.
    Property ID As String
        Get
            Return _iD
        End Get
        Set(value As String)
            _iD = value
        End Set
    End Property

    Private _type As ADVL_Utilities_Library_1.Project.Types = ADVL_Utilities_Library_1.Project.Types.Directory 'The type of location (None, Directory, Archive, Hybrid).
    Property Type As ADVL_Utilities_Library_1.Project.Types
        Get
            Return _type
        End Get
        Set(value As ADVL_Utilities_Library_1.Project.Types)
            _type = value
        End Set
    End Property

    Private _path As String = "" 'The path to the Project directory or archive.
    Property Path As String
        Get
            Return _path
        End Get
        Set(value As String)
            _path = value
        End Set
    End Property

    Private _description As String = "" 'A description of the project.
    Property Description As String
        Get
            Return _description
        End Get
        Set(value As String)
            _description = value
        End Set
    End Property

    Private _applicationName As String = "" 'The name of the application that created the project.
    Property ApplicationName As String
        Get
            Return _applicationName
        End Get
        Set(value As String)
            _applicationName = value
        End Set
    End Property

    Private _parentProjectName As String = "" 'The Name of the Parent Project.
    Property ParentProjectName As String
        Get
            Return _parentProjectName
        End Get
        Set(value As String)
            _parentProjectName = value
        End Set
    End Property

    Private _parentProjectID As String = "" 'The parent project ID.
    Property ParentProjectID As String
        Get
            Return _parentProjectID
        End Get
        Set(value As String)
            _parentProjectID = value
        End Set
    End Property

End Class


