
Imports System.Security.Permissions
<PermissionSet(SecurityAction.Demand, Name:="FullTrust")>
<System.Runtime.InteropServices.ComVisibleAttribute(True)>
Public Class frmWebPage
    'This form displays a Web view of the HTML Code.

#Region " Variable Declarations - All the variables used in this form and this application." '=================================================================================================

    Public WithEvents XSeq As New ADVL_Utilities_Library_1.XSequence
    'To run and XSequence:
    '  XSeq.RunXSequence(xDoc, Status) 'ImportStatus in Import
    '    Handle events:
    '      XSeq.ErrorMsg
    '      XSeq.Instruction(Info, Locn)

    Private Status As New System.Collections.Specialized.StringCollection

    'Variables used to restore Item values on a web page.
    Private FormName As String
    Private ItemName As String
    Private SelectId As String

#End Region 'Variable Declarations ------------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Properties - All the properties used in this form and this application" '============================================================================================================

    'PROPERTIES:
    'FormNo - the index number of the web page form in WebPageFormList().
    'FileName - the file name of the displayed web document.
    'Description - a description of the displayed web page.
    'FileLocationType - Project or FileSystem - In this application, the web page file will always be located in the Project.
    'FileDirectory - the path of the directory for a file in the FileSystem - Not used in this application.
    'ParentWebPageFileName - 'The file name of the web page that opened this web page.
    'ParentWebPageFormNo - 'The form index number of the web page that opened this web page.

    Private _formNo As Integer = -1 'Multiple instances of this form can be displayed. FormNo is the index number of the form in XmlDisplayFormList.
    'If the form is included in Main.WebViewFormList() then FormNo will be > -1 --> when exiting set Main.ClosedFormNo and call Main.WebViewFormClosed(). 
    Public Property FormNo As Integer
        Get
            Return _formNo
        End Get
        Set(ByVal value As Integer)
            _formNo = value
            Debug.Print("FormNo = " & _formNo)
        End Set
    End Property

    Private _fileName As String = "" 'The file name of the displayed document.
    Public Property FileName As String
        Get
            Return _fileName
        End Get
        Set(value As String)
            _fileName = value
            txtDocumentFile.Text = _fileName 'Display the document filename on the form.
            Me.Text = Main.ApplicationInfo.Name & " - Workflow - " & IO.Path.GetFileNameWithoutExtension(_fileName) 'Update the text at the top of the window.
            RestoreFormSettings() 'Resore the form settings used to display this web page.
        End Set
    End Property

    Private _description As String = "" 'A description of the displayed document.
    Public Property Description As String
        Get
            Return _description
        End Get
        Set(value As String)
            _description = value
        End Set
    End Property

    Enum LocationTypes
        Project
        FileSystem
    End Enum

    Private _fileLocationType As LocationTypes = LocationTypes.Project 'The location type of the Document File. (Either the current project or the file system.)
    Property FileLocationType As LocationTypes
        Get
            Return _fileLocationType
        End Get
        Set(value As LocationTypes)
            _fileLocationType = value
        End Set
    End Property

    Private _fileDirectory As String = "" 'The path of the directory containing the current file.
    Property FileDirectory As String
        Get
            Return _fileDirectory
        End Get
        Set(value As String)
            _fileDirectory = value
        End Set
    End Property

    Private _parentWebPageFileName As String = "" 'The file name of the web page that opened this web page.
    Property ParentWebPageFileName As String
        Get
            Return _parentWebPageFileName
        End Get
        Set(value As String)
            _parentWebPageFileName = value
        End Set
    End Property

    Private _parentWebPageFormNo As Integer = -2 'The form index number of the web page that opened this web page.
    Property ParentWebPageFormNo As Integer
        Get
            Return _parentWebPageFormNo
        End Get
        Set(value As Integer)
            _parentWebPageFormNo = value
        End Set
    End Property


#End Region 'Properties -----------------------------------------------------------------------------------------------------------------------------------------------------------------------

#Region " Process XML files - Read and write XML files." '=====================================================================================================================================

    Private Sub SaveFormSettings()
        'Save the form settings in an XML document.
        Dim settingsData = <?xml version="1.0" encoding="utf-8"?>
                           <!---->
                           <FormSettings>
                               <Left><%= Me.Left %></Left>
                               <Top><%= Me.Top %></Top>
                               <Width><%= Me.Width %></Width>
                               <Height><%= Me.Height %></Height>
                               <!---->
                           </FormSettings>

        'Add code to include other settings to save after the comment line <!---->

        'Dim SettingsFileName As String = "FormSettings_" & Main.ApplicationInfo.Name & "_" & Me.Text & ".xml"

        'NOTE: After a workflow is loaded, Me.Text is changed to AppName - Workflow - WorkflowName
        Dim SettingsFileName As String = "FormSettings_" & Me.Text & ".xml"
        Main.Project.SaveXmlSettings(SettingsFileName, settingsData)
    End Sub

    Private Sub RestoreFormSettings()
        'Read the form settings from an XML document.

        'Dim SettingsFileName As String = "FormSettings_" & Main.ApplicationInfo.Name & "_" & Me.Text & ".xml"

        'NOTE: After a workflow is loaded, Me.Text is changed to AppName - Workflow - WorkflowName
        Dim SettingsFileName As String = "FormSettings_" & Me.Text & ".xml"

        If Main.Project.SettingsFileExists(SettingsFileName) Then
            Dim Settings As System.Xml.Linq.XDocument
            Main.Project.ReadXmlSettings(SettingsFileName, Settings)

            If IsNothing(Settings) Then 'There is no Settings XML data.
                Exit Sub
            End If

            'Restore form position and size:
            If Settings.<FormSettings>.<Left>.Value <> Nothing Then Me.Left = Settings.<FormSettings>.<Left>.Value
            If Settings.<FormSettings>.<Top>.Value <> Nothing Then Me.Top = Settings.<FormSettings>.<Top>.Value
            If Settings.<FormSettings>.<Height>.Value <> Nothing Then Me.Height = Settings.<FormSettings>.<Height>.Value
            If Settings.<FormSettings>.<Width>.Value <> Nothing Then Me.Width = Settings.<FormSettings>.<Width>.Value

            'Add code to read other saved setting here:

            CheckFormPos()
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

    Protected Overrides Sub WndProc(ByRef m As Message) 'Save the form settings before the form is minimised:
        If m.Msg = &H112 Then 'SysCommand
            If m.WParam.ToInt32 = &HF020 Then 'Form is being minimised
                SaveFormSettings()
            End If
        End If
        MyBase.WndProc(m)
    End Sub

#End Region 'Process XML Files ----------------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Form Display Methods - Code used to display this form." '============================================================================================================================

    Private Sub Form_Load(sender As Object, e As EventArgs) Handles Me.Load
        'RestoreFormSettings()   'Restore the form settings

        Me.WebBrowser1.ObjectForScripting = Me

        'Add page title:
        'Me.Text = Main.ApplicationInfo.Name & " - Workflow Web Page"
        Me.Text = Main.ApplicationInfo.Name & " - Workflow"
        RestoreFormSettings()   'Restore the form settings
    End Sub

    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        'Exit the Form

        SaveWebPageSettings() 'Save the web page settings.

        If FormNo > -1 Then
            Main.ClosedFormNo = FormNo 'The Main form property ClosedFormNo is set to this form number. This is used in the RtfDisplayFormClosed method to select the correct form to set to nothing.
        End If

        Me.Close() 'Close the form
    End Sub

    Private Sub Form_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        If WindowState = FormWindowState.Normal Then
            SaveFormSettings()
        Else
            'Dont save settings if the form is minimised.
        End If
    End Sub


    Private Sub frmWebPage_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        If FormNo > -1 Then
            Main.WebPageFormClosed()
        End If
    End Sub

#End Region 'Form Display Methods -------------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Open and Close Forms - Code used to open and close other forms." '===================================================================================================================

#End Region 'Open and Close Forms -------------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Form Methods - The main actions performed by this form." '===========================================================================================================================

    Private Sub WebBrowser1_DocumentCompleted(sender As Object, e As WebBrowserDocumentCompletedEventArgs) Handles WebBrowser1.DocumentCompleted
        AddHandler WebBrowser1.Document.MouseOver, AddressOf Me.DisplayHyperlinks
    End Sub

    Public Sub DisplayHyperlinks(ByVal sender As Object, ByVal e As System.Windows.Forms.HtmlElementEventArgs)
        txtLink.Text = e.ToElement.GetAttribute("href")
    End Sub

    Public Sub OpenDocument()
        'Open the document specified by FileName, FileLocationType and FileDirectory.

        If FileLocationType = LocationTypes.Project Then

            Dim rtbData As New IO.MemoryStream
            Main.Project.ReadData(FileName, rtbData)

            rtbData.Position = 0

            Dim sr As New IO.StreamReader(rtbData)

            WebBrowser1.DocumentText = sr.ReadToEnd()

        Else

            WebBrowser1.Navigate("file:///" & FileDirectory & "\" & FileName)

        End If

    End Sub

#Region " Methods Called by JavaScript - A collection of methods that can be called by JavaScript in a web page shown in WebBrowser1" '========================================================

    Public Sub JSMethodTest1()
        'Test method that is called from JavaScript.
        Main.Message.Add("JSMethodTest1 called OK." & vbCrLf)
    End Sub

    Public Sub JSMethodTest2(ByVal Var1 As String, ByVal Var2 As String)
        'Test method that is called from JavaScript.
        Main.Message.Add("Var1 = " & Var1 & " Var2 = " & Var2 & vbCrLf)
    End Sub

    Public Sub JSDisplayXml(ByRef XDoc As XDocument)
        Main.Message.Add(XDoc.ToString & vbCrLf & vbCrLf)
    End Sub

    Public Sub ShowMessage(ByVal Msg As String)
        Main.Message.Add(Msg)
    End Sub

    Public Sub SaveHtmlSettings_Old2(ByVal xSettings As String, ByVal FileName As String)
        'Save the Html settings for a web page.

        'Convert the XSettings to XML format:
        Dim XmlHeader As String = "<?xml version=""1.0"" encoding=""utf-8"" standalone=""yes""?>"

        Dim XDocSettings As New System.Xml.Linq.XDocument

        Try
            XDocSettings = System.Xml.Linq.XDocument.Parse(XmlHeader & vbCrLf & xSettings)
        Catch ex As Exception
            Main.Message.AddWarning("Error saving HTML settings file. " & ex.Message & vbCrLf)
        End Try

        Main.Project.SaveXmlData(FileName, XDocSettings)

    End Sub

    Public Sub SaveHtmlSettings(ByVal xSettings As String)
        'Save the Html settings for a web page.

        Dim SettingsFileName As String = FileName & "Settings"

        If ParentWebPageFileName <> "" Then
            SettingsFileName = SettingsFileName & ParentWebPageFileName.GetHashCode
        End If

        'Convert the XSettings to XML format:
        Dim XmlHeader As String = "<?xml version=""1.0"" encoding=""utf-8"" standalone=""yes""?>"

        Dim XDocSettings As New System.Xml.Linq.XDocument

        Try
            XDocSettings = System.Xml.Linq.XDocument.Parse(XmlHeader & vbCrLf & xSettings)
        Catch ex As Exception
            Main.Message.AddWarning("Error saving HTML settings file. " & ex.Message & vbCrLf)
        End Try

        Main.Project.SaveXmlData(SettingsFileName, XDocSettings)

    End Sub

    Public Sub RestoreHtmlSettings_Old2(ByVal FileName As String)
        'Restore the Html settings for a web page.

        Dim XDocSettings As New System.Xml.Linq.XDocument
        Main.Project.ReadXmlData(FileName, XDocSettings)

        If XDocSettings Is Nothing Then
            Main.Message.Add("No HTML Settings file : " & FileName & vbCrLf)
        Else

            Dim XSettings As New System.Xml.XmlDocument
            Try
                XSettings.LoadXml(XDocSettings.ToString)

                'Run the Settings file:
                XSeq.RunXSequence(XSettings, Status)
            Catch ex As Exception
                Main.Message.AddWarning("Error restoring HTML settings. " & ex.Message & vbCrLf)
            End Try

        End If
    End Sub

    Public Sub RestoreHtmlSettings()
        'Restore the Html settings for a web page.

        Dim SettingsFileName As String = FileName & "Settings"

        If ParentWebPageFileName <> "" Then
            SettingsFileName = SettingsFileName & ParentWebPageFileName.GetHashCode
        End If

        Dim XDocSettings As New System.Xml.Linq.XDocument
        Main.Project.ReadXmlData(SettingsFileName, XDocSettings)

        If XDocSettings Is Nothing Then
            'Main.Message.Add("No HTML Settings file : " & SettingsFileName & vbCrLf) 'THIS WARNING MESSAGE IS NOT REALLY NEEDED!
        Else

            Dim XSettings As New System.Xml.XmlDocument
            Try
                XSettings.LoadXml(XDocSettings.ToString)

                'Run the Settings file:
                XSeq.RunXSequence(XSettings, Status)
            Catch ex As Exception
                Main.Message.AddWarning("Error restoring HTML settings. " & ex.Message & vbCrLf)
            End Try

        End If
    End Sub

    Private Sub XSeq_ErrorMsg(ErrMsg As String) Handles XSeq.ErrorMsg
        Main.Message.AddWarning(ErrMsg & vbCrLf)
    End Sub

    Public Sub RunClipboardXSeq()
        'Run the XSequence instructions in the clipboard.

        Dim XDocSeq As System.Xml.Linq.XDocument
        Try
            XDocSeq = XDocument.Parse(My.Computer.Clipboard.GetText)
        Catch ex As Exception
            Main.Message.AddWarning("Error reading Clipboard data. " & ex.Message & vbCrLf)
            Exit Sub
        End Try

        If IsNothing(XDocSeq) Then
            Main.Message.Add("No XSequence instructions were found in the clipboard.")
            'Exit Sub
        Else
            Dim XmlSeq As New System.Xml.XmlDocument
            Try
                XmlSeq.LoadXml(XDocSeq.ToString) 'Convert XDocSeq to an XmlDocument to process with XSeq.

                'Run the sequence:
                XSeq.RunXSequence(XmlSeq, Status)
            Catch ex As Exception
                Main.Message.AddWarning("Error restoring HTML settings. " & ex.Message & vbCrLf)
            End Try
        End If

    End Sub


    Private Sub XSeq_Instruction(Data As String, Locn As String) Handles XSeq.Instruction
        'Execute each instruction produced by running the XSeq file.

        Select Case Locn
            Case "Settings:SendData:LatDegrees" 'REDUNDANT!
                RestoreSetting("SendData", "LatDegrees", Data)

            Case "Settings:SendData:LongDegrees" 'REDUNDANT!
                RestoreSetting("SendData", "LongDegrees", Data)

            Case "Settings:Form:Name"
                FormName = Data

            Case "Settings:Form:Item:Name"
                ItemName = Data

            Case "Settings:Form:Item:Value"
                RestoreSetting(FormName, ItemName, Data)

            Case "Settings:Form:SelectId"
                SelectId = Data

            Case "Settings:Form:OptionText"
                'RestoreOption(FormName, SelectId, Data)
                RestoreOption(SelectId, Data)

            Case "Settings"

            Case "EndOfSequence"

            Case Else

                'If the instructions are not saved web page settings identified above, send them directly to the web page:
                XMsgInstruction(Data, Locn) 'The JavaScript function (also called XMsgInstruction) will attempt to process this instruction.

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
            Main.Message.AddWarning("Web page settings not saved: " & ex.Message & vbCrLf)
        End Try

    End Sub

    Public Function GetFormNo() As String
        'Return the Form Number of the current instance of the WebView form.
        Return FormNo.ToString
    End Function

    Public Function GetParentFormNo() As String
        'Return the Form Number of the Parent Form (that called this form).
        Return ParentWebPageFormNo.ToString
    End Function

    Public Sub RunXMessage(ByVal XMsg As String)
        'Run the XMessage by sending it to Main.InstrReceived.
        Main.InstrReceived = XMsg
    End Sub

    Public Sub RunXSequence(ByVal XSequence As String)
        'Run the XMSequence
        Dim XmlSeq As New System.Xml.XmlDocument
        XmlSeq.LoadXml(XSequence)
        Main.XSeq.RunXSequence(XmlSeq, Status)
    End Sub

    Public Function GetConnectionName() As String
        'Return the Connection Name of the project.
        Return Main.ConnectionName
    End Function

    'Public Function GetAppNetName() As String
    '    'Return the Application Network Name of the application.

    '    'Main.Message.Add("AppNetName = " & Main.AppNetName & vbCrLf)

    '    Return Main.AppNetName

    'End Function

    Public Function GetProNetName() As String
        'Return the Project Network Name of the project.

        Return Main.ProNetName

    End Function

    'Public Sub SendXMessage(ByVal AppNetName As String, ByVal ConnName As String, ByVal XMsg As String)
    '    'Send the XMsg to the application with the connection name ConnName.
    '    If IsNothing(Main.client) Then
    '        Main.Message.Add("No client connection available!" & vbCrLf)
    '    Else
    '        If Main.client.State = ServiceModel.CommunicationState.Faulted Then
    '            Main.Message.Add("client state is faulted. Message not sent!" & vbCrLf)
    '        Else
    '            Main.client.SendMessageAsync(AppNetName, ConnName, XMsg)
    '            Main.Message.XAddText("Message sent to " & ConnName & " (AppNet: " & AppNetName & ") " & ":" & vbCrLf, "XmlSentNotice") 'UPDATED 2Feb19
    '            Main.Message.XAddXml(XMsg)
    '            Main.Message.XAddText(vbCrLf, "Normal") 'Add extra line
    '        End If
    '    End If

    'End Sub

    Public Sub SendXMessage(ByVal ConnName As String, ByVal XMsg As String)
        'Send the XMsg to the application with the connection name ConnName.
        If IsNothing(Main.client) Then
            Main.Message.Add("No client connection available!" & vbCrLf)
        Else
            If Main.client.State = ServiceModel.CommunicationState.Faulted Then
                Main.Message.Add("client state is faulted. Message not sent!" & vbCrLf)
            Else
                'Main.client.SendMessageAsync(Main.ProNetName, ConnName, XMsg)
                If Main.bgwSendMessage.IsBusy Then
                    Application.DoEvents()
                    System.Threading.Thread.Sleep(100) 'Pause for 100ms
                    If Main.bgwSendMessage.IsBusy Then
                        Application.DoEvents()
                        System.Threading.Thread.Sleep(200) 'Pause for 200ms
                        If Main.bgwSendMessage.IsBusy Then
                            Application.DoEvents()
                            System.Threading.Thread.Sleep(400) 'Pause for 400ms
                        End If
                    End If
                End If
                If Main.bgwSendMessage.IsBusy Then
                    Main.Message.AddWarning("Send Message backgroundworker is busy." & vbCrLf)
                Else
                    Dim SendMessageParams As New Main.clsSendMessageParams
                    SendMessageParams.ProjectNetworkName = Main.ProNetName
                    SendMessageParams.ConnectionName = ConnName
                    SendMessageParams.Message = XMsg
                    Main.bgwSendMessage.RunWorkerAsync(SendMessageParams)
                    'Main.Message.XAddText("Message sent to " & ConnName & " (ProNet: " & Main.ProNetName & ") " & ":" & vbCrLf, "XmlSentNotice") 'UPDATED 2Feb19
                    If Main.ShowXMessages Then
                        Main.Message.XAddText("Message sent to " & "[" & Main.ProNetName & "]." & ConnName & ":" & vbCrLf, "XmlSentNotice")
                        Main.Message.XAddXml(XMsg)
                        Main.Message.XAddText(vbCrLf, "Normal") 'Add extra line
                    End If
                End If
            End If
        End If
    End Sub

    Public Sub SendXMessageExt(ByVal ProNetName As String, ByVal ConnName As String, ByVal XMsg As String)
        'Send the XMsg to the application with the connection name ConnName and Project Network Name ProNetname.
        'This version can send the XMessage to a connection external to the current Project Network.
        If IsNothing(Main.client) Then
            Main.Message.Add("No client connection available!" & vbCrLf)
        Else
            If Main.client.State = ServiceModel.CommunicationState.Faulted Then
                Main.Message.Add("client state is faulted. Message not sent!" & vbCrLf)
            Else
                'Main.client.SendMessageAsync(ProNetName, ConnName, XMsg)
                If Main.bgwSendMessage.IsBusy Then
                    Application.DoEvents()
                    System.Threading.Thread.Sleep(100) 'Pause for 100ms
                    If Main.bgwSendMessage.IsBusy Then
                        Application.DoEvents()
                        System.Threading.Thread.Sleep(200) 'Pause for 200ms
                        If Main.bgwSendMessage.IsBusy Then
                            Application.DoEvents()
                            System.Threading.Thread.Sleep(400) 'Pause for 400ms
                        End If
                    End If
                End If
                If Main.bgwSendMessage.IsBusy Then
                    Main.Message.AddWarning("Send Message backgroundworker is busy." & vbCrLf)
                Else
                    Dim SendMessageParams As New Main.clsSendMessageParams
                    SendMessageParams.ProjectNetworkName = ProNetName
                    SendMessageParams.ConnectionName = ConnName
                    SendMessageParams.Message = XMsg
                    Main.bgwSendMessage.RunWorkerAsync(SendMessageParams)
                    'Main.Message.XAddText("Message sent to " & ConnName & " (ProNet: " & Main.ProNetName & ") " & ":" & vbCrLf, "XmlSentNotice") 'UPDATED 2Feb19
                    'Main.Message.XAddText("Message sent to " & ConnName & " (ProNet: " & ProNetName & ") " & ":" & vbCrLf, "XmlSentNotice")
                    If Main.ShowXMessages Then
                        Main.Message.XAddText("Message sent to " & "[" & ProNetName & "]." & ConnName & ":" & vbCrLf, "XmlSentNotice")
                        Main.Message.XAddXml(XMsg)
                        Main.Message.XAddText(vbCrLf, "Normal") 'Add extra line
                    End If
                End If
            End If
        End If
    End Sub

    Public Sub SendXMessageWait(ByVal ConnName As String, ByVal XMsg As String)
        'Send the XMsg to the application with the connection name ConnName.
        'Wait for the connection to be made.
        If IsNothing(Main.client) Then
            Main.Message.Add("No client connection available!" & vbCrLf)
        Else
            Try
                Application.DoEvents()

                If Main.client.State = ServiceModel.CommunicationState.Faulted Then
                    Main.Message.Add("client state is faulted. Message not sent!" & vbCrLf)
                Else
                    Dim StartTime As Date = Now
                    Dim Duration As TimeSpan
                    'Wait up to 16 seconds for the connection ConnName to be established
                    'While Main.client.ConnectionAvailable(Main.ProNetName, ConnName) '(When the specified connection is no longer available, the connection has been made and can be used.) - Consider changing this to ConnectionExists!
                    While Main.client.ConnectionExists(Main.ProNetName, ConnName) = False 'Wait until the required connection is made.
                        System.Threading.Thread.Sleep(1000) 'Pause for 1000ms
                        Duration = Now - StartTime
                        If Duration.Seconds > 16 Then Exit While
                    End While

                    'If Main.client.ConnectionAvailable(Main.ProNetName, ConnName) Then
                    If Main.client.ConnectionExists(Main.ProNetName, ConnName) = False Then
                        Main.Message.AddWarning("Connection not available: " & ConnName & " in application network: " & Main.ProNetName & vbCrLf)
                    Else
                        'Main.client.SendMessageAsync(Main.ProNetName, ConnName, XMsg)
                        If Main.bgwSendMessage.IsBusy Then
                            Application.DoEvents()
                            System.Threading.Thread.Sleep(100) 'Pause for 100ms
                            If Main.bgwSendMessage.IsBusy Then
                                Application.DoEvents()
                                System.Threading.Thread.Sleep(200) 'Pause for 200ms
                                If Main.bgwSendMessage.IsBusy Then
                                    Application.DoEvents()
                                    System.Threading.Thread.Sleep(400) 'Pause for 400ms
                                End If
                            End If
                        End If
                        If Main.bgwSendMessage.IsBusy Then
                            Main.Message.AddWarning("Send Message backgroundworker is busy." & vbCrLf)
                        Else
                            Dim SendMessageParams As New Main.clsSendMessageParams
                            SendMessageParams.ProjectNetworkName = Main.ProNetName
                            SendMessageParams.ConnectionName = ConnName
                            SendMessageParams.Message = XMsg
                            Main.bgwSendMessage.RunWorkerAsync(SendMessageParams)
                            'Main.Message.XAddText("Message sent to " & ConnName & " (ProNet: " & Main.ProNetName & ") " & ":" & vbCrLf, "XmlSentNotice")
                            If Main.ShowXMessages Then
                                Main.Message.XAddText("Message sent to " & "[" & Main.ProNetName & "]." & ConnName & ":" & vbCrLf, "XmlSentNotice")
                                Main.Message.XAddXml(XMsg)
                                Main.Message.XAddText(vbCrLf, "Normal") 'Add extra line
                            End If
                        End If
                    End If
                End If
            Catch ex As Exception
                Main.Message.AddWarning(ex.Message & vbCrLf)
            End Try
        End If
    End Sub

    Public Sub SendXMessageExtWait(ByVal ProNetName As String, ByVal ConnName As String, ByVal XMsg As String)
        'Send the XMsg to the application with the connection name ConnName and Project Network Name ProNetName.
        'Wait for the connection to be made.
        'This version can send the XMessage to a connection external to the current Project Network.
        If IsNothing(Main.client) Then
            Main.Message.Add("No client connection available!" & vbCrLf)
        Else
            If Main.client.State = ServiceModel.CommunicationState.Faulted Then
                Main.Message.Add("client state is faulted. Message not sent!" & vbCrLf)
            Else
                Dim StartTime As Date = Now
                Dim Duration As TimeSpan
                'Wait up to 16 seconds for the connection ConnName to be established
                'While Main.client.ConnectionAvailable(Main.ProNetName, ConnName)
                'While Main.client.ConnectionAvailable(ProNetName, ConnName)
                While Main.client.ConnectionExists(ProNetName, ConnName) = False
                    System.Threading.Thread.Sleep(1000) 'Pause for 1000ms
                    Duration = Now - StartTime
                    If Duration.Seconds > 16 Then Exit While
                End While

                'If Main.client.ConnectionAvailable(Main.ProNetName, ConnName) Then
                'If Main.client.ConnectionAvailable(ProNetName, ConnName) Then
                If Main.client.ConnectionExists(ProNetName, ConnName) = False Then
                    'Main.Message.AddWarning("Connection not available: " & ConnName & " in application network: " & Main.ProNetName & vbCrLf)
                    Main.Message.AddWarning("Connection not available: " & ConnName & " in application network: " & ProNetName & vbCrLf)
                Else
                    'Main.client.SendMessageAsync(ProNetName, ConnName, XMsg)
                    If Main.bgwSendMessage.IsBusy Then
                        Application.DoEvents()
                        System.Threading.Thread.Sleep(100) 'Pause for 100ms
                        If Main.bgwSendMessage.IsBusy Then
                            Application.DoEvents()
                            System.Threading.Thread.Sleep(200) 'Pause for 200ms
                            If Main.bgwSendMessage.IsBusy Then
                                Application.DoEvents()
                                System.Threading.Thread.Sleep(400) 'Pause for 400ms
                            End If
                        End If
                    End If
                    If Main.bgwSendMessage.IsBusy Then
                        Main.Message.AddWarning("Send Message backgroundworker is busy." & vbCrLf)
                    Else
                        Dim SendMessageParams As New Main.clsSendMessageParams
                        SendMessageParams.ProjectNetworkName = ProNetName
                        SendMessageParams.ConnectionName = ConnName
                        SendMessageParams.Message = XMsg
                        Main.bgwSendMessage.RunWorkerAsync(SendMessageParams)
                        'Main.Message.XAddText("Message sent to " & ConnName & " (ProNet: " & ProNetName & ") " & ":" & vbCrLf, "XmlSentNotice")
                        If Main.ShowXMessages Then
                            Main.Message.XAddText("Message sent to " & "[" & ProNetName & "]." & ConnName & ":" & vbCrLf, "XmlSentNotice")
                            Main.Message.XAddXml(XMsg)
                            Main.Message.XAddText(vbCrLf, "Normal") 'Add extra line
                        End If
                    End If
                End If
            End If
        End If
    End Sub

    Public Sub AddText(ByVal Msg As String, ByVal TextType As String)
        Main.Message.AddText(Msg, TextType)
    End Sub

    Public Sub AddMessage(ByVal Msg As String)
        Main.Message.Add(Msg)
    End Sub

    Public Sub AddWarning(ByVal Msg As String)
        Main.Message.AddWarning(Msg)
    End Sub

    Public Sub XMsgInstruction(ByVal Info As String, ByVal Locn As String)
        'Send the XMessage Instruction to the JavaScript function XMsgInstruction for processing.
        Me.WebBrowser1.Document.InvokeScript("XMsgInstruction", New String() {Info, Locn})
    End Sub

    'Public Sub OpenWebPage(ByVal WebPageFileName As String, ByVal ParentWebPageName As String, ByVal ParentWebPageFormNo As Integer)
    Public Sub OpenWebPage(ByVal WebPageFileName As String)
        'Open a Web Page from the WebPageFileName.
        '  Pass the ParentName Property to the new web page. The is the name of this web page that is opening the new page.
        '  Pass the ParentWebPageFormNo Property to the new web page. This is the FormNo of this web page that is opening the new page.
        '    A hash code is generated from the ParentName. This is used to define a file name to save and restore the Web Page settings.
        '    The new web page can pass instructions back to the ParentWebPage using its ParentWebPageFormNo.

        Dim NewFormNo As Integer = Main.OpenNewWebPage()

        Main.WebPageFormList(NewFormNo).ParentWebPageFileName = FileName 'Set the Parent Web Page property.
        Main.WebPageFormList(NewFormNo).ParentWebPageFormNo = FormNo 'Set the Parent Form Number property.
        Main.WebPageFormList(NewFormNo).Description = ""             'The web page description can be blank.
        Main.WebPageFormList(NewFormNo).FileDirectory = ""           'Only Web files in the Project directory can be opened from another Web Page Form.
        Main.WebPageFormList(NewFormNo).FileLocationType = LocationTypes.Project 'Only Web files in the Project directory can be opened from another Web Page Form.
        Main.WebPageFormList(NewFormNo).FileName = WebPageFileName  'Set the web page file name to be opened.
        Main.WebPageFormList(NewFormNo).OpenDocument                'Open the web page file name.

    End Sub

    Private Sub btnEdit_Click(sender As Object, e As EventArgs) Handles btnEdit.Click
        'Edit the html file.

        Dim FileName As String
        FileName = txtDocumentFile.Text

        If FileName = "" Then
            Main.Message.AddWarning("No page selected." & vbCrLf)
        Else
            Dim FileFound As Boolean = False
            If Main.HtmlDisplayFormList.Count = 0 Then

            Else
                Dim I As Integer
                For I = 0 To Main.HtmlDisplayFormList.Count - 1
                    If Main.HtmlDisplayFormList(I) Is Nothing Then

                    Else
                        If Main.HtmlDisplayFormList(I).FileName = FileName Then
                            FileFound = True
                            Main.HtmlDisplayFormList(I).BringToFront
                        End If
                    End If
                Next
            End If

            If FileFound = False Then
                Dim FormNo As Integer = Main.OpenNewHtmlDisplayPage()
                Main.HtmlDisplayFormList(FormNo).FileName = FileName
                Main.HtmlDisplayFormList(FormNo).OpenDocument
            End If

        End If

    End Sub

    Public Sub ParentProjectName(ByVal FormName As String, ByVal ItemName As String)
        'Return the Parent Project name:
        RestoreSetting(FormName, ItemName, Main.Project.ParentProjectName)
    End Sub

    Public Sub ParentProjectPath(ByVal FormName As String, ByVal ItemName As String)
        'Return the Parent Project path:
        RestoreSetting(FormName, ItemName, Main.Project.ParentProjectPath)
    End Sub

    Public Sub ParentProjectParameterValue(ByVal FormName As String, ByVal ItemName As String, ByVal ParameterName As String)
        'Return the specified Parent Project parameter value:
        'RestoreSetting(FormName, ItemName, Main.Project.ParentParameter(ParameterName).Value)
        RestoreSetting(FormName, ItemName, Main.Project.GetParentParameter(ParameterName))
    End Sub

    Public Sub ProjectParameterValue(ByVal FormName As String, ByVal ItemName As String, ByVal ParameterName As String)
        'Return the specified Project parameter value:
        'RestoreSetting(FormName, ItemName, Main.Project.Parameter(ParameterName).Value)
        RestoreSetting(FormName, ItemName, Main.Project.GetParameter(ParameterName))
    End Sub

    Public Function ProjectParameterValue(ByVal ParameterName As String) As String
        'Return the specified Project parameter value:
        'RestoreSetting(FormName, ItemName, Main.Project.Parameter(ParameterName).Value)
        Return Main.Project.Parameter(ParameterName).Value
    End Function

    Public Function GetProjectParameterValue(ByVal ParameterName As String) As String
        'Return the specified Project parameter value:
        'RestoreSetting(FormName, ItemName, Main.Project.Parameter(ParameterName).Value)
        Return Main.Project.Parameter(ParameterName).Value
    End Function

    'Public Sub ApplicationNetworkName(ByVal FormName As String, ByVal ItemName As String)
    '    'Return the name of the Application Network:
    '    'RestoreSetting(FormName, ItemName, Main.Project.Parameter("AppNetName").Value)
    '    RestoreSetting(FormName, ItemName, Main.Project.GetParameter("AppNetName"))
    'End Sub

    Public Sub ProjectNetworkName(ByVal FormName As String, ByVal ItemName As String)
        'Return the name of the Project Network:
        'RestoreSetting(FormName, ItemName, Main.Project.Parameter("AppNetName").Value)
        RestoreSetting(FormName, ItemName, Main.Project.GetParameter("ProNetName"))
    End Sub



    Public Sub CalcValsTable(ByVal TableName As String, ByVal FormName As String, ByVal ItemName As String)
        'Search for TableName in the Calculated Values database. 
        'If Found return TableName to FormName/ItemName.
        'If not found return "" to FormName/ItemName.

        'If Main.TableExists(Main.Project.Parameter("CalculationDatabasePath").Value, TableName) Then
        If Main.TableExists(Main.Project.GetParameter("CalculationsDatabasePath"), TableName) Then
            RestoreSetting(FormName, ItemName, TableName)
        Else
            RestoreSetting(FormName, ItemName, "")
        End If
    End Sub

    Public Sub SharePriceTable(ByVal TableName As String, ByVal FormName As String, ByVal ItemName As String)
        'Search for TableName in the Share Price database. 
        'If Found return TableName to FormName/ItemName.
        'If not found return "" to FormName/ItemName.

        'If Main.TableExists(Main.Project.Parameter("CalculationDatabasePath").Value, TableName) Then
        If Main.TableExists(Main.Project.GetParameter("SharePriceDatabasePath"), TableName) Then
            RestoreSetting(FormName, ItemName, TableName)
        Else
            RestoreSetting(FormName, ItemName, "")
        End If
    End Sub

    Public Sub Speak(ByVal Sentence As String)
        'Speak the sentence through the computer speaker:

        Dim mySynth As New System.Speech.Synthesis.SpeechSynthesizer

        mySynth.SetOutputToDefaultAudioDevice()

        'mySynth.SelectVoiceByHints(Speech.Synthesis.VoiceGender.Female, Speech.Synthesis.VoiceAge.Teen)
        'mySynth.SelectVoiceByHints(Speech.Synthesis.VoiceGender.Male, Speech.Synthesis.VoiceAge.Teen, 0, System.Globalization.CultureInfo.GetCultureInfo("fr-FR"))

        'mySynth.SelectVoiceByHints(Speech.Synthesis.VoiceGender.Male)
        'mySynth.Rate = 1

        mySynth.Speak(Sentence)

    End Sub

    'Public Sub CalculateDailyStatistics(ByVal TableName As String, ByVal DateValue As Date)
    Public Sub CalculateDailyStatistics(ByVal TableName As String, ByVal DateString As String)
        'Calculate the Daily share trading Statistics for the date in DateValue.
        'The results are stored in the Calculation database in the table named TableName.
        Dim DateValue As Date

        DateValue = CDate(DateString)

        Main.CalculateDailyStatistics(TableName, DateValue)
    End Sub

    Public Sub DailyStatsBetweenDates(ByVal TableName As String, ByVal FromDateString As String, ByVal ToDateString As String)
        'Calculate the Daily share trading Statistics between the dates in FromDateValue and ToDateValue.
        'The results are stored in the Calculation database in the table named TableName.

        Dim FromDateValue As Date = CDate(FromDateString)
        Dim ToDateValue As Date = CDate(ToDateString)

        Main.DailyStatsBetweenDates(TableName, FromDateValue, ToDateValue)

    End Sub

    Public Sub CompanyDailyStatsBetweenDates(ByVal TableName As String, ByVal AsxCode As String, ByVal FromDateString As String, ByVal ToDateString As String)
        'Calculate the Daily share trading Statistics between the dates in FromDateString and ToDateString for the single company with the code AsxCode.
        'The results are stored in the Share Price database in the table named TableName.

        'Dim FromDateValue As Date = CDate(FromDateString)
        'Dim ToDateValue As Date = CDate(ToDateString)

        'Main.CompanyDailyStatsBetweenDates(TableName, AsxCode, FromDateValue, ToDateValue)
        Main.CompanyDailyStatsBetweenDates(TableName, AsxCode, FromDateString, ToDateString)

    End Sub

    Public Sub CalcSharePriceDistAllDays(ByVal TableName As String, ByVal AsxCode As String, ByVal FromDateString As String, ByVal ToDateString As String, ByVal WindowLen As Integer)
        'Calculate the closing share price distributions between the dates in FromDateString and ToDateString for the single company with the code AsxCode.
        'The mean and standard deviation of the closing prices are calculated over a woindow length of side WindowLen days.
        'The trade volume is also analysed.
        'The WindowLength includes all days, trade days and non-trade days. (Non-trade days include weekends, holidays and trading halt days.)

        Main.CalcSharePriceDistAllDays(TableName, AsxCode, FromDateString, ToDateString, WindowLen)
    End Sub

    Public Sub CalcSharePriceDistTradeDays(ByVal TableName As String, ByVal AsxCode As String, ByVal FromDateString As String, ByVal ToDateString As String, ByVal WindowLen As Integer)
        'Calculate the closing share price distributions between the dates in FromDateString and ToDateString for the single company with the code AsxCode.
        'The mean and standard deviation of the closing prices are calculated over a woindow length of side WindowLen days.
        'The trade volume is also analysed.
        'The WindowLength includes only trade days. Non-trade days are excluded from the statistics. (Non-trade days include weekends, holidays and trading halt days.)

        'FOR DEBUGGING:
        Main.Message.Add("Running CalcSharePriceDistTradeDays " & vbCrLf)
        Main.Message.Add("TableName = " & TableName & vbCrLf)
        Main.Message.Add("AsxCode = " & AsxCode & vbCrLf)
        Main.Message.Add("FromDateString = " & FromDateString & vbCrLf)
        Main.Message.Add("ToDateString = " & ToDateString & vbCrLf)
        Main.Message.Add("WindowLen = " & WindowLen & vbCrLf)

        Main.CalcSharePriceDistTradeDays(TableName, AsxCode, FromDateString, ToDateString, WindowLen)
    End Sub

    Public Sub AllCompanyDailyStatsBetweenDates(ByVal TableName As String, ByVal FromDateString As String, ByVal ToDateString As String)
        'Calculate the Daily share trading Statistics between the dates in FromDateValue and ToDateValue for all companies traded.
        'The results are stored in the Share Price database in the table named TableName.

        Dim FromDateValue As Date = CDate(FromDateString)
        Dim ToDateValue As Date = CDate(ToDateString)

        Main.AllCompanyDailyStatsBetweenDates(TableName, FromDateValue, ToDateValue)

    End Sub

    Public Sub DailyStatsDayNos(ByVal TableName As String)
        'Caluclate the Rel_Day_No and Trade_Day_No values in the Daily share trading Statistics table.

        Main.DailyStatsDayNos(TableName)

    End Sub

    'Public Sub GetDividendPaymentInfo(ByVal AsxCode As String)
    '    'Get dividend payment information for the company with code AsxCode.
    '    'The dividend info is stored in the table Dividend_Payment_Info in the Calculations database.

    '    Main.GetDividendPaymentInfo(AsxCode)

    'End Sub

    Public Sub DividendReboundAnalysis(ByVal AsxCode As String, ByVal MinBuyDelay As Integer, ByVal MaxBuyDelay As Integer, ByVal MinSellDelay As Integer, ByVal MaxSellDelay As Integer, ByVal BuyBrokPct As Single, ByVal SellBrokPct As Single)
        'Analyse the dividend rebound effect.
        Main.DividendReboundAnalysis(AsxCode, MinBuyDelay, MaxBuyDelay, MinSellDelay, MaxSellDelay, BuyBrokPct, SellBrokPct)
    End Sub

    Public Sub CalcDividendPaymentInfo(ByVal AsxCode As String)
        'Genrate the dividend payment records for the AsxCode.
        Main.CalcDividendPaymentInfo(AsxCode)
    End Sub


    Public Sub OpenProjectAtRelativePath(ByVal RelativePath As String, ByVal ConnectionName As String)
        'Open the Project at the specified Relative Path using the specified Connection Name.

        Dim ProjectPath As String
        If RelativePath.StartsWith("\") Then
            ProjectPath = Main.Project.Path & RelativePath
            Main.client.StartProjectAtPath(ProjectPath, ConnectionName)
        Else
            ProjectPath = Main.Project.Path & "\" & RelativePath
            Main.client.StartProjectAtPath(ProjectPath, ConnectionName)
        End If
    End Sub

    Public Sub CheckOpenProjectAtRelativePath(ByVal RelativePath As String, ByVal ConnectionName As String)
        'Check if the project at the specified Relative Path is open.
        'Open it if it is not already open.
        'Open the Project at the specified Relative Path using the specified Connection Name.
        Try
            Application.DoEvents()
            Dim ProjectPath As String
            If RelativePath.StartsWith("\") Then
                ProjectPath = Main.Project.Path & RelativePath
                If Main.client.ProjectOpen(ProjectPath) Then
                    'Project is already open.
                Else
                    Main.client.StartProjectAtPath(ProjectPath, ConnectionName)
                End If
            Else
                ProjectPath = Main.Project.Path & "\" & RelativePath
                If Main.client.ProjectOpen(ProjectPath) Then 'An exception of type 'System.InvalidOperationException' occurred in mscorlib.dll but was not handled in user code
                    'Additional Information: This operation would deadlock because the reply cannot be received until the current Message completes processing. If you want To allow out-Of-order message processing, specify ConcurrencyMode of Reentrant Or Multiple on CallbackBehaviorAttribute.
                    'Project is already open.
                Else
                    Main.client.StartProjectAtPath(ProjectPath, ConnectionName)
                End If
            End If
        Catch ex As Exception
            Main.Message.AddWarning(ex.Message & vbCrLf)
        End Try
    End Sub

    'Public Sub OpenProjectAtAppNetPath(ByVal RelativePath As String, ByVal ConnectionName As String)
    '    'Open the Project at the specified Path (relative to the AppNet Path) using the specified Connection Name.

    '    Dim ProjectPath As String
    '    If RelativePath.StartsWith("\") Then
    '        If Main.Project.ParameterExists("AppNetPath") Then
    '            ProjectPath = Main.Project.GetParameter("AppNetPath") & RelativePath
    '            Main.client.StartProjectAtPath(ProjectPath, ConnectionName)
    '        Else
    '            Main.Message.AddWarning("The Application Network Path is not known." & vbCrLf)
    '        End If
    '    Else
    '        If Main.Project.ParameterExists("AppNetPath") Then
    '            ProjectPath = Main.Project.GetParameter("AppNetPath") & "\" & RelativePath
    '            Main.client.StartProjectAtPath(ProjectPath, ConnectionName)
    '        Else
    '            Main.Message.AddWarning("The Application Network Path is not known." & vbCrLf)
    '        End If
    '    End If
    'End Sub

    Public Sub OpenProjectAtProNetPath(ByVal RelativePath As String, ByVal ConnectionName As String)
        'Open the Project at the specified Path (relative to the ProNet Path) using the specified Connection Name.

        Dim ProjectPath As String
        If RelativePath.StartsWith("\") Then
            If Main.Project.ParameterExists("ProNetPath") Then
                ProjectPath = Main.Project.GetParameter("ProNetPath") & RelativePath
                Main.client.StartProjectAtPath(ProjectPath, ConnectionName)
            Else
                Main.Message.AddWarning("The Project Network Path is not known." & vbCrLf)
            End If
        Else
            If Main.Project.ParameterExists("ProNetPath") Then
                ProjectPath = Main.Project.GetParameter("ProNetPath") & "\" & RelativePath
                Main.client.StartProjectAtPath(ProjectPath, ConnectionName)
            Else
                Main.Message.AddWarning("The Project Network Path is not known." & vbCrLf)
            End If
        End If
    End Sub

    'Public Sub CheckOpenProjectAtAppNetPath(ByVal RelativePath As String, ByVal ConnectionName As String)
    '    'Check if the project at the specified Path (relative to the AppNet Path) is open.
    '    'Open it if it is not already open.
    '    'Open the Project at the specified Path using the specified Connection Name.

    '    Dim ProjectPath As String
    '    If RelativePath.StartsWith("\") Then
    '        If Main.Project.ParameterExists("AppNetPath") Then
    '            ProjectPath = Main.Project.GetParameter("AppNetPath") & RelativePath
    '            'Main.client.StartProjectAtPath(ProjectPath, ConnectionName)
    '            If Main.client.ProjectOpen(ProjectPath) Then
    '                'Project is already open.
    '            Else
    '                Main.client.StartProjectAtPath(ProjectPath, ConnectionName)
    '            End If
    '        Else
    '            Main.Message.AddWarning("The Application Network Path is not known." & vbCrLf)
    '        End If
    '    Else
    '        If Main.Project.ParameterExists("AppNetPath") Then
    '            ProjectPath = Main.Project.GetParameter("AppNetPath") & "\" & RelativePath
    '            'Main.client.StartProjectAtPath(ProjectPath, ConnectionName)CheckOpen
    '            If Main.client.ProjectOpen(ProjectPath) Then
    '                'Project is already open.
    '            Else
    '                Main.client.StartProjectAtPath(ProjectPath, ConnectionName)
    '            End If
    '        Else
    '            Main.Message.AddWarning("The Application Network Path is not known." & vbCrLf)
    '        End If
    '    End If
    'End Sub

    Public Sub CheckOpenProjectAtProNetPath(ByVal RelativePath As String, ByVal ConnectionName As String)
        'Check if the project at the specified Path (relative to the ProNet Path) is open.
        'Open it if it is not already open.
        'Open the Project at the specified Path using the specified Connection Name.

        Dim ProjectPath As String
        If RelativePath.StartsWith("\") Then
            If Main.Project.ParameterExists("ProNetPath") Then
                ProjectPath = Main.Project.GetParameter("ProNetPath") & RelativePath
                If Main.client.ProjectOpen(ProjectPath) Then
                    'Project is already open.
                Else
                    Main.client.StartProjectAtPath(ProjectPath, ConnectionName)
                End If
            Else
                Main.Message.AddWarning("The Project Network Path is not known." & vbCrLf)
            End If
        Else
            If Main.Project.ParameterExists("ProNetPath") Then
                ProjectPath = Main.Project.GetParameter("ProNetPath") & "\" & RelativePath
                If Main.client.ProjectOpen(ProjectPath) Then
                    'Project is already open.
                Else
                    Main.client.StartProjectAtPath(ProjectPath, ConnectionName)
                End If
            Else
                Main.Message.AddWarning("The Project Network Path is not known." & vbCrLf)
            End If
        End If
    End Sub

    Public Function GetSharePriceValue(ByVal AsxCode As String, ByVal DateString As String, ByVal PriceType As String) As Single
        'Return a share price value at the specified date.
        Return Main.GetSharePriceValue(AsxCode, DateString, PriceType)

    End Function

    'Public Sub ShowFirstSharePriceValue(ByVal AsxCode As String, ByVal DateString As String, ByVal NDays As Integer, ByRef DateValue As Date, ByVal DateFormName As String, ByVal DateItemName As String, ByVal PriceType As String, ByRef PriceValue As Single, ByVal PriceFormName As String, ByVal PriceItemName As String, ByRef Status As String, ByVal StatusFormName As String, ByVal StatusItemName As String)
    Public Sub ShowFirstSharePriceValue(ByVal AsxCode As String, ByVal DateString As String, ByVal NDays As Integer, ByVal PriceType As String, ByVal PriceFormName As String, ByVal PriceItemName As String, ByVal DateFormName As String, ByVal DateItemName As String, ByVal StatusFormName As String, ByVal StatusItemName As String)
        'Get the first available Share price in the time window defined by the DateString and NDays.
        'Return the price date and status in the PriceValue, DateValue and Status variables.
        'The price, date and status values are then sent to the specified web page locations (PriceFormName, PriceItemName etc).

        Dim PriceValue As Single 'Stores the value of the first share price found in the date window
        Dim DateValue As Date    'Stores the date of the first share price found in the date window
        Dim Status As String     'Stores the status of the search for the first share price

        Main.GetFirstSharePriceValue(AsxCode, DateString, NDays, DateValue, PriceType, PriceValue, Status)

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


    'Public Sub CloseAppAtConnection(ByVal AppNetName As String, ByVal ConnectionName As String)
    '    'Close the application at the specified connection.
    '    Main.CloseAppAtConnection(AppNetName, ConnectionName)
    'End Sub

    'Public Sub CloseAppAtConnection(ByVal ProNetName As String, ByVal ConnectionName As String)
    Public Sub CloseProjectAtConnection(ByVal ProNetName As String, ByVal ConnectionName As String)
        'Close the application at the specified connection.
        'Main.CloseAppAtConnection(ProNetName, ConnectionName)
        Main.CloseProjectAtConnection(ProNetName, ConnectionName)
    End Sub


    'Public Sub SetEndInstructionInfo(ByVal EndInstructionInfo As String)
    '    Main.EndInstructionInfo = EndInstructionInfo
    'End Sub

    'Public Sub SetEndInstructionLocn(ByVal EndInstructionLocn As String)
    '    Main.EndInstructionLocn = EndInstructionLocn
    'End Sub

#End Region 'Methods Called by JavaScript -----------------------------------------------------------------------------------------------------------------------------------------------------

#End Region 'Form Methods ---------------------------------------------------------------------------------------------------------------------------------------------------------------------

End Class