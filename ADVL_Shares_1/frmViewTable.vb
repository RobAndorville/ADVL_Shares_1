Public Class frmViewTable
    'The View Table form is used to view tables in the Financials, Share Prices, News or Calculations databases.

#Region " Variable Declarations - All the variables used in this form and this application." '-------------------------------------------------------------------------------------------------

    'Variables used to connect to a database and open a table:
    Dim connString As String
    Public myConnection As OleDb.OleDbConnection = New OleDb.OleDbConnection
    Public ds As DataSet = New DataSet
    Dim da As OleDb.OleDbDataAdapter
    Dim tables As DataTableCollection = ds.Tables

    Dim UpdateNeeded As Boolean

    Public WithEvents DesignQuery As frmDesignQuery

#End Region 'Variable Declarations ------------------------------------------------------------------------------------------------------------------------------------------------------------

#Region " Properties - All the properties used in this form and this application" '------------------------------------------------------------------------------------------------------------

    'The FormNo property stores the number of the instance of this form.
    'This form can have multipe instances, which are stored in the SharePricesList ArrayList in the ADVL_Shares_1 Main form.
    'When this form is closed, the FormNo is used to update the ClosedFormNo property of the Main form.
    'ClosedFormNo is then used by a method to set the corresponding form element in SharePricesList to Nothing.
    Private _formNo As Integer
    Public Property FormNo As Integer
        Get
            Return _formNo
        End Get
        Set(ByVal value As Integer)
            _formNo = value
        End Set
    End Property

    Private _query As String = "" 'The Query property stores the text of the SQL query used to display table values in DataGridView1
    'Public Property Query() As String
    Public Property Query As String
        Get
            Return _query
        End Get
        Set(ByVal value As String)
            _query = value
            txtQuery.Text = _query
        End Set
    End Property

#End Region 'Properties -----------------------------------------------------------------------------------------------------------------------------------------------------------------------

#Region " Process XML files - Read and write XML files." '-------------------------------------------------------------------------------------------------------------------------------------

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
                               <Query><%= Query %></Query>
                               <Description><%= txtDataDescr.Text %></Description>
                               <AutoApplyQuery><%= chkAutoApply.Checked.ToString %></AutoApplyQuery>
                           </FormSettings>

        'Add code to include other settings to save after the comment line <!---->

        'Dim SettingsFileName As String = "FormSettings_" & Main.ApplicationInfo.Name & "_" & Me.Text & ".xml"
        'Multiple form version of the SettingsFileName:
        Dim SettingsFileName As String = "FormSettings_" & Main.ApplicationInfo.Name & "_" & Me.Text & "_" & FormNo & ".xml"
        Main.Project.SaveXmlSettings(SettingsFileName, settingsData)
    End Sub

    Private Sub RestoreFormSettings()
        'Read the form settings from an XML document.

        'Dim SettingsFileName As String = "FormSettings_" & Main.ApplicationInfo.Name & "_" & Me.Text & ".xml"
        'Multiple form version of the SettingsFileName:
        Dim SettingsFileName As String = "FormSettings_" & Main.ApplicationInfo.Name & "_" & Me.Text & "_" & FormNo & ".xml"

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
            If Settings.<FormSettings>.<Query>.Value <> Nothing Then Query = Settings.<FormSettings>.<Query>.Value
            If Settings.<FormSettings>.<Description>.Value <> Nothing Then txtDataDescr.Text = Settings.<FormSettings>.<Description>.Value
            If Settings.<FormSettings>.<AutoApplyQuery>.Value <> Nothing Then chkAutoApply.Checked = Settings.<FormSettings>.<AutoApplyQuery>.Value

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

#Region " Form Display Methods - Code used to display this form." '----------------------------------------------------------------------------------------------------------------------------

    'Private Sub frmTemplate_Load(sender As Object, e As EventArgs) Handles Me.Load
    Private Sub Form_Load(sender As Object, e As EventArgs) Handles Me.Load
        FillCmbSelectTable()
        RestoreFormSettings()   'Restore the form settings
        If chkAutoApply.Checked Then
            ApplyQuery()
        End If
    End Sub

    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        'Exit the Form
        Main.ClosedFormNo = FormNo 'The Main form property ClosedFormNo is set to this form number. This is used in the SharePricesFormClosed method to select the correct form to set to nothing.
        Me.Close() 'Close the form
    End Sub

    'Private Sub frmTemplate_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
    Private Sub Form_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        If WindowState = FormWindowState.Normal Then
            SaveFormSettings()
        Else
            'Dont save settings if form is minimised.
        End If
    End Sub

    'Private Sub frmSharePrices_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
    '    Main.FinancialsFormClosed()
    'End Sub

    'Private Sub frmViewTable_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
    '    Main.ViewTableFormClosed()
    'End Sub

    Public Sub CloseForm()
        'Used to close the form remotely.
        Main.ClosedFormNo = FormNo 'The Main form property ClosedFormNo is set to this form number. This is used in the SharePricesFormClosed method to select the correct form to set to nothing.
        Me.Close() 'Close the form
    End Sub

#End Region 'Form Display Methods -------------------------------------------------------------------------------------------------------------------------------------------------------------

#Region " Form Methods - The main actions performed by this form." '---------------------------------------------------------------------------------------------------------------------------

    'Private Sub cmbSelectTable_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbSelectTable.SelectedIndexChanged
    '    'Update DataGridView1:

    '    If IsNothing(cmbSelectTable.SelectedItem) Then
    '        Exit Sub
    '    End If

    '    TableName = cmbSelectTable.SelectedItem.ToString
    '    'Query = "Select Top 500 * From " & TableName
    '    _query = "Select Top 500 * From " & TableName 'This sets the value of  the Query property without running the associated method.
    '    txtQuery.Text = Query

    '    If cmbSelectTable.Focused Then
    '        connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source =" & Main.SharePriceDbPath 'DatabasePath
    '        myConnection.ConnectionString = connString
    '        myConnection.Open()

    '        da = New OleDb.OleDbDataAdapter(Query, myConnection)

    '        da.MissingSchemaAction = MissingSchemaAction.AddWithKey 'This statement is required to obtain the correct result from the statement: ds.Tables(0).Columns(0).MaxLength (This fixes a Microsoft bug: http://support.microsoft.com/kb/317175 )

    '        ds.Clear()
    '        ds.Reset()

    '        da.FillSchema(ds, SchemaType.Source, TableName)

    '        da.Fill(ds, TableName)

    '        DataGridView1.AutoGenerateColumns = True

    '        DataGridView1.EditMode = DataGridViewEditMode.EditOnKeystroke

    '        DataGridView1.DataSource = ds.Tables(0)
    '        DataGridView1.AutoResizeColumns()

    '        DataGridView1.Update()
    '        DataGridView1.Refresh()
    '        myConnection.Close()
    '    End If

    'End Sub

    Public Sub FillCmbSelectTable()
        'Fill the cmbSelectTable listbox with the available tables in the selected database.

        'If DatabasePath = "" Then
        If Main.FinancialsDbPath = "" Then
            Main.Message.AddWarning("No database selected!" & vbCrLf)
            Exit Sub
        End If

        'Database access for MS Access:
        Dim connectionString As String 'Declare a connection string for MS Access - defines the database or server to be used.
        Dim conn As System.Data.OleDb.OleDbConnection 'Declare a connection for MS Access - used by the Data Adapter to connect to and disconnect from the database.
        Dim dt As DataTable

        cmbSelectTable.Text = ""
        cmbSelectTable.Items.Clear()
        ds.Clear()
        ds.Reset()
        DataGridView1.Columns.Clear()

        'Specify the connection string:
        'Access 2003
        'connectionString = "provider=Microsoft.Jet.OLEDB.4.0;" + _
        '"data source = " + txtDatabase.Text

        'Access 2007:
        connectionString = "provider=Microsoft.ACE.OLEDB.12.0;" +
        "data source = " + Main.FinancialsDbPath

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

    Private Sub txtQuery_LostFocus(sender As Object, e As EventArgs) Handles txtQuery.LostFocus
        'Update the _query value:
        _query = txtQuery.Text
    End Sub

    Private Sub btnApplyQuery_Click(sender As Object, e As EventArgs) Handles btnApplyQuery.Click
        'Update DataGridView1:
        ApplyQuery()
    End Sub

    Private Sub ApplyQuery()
        'If IsNothing(cmbSelectTable.SelectedItem) Then
        '    Exit Sub
        'End If

        If Main.FinancialsDbPath = "" Then
            Main.Message.AddWarning("A Financials database has not been selected!" & vbCrLf)
            Exit Sub
        End If

        connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source =" & Main.FinancialsDbPath 'DatabasePath
        myConnection.ConnectionString = connString
        myConnection.Open()

        'da = New OleDb.OleDbDataAdapter(txtQuery.Text, myConnection)
        da = New OleDb.OleDbDataAdapter(Query, myConnection)

        da.MissingSchemaAction = MissingSchemaAction.AddWithKey

        ds.Clear()
        ds.Reset()
        Try

            'da.FillSchema(ds, SchemaType.Source, "myData")

            'da.Fill(ds, TableName)
            da.Fill(ds, "myData")

            DataGridView1.AutoGenerateColumns = True

            DataGridView1.EditMode = DataGridViewEditMode.EditOnKeystroke

            DataGridView1.DataSource = ds.Tables(0)
            DataGridView1.AutoResizeColumns()

            DataGridView1.Update()
            DataGridView1.Refresh()
        Catch ex As Exception
            Main.Message.Add("Error applying query." & vbCrLf)
            Main.Message.Add(ex.Message & vbCrLf & vbCrLf)
        End Try

        myConnection.Close()
    End Sub

    'Private Sub txtDataDescr_LostFocus(sender As Object, e As EventArgs) Handles txtDataDescr.LostFocus
    '    'Update the description of the data shown on this Comapny Financials form:
    '    'Main.FinancialsData(FormNo, txtDataDescr.Text)
    '    Main.UpdateFinancialsDataDescr(FormNo, txtFinancialDataDescr.Text)
    'End Sub

    Private Sub btnDisplay_Click(sender As Object, e As EventArgs) Handles btnDisplay.Click
        'Update DataGridView1:

        If IsNothing(cmbSelectTable.SelectedItem) Then
            Exit Sub
        End If

        Dim TableName As String = cmbSelectTable.SelectedItem.ToString
        'Query = "Select Top 500 * From " & TableName
        _query = "Select Top 500 * From " & TableName 'This sets the value of  the Query property without running the associated method.
        txtQuery.Text = Query

        If cmbSelectTable.Focused Then
            connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source =" & Main.FinancialsDbPath 'DatabasePath
            myConnection.ConnectionString = connString
            myConnection.Open()

            da = New OleDb.OleDbDataAdapter(Query, myConnection)

            da.MissingSchemaAction = MissingSchemaAction.AddWithKey 'This statement is required to obtain the correct result from the statement: ds.Tables(0).Columns(0).MaxLength (This fixes a Microsoft bug: http://support.microsoft.com/kb/317175 )

            ds.Clear()
            ds.Reset()

            da.FillSchema(ds, SchemaType.Source, TableName)

            da.Fill(ds, TableName)

            DataGridView1.AutoGenerateColumns = True

            DataGridView1.EditMode = DataGridViewEditMode.EditOnKeystroke

            DataGridView1.DataSource = ds.Tables(0)
            DataGridView1.AutoResizeColumns()

            DataGridView1.Update()
            DataGridView1.Refresh()
            myConnection.Close()
        End If
    End Sub

    Private Sub btnSaveChanges_Click(sender As Object, e As EventArgs) Handles btnSaveChanges.Click
        'Save the changes made to the data in DataGridView1 to the corresponding table in the database:
        Dim cb = New OleDb.OleDbCommandBuilder(da)
        Try
            DataGridView1.EndEdit()
            da.Update(ds.Tables(0))
            ds.Tables(0).AcceptChanges()
            UpdateNeeded = False
            DataGridView1.EditMode = DataGridViewEditMode.EditOnKeystroke
        Catch ex As Exception
            Main.Message.AddWarning("Error saving changes." & vbCrLf)
            Main.Message.AddWarning(ex.Message & vbCrLf & vbCrLf)
        End Try
    End Sub

#End Region 'Form Methods ---------------------------------------------------------------------------------------------------------------------------------------------------------------------




End Class