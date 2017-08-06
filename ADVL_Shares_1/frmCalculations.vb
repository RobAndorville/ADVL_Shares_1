Public Class frmCalculations
    'The Calculations form is used to view Calculations tables from the Calculations database.

#Region " Variable Declarations - All the variables used in this form and this application." '=================================================================================================

    'Declare forms opened from this form:
    Public WithEvents DesignCalculationsQuery As frmDesignQuery

    'Variables used to connect to a database and open a table:
    Dim connString As String
    Public myConnection As OleDb.OleDbConnection = New OleDb.OleDbConnection
    Public ds As DataSet = New DataSet
    Dim da As OleDb.OleDbDataAdapter
    Dim tables As DataTableCollection = ds.Tables

    Dim UpdateNeeded As Boolean

#End Region 'Variable Declarations ------------------------------------------------------------------------------------------------------------------------------------------------------------

#Region " Properties - All the properties used in this form and this application" '============================================================================================================

    'The FormNo property stores the number of the instance of this form.
    'This form can have multipe instances, which are stored in the CalculationsFormList ArrayList in the ADVL_Shares_1 Main form.
    'When this form is closed, the FormNo is used to update the ClosedFormNo property of the Main form.
    'ClosedFormNo is then used by a method to set the corresponding form element in CalculationsList to Nothing.

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

#Region " Process XML files - Read and write XML files." '=====================================================================================================================================

    Private Sub SaveFormSettings()
        'This SaveFormSettings method saves the settings in Main.CalculationsSettingsList

        If FormNo + 1 > Main.CalculationsSettings.List.Count Then
            Main.Message.AddWarning("Form number: " & FormNo & " does not exist in the Calculations Settings List!" & vbCrLf)
        Else
            'Save the form settings:
            Main.CalculationsSettings.List(FormNo).Left = Me.Left
            Main.CalculationsSettings.List(FormNo).Top = Me.Top
            Main.CalculationsSettings.List(FormNo).Width = Me.Width
            Main.CalculationsSettings.List(FormNo).Height = Me.Height
            Main.CalculationsSettings.List(FormNo).Query = Query
            Main.CalculationsSettings.List(FormNo).Description = txtDataDescr.Text
            Main.CalculationsSettings.List(FormNo).VersionName = txtVersionName.Text
            Main.CalculationsSettings.List(FormNo).VersionDesc = txtVersionDesc.Text
            Main.CalculationsSettings.List(FormNo).AutoApplyQuery = chkAutoApply.Checked.ToString
            Main.CalculationsSettings.List(FormNo).SelectedTab = TabControl1.SelectedIndex
            Main.CalculationsSettings.List(FormNo).SaveFileDir = Trim(txtDirectory.Text)
            Main.CalculationsSettings.List(FormNo).XmlFileName = Trim(txtXmlFileName.Text)
        End If
    End Sub

    Private Sub RestoreFormSettings()
        'This RestoreFormSettings method restores the settings from Main.CalculationsSettings.List

        If FormNo + 1 > Main.CalculationsSettings.List.Count Then
            'Add form entry to the Calculations Settings list.
            Dim NewSettings As New DataViewSettings
            Main.CalculationsSettings.InsertSettings(FormNo, NewSettings)
        Else
            'Restore the form settings:
            Me.Left = Main.CalculationsSettings.List(FormNo).Left
            Me.Top = Main.CalculationsSettings.List(FormNo).Top
            Me.Width = Main.CalculationsSettings.List(FormNo).Width
            Me.Height = Main.CalculationsSettings.List(FormNo).Height
            Query = Main.CalculationsSettings.List(FormNo).Query
            txtDataDescr.Text = Main.CalculationsSettings.List(FormNo).Description
            txtVersionName.Text = Main.CalculationsSettings.List(FormNo).VersionName
            txtDataVersion.Text = Main.CalculationsSettings.List(FormNo).VersionName
            txtVersionDesc.Text = Main.CalculationsSettings.List(FormNo).VersionDesc
            chkAutoApply.Checked = Main.CalculationsSettings.List(FormNo).AutoApplyQuery
            TabControl1.SelectedIndex = Main.CalculationsSettings.List(FormNo).SelectedTab
            txtDirectory.Text = Main.CalculationsSettings.List(FormNo).SaveFileDir
            txtXmlFileName.Text = Main.CalculationsSettings.List(FormNo).XmlFileName
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
        FillCmbSelectTable()
        RestoreFormSettings()   'Restore the form settings
        If chkAutoApply.Checked Then
            ApplyQuery()
        End If

        ShowVersionList()

    End Sub

    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        'Exit the Form
        Main.ClosedFormNo = FormNo 'The Main form property ClosedFormNo is set to this form number. This is used in the CalculationsFormClosed method to select the correct form to set to nothing.
        Me.Close() 'Close the form
    End Sub

    Private Sub Form_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        If WindowState = FormWindowState.Normal Then
            SaveFormSettings()
        Else
            'Dont save settings if form is minimised.
        End If
    End Sub

    Private Sub frmCalculations_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        Main.CalculationsFormClosed()
    End Sub

    Public Sub CloseForm()
        'Used to close the form remotely.
        Main.ClosedFormNo = FormNo 'The Main form property ClosedFormNo is set to this form number. This is used in the CalculationsFormClosed method to select the correct form to set to nothing.
        Me.Close() 'Close the form
    End Sub

#End Region 'Form Display Methods -------------------------------------------------------------------------------------------------------------------------------------------------------------

#Region " Open and Close Forms - Code used to open and close other forms." '===================================================================================================================

    Private Sub btnDesign_Click(sender As Object, e As EventArgs) Handles btnDesign.Click
        'Open the Design Query form
        If IsNothing(DesignCalculationsQuery) Then
            DesignCalculationsQuery = New frmDesignQuery
            DesignCalculationsQuery.Text = "Design Calculations Query"
            DesignCalculationsQuery.Show()
            DesignCalculationsQuery.DatabasePath = Main.CalculationsDbPath
        Else
            DesignCalculationsQuery.Show()
        End If
        Main.CalculationsSettingsChanged = True
    End Sub



    Private Sub DesignCalculationsQuery_FormClosed(sender As Object, e As FormClosedEventArgs) Handles DesignCalculationsQuery.FormClosed
        DesignCalculationsQuery = Nothing
    End Sub

#End Region 'Open and Close Forms -------------------------------------------------------------------------------------------------------------------------------------------------------------

#Region " Form Methods - The main actions performed by this form." '===========================================================================================================================

    Public Sub FillCmbSelectTable()
        'Fill the cmbSelectTable listbox with the available tables in the selected database.

        'If Main.SharePriceDbPath = "" Then
        If Main.CalculationsDbPath = "" Then
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
        "data source = " + Main.CalculationsDbPath

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

        If Main.CalculationsDbPath = "" Then
            Main.Message.AddWarning("A Calculations database has not been selected!" & vbCrLf)
            Exit Sub
        End If

        connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source =" & Main.CalculationsDbPath 'DatabasePath
        myConnection.ConnectionString = connString
        myConnection.Open()

        da = New OleDb.OleDbDataAdapter(Query, myConnection)

        da.MissingSchemaAction = MissingSchemaAction.AddWithKey

        ds.Clear()
        ds.Reset()

        Try
            da.Fill(ds, "myData")

            DataGridView1.AutoGenerateColumns = True

            DataGridView1.EditMode = DataGridViewEditMode.EditOnKeystroke

            DataGridView1.DataSource = ds.Tables(0)
            DataGridView1.AutoResizeColumns()

            DataGridView1.Update()
            DataGridView1.Refresh()
            Main.CalculationsSettings.List(FormNo).Query = Query

            'Update list of TableColumns
            Main.CalculationsSettings.List(FormNo).TableCols.Clear() 'Clear the old list of TableCols
            For Each item In ds.Tables(0).Columns
                Main.CalculationsSettings.List(FormNo).TableCols.Add(item.Columnname)
            Next

        Catch ex As Exception
            Main.Message.Add("Error applying query." & vbCrLf)
            Main.Message.Add(ex.Message & vbCrLf & vbCrLf)
        End Try

        If ds.Tables.Count > 0 Then
            txtNRecords.Text = ds.Tables(0).Rows.Count
            If DataGridView1.SelectedCells.Count > 0 Then
                txtSelectedRecord.Text = DataGridView1.SelectedCells.Item(0).RowIndex
            Else
                txtSelectedRecord.Text = ""
            End If
        Else
            txtSelectedRecord.Text = ""
        End If

        myConnection.Close()
    End Sub

    Private Sub txtDataDescr_LostFocus(sender As Object, e As EventArgs) Handles txtDataDescr.LostFocus
        'Update the description of the data shown on this Calculations form:
        Main.UpdateCalculationsDataDescr(FormNo, txtDataDescr.Text)
    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        If DataGridView1.SelectedCells.Count > 0 Then
            txtSelectedRecord.Text = DataGridView1.SelectedCells.Item(0).RowIndex
        Else
            txtSelectedRecord.Text = ""
        End If
    End Sub

    Private Sub btnDisplay_Click(sender As Object, e As EventArgs) Handles btnDisplay.Click
        'Update DataGridView1:

        If IsNothing(cmbSelectTable.SelectedItem) Then
            Exit Sub
        End If

        Dim TableName As String = cmbSelectTable.SelectedItem.ToString
        _query = "Select Top 500 * From " & TableName 'Modifying _query instead of Query sets the value of  the Query property without running the associated method.
        txtQuery.Text = Query
        Main.CalculationsSettingsChanged = True

        'Display the data in DataGridView1
        'If cmbSelectTable.Focused Then
        connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source =" & Main.CalculationsDbPath 'DatabasePath
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
        'End If

        'Save as a new version

        'Update the entry in CalculationsSettings 
        Main.CalculationsSettings.List(FormNo).Left = Me.Left
        Main.CalculationsSettings.List(FormNo).Top = Me.Top
        Main.CalculationsSettings.List(FormNo).Width = Me.Width
        Main.CalculationsSettings.List(FormNo).Height = Me.Height
        Main.CalculationsSettings.List(FormNo).Query = Query
        Main.CalculationsSettings.List(FormNo).Description = txtDataDescr.Text
        'Main.CalculationsSettings.List(FormNo).VersionNo 'This is only changed when a different version is selected.

        If Trim(txtVersionName.Text) = "" Then
            txtVersionName.Text = "Version 1"
            txtDataVersion.Text = "Version 1"
        Else
            txtDataVersion.Text = txtVersionName.Text
        End If

        Main.CalculationsSettings.List(FormNo).VersionName = txtVersionName.Text
        Main.CalculationsSettings.List(FormNo).VersionDesc = txtVersionDesc.Text
        Main.CalculationsSettings.List(FormNo).AutoApplyQuery = chkAutoApply.Checked.ToString
        Main.CalculationsSettings.List(FormNo).SelectedTab = TabControl1.SelectedIndex
        Main.CalculationsSettings.List(FormNo).SaveFileDir = Trim(txtDirectory.Text)
        Main.CalculationsSettings.List(FormNo).XmlFileName = Trim(txtXmlFileName.Text)

        Main.CalculationsSettings.List(FormNo).VersionNo = Main.CalculationsSettings.List(FormNo).Versions.Count
        Dim NewVersion As New DataViewVersionInfo
        NewVersion.AutoApplyQuery = Main.CalculationsSettings.List(FormNo).AutoApplyQuery
        NewVersion.Query = Main.CalculationsSettings.List(FormNo).Query
        NewVersion.VersionName = Main.CalculationsSettings.List(FormNo).VersionName
        NewVersion.VersionDesc = Main.CalculationsSettings.List(FormNo).VersionDesc
        For I = 1 To Main.CalculationsSettings.List(FormNo).TableCols.Count
            NewVersion.TableCols.Add(Main.CalculationsSettings.List(FormNo).TableCols(I - 1))
        Next
        Main.CalculationsSettings.List(FormNo).Versions.Add(NewVersion)
        ShowVersionList()
        ListBox1.SelectedIndex = Main.CalculationsSettings.List(FormNo).VersionNo
        Main.CalculationsSettingsChanged = True

    End Sub

    Private Sub btnSaveChanges_Click(sender As Object, e As EventArgs) Handles btnSaveChanges.Click
        'Save the changes made to the data in DataGridView1 to the corresponding table in the database:

        'If MessageBox.Show("Do you want to apply the changes to the table in the database?", "Confirm Changes", MessageBoxButtons.YesNoCancel) = DialogResult.OK Then
        If MessageBox.Show("Do you want to apply the changes to the table in the database?", "Confirm Changes", MessageBoxButtons.YesNoCancel) = DialogResult.Yes Then
                'Apply the edits.
            Else
                'Cancel the Save Changes.
                Exit Sub
        End If

        Dim cb = New OleDb.OleDbCommandBuilder(da)
        Try
            DataGridView1.EndEdit()
            da.Update(ds.Tables(0))
            ds.Tables(0).AcceptChanges()
            UpdateNeeded = False
            DataGridView1.EditMode = DataGridViewEditMode.EditOnKeystroke
            Main.Message.Add("Table update complete." & vbCrLf)
        Catch ex As Exception
            Main.Message.AddWarning("Error saving changes." & vbCrLf)
            Main.Message.AddWarning(ex.Message & vbCrLf & vbCrLf)
        End Try
    End Sub

    Private Sub DesignCalculationsQuery_Apply(myQuery As String) Handles DesignCalculationsQuery.Apply
        'Apply the Query designed in the Design Query form:
        Query = myQuery
    End Sub


    Private Sub btnSaveAsXml_Click(sender As Object, e As EventArgs) Handles btnSaveAsXml.Click
        'Save the data shown on DataGridView1 in an XML file.

        If Trim(txtXmlFileName.Text) = "" Then
            Main.Message.AddWarning("File name not specified!" & vbCrLf)
            Exit Sub
        End If

        If Trim(txtDirectory.Text) = "" Then
            Main.Message.AddWarning("File directory not specified!" & vbCrLf)
            Exit Sub
        End If

        Dim FilePath As String = Trim(txtDirectory.Text)

        If FilePath.EndsWith("\") Then
            FilePath = FilePath & Trim(txtXmlFileName.Text)
        Else
            FilePath = FilePath & "\" & Trim(txtXmlFileName.Text)
        End If

        If System.IO.File.Exists(FilePath) Then
            If MessageBox.Show("Overwrite existing file?", "Notice") = DialogResult.OK Then
                WriteXmlData(FilePath)
            End If
        Else
            WriteXmlData(FilePath)
        End If

    End Sub

    Private Sub WriteXmlData(ByRef FilePath As String)
        'Write the contents onf DataGridView1 in an XML file with path FielPath.
        ds.WriteXml(FilePath, XmlWriteMode.WriteSchema)
    End Sub


    Private Sub btnFind_Click(sender As Object, e As EventArgs) Handles btnFind.Click
        'Find a directory to save an XML file.

        Dim Directory As String = ""

        If Trim(txtDirectory.Text) = "" Then
            Directory = My.Computer.FileSystem.SpecialDirectories.MyDocuments
        Else
            Directory = Trim(txtDirectory.Text)
        End If

        FolderBrowserDialog1.SelectedPath = Directory
        If FolderBrowserDialog1.ShowDialog = DialogResult.OK Then
            txtDirectory.Text = FolderBrowserDialog1.SelectedPath
        End If
        Main.CalculationsSettingsChanged = True
    End Sub

    Private Sub btnSaveVersionChanges_Click(sender As Object, e As EventArgs) Handles btnSaveVersionChanges.Click

        ApplyQuery() 'Apply the query to update the list of TableCols.

        'Update the entry in CalculationsSettings 
        Main.CalculationsSettings.List(FormNo).Left = Me.Left
        Main.CalculationsSettings.List(FormNo).Top = Me.Top
        Main.CalculationsSettings.List(FormNo).Width = Me.Width
        Main.CalculationsSettings.List(FormNo).Height = Me.Height
        Main.CalculationsSettings.List(FormNo).Query = Query
        Main.CalculationsSettings.List(FormNo).Description = txtDataDescr.Text
        'Main.CalculationsSettings.List(FormNo).VersionNo 'This is only changed when a different version is selected.
        Main.CalculationsSettings.List(FormNo).VersionName = txtVersionName.Text
        Main.CalculationsSettings.List(FormNo).VersionDesc = txtVersionDesc.Text
        Main.CalculationsSettings.List(FormNo).AutoApplyQuery = chkAutoApply.Checked.ToString
        Main.CalculationsSettings.List(FormNo).SelectedTab = TabControl1.SelectedIndex
        Main.CalculationsSettings.List(FormNo).SaveFileDir = Trim(txtDirectory.Text)
        Main.CalculationsSettings.List(FormNo).XmlFileName = Trim(txtXmlFileName.Text)

        txtDataVersion.Text = Main.CalculationsSettings.List(FormNo).VersionName

        Dim I As Integer

        If Main.CalculationsSettings.List(FormNo).Versions.Count = 0 Then 'There are no versions to update. Create a new version.
            Main.CalculationsSettings.List(FormNo).VersionNo = 0
            Dim NewVersion As New DataViewVersionInfo
            NewVersion.AutoApplyQuery = Main.CalculationsSettings.List(FormNo).AutoApplyQuery
            NewVersion.Query = Main.CalculationsSettings.List(FormNo).Query
            NewVersion.VersionName = Main.CalculationsSettings.List(FormNo).VersionName
            'NewVersion.AutoApplyQuery = Main.CalculationsSettings.List(FormNo).VersionDesc
            NewVersion.VersionDesc = Main.CalculationsSettings.List(FormNo).VersionDesc
            For I = 1 To Main.CalculationsSettings.List(FormNo).TableCols.Count
                NewVersion.TableCols.Add(Main.CalculationsSettings.List(FormNo).TableCols(I - 1))
            Next
            Main.CalculationsSettings.List(FormNo).Versions.Add(NewVersion)
        ElseIf Main.CalculationsSettings.List(FormNo).VersionNo > Main.CalculationsSettings.List(FormNo).Versions.Count Then 'The current version number is too high.
            Main.Message.AddWarning("Version number: " & Main.CalculationsSettings.List(FormNo).VersionNo & " is larger than the number of versions: " & Main.CalculationsSettings.List(FormNo).Versions.Count & vbCrLf)
            Main.Message.Add("A new version will be appended to the list. " & vbCrLf)

            Main.CalculationsSettings.List(FormNo).VersionNo = Main.CalculationsSettings.List(FormNo).Versions.Count
            Dim NewVersion As New DataViewVersionInfo
            NewVersion.AutoApplyQuery = Main.CalculationsSettings.List(FormNo).AutoApplyQuery
            NewVersion.Query = Main.CalculationsSettings.List(FormNo).Query
            NewVersion.VersionName = Main.CalculationsSettings.List(FormNo).VersionName
            'NewVersion.AutoApplyQuery = Main.CalculationsSettings.List(FormNo).VersionDesc
            NewVersion.VersionDesc = Main.CalculationsSettings.List(FormNo).VersionDesc
            For I = 1 To Main.CalculationsSettings.List(FormNo).TableCols.Count
                NewVersion.TableCols.Add(Main.CalculationsSettings.List(FormNo).TableCols(I - 1))
            Next
            Main.CalculationsSettings.List(FormNo).Versions.Add(NewVersion)
        Else
            'Update the selected version settings:
            Main.CalculationsSettings.List(FormNo).Versions(Main.CalculationsSettings.List(FormNo).VersionNo).AutoApplyQuery = Main.CalculationsSettings.List(FormNo).AutoApplyQuery
            Main.CalculationsSettings.List(FormNo).Versions(Main.CalculationsSettings.List(FormNo).VersionNo).Query = Main.CalculationsSettings.List(FormNo).Query
            Main.CalculationsSettings.List(FormNo).Versions(Main.CalculationsSettings.List(FormNo).VersionNo).VersionName = Main.CalculationsSettings.List(FormNo).VersionName
            Main.CalculationsSettings.List(FormNo).Versions(Main.CalculationsSettings.List(FormNo).VersionNo).VersionDesc = Main.CalculationsSettings.List(FormNo).VersionDesc
            Main.CalculationsSettings.List(FormNo).Versions(Main.CalculationsSettings.List(FormNo).VersionNo).TableCols.Clear()

            For I = 1 To Main.CalculationsSettings.List(FormNo).TableCols.Count
                Main.CalculationsSettings.List(FormNo).Versions(Main.CalculationsSettings.List(FormNo).VersionNo).TableCols.Add(Main.CalculationsSettings.List(FormNo).TableCols(I - 1))
            Next
        End If

        ShowVersionList()
        ListBox1.SelectedIndex = Main.CalculationsSettings.List(FormNo).VersionNo

    End Sub

    Private Sub btnCancelVersionChanges_Click(sender As Object, e As EventArgs) Handles btnCancelVersionChanges.Click
        'Cancel the latest changes.
        'Restore the old settings fromthe stored version settings.

        If Main.CalculationsSettings.List(FormNo).Versions.Count = 0 Then 'There are no versions to restore from. 
            Main.Message.AddWarning("These are no versions available to restore from!" & vbCrLf)
        ElseIf Main.CalculationsSettings.List(FormNo).VersionNo > Main.CalculationsSettings.List(FormNo).Versions.Count Then 'The current version number is too high.
            Main.Message.AddWarning("Version number: " & Main.CalculationsSettings.List(FormNo).VersionNo & " is larger than the number of versions: " & Main.CalculationsSettings.List(FormNo).Versions.Count & vbCrLf)
        Else
            'Restore the curret settings from the stored version settings:
            Query = Main.CalculationsSettings.List(FormNo).Versions(Main.CalculationsSettings.List(FormNo).VersionNo).Query
            txtVersionName.Text = Main.CalculationsSettings.List(FormNo).Versions(Main.CalculationsSettings.List(FormNo).VersionNo).VersionName
            txtVersionDesc.Text = Main.CalculationsSettings.List(FormNo).Versions(Main.CalculationsSettings.List(FormNo).VersionNo).VersionDesc
            chkAutoApply.Checked = Main.CalculationsSettings.List(FormNo).Versions(Main.CalculationsSettings.List(FormNo).VersionNo).AutoApplyQuery
        End If
    End Sub

    Private Sub btnNewVersion_Click(sender As Object, e As EventArgs) Handles btnNewVersion.Click
        'Save the current settings as a new version:

        ApplyQuery() 'Apply the query to update the list of TableCols.

        'Update the entry in CalculationsSettings 
        Main.CalculationsSettings.List(FormNo).Left = Me.Left
        Main.CalculationsSettings.List(FormNo).Top = Me.Top
        Main.CalculationsSettings.List(FormNo).Width = Me.Width
        Main.CalculationsSettings.List(FormNo).Height = Me.Height
        Main.CalculationsSettings.List(FormNo).Query = Query
        Main.CalculationsSettings.List(FormNo).Description = txtDataDescr.Text
        'Main.CalculationsSettings.List(FormNo).VersionNo 'This is only changed when a different version is selected.
        Main.CalculationsSettings.List(FormNo).VersionName = txtVersionName.Text
        Main.CalculationsSettings.List(FormNo).VersionDesc = txtVersionDesc.Text
        Main.CalculationsSettings.List(FormNo).AutoApplyQuery = chkAutoApply.Checked.ToString
        Main.CalculationsSettings.List(FormNo).SelectedTab = TabControl1.SelectedIndex
        Main.CalculationsSettings.List(FormNo).SaveFileDir = Trim(txtDirectory.Text)
        Main.CalculationsSettings.List(FormNo).XmlFileName = Trim(txtXmlFileName.Text)

        Main.CalculationsSettings.List(FormNo).VersionNo = Main.CalculationsSettings.List(FormNo).Versions.Count
        Dim NewVersion As New DataViewVersionInfo
        NewVersion.AutoApplyQuery = Main.CalculationsSettings.List(FormNo).AutoApplyQuery
        NewVersion.Query = Main.CalculationsSettings.List(FormNo).Query
        NewVersion.VersionName = Main.CalculationsSettings.List(FormNo).VersionName
        NewVersion.VersionDesc = Main.CalculationsSettings.List(FormNo).VersionDesc
        For I = 1 To Main.CalculationsSettings.List(FormNo).TableCols.Count
            NewVersion.TableCols.Add(Main.CalculationsSettings.List(FormNo).TableCols(I - 1))
        Next
        Main.CalculationsSettings.List(FormNo).Versions.Add(NewVersion)
        ShowVersionList()
        ListBox1.SelectedIndex = Main.CalculationsSettings.List(FormNo).VersionNo
        Main.CalculationsSettingsChanged = True
    End Sub

    Private Sub ShowVersionList()
        'Show the list of DataView Versions in the list:

        ListBox1.Items.Clear()

        Dim I As Integer

        For I = 1 To Main.CalculationsSettings.List(FormNo).Versions.Count
            ListBox1.Items.Add(Main.CalculationsSettings.List(FormNo).Versions(I - 1).VersionName)
        Next

    End Sub

    Private Sub ListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox1.SelectedIndexChanged

        Dim SelRow As Integer = ListBox1.SelectedIndex

        If SelRow > -1 Then
            txtSelVersionQuery.Text = Main.CalculationsSettings.List(FormNo).Versions(SelRow).Query
            txtSelVersionDesc.Text = Main.CalculationsSettings.List(FormNo).Versions(SelRow).VersionDesc
        End If



    End Sub

    Private Sub btnSelect_Click(sender As Object, e As EventArgs) Handles btnSelect.Click
        'Select the selected DataView Version

        Dim SelRow As Integer = ListBox1.SelectedIndex
        Dim I As Integer

        Main.CalculationsSettings.List(FormNo).Query = Query
        txtQuery.Text = Query

        Main.CalculationsSettings.List(FormNo).VersionNo = SelRow
        Main.CalculationsSettings.List(FormNo).VersionDesc = Main.CalculationsSettings.List(FormNo).Versions(SelRow).VersionDesc
        Main.CalculationsSettings.List(FormNo).VersionName = Main.CalculationsSettings.List(FormNo).Versions(SelRow).VersionName
        Main.CalculationsSettings.List(FormNo).Query = Main.CalculationsSettings.List(FormNo).Versions(SelRow).Query
        Query = Main.CalculationsSettings.List(FormNo).Versions(SelRow).Query
        Main.CalculationsSettings.List(FormNo).AutoApplyQuery = Main.CalculationsSettings.List(FormNo).Versions(SelRow).AutoApplyQuery
        Main.CalculationsSettings.List(FormNo).TableCols.Clear()
        For I = 1 To Main.CalculationsSettings.List(FormNo).Versions(SelRow).TableCols.Count
            Main.CalculationsSettings.List(FormNo).TableCols.Add(Main.CalculationsSettings.List(FormNo).Versions(SelRow).TableCols(I - 1))
        Next

        txtDataVersion.Text = Main.CalculationsSettings.List(FormNo).VersionName
        'txtQuery.Text is updated when Query is changed
        txtVersionName.Text = Main.CalculationsSettings.List(FormNo).VersionName
        txtVersionDesc.Text = Main.CalculationsSettings.List(FormNo).VersionDesc
        chkAutoApply.Checked = Main.CalculationsSettings.List(FormNo).AutoApplyQuery

        If chkAutoApply.Checked Then
            ApplyQuery()
        Else
            ds.Clear()
            ds.Reset()
        End If
        Main.CalculationsSettingsChanged = True
    End Sub

    Private Sub btnDelete_Click(sender As Object, e As EventArgs) Handles btnDelete.Click
        'Select the selected DataView Version

        Dim SelRow As Integer = ListBox1.SelectedIndex

        If Main.CalculationsSettings.List(FormNo).VersionNo = SelRow Then
            Main.Message.AddWarning("Cannot delete current version of the Data View." & vbCrLf)
        Else
            If SelRow = -1 Then
                Main.Message.AddWarning("No version has been selected for deletion." & vbCrLf)
            Else
                Main.CalculationsSettings.List(FormNo).Versions.RemoveAt(SelRow)
                If Main.CalculationsSettings.List(FormNo).VersionNo > SelRow Then
                    Main.CalculationsSettings.List(FormNo).VersionNo -= 1
                End If

                ShowVersionList()
                ListBox1.SelectedIndex = Main.CalculationsSettings.List(FormNo).VersionNo
                Main.CalculationsSettingsChanged = True
            End If
        End If
    End Sub

    Private Sub btnMoveUp_Click(sender As Object, e As EventArgs) Handles btnMoveUp.Click
        'Move the version entry up in the list.

        Dim SelRow As Integer = ListBox1.SelectedIndex
        Dim I As Integer

        If SelRow = 0 Then
            'Already at the top of the list.
        Else
            'Save version info at SelRow
            Dim TempVersion As New DataViewVersionInfo
            TempVersion.AutoApplyQuery = Main.CalculationsSettings.List(FormNo).Versions(SelRow).AutoApplyQuery
            TempVersion.Query = Main.CalculationsSettings.List(FormNo).Versions(SelRow).Query
            TempVersion.VersionDesc = Main.CalculationsSettings.List(FormNo).Versions(SelRow).VersionDesc
            TempVersion.VersionName = Main.CalculationsSettings.List(FormNo).Versions(SelRow).VersionName
            For I = 1 To Main.CalculationsSettings.List(FormNo).Versions(SelRow).TableCols.Count
                TempVersion.TableCols.Add(Main.CalculationsSettings.List(FormNo).Versions(SelRow).TableCols(I - 1))
            Next
            'Copy version info at SelRow - 1 to SelRow
            Main.CalculationsSettings.List(FormNo).Versions(SelRow).AutoApplyQuery = Main.CalculationsSettings.List(FormNo).Versions(SelRow - 1).AutoApplyQuery
            Main.CalculationsSettings.List(FormNo).Versions(SelRow).Query = Main.CalculationsSettings.List(FormNo).Versions(SelRow - 1).Query
            Main.CalculationsSettings.List(FormNo).Versions(SelRow).VersionDesc = Main.CalculationsSettings.List(FormNo).Versions(SelRow - 1).VersionDesc
            Main.CalculationsSettings.List(FormNo).Versions(SelRow).VersionName = Main.CalculationsSettings.List(FormNo).Versions(SelRow - 1).VersionName
            Main.CalculationsSettings.List(FormNo).Versions(SelRow).TableCols.Clear()
            For I = 1 To Main.CalculationsSettings.List(FormNo).Versions(SelRow - 1).TableCols.Count
                Main.CalculationsSettings.List(FormNo).Versions(SelRow).TableCols.Add(Main.CalculationsSettings.List(FormNo).Versions(SelRow - 1).TableCols(I - 1))
            Next
            'Copy version info in TempVersion to SelRow - 1
            Main.CalculationsSettings.List(FormNo).Versions(SelRow - 1).AutoApplyQuery = TempVersion.AutoApplyQuery
            Main.CalculationsSettings.List(FormNo).Versions(SelRow - 1).Query = TempVersion.Query
            Main.CalculationsSettings.List(FormNo).Versions(SelRow - 1).VersionDesc = TempVersion.VersionDesc
            Main.CalculationsSettings.List(FormNo).Versions(SelRow - 1).VersionName = TempVersion.VersionName
            Main.CalculationsSettings.List(FormNo).Versions(SelRow - 1).TableCols.Clear()
            For I = 1 To TempVersion.TableCols.Count
                Main.CalculationsSettings.List(FormNo).Versions(SelRow - 1).TableCols.Add(TempVersion.TableCols(I - 1))
            Next

            If Main.CalculationsSettings.List(FormNo).VersionNo = SelRow Then Main.CalculationsSettings.List(FormNo).VersionNo -= 1
            If Main.CalculationsSettings.List(FormNo).VersionNo = SelRow - 1 Then Main.CalculationsSettings.List(FormNo).VersionNo += 1

            ShowVersionList()
            ListBox1.SelectedIndex = SelRow - 1
            Main.CalculationsSettingsChanged = True
        End If
    End Sub

    Private Sub btnMoveDown_Click(sender As Object, e As EventArgs) Handles btnMoveDown.Click
        'Move the version entry down in the list.

        Dim SelRow As Integer = ListBox1.SelectedIndex

        If SelRow = Main.CalculationsSettings.List(FormNo).Versions.Count - 1 Then
            'Already at the end of the list.
        Else
            'Save version info at SelRow
            Dim TempVersion As New DataViewVersionInfo
            TempVersion.AutoApplyQuery = Main.CalculationsSettings.List(FormNo).Versions(SelRow).AutoApplyQuery
            TempVersion.Query = Main.CalculationsSettings.List(FormNo).Versions(SelRow).Query
            TempVersion.VersionDesc = Main.CalculationsSettings.List(FormNo).Versions(SelRow).VersionDesc
            TempVersion.VersionName = Main.CalculationsSettings.List(FormNo).Versions(SelRow).VersionName
            For I = 1 To Main.CalculationsSettings.List(FormNo).Versions(SelRow).TableCols.Count
                TempVersion.TableCols.Add(Main.CalculationsSettings.List(FormNo).Versions(SelRow).TableCols(I - 1))
            Next
            'Copy version info at SelRow + 1 to SelRow
            Main.CalculationsSettings.List(FormNo).Versions(SelRow).AutoApplyQuery = Main.CalculationsSettings.List(FormNo).Versions(SelRow + 1).AutoApplyQuery
            Main.CalculationsSettings.List(FormNo).Versions(SelRow).Query = Main.CalculationsSettings.List(FormNo).Versions(SelRow + 1).Query
            Main.CalculationsSettings.List(FormNo).Versions(SelRow).VersionDesc = Main.CalculationsSettings.List(FormNo).Versions(SelRow + 1).VersionDesc
            Main.CalculationsSettings.List(FormNo).Versions(SelRow).VersionName = Main.CalculationsSettings.List(FormNo).Versions(SelRow + 1).VersionName
            Main.CalculationsSettings.List(FormNo).Versions(SelRow).TableCols.Clear()
            For I = 1 To Main.CalculationsSettings.List(FormNo).Versions(SelRow + 1).TableCols.Count
                Main.CalculationsSettings.List(FormNo).Versions(SelRow).TableCols.Add(Main.CalculationsSettings.List(FormNo).Versions(SelRow + 1).TableCols(I - 1))
            Next
            'Copy version info in TempVersion to SelRow + 1
            Main.CalculationsSettings.List(FormNo).Versions(SelRow + 1).AutoApplyQuery = TempVersion.AutoApplyQuery
            Main.CalculationsSettings.List(FormNo).Versions(SelRow + 1).Query = TempVersion.Query
            Main.CalculationsSettings.List(FormNo).Versions(SelRow + 1).VersionDesc = TempVersion.VersionDesc
            Main.CalculationsSettings.List(FormNo).Versions(SelRow + 1).VersionName = TempVersion.VersionName
            Main.CalculationsSettings.List(FormNo).Versions(SelRow + 1).TableCols.Clear()
            For I = 1 To TempVersion.TableCols.Count
                Main.CalculationsSettings.List(FormNo).Versions(SelRow + 1).TableCols.Add(TempVersion.TableCols(I - 1))
            Next

            If Main.CalculationsSettings.List(FormNo).VersionNo = SelRow Then Main.CalculationsSettings.List(FormNo).VersionNo += 1
            If Main.CalculationsSettings.List(FormNo).VersionNo = SelRow + 1 Then Main.CalculationsSettings.List(FormNo).VersionNo -= 1

            ShowVersionList()
            ListBox1.SelectedIndex = SelRow + 1
            Main.CalculationsSettingsChanged = True
        End If

    End Sub

    Private Sub txtVersionName_TextChanged(sender As Object, e As EventArgs) Handles txtVersionName.TextChanged
        Main.CalculationsSettingsChanged = True
    End Sub

    Private Sub txtVersionDesc_TextChanged(sender As Object, e As EventArgs) Handles txtVersionDesc.TextChanged
        Main.CalculationsSettingsChanged = True
    End Sub

    Private Sub txtXmlFileName_TextChanged(sender As Object, e As EventArgs) Handles txtXmlFileName.TextChanged
        Main.CalculationsSettingsChanged = True
    End Sub

















#End Region 'Form Methods ---------------------------------------------------------------------------------------------------------------------------------------------------------------------

End Class