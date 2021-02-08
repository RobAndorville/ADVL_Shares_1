Public Class frmCompanyList
    'Gets a list of companies for the Share Price Chart tab.

#Region " Variable Declarations - All the variables used in this form and this application." '=================================================================================================

    Dim WithEvents Zip As ADVL_Utilities_Library_1.ZipComp
    'Dim XDoc As New System.Xml.XmlDocument '4/6/2018

#End Region 'Variable Declarations ------------------------------------------------------------------------------------------------------------------------------------------------------------

#Region " Properties - All the properties used in this form and this application" '============================================================================================================
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
                               <!--Select Company Info Tab-->
                               <SelectedDatabase><%= cmbCompanyListDb.SelectedItem.ToString %></SelectedDatabase>
                               <CompanyInfoQuery><%= txtCompanyInfoQuery.Text %></CompanyInfoQuery>
                               <%= If(cmbCompanyCodeCol.SelectedIndex = -1,
                                   <SetCompanyCodeValueTo></SetCompanyCodeValueTo>,
                                   <SetCompanyCodeValueTo><%= cmbCompanyCodeCol.SelectedItem.ToString %></SetCompanyCodeValueTo>) %>
                               <CompanyListFile><%= txtCompanyListFile.Text %></CompanyListFile>
                               <!---->
                           </FormSettings>

        'Add code to include other settings to save after the comment line <!---->

        Dim SettingsFileName As String = "FormSettings_" & Main.ApplicationInfo.Name & "_" & Me.Text & ".xml"
        Main.Project.SaveXmlSettings(SettingsFileName, settingsData)
    End Sub

    Private Sub RestoreFormSettings()
        'Read the form settings from an XML document.

        Dim SettingsFileName As String = "FormSettings_" & Main.ApplicationInfo.Name & "_" & Me.Text & ".xml"

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

            'Restore Select Campany Info Tab:
            If Settings.<FormSettings>.<SelectedDatabase>.Value <> Nothing Then cmbCompanyListDb.SelectedIndex = cmbCompanyListDb.FindStringExact(Settings.<FormSettings>.<SelectedDatabase>.Value)
            If Settings.<FormSettings>.<CompanyInfoQuery>.Value <> Nothing Then
                txtCompanyInfoQuery.Text = Settings.<FormSettings>.<CompanyInfoQuery>.Value
                ApplyQuery()
            End If
            If Settings.<FormSettings>.<SetCompanyCodeValueTo>.Value <> Nothing Then cmbCompanyCodeCol.SelectedIndex = cmbCompanyCodeCol.FindStringExact(Settings.<FormSettings>.<SetCompanyCodeValueTo>.Value)
            If Settings.<FormSettings>.<CompanyListFile>.Value <> Nothing Then
                txtCompanyListFile.Text = Settings.<FormSettings>.<CompanyListFile>.Value
                LoadSettingsFile()
            End If
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
        'RestoreFormSettings()   'Restore the form settings NOTE: THE COMBOBOXES MUST BE SET UP BEFORE RESTOREFORMSETTINGS IS RUN!!!

        'Initialise Calculations - Copy Data Tab
        cmbCompanyListDb.Items.Add("Share Prices")
        cmbCompanyListDb.Items.Add("Financials")
        cmbCompanyListDb.Items.Add("Calculations")
        cmbCompanyListDb.SelectedIndex = 0 'Select the first item as default.

        If Main.cmbChartDataTable.SelectedIndex = -1 Then
            txtChartDataTable.Text = ""
            Main.Message.AddWarning("Select the Chart Data Table on the Share Prices \ Input Data Tab on the Shares application." & vbCrLf)
            Main.Message.AddWarning("  Then restart this form." & vbCrLf)
        Else
            txtChartDataTable.Text = Main.cmbChartDataTable.SelectedItem.ToString()
        End If

        If Main.cmbCompanyCodeCol.SelectedIndex = -1 Then
            txtCompanyCodeColumn.Text = ""
            Main.Message.AddWarning("Select the Company Code Column on the Share Prices \ Input Data Tab on the Shares application." & vbCrLf)
            Main.Message.AddWarning("  Then restart this form." & vbCrLf)
        Else
            txtCompanyCodeColumn.Text = Main.cmbCompanyCodeCol.SelectedItem.ToString
        End If


        RestoreFormSettings()   'Restore the form settings

    End Sub

    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        'Exit the Form
        Me.Close() 'Close the form
    End Sub

    Private Sub Form_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        If WindowState = FormWindowState.Normal Then
            SaveFormSettings()
        Else
            'Dont save settings if the form is minimised.
        End If
    End Sub



#End Region 'Form Display Methods -------------------------------------------------------------------------------------------------------------------------------------------------------------

#Region " Open and Close Forms - Code used to open and close other forms." '===================================================================================================================
#End Region 'Open and Close Forms -------------------------------------------------------------------------------------------------------------------------------------------------------------

#Region " Form Methods - The main actions performed by this form." '===========================================================================================================================

    Private Sub btnApplyQuery_Click(sender As Object, e As EventArgs) Handles btnApplyQuery.Click
        'Apply the Select Company Info query

        ApplyQuery()

    End Sub

    Private Sub ApplyQuery()
        Dim connString As String
        Dim myConnection As OleDb.OleDbConnection = New OleDb.OleDbConnection
        Dim ds As DataSet = New DataSet
        Dim da As OleDb.OleDbDataAdapter

        Dim DatabasePath As String = ""

        Select Case cmbCompanyListDb.SelectedItem.ToString
            Case "Share Prices"
                DatabasePath = Main.SharePriceDbPath
            Case "Financials"
                DatabasePath = Main.FinancialsDbPath
            Case "Calculations"
                DatabasePath = Main.CalculationsDbPath
        End Select

        If DatabasePath = "" Then
            Main.Message.AddWarning("A database has not been selected!" & vbCrLf)
            Exit Sub
        End If

        connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source =" & DatabasePath
        myConnection.ConnectionString = connString
        myConnection.Open()

        Dim Query As String = txtCompanyInfoQuery.Text

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

            cmbCompanyCodeCol.Items.Clear()
            Dim NCols As Integer = DataGridView1.ColumnCount
            Dim I As Integer

            For I = 0 To NCols - 1
                cmbCompanyCodeCol.Items.Add(DataGridView1.Columns(I).HeaderText)
            Next
            SetUpSeriesNameGrid()
            SetUpChartTitleGrid()
        Catch ex As Exception
            Main.Message.Add("Error applying query." & vbCrLf)
            Main.Message.Add(ex.Message & vbCrLf & vbCrLf)
        End Try
    End Sub

    Private Sub cmbCompanyCodeCol_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbCompanyCodeCol.SelectedIndexChanged
        UpdateChartDataQuery()
    End Sub

    Private Sub UpdateChartDataQuery()
        If cmbCompanyCodeCol.SelectedIndex = -1 Then
            'No item selected.
            txtCompanyCode.Text = ""
            txtChartDataQuery.Text = ""
        Else
            Dim CompanyCode As String = ""
            If DataGridView1.SelectedRows.Count > 0 Then
                If cmbCompanyCodeCol.SelectedIndex = -1 Then
                    CompanyCode = ""
                Else
                    CompanyCode = DataGridView1.SelectedRows(0).Cells(cmbCompanyCodeCol.SelectedIndex).Value
                End If
            Else
                CompanyCode = ""
            End If
            txtCompanyCode.Text = CompanyCode
            If Main.chkSPChartUseDateRange.Checked Then 'Include a date range in the query:
                txtChartDataQuery.Text = "SELECT * FROM " & txtChartDataTable.Text & " WHERE " & txtCompanyCodeColumn.Text & " = '" & CompanyCode & "'" & " AND " & Main.cmbXValues.SelectedItem.ToString & " BETWEEN #" & Format(Main.dtpSPChartFromDate.Value, "MM-dd-yyyy") & "# AND #" & Format(Main.dtpSPChartToDate.Value, "MM-dd-yyyy") & "#"

            Else 'Dont include a date range in the query:
                txtChartDataQuery.Text = "SELECT * FROM " & txtChartDataTable.Text & " WHERE " & txtCompanyCodeColumn.Text & " = '" & CompanyCode & "'"
            End If

        End If
    End Sub

    Private Sub SetUpSeriesNameGrid()
        'Set up the dgvSeriesName data grid view.

        'Save the current settings:
        Dim I As Integer 'Loop index
        dgvSeriesName.AllowUserToAddRows = False
        Dim RowCount As Integer = dgvSeriesName.RowCount
        Dim Items(0 To RowCount - 1) As String
        Dim Values(0 To RowCount - 1) As String

        For I = 0 To RowCount - 1
            Items(I) = dgvSeriesName.Rows(I).Cells(0).Value
            Values(I) = dgvSeriesName.Rows(I).Cells(1).Value
        Next

        dgvSeriesName.Rows.Clear()
        dgvSeriesName.Columns.Clear()
        Dim ItemType As New DataGridViewComboBoxColumn
        dgvSeriesName.Columns.Add(ItemType)
        dgvSeriesName.Columns(0).HeaderText = "Add Item"
        ItemType.Items.Add("Text")
        Dim NColItems As Integer = cmbCompanyCodeCol.Items.Count

        For I = 1 To NColItems
            ItemType.Items.Add(cmbCompanyCodeCol.Items(I - 1).ToString)
        Next

        Dim ItemValue As New DataGridViewTextBoxColumn
        dgvSeriesName.Columns.Add(ItemValue)
        dgvSeriesName.Columns(1).HeaderText = "Value"

        'Restore the settings:
        For I = 0 To RowCount - 1
            If ItemType.Items.Contains(Items(I)) Then
                dgvSeriesName.Rows.Add()
                dgvSeriesName.Rows(I).Cells(0).Value = Items(I)
                dgvSeriesName.Rows(I).Cells(1).Value = Values(I)
            Else
                dgvSeriesName.Rows.Add()
            End If
        Next

        dgvSeriesName.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        dgvSeriesName.AutoResizeColumns()

        dgvSeriesName.AllowUserToAddRows = True

    End Sub

    Private Sub dgvSeriesName_CellMouseLeave(sender As Object, e As DataGridViewCellEventArgs) Handles dgvSeriesName.CellMouseLeave
        UpdateSeriesName()
    End Sub

    Private Sub UpdateSeriesName()
        txtSeriesName.Text = ""
        Dim I As Integer
        For I = 0 To dgvSeriesName.Rows.Count - 1
            If dgvSeriesName.Rows(I).IsNewRow Then

            Else
                If dgvSeriesName.Rows(I).Cells(0).Value = "" Then

                ElseIf dgvSeriesName.Rows(I).Cells(0).Value = "Text" Then
                    txtSeriesName.Text &= dgvSeriesName.Rows(I).Cells(1).Value
                Else
                    If DataGridView1.SelectedRows.Count = 0 Then
                        'No item selected in DataGridView1
                        dgvSeriesName.Rows(I).Cells(1).Value = ""
                    Else
                        dgvSeriesName.Rows(I).Cells(1).Value = DataGridView1.SelectedRows(0).Cells(dgvSeriesName.Rows(I).Cells(0).Value).Value
                        txtSeriesName.Text = txtSeriesName.Text & dgvSeriesName.Rows(I).Cells(1).Value
                    End If
                End If
            End If

        Next

        dgvSeriesName.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        dgvSeriesName.AutoResizeColumns()

    End Sub

    Private Sub SetUpChartTitleGrid()
        'Set up the dgvChartTitle data grid view.

        'Save the current settings:
        Dim I As Integer 'Loop index
        dgvChartTitle.AllowUserToAddRows = False
        Dim RowCount As Integer = dgvChartTitle.RowCount
        Dim Items(0 To RowCount - 1) As String
        Dim Values(0 To RowCount - 1) As String

        For I = 0 To RowCount - 1
            Items(I) = dgvChartTitle.Rows(I).Cells(0).Value
            Values(I) = dgvChartTitle.Rows(I).Cells(1).Value
        Next

        dgvChartTitle.Rows.Clear()
        dgvChartTitle.Columns.Clear()
        Dim ItemType As New DataGridViewComboBoxColumn
        dgvChartTitle.Columns.Add(ItemType)
        dgvChartTitle.Columns(0).HeaderText = "Add Item"
        ItemType.Items.Add("Text")
        Dim NColItems As Integer = cmbCompanyCodeCol.Items.Count

        For I = 1 To NColItems
            ItemType.Items.Add(cmbCompanyCodeCol.Items(I - 1).ToString)
        Next

        Dim ItemValue As New DataGridViewTextBoxColumn
        dgvChartTitle.Columns.Add(ItemValue)
        dgvChartTitle.Columns(1).HeaderText = "Value"

        'Restore the settings:
        For I = 0 To RowCount - 1
            If ItemType.Items.Contains(Items(I)) Then
                dgvChartTitle.Rows.Add()
                dgvChartTitle.Rows(I).Cells(0).Value = Items(I)
                dgvChartTitle.Rows(I).Cells(1).Value = Values(I)
            Else
                dgvChartTitle.Rows.Add()
            End If
        Next

        dgvChartTitle.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        dgvChartTitle.AutoResizeColumns()

        dgvChartTitle.AllowUserToAddRows = True

    End Sub

    Private Sub dgvChartTitle_CellMouseLeave(sender As Object, e As DataGridViewCellEventArgs) Handles dgvChartTitle.CellMouseLeave
        UpdateChartTitle()
    End Sub

    Private Sub UpdateChartTitle()
        txtChartTitle.Text = ""
        Dim I As Integer
        For I = 0 To dgvChartTitle.Rows.Count - 1
            If dgvChartTitle.Rows(I).IsNewRow Then

            Else
                If dgvChartTitle.Rows(I).Cells(0).Value = "" Then

                ElseIf dgvChartTitle.Rows(I).Cells(0).Value = "Text" Then
                    txtChartTitle.Text &= dgvChartTitle.Rows(I).Cells(1).Value
                Else
                    If DataGridView1.SelectedRows.Count = 0 Then
                        'No item selected in dgvChartTitle
                        dgvChartTitle.Rows(I).Cells(1).Value = ""
                    Else
                        dgvChartTitle.Rows(I).Cells(1).Value = DataGridView1.SelectedRows(0).Cells(dgvChartTitle.Rows(I).Cells(0).Value).Value
                        txtChartTitle.Text = txtChartTitle.Text & dgvChartTitle.Rows(I).Cells(1).Value
                    End If
                End If
            End If

        Next

        dgvChartTitle.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        dgvChartTitle.AutoResizeColumns()

    End Sub

    Private Sub btnSaveCompanyList_Click(sender As Object, e As EventArgs) Handles btnSaveCompanyList.Click
        'Save SP Chart Company List file.

        If Trim(txtCompanyListFile.Text) = "" Then
            Main.Message.AddWarning("Enter a name for the Company List file." & vbCrLf)
        Else
            If txtCompanyListFile.Text.EndsWith(".SPChartCompanyList") Then
                txtCompanyListFile.Text = Trim(txtCompanyListFile.Text)
            Else
                txtCompanyListFile.Text = Trim(txtCompanyListFile.Text) & ".SPChartCompanyList"
            End If

            dgvSeriesName.AllowUserToAddRows = False 'This removes the last edit row from the DataGridView.
            dgvChartTitle.AllowUserToAddRows = False 'This removes the last edit row from the DataGridView.

            Try
                Dim XDoc = <?xml version="1.0" encoding="utf-8"?>
                           <SharePriceChartCompanyList>
                               <HostApplication><%= Main.ApplicationInfo.Name %></HostApplication>
                               <SelectedDatabase><%= cmbCompanyListDb.SelectedItem.ToString %></SelectedDatabase>
                               <Query><%= txtCompanyInfoQuery.Text %></Query>
                               <SetCompanyCodeValueTo><%= cmbCompanyCodeCol.SelectedItem.ToString %></SetCompanyCodeValueTo>
                               <SeriesNameItems>
                                   <%= From item In dgvSeriesName.Rows
                                       Select
                                           <Item>
                                               <Type><%= item.Cells(0).Value %></Type>
                                               <Value><%= item.Cells(1).Value %></Value>
                                           </Item>
                                   %>
                               </SeriesNameItems>
                               <ChartTitleItems>
                                   <%= From item In dgvChartTitle.Rows
                                       Select
                                           <Item>
                                               <Type><%= item.Cells(0).Value %></Type>
                                               <Value><%= item.Cells(1).Value %></Value>
                                           </Item>
                                   %>
                               </ChartTitleItems>
                           </SharePriceChartCompanyList>

                Main.Project.SaveXmlData(txtCompanyListFile.Text, XDoc)
                dgvSeriesName.AllowUserToAddRows = True 'Allow user to add rows again.
                dgvChartTitle.AllowUserToAddRows = True 'Allow user to add rows again.
            Catch ex As Exception
                Main.Message.AddWarning("Error saving company list: " & ex.Message & vbCrLf)
            End Try

        End If
    End Sub

    Private Sub btnFindCompanyList_Click(sender As Object, e As EventArgs) Handles btnFindCompanyList.Click
        'Find and open a SP Chart Company List file.

        Select Case Main.Project.DataLocn.Type
            Case ADVL_Utilities_Library_1.FileLocation.Types.Directory
                'Select a SP Chart Company List file from the project directory:
                OpenFileDialog1.InitialDirectory = Main.Project.DataLocn.Path
                OpenFileDialog1.Filter = "Company List settings file | *.SPChartCompanyList"
                If OpenFileDialog1.ShowDialog() = DialogResult.OK Then
                    Dim DataFileName As String = System.IO.Path.GetFileName(OpenFileDialog1.FileName)
                    txtCompanyListFile.Text = DataFileName
                    LoadSettingsFile
                End If

            Case ADVL_Utilities_Library_1.FileLocation.Types.Archive
                'Select a SP Chart Company List file from the project archive:
                'Show the zip archive file selection form:
                Zip = New ADVL_Utilities_Library_1.ZipComp
                Zip.ArchivePath = Main.Project.DataLocn.Path
                Zip.SelectFile() 'Show the SelectFile form.
                'Zip.SelectFileForm.ApplicationName = Main.Project.ApplicationName
                Zip.SelectFileForm.ApplicationName = Main.Project.Application.Name
                Zip.SelectFileForm.SettingsLocn = Main.Project.SettingsLocn
                Zip.SelectFileForm.Show()
                Zip.SelectFileForm.RestoreFormSettings()
                Zip.SelectFileForm.FileExtension = ".SPChartCompanyList"
                Zip.SelectFileForm.GetFileList()
                If Zip.SelectedFile <> "" Then
                    'A file has been selected
                    txtCompanyListFile.Text = Zip.SelectedFile
                    LoadSettingsFile
                End If
        End Select
    End Sub

    Private Sub LoadSettingsFile()
        'Load the SP Chart Company List file.
        'txtCompanyListFile.Text contains the file name.

        If Trim(txtCompanyListFile.Text) = "" Then
            Main.Message.AddWarning("No SP Chart Company List file name has been specified!" & vbCrLf)
        Else
            Dim XDoc As System.Xml.Linq.XDocument
            Main.Project.ReadXmlData(txtCompanyListFile.Text, XDoc)
            If XDoc Is Nothing Then

            Else
                If XDoc.<SharePriceChartCompanyList>.<SelectedDatabase>.Value <> Nothing Then cmbCompanyListDb.SelectedIndex = cmbCompanyListDb.FindStringExact(XDoc.<SharePriceChartCompanyList>.<SelectedDatabase>.Value)
                If XDoc.<SharePriceChartCompanyList>.<Query>.Value <> Nothing Then txtCompanyInfoQuery.Text = XDoc.<SharePriceChartCompanyList>.<Query>.Value
                If XDoc.<SharePriceChartCompanyList>.<SetCompanyCodeValueTo>.Value <> Nothing Then cmbCompanyCodeCol.SelectedIndex = cmbCompanyCodeCol.FindStringExact(XDoc.<SharePriceChartCompanyList>.<SetCompanyCodeValueTo>.Value)

                dgvSeriesName.Rows.Clear()
                Dim SeriesNameItems = From item In XDoc.<SharePriceChartCompanyList>.<SeriesNameItems>.<Item>
                For Each Item In SeriesNameItems
                    dgvSeriesName.Rows.Add(Item.<Type>.Value, Item.<Value>.Value)
                Next

                dgvChartTitle.Rows.Clear()
                Dim ChartTitleItems = From item In XDoc.<SharePriceChartCompanyList>.<ChartTitleItems>.<Item>
                For Each Item In ChartTitleItems
                    dgvChartTitle.Rows.Add(Item.<Type>.Value, Item.<Value>.Value)
                Next

                UpdateChartDataQuery()
                UpdateSeriesName()
                UpdateChartTitle()
            End If
        End If
    End Sub

    Private Sub btnChartSelected_Click(sender As Object, e As EventArgs) Handles btnChartSelected.Click
        ChartSelectedCompany()
    End Sub

    Private Sub ChartSelectedCompany()

        Main.txtSeriesName.Text = txtSeriesName.Text
        Main.txtSPChartQuery.Text = txtChartDataQuery.Text
        Main.txtChartTitle.Text = txtChartTitle.Text
        Main.txtSPChartCompanyCode.Text = txtCompanyCode.Text
        Main.DisplayStockChart()

    End Sub



    Private Sub DataGridView1_SelectionChanged(sender As Object, e As EventArgs) Handles DataGridView1.SelectionChanged

        UpdateChartDataQuery()
        UpdateSeriesName()
        UpdateChartTitle()

    End Sub

    Private Sub btnChartNext_Click(sender As Object, e As EventArgs) Handles btnChartNext.Click
        If DataGridView1.SelectedRows.Count > 0 Then
            Dim SelRow As Integer = DataGridView1.SelectedRows(0).Index
            If SelRow = DataGridView1.Rows.Count - 1 Then
                'At last row.
            Else
                DataGridView1.Rows(SelRow).Selected = False
                DataGridView1.Rows(SelRow + 1).Selected = True
                ChartSelectedCompany()
            End If
        Else
            'No row has been selected
        End If
    End Sub

    Private Sub btnChartPrev_Click(sender As Object, e As EventArgs) Handles btnChartPrev.Click
        If DataGridView1.SelectedRows.Count > 0 Then
            Dim SelRow As Integer = DataGridView1.SelectedRows(0).Index
            If SelRow = 0 Then
                'At first row.
            Else
                DataGridView1.Rows(SelRow).Selected = False
                DataGridView1.Rows(SelRow - 1).Selected = True
                ChartSelectedCompany()
            End If
        Else
            'No row has been selected
        End If
    End Sub

    Private Sub txtCompanyInfoQuery_TextChanged(sender As Object, e As EventArgs) Handles txtCompanyInfoQuery.TextChanged

    End Sub

    Private Sub txtCompanyInfoQuery_LostFocus(sender As Object, e As EventArgs) Handles txtCompanyInfoQuery.LostFocus

    End Sub

#End Region 'Form Methods ---------------------------------------------------------------------------------------------------------------------------------------------------------------------

#Region " Form Events - Events that can be triggered by this form." '==========================================================================================================================
#End Region 'Form Events ----------------------------------------------------------------------------------------------------------------------------------------------------------------------

End Class