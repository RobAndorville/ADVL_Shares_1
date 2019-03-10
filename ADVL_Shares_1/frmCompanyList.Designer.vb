<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmCompanyList
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.btnExit = New System.Windows.Forms.Button()
        Me.TabControl1 = New System.Windows.Forms.TabControl()
        Me.TabPage1 = New System.Windows.Forms.TabPage()
        Me.btnApplyQuery = New System.Windows.Forms.Button()
        Me.btnChartNext = New System.Windows.Forms.Button()
        Me.btnChartPrev = New System.Windows.Forms.Button()
        Me.btnChartSelected = New System.Windows.Forms.Button()
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.cmbCompanyListDb = New System.Windows.Forms.ComboBox()
        Me.Label29 = New System.Windows.Forms.Label()
        Me.txtCompanyInfoQuery = New System.Windows.Forms.TextBox()
        Me.Label97 = New System.Windows.Forms.Label()
        Me.TabPage2 = New System.Windows.Forms.TabPage()
        Me.txtCompanyCode = New System.Windows.Forms.TextBox()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.txtCompanyCodeColumn = New System.Windows.Forms.TextBox()
        Me.txtChartDataTable = New System.Windows.Forms.TextBox()
        Me.txtChartDataQuery = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.cmbCompanyCodeCol = New System.Windows.Forms.ComboBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.TabPage3 = New System.Windows.Forms.TabPage()
        Me.txtSeriesName = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.dgvSeriesName = New System.Windows.Forms.DataGridView()
        Me.TabPage4 = New System.Windows.Forms.TabPage()
        Me.txtChartTitle = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.dgvChartTitle = New System.Windows.Forms.DataGridView()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.btnSaveCompanyList = New System.Windows.Forms.Button()
        Me.txtCompanyListFile = New System.Windows.Forms.TextBox()
        Me.btnFindCompanyList = New System.Windows.Forms.Button()
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage2.SuspendLayout()
        Me.TabPage3.SuspendLayout()
        CType(Me.dgvSeriesName, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage4.SuspendLayout()
        CType(Me.dgvChartTitle, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'btnExit
        '
        Me.btnExit.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnExit.Location = New System.Drawing.Point(580, 12)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(64, 22)
        Me.btnExit.TabIndex = 8
        Me.btnExit.Text = "Exit"
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'TabControl1
        '
        Me.TabControl1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Controls.Add(Me.TabPage3)
        Me.TabControl1.Controls.Add(Me.TabPage4)
        Me.TabControl1.Location = New System.Drawing.Point(12, 40)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(632, 475)
        Me.TabControl1.TabIndex = 9
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.btnApplyQuery)
        Me.TabPage1.Controls.Add(Me.btnChartNext)
        Me.TabPage1.Controls.Add(Me.btnChartPrev)
        Me.TabPage1.Controls.Add(Me.btnChartSelected)
        Me.TabPage1.Controls.Add(Me.DataGridView1)
        Me.TabPage1.Controls.Add(Me.cmbCompanyListDb)
        Me.TabPage1.Controls.Add(Me.Label29)
        Me.TabPage1.Controls.Add(Me.txtCompanyInfoQuery)
        Me.TabPage1.Controls.Add(Me.Label97)
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(624, 449)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Select Company Info"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'btnApplyQuery
        '
        Me.btnApplyQuery.Location = New System.Drawing.Point(6, 56)
        Me.btnApplyQuery.Name = "btnApplyQuery"
        Me.btnApplyQuery.Size = New System.Drawing.Size(42, 22)
        Me.btnApplyQuery.TabIndex = 150
        Me.btnApplyQuery.Text = "Apply"
        Me.btnApplyQuery.UseVisualStyleBackColor = True
        '
        'btnChartNext
        '
        Me.btnChartNext.Location = New System.Drawing.Point(471, 6)
        Me.btnChartNext.Name = "btnChartNext"
        Me.btnChartNext.Size = New System.Drawing.Size(42, 22)
        Me.btnChartNext.TabIndex = 149
        Me.btnChartNext.Text = "Next"
        Me.btnChartNext.UseVisualStyleBackColor = True
        '
        'btnChartPrev
        '
        Me.btnChartPrev.Location = New System.Drawing.Point(272, 6)
        Me.btnChartPrev.Name = "btnChartPrev"
        Me.btnChartPrev.Size = New System.Drawing.Size(42, 22)
        Me.btnChartPrev.TabIndex = 148
        Me.btnChartPrev.Text = "Prev"
        Me.btnChartPrev.UseVisualStyleBackColor = True
        '
        'btnChartSelected
        '
        Me.btnChartSelected.Location = New System.Drawing.Point(320, 6)
        Me.btnChartSelected.Name = "btnChartSelected"
        Me.btnChartSelected.Size = New System.Drawing.Size(145, 22)
        Me.btnChartSelected.TabIndex = 147
        Me.btnChartSelected.Text = "Chart selected company"
        Me.btnChartSelected.UseVisualStyleBackColor = True
        '
        'DataGridView1
        '
        Me.DataGridView1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Location = New System.Drawing.Point(9, 84)
        Me.DataGridView1.MultiSelect = False
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.DataGridView1.Size = New System.Drawing.Size(609, 359)
        Me.DataGridView1.TabIndex = 142
        '
        'cmbCompanyListDb
        '
        Me.cmbCompanyListDb.FormattingEnabled = True
        Me.cmbCompanyListDb.Location = New System.Drawing.Point(111, 6)
        Me.cmbCompanyListDb.Name = "cmbCompanyListDb"
        Me.cmbCompanyListDb.Size = New System.Drawing.Size(155, 21)
        Me.cmbCompanyListDb.TabIndex = 141
        '
        'Label29
        '
        Me.Label29.AutoSize = True
        Me.Label29.Location = New System.Drawing.Point(6, 9)
        Me.Label29.Name = "Label29"
        Me.Label29.Size = New System.Drawing.Size(99, 13)
        Me.Label29.TabIndex = 140
        Me.Label29.Text = "Selected database:"
        '
        'txtCompanyInfoQuery
        '
        Me.txtCompanyInfoQuery.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtCompanyInfoQuery.Location = New System.Drawing.Point(54, 33)
        Me.txtCompanyInfoQuery.Multiline = True
        Me.txtCompanyInfoQuery.Name = "txtCompanyInfoQuery"
        Me.txtCompanyInfoQuery.Size = New System.Drawing.Size(564, 45)
        Me.txtCompanyInfoQuery.TabIndex = 139
        '
        'Label97
        '
        Me.Label97.AutoSize = True
        Me.Label97.Location = New System.Drawing.Point(6, 36)
        Me.Label97.Name = "Label97"
        Me.Label97.Size = New System.Drawing.Size(38, 13)
        Me.Label97.TabIndex = 138
        Me.Label97.Text = "Query:"
        '
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.txtCompanyCode)
        Me.TabPage2.Controls.Add(Me.Label8)
        Me.TabPage2.Controls.Add(Me.Label7)
        Me.TabPage2.Controls.Add(Me.txtCompanyCodeColumn)
        Me.TabPage2.Controls.Add(Me.txtChartDataTable)
        Me.TabPage2.Controls.Add(Me.txtChartDataQuery)
        Me.TabPage2.Controls.Add(Me.Label4)
        Me.TabPage2.Controls.Add(Me.cmbCompanyCodeCol)
        Me.TabPage2.Controls.Add(Me.Label3)
        Me.TabPage2.Controls.Add(Me.Label2)
        Me.TabPage2.Location = New System.Drawing.Point(4, 22)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage2.Size = New System.Drawing.Size(624, 449)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "Chart Data Query"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'txtCompanyCode
        '
        Me.txtCompanyCode.Location = New System.Drawing.Point(152, 85)
        Me.txtCompanyCode.Name = "txtCompanyCode"
        Me.txtCompanyCode.Size = New System.Drawing.Size(296, 20)
        Me.txtCompanyCode.TabIndex = 145
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(6, 88)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(81, 13)
        Me.Label8.TabIndex = 144
        Me.Label8.Text = "Company code:"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(6, 61)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(140, 13)
        Me.Label7.TabIndex = 143
        Me.Label7.Text = "Set company code value to:"
        '
        'txtCompanyCodeColumn
        '
        Me.txtCompanyCodeColumn.Location = New System.Drawing.Point(152, 32)
        Me.txtCompanyCodeColumn.Name = "txtCompanyCodeColumn"
        Me.txtCompanyCodeColumn.Size = New System.Drawing.Size(296, 20)
        Me.txtCompanyCodeColumn.TabIndex = 142
        '
        'txtChartDataTable
        '
        Me.txtChartDataTable.Location = New System.Drawing.Point(152, 6)
        Me.txtChartDataTable.Name = "txtChartDataTable"
        Me.txtChartDataTable.Size = New System.Drawing.Size(296, 20)
        Me.txtChartDataTable.TabIndex = 141
        '
        'txtChartDataQuery
        '
        Me.txtChartDataQuery.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtChartDataQuery.Location = New System.Drawing.Point(50, 111)
        Me.txtChartDataQuery.Multiline = True
        Me.txtChartDataQuery.Name = "txtChartDataQuery"
        Me.txtChartDataQuery.Size = New System.Drawing.Size(568, 45)
        Me.txtChartDataQuery.TabIndex = 140
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(6, 111)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(38, 13)
        Me.Label4.TabIndex = 4
        Me.Label4.Text = "Query:"
        '
        'cmbCompanyCodeCol
        '
        Me.cmbCompanyCodeCol.FormattingEnabled = True
        Me.cmbCompanyCodeCol.Location = New System.Drawing.Point(152, 58)
        Me.cmbCompanyCodeCol.Name = "cmbCompanyCodeCol"
        Me.cmbCompanyCodeCol.Size = New System.Drawing.Size(296, 21)
        Me.cmbCompanyCodeCol.TabIndex = 3
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(6, 36)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(118, 13)
        Me.Label3.TabIndex = 2
        Me.Label3.Text = "Company code column:"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(6, 9)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(85, 13)
        Me.Label2.TabIndex = 0
        Me.Label2.Text = "Chart data table:"
        '
        'TabPage3
        '
        Me.TabPage3.Controls.Add(Me.txtSeriesName)
        Me.TabPage3.Controls.Add(Me.Label5)
        Me.TabPage3.Controls.Add(Me.dgvSeriesName)
        Me.TabPage3.Location = New System.Drawing.Point(4, 22)
        Me.TabPage3.Name = "TabPage3"
        Me.TabPage3.Size = New System.Drawing.Size(624, 449)
        Me.TabPage3.TabIndex = 2
        Me.TabPage3.Text = "Series Name"
        Me.TabPage3.UseVisualStyleBackColor = True
        '
        'txtSeriesName
        '
        Me.txtSeriesName.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtSeriesName.Location = New System.Drawing.Point(74, 423)
        Me.txtSeriesName.Name = "txtSeriesName"
        Me.txtSeriesName.Size = New System.Drawing.Size(544, 20)
        Me.txtSeriesName.TabIndex = 2
        '
        'Label5
        '
        Me.Label5.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(0, 426)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(68, 13)
        Me.Label5.TabIndex = 1
        Me.Label5.Text = "Series name:"
        '
        'dgvSeriesName
        '
        Me.dgvSeriesName.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvSeriesName.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvSeriesName.Location = New System.Drawing.Point(3, 3)
        Me.dgvSeriesName.Name = "dgvSeriesName"
        Me.dgvSeriesName.Size = New System.Drawing.Size(618, 414)
        Me.dgvSeriesName.TabIndex = 0
        '
        'TabPage4
        '
        Me.TabPage4.Controls.Add(Me.txtChartTitle)
        Me.TabPage4.Controls.Add(Me.Label6)
        Me.TabPage4.Controls.Add(Me.dgvChartTitle)
        Me.TabPage4.Location = New System.Drawing.Point(4, 22)
        Me.TabPage4.Name = "TabPage4"
        Me.TabPage4.Size = New System.Drawing.Size(624, 449)
        Me.TabPage4.TabIndex = 3
        Me.TabPage4.Text = "Chart Title"
        Me.TabPage4.UseVisualStyleBackColor = True
        '
        'txtChartTitle
        '
        Me.txtChartTitle.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtChartTitle.Location = New System.Drawing.Point(61, 387)
        Me.txtChartTitle.Multiline = True
        Me.txtChartTitle.Name = "txtChartTitle"
        Me.txtChartTitle.Size = New System.Drawing.Size(558, 57)
        Me.txtChartTitle.TabIndex = 5
        '
        'Label6
        '
        Me.Label6.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(1, 390)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(54, 13)
        Me.Label6.TabIndex = 4
        Me.Label6.Text = "Chart title:"
        '
        'dgvChartTitle
        '
        Me.dgvChartTitle.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvChartTitle.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvChartTitle.Location = New System.Drawing.Point(4, 4)
        Me.dgvChartTitle.Name = "dgvChartTitle"
        Me.dgvChartTitle.Size = New System.Drawing.Size(618, 377)
        Me.dgvChartTitle.TabIndex = 3
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(9, 17)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(42, 13)
        Me.Label1.TabIndex = 144
        Me.Label1.Text = "List file:"
        '
        'btnSaveCompanyList
        '
        Me.btnSaveCompanyList.Location = New System.Drawing.Point(57, 12)
        Me.btnSaveCompanyList.Name = "btnSaveCompanyList"
        Me.btnSaveCompanyList.Size = New System.Drawing.Size(42, 22)
        Me.btnSaveCompanyList.TabIndex = 146
        Me.btnSaveCompanyList.Text = "Save"
        Me.btnSaveCompanyList.UseVisualStyleBackColor = True
        '
        'txtCompanyListFile
        '
        Me.txtCompanyListFile.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtCompanyListFile.Location = New System.Drawing.Point(105, 12)
        Me.txtCompanyListFile.Name = "txtCompanyListFile"
        Me.txtCompanyListFile.Size = New System.Drawing.Size(421, 20)
        Me.txtCompanyListFile.TabIndex = 147
        '
        'btnFindCompanyList
        '
        Me.btnFindCompanyList.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnFindCompanyList.Location = New System.Drawing.Point(532, 12)
        Me.btnFindCompanyList.Name = "btnFindCompanyList"
        Me.btnFindCompanyList.Size = New System.Drawing.Size(42, 22)
        Me.btnFindCompanyList.TabIndex = 148
        Me.btnFindCompanyList.Text = "Find"
        Me.btnFindCompanyList.UseVisualStyleBackColor = True
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.FileName = "OpenFileDialog1"
        '
        'frmCompanyList
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(656, 527)
        Me.Controls.Add(Me.btnFindCompanyList)
        Me.Controls.Add(Me.txtCompanyListFile)
        Me.Controls.Add(Me.btnSaveCompanyList)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.TabControl1)
        Me.Controls.Add(Me.btnExit)
        Me.Name = "frmCompanyList"
        Me.Text = "Company List"
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.TabPage1.PerformLayout()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage2.ResumeLayout(False)
        Me.TabPage2.PerformLayout()
        Me.TabPage3.ResumeLayout(False)
        Me.TabPage3.PerformLayout()
        CType(Me.dgvSeriesName, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage4.ResumeLayout(False)
        Me.TabPage4.PerformLayout()
        CType(Me.dgvChartTitle, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents btnExit As Button
    Friend WithEvents TabControl1 As TabControl
    Friend WithEvents TabPage1 As TabPage
    Friend WithEvents txtCompanyInfoQuery As TextBox
    Friend WithEvents Label97 As Label
    Friend WithEvents TabPage2 As TabPage
    Friend WithEvents DataGridView1 As DataGridView
    Friend WithEvents cmbCompanyListDb As ComboBox
    Friend WithEvents Label29 As Label
    Friend WithEvents Label1 As Label
    Friend WithEvents btnSaveCompanyList As Button
    Friend WithEvents txtCompanyListFile As TextBox
    Friend WithEvents btnFindCompanyList As Button
    Friend WithEvents btnChartNext As Button
    Friend WithEvents btnChartPrev As Button
    Friend WithEvents btnChartSelected As Button
    Friend WithEvents txtChartDataQuery As TextBox
    Friend WithEvents Label4 As Label
    Friend WithEvents cmbCompanyCodeCol As ComboBox
    Friend WithEvents Label3 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents TabPage3 As TabPage
    Friend WithEvents TabPage4 As TabPage
    Friend WithEvents txtSeriesName As TextBox
    Friend WithEvents Label5 As Label
    Friend WithEvents dgvSeriesName As DataGridView
    Friend WithEvents txtChartTitle As TextBox
    Friend WithEvents Label6 As Label
    Friend WithEvents dgvChartTitle As DataGridView
    Friend WithEvents btnApplyQuery As Button
    Friend WithEvents Label7 As Label
    Friend WithEvents txtCompanyCodeColumn As TextBox
    Friend WithEvents txtChartDataTable As TextBox
    Friend WithEvents OpenFileDialog1 As OpenFileDialog
    Friend WithEvents txtCompanyCode As TextBox
    Friend WithEvents Label8 As Label
End Class
