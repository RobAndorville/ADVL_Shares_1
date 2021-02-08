<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmCalculations
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
        Me.components = New System.ComponentModel.Container()
        Me.btnSaveChanges = New System.Windows.Forms.Button()
        Me.txtDataName = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.btnExit = New System.Windows.Forms.Button()
        Me.TabControl1 = New System.Windows.Forms.TabControl()
        Me.TabPage1 = New System.Windows.Forms.TabPage()
        Me.TabPage2 = New System.Windows.Forms.TabPage()
        Me.btnSaveVersionChanges = New System.Windows.Forms.Button()
        Me.btnCancelVersionChanges = New System.Windows.Forms.Button()
        Me.btnNewVersion = New System.Windows.Forms.Button()
        Me.txtVersionName = New System.Windows.Forms.TextBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.txtVersionDesc = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.btnDesign = New System.Windows.Forms.Button()
        Me.btnApplyQuery = New System.Windows.Forms.Button()
        Me.chkAutoApply = New System.Windows.Forms.CheckBox()
        Me.btnDisplay = New System.Windows.Forms.Button()
        Me.txtQuery = New System.Windows.Forms.TextBox()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.cmbSelectTable = New System.Windows.Forms.ComboBox()
        Me.TabPage3 = New System.Windows.Forms.TabPage()
        Me.btnDelete = New System.Windows.Forms.Button()
        Me.btnMoveDown = New System.Windows.Forms.Button()
        Me.btnMoveUp = New System.Windows.Forms.Button()
        Me.btnSelect = New System.Windows.Forms.Button()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.ListBox1 = New System.Windows.Forms.ListBox()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.txtSelVersionDesc = New System.Windows.Forms.TextBox()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.txtSelVersionQuery = New System.Windows.Forms.TextBox()
        Me.TabPage4 = New System.Windows.Forms.TabPage()
        Me.btnFind = New System.Windows.Forms.Button()
        Me.txtDirectory = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.txtXmlFileName = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.btnSaveAsXml = New System.Windows.Forms.Button()
        Me.txtSelectedRecord = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.txtNRecords = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.FolderBrowserDialog1 = New System.Windows.Forms.FolderBrowserDialog()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.txtDataVersion = New System.Windows.Forms.TextBox()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.TabPage2.SuspendLayout()
        Me.TabPage3.SuspendLayout()
        Me.TabPage4.SuspendLayout()
        Me.SuspendLayout()
        '
        'btnSaveChanges
        '
        Me.btnSaveChanges.Location = New System.Drawing.Point(12, 11)
        Me.btnSaveChanges.Name = "btnSaveChanges"
        Me.btnSaveChanges.Size = New System.Drawing.Size(92, 22)
        Me.btnSaveChanges.TabIndex = 72
        Me.btnSaveChanges.Text = "Save Changes"
        Me.ToolTip1.SetToolTip(Me.btnSaveChanges, "Save the changes made in the Data View to the corresponding table in the database" &
        ".")
        Me.btnSaveChanges.UseVisualStyleBackColor = True
        '
        'txtDataName
        '
        Me.txtDataName.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtDataName.Location = New System.Drawing.Point(178, 13)
        Me.txtDataName.Name = "txtDataName"
        Me.txtDataName.Size = New System.Drawing.Size(443, 20)
        Me.txtDataName.TabIndex = 68
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(110, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(62, 13)
        Me.Label1.TabIndex = 67
        Me.Label1.Text = "Data name:"
        '
        'DataGridView1
        '
        Me.DataGridView1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Location = New System.Drawing.Point(6, 6)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.Size = New System.Drawing.Size(659, 419)
        Me.DataGridView1.TabIndex = 66
        '
        'btnExit
        '
        Me.btnExit.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnExit.Location = New System.Drawing.Point(627, 12)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(64, 22)
        Me.btnExit.TabIndex = 60
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
        Me.TabControl1.Location = New System.Drawing.Point(12, 65)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(679, 457)
        Me.TabControl1.TabIndex = 73
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.DataGridView1)
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(671, 431)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Data"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.btnSaveVersionChanges)
        Me.TabPage2.Controls.Add(Me.btnCancelVersionChanges)
        Me.TabPage2.Controls.Add(Me.btnNewVersion)
        Me.TabPage2.Controls.Add(Me.txtVersionName)
        Me.TabPage2.Controls.Add(Me.Label7)
        Me.TabPage2.Controls.Add(Me.txtVersionDesc)
        Me.TabPage2.Controls.Add(Me.Label6)
        Me.TabPage2.Controls.Add(Me.btnDesign)
        Me.TabPage2.Controls.Add(Me.btnApplyQuery)
        Me.TabPage2.Controls.Add(Me.chkAutoApply)
        Me.TabPage2.Controls.Add(Me.btnDisplay)
        Me.TabPage2.Controls.Add(Me.txtQuery)
        Me.TabPage2.Controls.Add(Me.Label14)
        Me.TabPage2.Controls.Add(Me.Label15)
        Me.TabPage2.Controls.Add(Me.cmbSelectTable)
        Me.TabPage2.Location = New System.Drawing.Point(4, 22)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage2.Size = New System.Drawing.Size(671, 431)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "Query"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'btnSaveVersionChanges
        '
        Me.btnSaveVersionChanges.Location = New System.Drawing.Point(75, 118)
        Me.btnSaveVersionChanges.Name = "btnSaveVersionChanges"
        Me.btnSaveVersionChanges.Size = New System.Drawing.Size(133, 22)
        Me.btnSaveVersionChanges.TabIndex = 86
        Me.btnSaveVersionChanges.Text = "Save Version Changes"
        Me.btnSaveVersionChanges.UseVisualStyleBackColor = True
        '
        'btnCancelVersionChanges
        '
        Me.btnCancelVersionChanges.Location = New System.Drawing.Point(214, 118)
        Me.btnCancelVersionChanges.Name = "btnCancelVersionChanges"
        Me.btnCancelVersionChanges.Size = New System.Drawing.Size(133, 22)
        Me.btnCancelVersionChanges.TabIndex = 85
        Me.btnCancelVersionChanges.Text = "Cancel Version Changes"
        Me.btnCancelVersionChanges.UseVisualStyleBackColor = True
        '
        'btnNewVersion
        '
        Me.btnNewVersion.Location = New System.Drawing.Point(353, 118)
        Me.btnNewVersion.Name = "btnNewVersion"
        Me.btnNewVersion.Size = New System.Drawing.Size(133, 22)
        Me.btnNewVersion.TabIndex = 84
        Me.btnNewVersion.Text = "Save As New Version"
        Me.btnNewVersion.UseVisualStyleBackColor = True
        '
        'txtVersionName
        '
        Me.txtVersionName.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtVersionName.Location = New System.Drawing.Point(75, 146)
        Me.txtVersionName.Name = "txtVersionName"
        Me.txtVersionName.Size = New System.Drawing.Size(590, 20)
        Me.txtVersionName.TabIndex = 83
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(6, 149)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(45, 13)
        Me.Label7.TabIndex = 82
        Me.Label7.Text = "Version:"
        '
        'txtVersionDesc
        '
        Me.txtVersionDesc.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtVersionDesc.Location = New System.Drawing.Point(75, 172)
        Me.txtVersionDesc.Multiline = True
        Me.txtVersionDesc.Name = "txtVersionDesc"
        Me.txtVersionDesc.Size = New System.Drawing.Size(590, 77)
        Me.txtVersionDesc.TabIndex = 81
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(6, 175)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(63, 13)
        Me.Label6.TabIndex = 80
        Me.Label6.Text = "Description:"
        '
        'btnDesign
        '
        Me.btnDesign.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnDesign.Location = New System.Drawing.Point(601, 59)
        Me.btnDesign.Name = "btnDesign"
        Me.btnDesign.Size = New System.Drawing.Size(64, 22)
        Me.btnDesign.TabIndex = 79
        Me.btnDesign.Text = "Design"
        Me.btnDesign.UseVisualStyleBackColor = True
        '
        'btnApplyQuery
        '
        Me.btnApplyQuery.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnApplyQuery.Location = New System.Drawing.Point(601, 33)
        Me.btnApplyQuery.Name = "btnApplyQuery"
        Me.btnApplyQuery.Size = New System.Drawing.Size(64, 22)
        Me.btnApplyQuery.TabIndex = 76
        Me.btnApplyQuery.Text = "Apply"
        Me.btnApplyQuery.UseVisualStyleBackColor = True
        '
        'chkAutoApply
        '
        Me.chkAutoApply.AutoSize = True
        Me.chkAutoApply.Location = New System.Drawing.Point(6, 63)
        Me.chkAutoApply.Name = "chkAutoApply"
        Me.chkAutoApply.Size = New System.Drawing.Size(77, 17)
        Me.chkAutoApply.TabIndex = 78
        Me.chkAutoApply.Text = "Auto Apply"
        Me.chkAutoApply.UseVisualStyleBackColor = True
        '
        'btnDisplay
        '
        Me.btnDisplay.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnDisplay.Location = New System.Drawing.Point(601, 6)
        Me.btnDisplay.Name = "btnDisplay"
        Me.btnDisplay.Size = New System.Drawing.Size(64, 22)
        Me.btnDisplay.TabIndex = 77
        Me.btnDisplay.Text = "Display"
        Me.btnDisplay.UseVisualStyleBackColor = True
        '
        'txtQuery
        '
        Me.txtQuery.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtQuery.Location = New System.Drawing.Point(89, 35)
        Me.txtQuery.Multiline = True
        Me.txtQuery.Name = "txtQuery"
        Me.txtQuery.Size = New System.Drawing.Size(506, 77)
        Me.txtQuery.TabIndex = 74
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.Location = New System.Drawing.Point(6, 11)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(42, 13)
        Me.Label14.TabIndex = 73
        Me.Label14.Text = "Tables:"
        '
        'Label15
        '
        Me.Label15.AutoSize = True
        Me.Label15.Location = New System.Drawing.Point(6, 38)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(38, 13)
        Me.Label15.TabIndex = 75
        Me.Label15.Text = "Query:"
        '
        'cmbSelectTable
        '
        Me.cmbSelectTable.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmbSelectTable.FormattingEnabled = True
        Me.cmbSelectTable.Location = New System.Drawing.Point(89, 8)
        Me.cmbSelectTable.Name = "cmbSelectTable"
        Me.cmbSelectTable.Size = New System.Drawing.Size(506, 21)
        Me.cmbSelectTable.TabIndex = 72
        '
        'TabPage3
        '
        Me.TabPage3.Controls.Add(Me.btnDelete)
        Me.TabPage3.Controls.Add(Me.btnMoveDown)
        Me.TabPage3.Controls.Add(Me.btnMoveUp)
        Me.TabPage3.Controls.Add(Me.btnSelect)
        Me.TabPage3.Controls.Add(Me.Label11)
        Me.TabPage3.Controls.Add(Me.ListBox1)
        Me.TabPage3.Controls.Add(Me.Label10)
        Me.TabPage3.Controls.Add(Me.txtSelVersionDesc)
        Me.TabPage3.Controls.Add(Me.Label9)
        Me.TabPage3.Controls.Add(Me.txtSelVersionQuery)
        Me.TabPage3.Location = New System.Drawing.Point(4, 22)
        Me.TabPage3.Name = "TabPage3"
        Me.TabPage3.Size = New System.Drawing.Size(671, 431)
        Me.TabPage3.TabIndex = 2
        Me.TabPage3.Text = "Versions"
        Me.TabPage3.UseVisualStyleBackColor = True
        '
        'btnDelete
        '
        Me.btnDelete.Location = New System.Drawing.Point(9, 310)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(69, 22)
        Me.btnDelete.TabIndex = 78
        Me.btnDelete.Text = "Delete"
        Me.btnDelete.UseVisualStyleBackColor = True
        '
        'btnMoveDown
        '
        Me.btnMoveDown.Location = New System.Drawing.Point(9, 254)
        Me.btnMoveDown.Name = "btnMoveDown"
        Me.btnMoveDown.Size = New System.Drawing.Size(69, 22)
        Me.btnMoveDown.TabIndex = 77
        Me.btnMoveDown.Text = "Move Dwn"
        Me.btnMoveDown.UseVisualStyleBackColor = True
        '
        'btnMoveUp
        '
        Me.btnMoveUp.Location = New System.Drawing.Point(9, 226)
        Me.btnMoveUp.Name = "btnMoveUp"
        Me.btnMoveUp.Size = New System.Drawing.Size(69, 22)
        Me.btnMoveUp.TabIndex = 76
        Me.btnMoveUp.Text = "Move Up"
        Me.btnMoveUp.UseVisualStyleBackColor = True
        '
        'btnSelect
        '
        Me.btnSelect.Location = New System.Drawing.Point(9, 198)
        Me.btnSelect.Name = "btnSelect"
        Me.btnSelect.Size = New System.Drawing.Size(69, 22)
        Me.btnSelect.TabIndex = 75
        Me.btnSelect.Text = "Select"
        Me.btnSelect.UseVisualStyleBackColor = True
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(6, 182)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(45, 13)
        Me.Label11.TabIndex = 74
        Me.Label11.Text = "Version:"
        '
        'ListBox1
        '
        Me.ListBox1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ListBox1.FormattingEnabled = True
        Me.ListBox1.Location = New System.Drawing.Point(84, 182)
        Me.ListBox1.Name = "ListBox1"
        Me.ListBox1.Size = New System.Drawing.Size(581, 225)
        Me.ListBox1.TabIndex = 73
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(6, 92)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(63, 13)
        Me.Label10.TabIndex = 72
        Me.Label10.Text = "Description:"
        '
        'txtSelVersionDesc
        '
        Me.txtSelVersionDesc.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtSelVersionDesc.Location = New System.Drawing.Point(75, 89)
        Me.txtSelVersionDesc.Multiline = True
        Me.txtSelVersionDesc.Name = "txtSelVersionDesc"
        Me.txtSelVersionDesc.Size = New System.Drawing.Size(590, 77)
        Me.txtSelVersionDesc.TabIndex = 71
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(6, 9)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(38, 13)
        Me.Label9.TabIndex = 70
        Me.Label9.Text = "Query:"
        '
        'txtSelVersionQuery
        '
        Me.txtSelVersionQuery.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtSelVersionQuery.Location = New System.Drawing.Point(75, 6)
        Me.txtSelVersionQuery.Multiline = True
        Me.txtSelVersionQuery.Name = "txtSelVersionQuery"
        Me.txtSelVersionQuery.Size = New System.Drawing.Size(590, 77)
        Me.txtSelVersionQuery.TabIndex = 69
        '
        'TabPage4
        '
        Me.TabPage4.Controls.Add(Me.btnFind)
        Me.TabPage4.Controls.Add(Me.txtDirectory)
        Me.TabPage4.Controls.Add(Me.Label5)
        Me.TabPage4.Controls.Add(Me.txtXmlFileName)
        Me.TabPage4.Controls.Add(Me.Label4)
        Me.TabPage4.Controls.Add(Me.btnSaveAsXml)
        Me.TabPage4.Controls.Add(Me.txtSelectedRecord)
        Me.TabPage4.Controls.Add(Me.Label3)
        Me.TabPage4.Controls.Add(Me.txtNRecords)
        Me.TabPage4.Controls.Add(Me.Label2)
        Me.TabPage4.Location = New System.Drawing.Point(4, 22)
        Me.TabPage4.Name = "TabPage4"
        Me.TabPage4.Size = New System.Drawing.Size(671, 431)
        Me.TabPage4.TabIndex = 3
        Me.TabPage4.Text = "Information"
        Me.TabPage4.UseVisualStyleBackColor = True
        '
        'btnFind
        '
        Me.btnFind.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnFind.Location = New System.Drawing.Point(600, 112)
        Me.btnFind.Name = "btnFind"
        Me.btnFind.Size = New System.Drawing.Size(64, 22)
        Me.btnFind.TabIndex = 39
        Me.btnFind.Text = "Find"
        Me.btnFind.UseVisualStyleBackColor = True
        '
        'txtDirectory
        '
        Me.txtDirectory.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtDirectory.Location = New System.Drawing.Point(66, 114)
        Me.txtDirectory.Name = "txtDirectory"
        Me.txtDirectory.Size = New System.Drawing.Size(528, 20)
        Me.txtDirectory.TabIndex = 38
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(4, 117)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(52, 13)
        Me.Label5.TabIndex = 37
        Me.Label5.Text = "Directory:"
        '
        'txtXmlFileName
        '
        Me.txtXmlFileName.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtXmlFileName.Location = New System.Drawing.Point(66, 140)
        Me.txtXmlFileName.Name = "txtXmlFileName"
        Me.txtXmlFileName.Size = New System.Drawing.Size(528, 20)
        Me.txtXmlFileName.TabIndex = 36
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(4, 143)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(55, 13)
        Me.Label4.TabIndex = 35
        Me.Label4.Text = "File name:"
        '
        'btnSaveAsXml
        '
        Me.btnSaveAsXml.Location = New System.Drawing.Point(66, 86)
        Me.btnSaveAsXml.Name = "btnSaveAsXml"
        Me.btnSaveAsXml.Size = New System.Drawing.Size(134, 22)
        Me.btnSaveAsXml.TabIndex = 34
        Me.btnSaveAsXml.Text = "Save Data in XML File"
        Me.btnSaveAsXml.UseVisualStyleBackColor = True
        '
        'txtSelectedRecord
        '
        Me.txtSelectedRecord.Location = New System.Drawing.Point(110, 40)
        Me.txtSelectedRecord.Name = "txtSelectedRecord"
        Me.txtSelectedRecord.Size = New System.Drawing.Size(195, 20)
        Me.txtSelectedRecord.TabIndex = 33
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(4, 43)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(100, 13)
        Me.Label3.TabIndex = 32
        Me.Label3.Text = "Selected row index:"
        '
        'txtNRecords
        '
        Me.txtNRecords.Location = New System.Drawing.Point(110, 12)
        Me.txtNRecords.Name = "txtNRecords"
        Me.txtNRecords.Size = New System.Drawing.Size(195, 20)
        Me.txtNRecords.TabIndex = 31
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(4, 15)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(50, 13)
        Me.Label2.TabIndex = 30
        Me.Label2.Text = "Records:"
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(9, 42)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(45, 13)
        Me.Label8.TabIndex = 75
        Me.Label8.Text = "Version:"
        '
        'txtDataVersion
        '
        Me.txtDataVersion.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtDataVersion.Location = New System.Drawing.Point(60, 39)
        Me.txtDataVersion.Name = "txtDataVersion"
        Me.txtDataVersion.Size = New System.Drawing.Size(631, 20)
        Me.txtDataVersion.TabIndex = 74
        '
        'frmCalculations
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(703, 534)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.txtDataVersion)
        Me.Controls.Add(Me.TabControl1)
        Me.Controls.Add(Me.btnSaveChanges)
        Me.Controls.Add(Me.txtDataName)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.btnExit)
        Me.Name = "frmCalculations"
        Me.Text = "Calculations"
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.TabPage2.ResumeLayout(False)
        Me.TabPage2.PerformLayout()
        Me.TabPage3.ResumeLayout(False)
        Me.TabPage3.PerformLayout()
        Me.TabPage4.ResumeLayout(False)
        Me.TabPage4.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents btnSaveChanges As Button
    Friend WithEvents txtDataName As TextBox
    Friend WithEvents Label1 As Label
    Friend WithEvents DataGridView1 As DataGridView
    Friend WithEvents btnExit As Button
    Friend WithEvents TabControl1 As TabControl
    Friend WithEvents TabPage1 As TabPage
    Friend WithEvents TabPage2 As TabPage
    Friend WithEvents FolderBrowserDialog1 As FolderBrowserDialog
    Friend WithEvents btnDesign As Button
    Friend WithEvents btnApplyQuery As Button
    Friend WithEvents chkAutoApply As CheckBox
    Friend WithEvents btnDisplay As Button
    Friend WithEvents txtQuery As TextBox
    Friend WithEvents Label14 As Label
    Friend WithEvents Label15 As Label
    Friend WithEvents cmbSelectTable As ComboBox
    Friend WithEvents TabPage3 As TabPage
    Friend WithEvents TabPage4 As TabPage
    Friend WithEvents btnFind As Button
    Friend WithEvents txtDirectory As TextBox
    Friend WithEvents Label5 As Label
    Friend WithEvents txtXmlFileName As TextBox
    Friend WithEvents Label4 As Label
    Friend WithEvents btnSaveAsXml As Button
    Friend WithEvents txtSelectedRecord As TextBox
    Friend WithEvents Label3 As Label
    Friend WithEvents txtNRecords As TextBox
    Friend WithEvents Label2 As Label
    Friend WithEvents Label8 As Label
    Friend WithEvents txtDataVersion As TextBox
    Friend WithEvents btnSaveVersionChanges As Button
    Friend WithEvents btnCancelVersionChanges As Button
    Friend WithEvents btnNewVersion As Button
    Friend WithEvents txtVersionName As TextBox
    Friend WithEvents Label7 As Label
    Friend WithEvents txtVersionDesc As TextBox
    Friend WithEvents Label6 As Label
    Friend WithEvents btnDelete As Button
    Friend WithEvents btnMoveDown As Button
    Friend WithEvents btnMoveUp As Button
    Friend WithEvents btnSelect As Button
    Friend WithEvents Label11 As Label
    Friend WithEvents ListBox1 As ListBox
    Friend WithEvents Label10 As Label
    Friend WithEvents txtSelVersionDesc As TextBox
    Friend WithEvents Label9 As Label
    Friend WithEvents txtSelVersionQuery As TextBox
    Friend WithEvents ToolTip1 As ToolTip
End Class
