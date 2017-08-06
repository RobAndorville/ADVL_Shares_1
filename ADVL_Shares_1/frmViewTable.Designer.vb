<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmViewTable
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
        Me.btnDesign = New System.Windows.Forms.Button()
        Me.btnSaveChanges = New System.Windows.Forms.Button()
        Me.chkAutoApply = New System.Windows.Forms.CheckBox()
        Me.btnDisplay = New System.Windows.Forms.Button()
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.btnApplyQuery = New System.Windows.Forms.Button()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.txtQuery = New System.Windows.Forms.TextBox()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.cmbSelectTable = New System.Windows.Forms.ComboBox()
        Me.txtDataDescr = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.btnExit = New System.Windows.Forms.Button()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'btnDesign
        '
        Me.btnDesign.Location = New System.Drawing.Point(666, 93)
        Me.btnDesign.Name = "btnDesign"
        Me.btnDesign.Size = New System.Drawing.Size(64, 22)
        Me.btnDesign.TabIndex = 70
        Me.btnDesign.Text = "Design"
        Me.btnDesign.UseVisualStyleBackColor = True
        '
        'btnSaveChanges
        '
        Me.btnSaveChanges.Location = New System.Drawing.Point(12, 12)
        Me.btnSaveChanges.Name = "btnSaveChanges"
        Me.btnSaveChanges.Size = New System.Drawing.Size(92, 22)
        Me.btnSaveChanges.TabIndex = 69
        Me.btnSaveChanges.Text = "Save Changes"
        Me.btnSaveChanges.UseVisualStyleBackColor = True
        '
        'chkAutoApply
        '
        Me.chkAutoApply.AutoSize = True
        Me.chkAutoApply.Location = New System.Drawing.Point(12, 96)
        Me.chkAutoApply.Name = "chkAutoApply"
        Me.chkAutoApply.Size = New System.Drawing.Size(77, 17)
        Me.chkAutoApply.TabIndex = 68
        Me.chkAutoApply.Text = "Auto Apply"
        Me.chkAutoApply.UseVisualStyleBackColor = True
        '
        'btnDisplay
        '
        Me.btnDisplay.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnDisplay.Location = New System.Drawing.Point(666, 40)
        Me.btnDisplay.Name = "btnDisplay"
        Me.btnDisplay.Size = New System.Drawing.Size(64, 22)
        Me.btnDisplay.TabIndex = 67
        Me.btnDisplay.Text = "Display"
        Me.btnDisplay.UseVisualStyleBackColor = True
        '
        'DataGridView1
        '
        Me.DataGridView1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Location = New System.Drawing.Point(15, 120)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.Size = New System.Drawing.Size(715, 421)
        Me.DataGridView1.TabIndex = 66
        '
        'btnApplyQuery
        '
        Me.btnApplyQuery.Location = New System.Drawing.Point(666, 67)
        Me.btnApplyQuery.Name = "btnApplyQuery"
        Me.btnApplyQuery.Size = New System.Drawing.Size(64, 22)
        Me.btnApplyQuery.TabIndex = 65
        Me.btnApplyQuery.Text = "Apply"
        Me.btnApplyQuery.UseVisualStyleBackColor = True
        '
        'Label15
        '
        Me.Label15.AutoSize = True
        Me.Label15.Location = New System.Drawing.Point(12, 73)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(38, 13)
        Me.Label15.TabIndex = 64
        Me.Label15.Text = "Query:"
        '
        'txtQuery
        '
        Me.txtQuery.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtQuery.Location = New System.Drawing.Point(95, 68)
        Me.txtQuery.Multiline = True
        Me.txtQuery.Name = "txtQuery"
        Me.txtQuery.Size = New System.Drawing.Size(565, 46)
        Me.txtQuery.TabIndex = 63
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.Location = New System.Drawing.Point(12, 45)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(42, 13)
        Me.Label14.TabIndex = 62
        Me.Label14.Text = "Tables:"
        '
        'cmbSelectTable
        '
        Me.cmbSelectTable.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmbSelectTable.FormattingEnabled = True
        Me.cmbSelectTable.Location = New System.Drawing.Point(95, 41)
        Me.cmbSelectTable.Name = "cmbSelectTable"
        Me.cmbSelectTable.Size = New System.Drawing.Size(565, 21)
        Me.cmbSelectTable.TabIndex = 61
        '
        'txtDataDescr
        '
        Me.txtDataDescr.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtDataDescr.Location = New System.Drawing.Point(203, 13)
        Me.txtDataDescr.Name = "txtDataDescr"
        Me.txtDataDescr.Size = New System.Drawing.Size(457, 20)
        Me.txtDataDescr.TabIndex = 60
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(110, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(87, 13)
        Me.Label1.TabIndex = 59
        Me.Label1.Text = "Data description:"
        '
        'btnExit
        '
        Me.btnExit.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnExit.Location = New System.Drawing.Point(666, 12)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(64, 22)
        Me.btnExit.TabIndex = 58
        Me.btnExit.Text = "Exit"
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'frmViewTable
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(742, 553)
        Me.Controls.Add(Me.btnDesign)
        Me.Controls.Add(Me.btnSaveChanges)
        Me.Controls.Add(Me.chkAutoApply)
        Me.Controls.Add(Me.btnDisplay)
        Me.Controls.Add(Me.DataGridView1)
        Me.Controls.Add(Me.btnApplyQuery)
        Me.Controls.Add(Me.Label15)
        Me.Controls.Add(Me.txtQuery)
        Me.Controls.Add(Me.Label14)
        Me.Controls.Add(Me.cmbSelectTable)
        Me.Controls.Add(Me.txtDataDescr)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.btnExit)
        Me.Name = "frmViewTable"
        Me.Text = "View Table"
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents btnDesign As Button
    Friend WithEvents btnSaveChanges As Button
    Friend WithEvents chkAutoApply As CheckBox
    Friend WithEvents btnDisplay As Button
    Friend WithEvents DataGridView1 As DataGridView
    Friend WithEvents btnApplyQuery As Button
    Friend WithEvents Label15 As Label
    Friend WithEvents txtQuery As TextBox
    Friend WithEvents Label14 As Label
    Friend WithEvents cmbSelectTable As ComboBox
    Friend WithEvents txtDataDescr As TextBox
    Friend WithEvents Label1 As Label
    Friend WithEvents btnExit As Button
End Class
