<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmWebPageList
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
        Me.btnDelete = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.txtNewHtmlFileTitle = New System.Windows.Forms.TextBox()
        Me.txtNewHtmlFileName = New System.Windows.Forms.TextBox()
        Me.btnNew = New System.Windows.Forms.Button()
        Me.btnEdit = New System.Windows.Forms.Button()
        Me.btnOpen = New System.Windows.Forms.Button()
        Me.lstWebPages = New System.Windows.Forms.ListBox()
        Me.btnExit = New System.Windows.Forms.Button()
        Me.btnOpenInMain = New System.Windows.Forms.Button()
        Me.btnHome = New System.Windows.Forms.Button()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.txtEdited = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.txtCreated = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.txtDescription = New System.Windows.Forms.TextBox()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'btnDelete
        '
        Me.btnDelete.Location = New System.Drawing.Point(120, 14)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(48, 22)
        Me.btnDelete.TabIndex = 25
        Me.btnDelete.Text = "Delete"
        Me.btnDelete.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(18, 52)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(30, 13)
        Me.Label1.TabIndex = 24
        Me.Label1.Text = "Title:"
        '
        'txtNewHtmlFileTitle
        '
        Me.txtNewHtmlFileTitle.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtNewHtmlFileTitle.Location = New System.Drawing.Point(63, 49)
        Me.txtNewHtmlFileTitle.Name = "txtNewHtmlFileTitle"
        Me.txtNewHtmlFileTitle.Size = New System.Drawing.Size(518, 20)
        Me.txtNewHtmlFileTitle.TabIndex = 23
        '
        'txtNewHtmlFileName
        '
        Me.txtNewHtmlFileName.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtNewHtmlFileName.Location = New System.Drawing.Point(63, 19)
        Me.txtNewHtmlFileName.Name = "txtNewHtmlFileName"
        Me.txtNewHtmlFileName.Size = New System.Drawing.Size(518, 20)
        Me.txtNewHtmlFileName.TabIndex = 22
        '
        'btnNew
        '
        Me.btnNew.Location = New System.Drawing.Point(9, 19)
        Me.btnNew.Name = "btnNew"
        Me.btnNew.Size = New System.Drawing.Size(48, 22)
        Me.btnNew.TabIndex = 21
        Me.btnNew.Text = "New"
        Me.btnNew.UseVisualStyleBackColor = True
        '
        'btnEdit
        '
        Me.btnEdit.Location = New System.Drawing.Point(66, 14)
        Me.btnEdit.Name = "btnEdit"
        Me.btnEdit.Size = New System.Drawing.Size(48, 22)
        Me.btnEdit.TabIndex = 20
        Me.btnEdit.Text = "Edit"
        Me.btnEdit.UseVisualStyleBackColor = True
        '
        'btnOpen
        '
        Me.btnOpen.Location = New System.Drawing.Point(12, 14)
        Me.btnOpen.Name = "btnOpen"
        Me.btnOpen.Size = New System.Drawing.Size(48, 22)
        Me.btnOpen.TabIndex = 19
        Me.btnOpen.Text = "Open"
        Me.btnOpen.UseVisualStyleBackColor = True
        '
        'lstWebPages
        '
        Me.lstWebPages.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lstWebPages.FormattingEnabled = True
        Me.lstWebPages.Location = New System.Drawing.Point(12, 222)
        Me.lstWebPages.Name = "lstWebPages"
        Me.lstWebPages.Size = New System.Drawing.Size(587, 264)
        Me.lstWebPages.TabIndex = 18
        '
        'btnExit
        '
        Me.btnExit.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnExit.Location = New System.Drawing.Point(551, 14)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(48, 22)
        Me.btnExit.TabIndex = 17
        Me.btnExit.Text = "Exit"
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'btnOpenInMain
        '
        Me.btnOpenInMain.Location = New System.Drawing.Point(174, 14)
        Me.btnOpenInMain.Name = "btnOpenInMain"
        Me.btnOpenInMain.Size = New System.Drawing.Size(109, 22)
        Me.btnOpenInMain.TabIndex = 36
        Me.btnOpenInMain.Text = "Open in Main Form"
        Me.ToolTip1.SetToolTip(Me.btnOpenInMain, "Open the Start Page on the Main form")
        Me.btnOpenInMain.UseVisualStyleBackColor = True
        '
        'btnHome
        '
        Me.btnHome.Location = New System.Drawing.Point(289, 14)
        Me.btnHome.Name = "btnHome"
        Me.btnHome.Size = New System.Drawing.Size(48, 22)
        Me.btnHome.TabIndex = 37
        Me.btnHome.Text = "Home"
        Me.ToolTip1.SetToolTip(Me.btnHome, "Open the Start Page on the Main form Workflow tab")
        Me.btnHome.UseVisualStyleBackColor = True
        '
        'txtEdited
        '
        Me.txtEdited.Location = New System.Drawing.Point(258, 128)
        Me.txtEdited.Name = "txtEdited"
        Me.txtEdited.Size = New System.Drawing.Size(140, 20)
        Me.txtEdited.TabIndex = 44
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(212, 131)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(40, 13)
        Me.Label3.TabIndex = 43
        Me.Label3.Text = "Edited:"
        '
        'txtCreated
        '
        Me.txtCreated.Location = New System.Drawing.Point(66, 128)
        Me.txtCreated.Name = "txtCreated"
        Me.txtCreated.Size = New System.Drawing.Size(140, 20)
        Me.txtCreated.TabIndex = 42
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(13, 131)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(47, 13)
        Me.Label2.TabIndex = 41
        Me.Label2.Text = "Created:"
        '
        'txtDescription
        '
        Me.txtDescription.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtDescription.Location = New System.Drawing.Point(12, 154)
        Me.txtDescription.Multiline = True
        Me.txtDescription.Name = "txtDescription"
        Me.txtDescription.Size = New System.Drawing.Size(587, 62)
        Me.txtDescription.TabIndex = 45
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.Controls.Add(Me.txtNewHtmlFileTitle)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Controls.Add(Me.btnNew)
        Me.GroupBox1.Controls.Add(Me.txtNewHtmlFileName)
        Me.GroupBox1.Location = New System.Drawing.Point(12, 42)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(587, 75)
        Me.GroupBox1.TabIndex = 46
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Create New Workflow:"
        '
        'frmWebPageList
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(611, 497)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.txtDescription)
        Me.Controls.Add(Me.txtEdited)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.txtCreated)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.btnHome)
        Me.Controls.Add(Me.btnOpenInMain)
        Me.Controls.Add(Me.btnDelete)
        Me.Controls.Add(Me.btnEdit)
        Me.Controls.Add(Me.btnOpen)
        Me.Controls.Add(Me.lstWebPages)
        Me.Controls.Add(Me.btnExit)
        Me.Name = "frmWebPageList"
        Me.Text = "Workflow Web Pages"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents btnDelete As Button
    Friend WithEvents Label1 As Label
    Friend WithEvents txtNewHtmlFileTitle As TextBox
    Friend WithEvents txtNewHtmlFileName As TextBox
    Friend WithEvents btnNew As Button
    Friend WithEvents btnEdit As Button
    Friend WithEvents btnOpen As Button
    Friend WithEvents lstWebPages As ListBox
    Friend WithEvents btnExit As Button
    Friend WithEvents btnOpenInMain As Button
    Friend WithEvents btnHome As Button
    Friend WithEvents ToolTip1 As ToolTip
    Friend WithEvents txtEdited As TextBox
    Friend WithEvents Label3 As Label
    Friend WithEvents txtCreated As TextBox
    Friend WithEvents Label2 As Label
    Friend WithEvents txtDescription As TextBox
    Friend WithEvents GroupBox1 As GroupBox
End Class
