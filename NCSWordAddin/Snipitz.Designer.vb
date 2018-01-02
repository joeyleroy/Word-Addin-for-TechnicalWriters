<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Snipitz
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
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
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Snipitz))
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.TreeView1 = New System.Windows.Forms.TreeView()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.PWLabel = New System.Windows.Forms.Label()
        Me.PWPanel = New System.Windows.Forms.Panel()
        Me.PWTextBox = New System.Windows.Forms.TextBox()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.Button3 = New System.Windows.Forms.PictureBox()
        Me.Button2 = New System.Windows.Forms.PictureBox()
        Me.Button1 = New System.Windows.Forms.PictureBox()
        Me.Button4 = New System.Windows.Forms.PictureBox()
        Me.PanelInfo = New System.Windows.Forms.Panel()
        Me.TextBoxInfo = New System.Windows.Forms.TextBox()
        Me.LabelInformation = New System.Windows.Forms.Label()
        Me.ImageList1 = New System.Windows.Forms.ImageList(Me.components)
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.LabelInsert = New System.Windows.Forms.Label()
        Me.LabelEdit = New System.Windows.Forms.Label()
        Me.LabelCreate = New System.Windows.Forms.Label()
        Me.LabelDelete = New System.Windows.Forms.Label()
        Me.Panel2.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.PWPanel.SuspendLayout()
        CType(Me.Button3, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Button2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Button1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Button4, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.PanelInfo.SuspendLayout()
        Me.SuspendLayout()
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.BackColor = System.Drawing.SystemColors.Window
        Me.Label2.Font = New System.Drawing.Font("Comic Sans MS", 10.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.SystemColors.GrayText
        Me.Label2.Location = New System.Drawing.Point(50, 123)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(96, 27)
        Me.Label2.TabIndex = 23
        Me.Label2.Text = "Selection"
        '
        'Panel2
        '
        Me.Panel2.BackgroundImage = Global.NCSWordAddin.My.Resources.Resources.Button1CheckFrame
        Me.Panel2.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.Panel2.Controls.Add(Me.TreeView1)
        Me.Panel2.Location = New System.Drawing.Point(0, 135)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(380, 485)
        Me.Panel2.TabIndex = 22
        '
        'TreeView1
        '
        Me.TreeView1.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.TreeView1.Font = New System.Drawing.Font("Calibri", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TreeView1.ForeColor = System.Drawing.SystemColors.GrayText
        Me.TreeView1.Location = New System.Drawing.Point(25, 20)
        Me.TreeView1.Name = "TreeView1"
        Me.TreeView1.Size = New System.Drawing.Size(330, 438)
        Me.TreeView1.TabIndex = 2
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.BackColor = System.Drawing.SystemColors.Window
        Me.Label1.Font = New System.Drawing.Font("Comic Sans MS", 10.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.SystemColors.GrayText
        Me.Label1.Location = New System.Drawing.Point(50, 628)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(55, 27)
        Me.Label1.TabIndex = 25
        Me.Label1.Text = "Title"
        '
        'Panel1
        '
        Me.Panel1.BackgroundImage = Global.NCSWordAddin.My.Resources.Resources.TemplatesFrame3
        Me.Panel1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.Panel1.Controls.Add(Me.PWLabel)
        Me.Panel1.Controls.Add(Me.PWPanel)
        Me.Panel1.Controls.Add(Me.TextBox1)
        Me.Panel1.Location = New System.Drawing.Point(0, 640)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(380, 60)
        Me.Panel1.TabIndex = 24
        '
        'PWLabel
        '
        Me.PWLabel.AutoSize = True
        Me.PWLabel.BackColor = System.Drawing.SystemColors.Window
        Me.PWLabel.Font = New System.Drawing.Font("Calibri", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.PWLabel.ForeColor = System.Drawing.SystemColors.GrayText
        Me.PWLabel.Location = New System.Drawing.Point(152, 7)
        Me.PWLabel.Name = "PWLabel"
        Me.PWLabel.Size = New System.Drawing.Size(144, 17)
        Me.PWLabel.TabIndex = 32
        Me.PWLabel.Text = "Confirm with Password"
        '
        'PWPanel
        '
        Me.PWPanel.BackgroundImage = Global.NCSWordAddin.My.Resources.Resources.TemplatesFrame3
        Me.PWPanel.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.PWPanel.Controls.Add(Me.PWTextBox)
        Me.PWPanel.Location = New System.Drawing.Point(129, 14)
        Me.PWPanel.Name = "PWPanel"
        Me.PWPanel.Size = New System.Drawing.Size(200, 60)
        Me.PWPanel.TabIndex = 32
        '
        'PWTextBox
        '
        Me.PWTextBox.BackColor = System.Drawing.SystemColors.Control
        Me.PWTextBox.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.PWTextBox.Font = New System.Drawing.Font("Comic Sans MS", 9.0!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.PWTextBox.Location = New System.Drawing.Point(25, 20)
        Me.PWTextBox.Name = "PWTextBox"
        Me.PWTextBox.Size = New System.Drawing.Size(150, 21)
        Me.PWTextBox.TabIndex = 0
        '
        'TextBox1
        '
        Me.TextBox1.BackColor = System.Drawing.SystemColors.Control
        Me.TextBox1.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.TextBox1.Font = New System.Drawing.Font("Calibri", 9.0!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox1.Location = New System.Drawing.Point(25, 20)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(330, 19)
        Me.TextBox1.TabIndex = 0
        '
        'Button3
        '
        Me.Button3.Image = Global.NCSWordAddin.My.Resources.Resources.Create1
        Me.Button3.Location = New System.Drawing.Point(187, 710)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(60, 60)
        Me.Button3.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.Button3.TabIndex = 28
        Me.Button3.TabStop = False
        Me.ToolTip1.SetToolTip(Me.Button3, "Create")
        '
        'Button2
        '
        Me.Button2.Image = Global.NCSWordAddin.My.Resources.Resources.EditPreview1
        Me.Button2.Location = New System.Drawing.Point(107, 710)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(60, 60)
        Me.Button2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.Button2.TabIndex = 27
        Me.Button2.TabStop = False
        Me.ToolTip1.SetToolTip(Me.Button2, "Edit / Preview")
        '
        'Button1
        '
        Me.Button1.Image = Global.NCSWordAddin.My.Resources.Resources.Insert1
        Me.Button1.Location = New System.Drawing.Point(28, 710)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(60, 60)
        Me.Button1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.Button1.TabIndex = 26
        Me.Button1.TabStop = False
        Me.ToolTip1.SetToolTip(Me.Button1, "Insert")
        '
        'Button4
        '
        Me.Button4.Image = Global.NCSWordAddin.My.Resources.Resources.Trash1
        Me.Button4.Location = New System.Drawing.Point(267, 710)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(60, 60)
        Me.Button4.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.Button4.TabIndex = 29
        Me.Button4.TabStop = False
        Me.ToolTip1.SetToolTip(Me.Button4, "Delete")
        '
        'PanelInfo
        '
        Me.PanelInfo.BackColor = System.Drawing.Color.Transparent
        Me.PanelInfo.BackgroundImage = Global.NCSWordAddin.My.Resources.Resources.MessageFrame
        Me.PanelInfo.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.PanelInfo.Controls.Add(Me.TextBoxInfo)
        Me.PanelInfo.Location = New System.Drawing.Point(0, 15)
        Me.PanelInfo.Name = "PanelInfo"
        Me.PanelInfo.Size = New System.Drawing.Size(380, 100)
        Me.PanelInfo.TabIndex = 30
        '
        'TextBoxInfo
        '
        Me.TextBoxInfo.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.TextBoxInfo.Font = New System.Drawing.Font("Calibri", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBoxInfo.ForeColor = System.Drawing.SystemColors.ControlDark
        Me.TextBoxInfo.Location = New System.Drawing.Point(25, 20)
        Me.TextBoxInfo.Multiline = True
        Me.TextBoxInfo.Name = "TextBoxInfo"
        Me.TextBoxInfo.Size = New System.Drawing.Size(330, 72)
        Me.TextBoxInfo.TabIndex = 2
        Me.TextBoxInfo.Text = "This is What the Text will look like in an Error and This also tests word wrappin" &
    "g." & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "This is a new line."
        '
        'LabelInformation
        '
        Me.LabelInformation.AutoSize = True
        Me.LabelInformation.Font = New System.Drawing.Font("Comic Sans MS", 10.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelInformation.ForeColor = System.Drawing.SystemColors.ControlDarkDark
        Me.LabelInformation.Location = New System.Drawing.Point(50, 3)
        Me.LabelInformation.Name = "LabelInformation"
        Me.LabelInformation.Size = New System.Drawing.Size(121, 27)
        Me.LabelInformation.TabIndex = 31
        Me.LabelInformation.Text = "Information"
        '
        'ImageList1
        '
        Me.ImageList1.ImageStream = CType(resources.GetObject("ImageList1.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.ImageList1.TransparentColor = System.Drawing.Color.Transparent
        Me.ImageList1.Images.SetKeyName(0, "box-closed-2.png")
        Me.ImageList1.Images.SetKeyName(1, "box12.png")
        Me.ImageList1.Images.SetKeyName(2, "File2.png")
        Me.ImageList1.Images.SetKeyName(3, "add4.png")
        Me.ImageList1.Images.SetKeyName(4, "add3.png")
        '
        'LabelInsert
        '
        Me.LabelInsert.AutoSize = True
        Me.LabelInsert.Location = New System.Drawing.Point(34, 773)
        Me.LabelInsert.Name = "LabelInsert"
        Me.LabelInsert.Size = New System.Drawing.Size(43, 17)
        Me.LabelInsert.TabIndex = 32
        Me.LabelInsert.Text = "Insert"
        '
        'LabelEdit
        '
        Me.LabelEdit.AutoSize = True
        Me.LabelEdit.Location = New System.Drawing.Point(94, 773)
        Me.LabelEdit.Name = "LabelEdit"
        Me.LabelEdit.Size = New System.Drawing.Size(85, 17)
        Me.LabelEdit.TabIndex = 33
        Me.LabelEdit.Text = "Edit/Preview"
        '
        'LabelCreate
        '
        Me.LabelCreate.AutoSize = True
        Me.LabelCreate.Location = New System.Drawing.Point(191, 773)
        Me.LabelCreate.Name = "LabelCreate"
        Me.LabelCreate.Size = New System.Drawing.Size(50, 17)
        Me.LabelCreate.TabIndex = 34
        Me.LabelCreate.Text = "Create"
        '
        'LabelDelete
        '
        Me.LabelDelete.AutoSize = True
        Me.LabelDelete.Location = New System.Drawing.Point(272, 773)
        Me.LabelDelete.Name = "LabelDelete"
        Me.LabelDelete.Size = New System.Drawing.Size(49, 17)
        Me.LabelDelete.TabIndex = 35
        Me.LabelDelete.Text = "Delete"
        '
        'Snipitz
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Window
        Me.ClientSize = New System.Drawing.Size(382, 791)
        Me.Controls.Add(Me.LabelDelete)
        Me.Controls.Add(Me.LabelCreate)
        Me.Controls.Add(Me.LabelEdit)
        Me.Controls.Add(Me.LabelInsert)
        Me.Controls.Add(Me.LabelInformation)
        Me.Controls.Add(Me.PanelInfo)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.Button4)
        Me.DataBindings.Add(New System.Windows.Forms.Binding("Location", Global.NCSWordAddin.MySettings.Default, "SnipitsLoc", True, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged))
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Location = Global.NCSWordAddin.MySettings.Default.SnipitsLoc
        Me.Name = "Snipitz"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Snipitz"
        Me.Panel2.ResumeLayout(False)
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.PWPanel.ResumeLayout(False)
        Me.PWPanel.PerformLayout()
        CType(Me.Button3, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Button2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Button1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Button4, System.ComponentModel.ISupportInitialize).EndInit()
        Me.PanelInfo.ResumeLayout(False)
        Me.PanelInfo.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents TreeView1 As System.Windows.Forms.TreeView
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents Button3 As System.Windows.Forms.PictureBox
    Friend WithEvents Button2 As System.Windows.Forms.PictureBox
    Friend WithEvents Button1 As System.Windows.Forms.PictureBox
    Friend WithEvents Button4 As System.Windows.Forms.PictureBox
    Friend WithEvents PanelInfo As System.Windows.Forms.Panel
    Friend WithEvents TextBoxInfo As System.Windows.Forms.TextBox
    Friend WithEvents LabelInformation As System.Windows.Forms.Label
    Friend WithEvents ImageList1 As System.Windows.Forms.ImageList
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents PWLabel As System.Windows.Forms.Label
    Friend WithEvents PWPanel As System.Windows.Forms.Panel
    Friend WithEvents PWTextBox As System.Windows.Forms.TextBox
    Friend WithEvents LabelInsert As System.Windows.Forms.Label
    Friend WithEvents LabelEdit As System.Windows.Forms.Label
    Friend WithEvents LabelCreate As System.Windows.Forms.Label
    Friend WithEvents LabelDelete As System.Windows.Forms.Label
End Class
