<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class FormTemplates
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FormTemplates))
        Me.RichTextBox1 = New System.Windows.Forms.RichTextBox()
        Me.TreeView1 = New System.Windows.Forms.TreeView()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.PanelInfo = New System.Windows.Forms.Panel()
        Me.TextBoxInfo = New System.Windows.Forms.TextBox()
        Me.LabelInformation = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.PictureBox2 = New System.Windows.Forms.PictureBox()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.ImageList1 = New System.Windows.Forms.ImageList(Me.components)
        Me.Panel1.SuspendLayout()
        Me.Panel2.SuspendLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.PanelInfo.SuspendLayout()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'RichTextBox1
        '
        Me.RichTextBox1.BackColor = System.Drawing.SystemColors.Window
        Me.RichTextBox1.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.RichTextBox1.Location = New System.Drawing.Point(26, 22)
        Me.RichTextBox1.Name = "RichTextBox1"
        Me.RichTextBox1.Size = New System.Drawing.Size(400, 300)
        Me.RichTextBox1.TabIndex = 0
        Me.RichTextBox1.Text = ""
        '
        'TreeView1
        '
        Me.TreeView1.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.TreeView1.Font = New System.Drawing.Font("Comic Sans MS", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TreeView1.ForeColor = System.Drawing.SystemColors.GrayText
        Me.TreeView1.Location = New System.Drawing.Point(26, 22)
        Me.TreeView1.Name = "TreeView1"
        Me.TreeView1.Size = New System.Drawing.Size(400, 300)
        Me.TreeView1.TabIndex = 2
        '
        'Panel1
        '
        Me.Panel1.BackgroundImage = Global.NCSWordAddin.My.Resources.Resources.TemplatesFrame1
        Me.Panel1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.Panel1.Controls.Add(Me.RichTextBox1)
        Me.Panel1.Location = New System.Drawing.Point(0, 130)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(450, 345)
        Me.Panel1.TabIndex = 3
        '
        'Panel2
        '
        Me.Panel2.BackgroundImage = Global.NCSWordAddin.My.Resources.Resources.TemplatesFrame1
        Me.Panel2.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.Panel2.Controls.Add(Me.TreeView1)
        Me.Panel2.Location = New System.Drawing.Point(459, 130)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(450, 345)
        Me.Panel2.TabIndex = 4
        '
        'PictureBox1
        '
        Me.PictureBox1.Image = Global.NCSWordAddin.My.Resources.Resources.sys_printer1
        Me.PictureBox1.Location = New System.Drawing.Point(500, 481)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(61, 61)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox1.TabIndex = 5
        Me.PictureBox1.TabStop = False
        Me.ToolTip1.SetToolTip(Me.PictureBox1, "Attach: Attaches the selected template to the current document.")
        '
        'PanelInfo
        '
        Me.PanelInfo.BackColor = System.Drawing.Color.Transparent
        Me.PanelInfo.BackgroundImage = Global.NCSWordAddin.My.Resources.Resources.TemplatesFrame3
        Me.PanelInfo.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.PanelInfo.Controls.Add(Me.TextBoxInfo)
        Me.PanelInfo.Location = New System.Drawing.Point(0, 13)
        Me.PanelInfo.Name = "PanelInfo"
        Me.PanelInfo.Size = New System.Drawing.Size(910, 103)
        Me.PanelInfo.TabIndex = 19
        '
        'TextBoxInfo
        '
        Me.TextBoxInfo.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.TextBoxInfo.Font = New System.Drawing.Font("Calibri", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBoxInfo.ForeColor = System.Drawing.SystemColors.ControlDark
        Me.TextBoxInfo.Location = New System.Drawing.Point(35, 26)
        Me.TextBoxInfo.Multiline = True
        Me.TextBoxInfo.Name = "TextBoxInfo"
        Me.TextBoxInfo.Size = New System.Drawing.Size(852, 56)
        Me.TextBoxInfo.TabIndex = 2
        Me.TextBoxInfo.Text = "This is What the Text will look like in an Error and This also tests word wrappin" &
    "g." & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "This is a new line."
        '
        'LabelInformation
        '
        Me.LabelInformation.AutoSize = True
        Me.LabelInformation.Font = New System.Drawing.Font("Calibri", 10.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelInformation.ForeColor = System.Drawing.SystemColors.ControlDarkDark
        Me.LabelInformation.Location = New System.Drawing.Point(48, -1)
        Me.LabelInformation.Name = "LabelInformation"
        Me.LabelInformation.Size = New System.Drawing.Size(103, 23)
        Me.LabelInformation.TabIndex = 3
        Me.LabelInformation.Text = "Information"
        Me.ToolTip1.SetToolTip(Me.LabelInformation, "Provides instructions and feedback.")
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.BackColor = System.Drawing.SystemColors.Window
        Me.Label1.Font = New System.Drawing.Font("Calibri", 10.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.SystemColors.GrayText
        Me.Label1.Location = New System.Drawing.Point(55, 117)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(152, 23)
        Me.Label1.TabIndex = 20
        Me.Label1.Text = "Template Preview"
        Me.ToolTip1.SetToolTip(Me.Label1, "Displays a preview of the selected template.")
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.BackColor = System.Drawing.SystemColors.Window
        Me.Label2.Font = New System.Drawing.Font("Calibri", 10.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.SystemColors.GrayText
        Me.Label2.Location = New System.Drawing.Point(515, 117)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(160, 23)
        Me.Label2.TabIndex = 21
        Me.Label2.Text = "Template Selection"
        Me.ToolTip1.SetToolTip(Me.Label2, "Select the desired template from the tree view below.")
        '
        'PictureBox2
        '
        Me.PictureBox2.Image = Global.NCSWordAddin.My.Resources.Resources.Refresh1
        Me.PictureBox2.Location = New System.Drawing.Point(40, 481)
        Me.PictureBox2.Name = "PictureBox2"
        Me.PictureBox2.Size = New System.Drawing.Size(61, 61)
        Me.PictureBox2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox2.TabIndex = 22
        Me.PictureBox2.TabStop = False
        Me.ToolTip1.SetToolTip(Me.PictureBox2, "Refresh: Refresh the current document template.")
        '
        'ImageList1
        '
        Me.ImageList1.ImageStream = CType(resources.GetObject("ImageList1.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.ImageList1.TransparentColor = System.Drawing.Color.Transparent
        Me.ImageList1.Images.SetKeyName(0, "box-closed-2.png")
        Me.ImageList1.Images.SetKeyName(1, "box-2.png")
        Me.ImageList1.Images.SetKeyName(2, "File2.png")
        '
        'FormTemplates
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Window
        Me.ClientSize = New System.Drawing.Size(909, 544)
        Me.Controls.Add(Me.PictureBox2)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.LabelInformation)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.PanelInfo)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.Panel1)
        Me.DataBindings.Add(New System.Windows.Forms.Binding("Location", Global.NCSWordAddin.MySettings.Default, "TempLoc", True, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged))
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Location = Global.NCSWordAddin.MySettings.Default.TempLoc
        Me.Name = "FormTemplates"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Templates"
        Me.Panel1.ResumeLayout(False)
        Me.Panel2.ResumeLayout(False)
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.PanelInfo.ResumeLayout(False)
        Me.PanelInfo.PerformLayout()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents RichTextBox1 As System.Windows.Forms.RichTextBox
    Friend WithEvents TreeView1 As System.Windows.Forms.TreeView
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents PanelInfo As System.Windows.Forms.Panel
    Friend WithEvents LabelInformation As System.Windows.Forms.Label
    Friend WithEvents TextBoxInfo As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents PictureBox2 As System.Windows.Forms.PictureBox
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents ImageList1 As System.Windows.Forms.ImageList
End Class
