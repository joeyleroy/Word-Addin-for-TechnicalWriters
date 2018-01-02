<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Configuration
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
        Me.LabelSnipits = New System.Windows.Forms.Label()
        Me.LabelTemplates = New System.Windows.Forms.Label()
        Me.ButtonSave = New System.Windows.Forms.Button()
        Me.LabelSCurrent = New System.Windows.Forms.Label()
        Me.LabelTCurrent = New System.Windows.Forms.Label()
        Me.TextBoxSnipits = New System.Windows.Forms.TextBox()
        Me.TextBoxTemplates = New System.Windows.Forms.TextBox()
        Me.ButtonSnipits = New System.Windows.Forms.Button()
        Me.ButtonTemplates = New System.Windows.Forms.Button()
        Me.FolderBrowserDialogSnipits = New System.Windows.Forms.FolderBrowserDialog()
        Me.FolderBrowserDialogTemplates = New System.Windows.Forms.FolderBrowserDialog()
        Me.TextBoxSnipitsNew = New System.Windows.Forms.TextBox()
        Me.LabelSNew = New System.Windows.Forms.Label()
        Me.TextBoxTemplatesNew = New System.Windows.Forms.TextBox()
        Me.LabelTNew = New System.Windows.Forms.Label()
        Me.IntroTextBox = New System.Windows.Forms.TextBox()
        Me.cDefaultSF = New System.Windows.Forms.Button()
        Me.cDefaultTF = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'LabelSnipits
        '
        Me.LabelSnipits.AutoSize = True
        Me.LabelSnipits.Location = New System.Drawing.Point(16, 128)
        Me.LabelSnipits.Name = "LabelSnipits"
        Me.LabelSnipits.Size = New System.Drawing.Size(319, 17)
        Me.LabelSnipits.TabIndex = 1
        Me.LabelSnipits.Text = "Select Snipits Directory (Where you store Snipits)"
        '
        'LabelTemplates
        '
        Me.LabelTemplates.AutoSize = True
        Me.LabelTemplates.Location = New System.Drawing.Point(388, 128)
        Me.LabelTemplates.Name = "LabelTemplates"
        Me.LabelTemplates.Size = New System.Drawing.Size(367, 17)
        Me.LabelTemplates.TabIndex = 2
        Me.LabelTemplates.Text = "Select Templates Directory (Where you store Templates)"
        '
        'ButtonSave
        '
        Me.ButtonSave.Location = New System.Drawing.Point(301, 349)
        Me.ButtonSave.Name = "ButtonSave"
        Me.ButtonSave.Size = New System.Drawing.Size(155, 35)
        Me.ButtonSave.TabIndex = 3
        Me.ButtonSave.Text = "Save Configuration"
        Me.ButtonSave.UseVisualStyleBackColor = True
        '
        'LabelSCurrent
        '
        Me.LabelSCurrent.AutoSize = True
        Me.LabelSCurrent.Location = New System.Drawing.Point(19, 193)
        Me.LabelSCurrent.Name = "LabelSCurrent"
        Me.LabelSCurrent.Size = New System.Drawing.Size(59, 17)
        Me.LabelSCurrent.TabIndex = 6
        Me.LabelSCurrent.Text = "Current:"
        '
        'LabelTCurrent
        '
        Me.LabelTCurrent.AutoSize = True
        Me.LabelTCurrent.Location = New System.Drawing.Point(397, 193)
        Me.LabelTCurrent.Name = "LabelTCurrent"
        Me.LabelTCurrent.Size = New System.Drawing.Size(59, 17)
        Me.LabelTCurrent.TabIndex = 7
        Me.LabelTCurrent.Text = "Current:"
        '
        'TextBoxSnipits
        '
        Me.TextBoxSnipits.Location = New System.Drawing.Point(84, 190)
        Me.TextBoxSnipits.Multiline = True
        Me.TextBoxSnipits.Name = "TextBoxSnipits"
        Me.TextBoxSnipits.ReadOnly = True
        Me.TextBoxSnipits.Size = New System.Drawing.Size(287, 53)
        Me.TextBoxSnipits.TabIndex = 9
        '
        'TextBoxTemplates
        '
        Me.TextBoxTemplates.Location = New System.Drawing.Point(464, 190)
        Me.TextBoxTemplates.Multiline = True
        Me.TextBoxTemplates.Name = "TextBoxTemplates"
        Me.TextBoxTemplates.ReadOnly = True
        Me.TextBoxTemplates.Size = New System.Drawing.Size(287, 53)
        Me.TextBoxTemplates.TabIndex = 10
        '
        'ButtonSnipits
        '
        Me.ButtonSnipits.Location = New System.Drawing.Point(16, 269)
        Me.ButtonSnipits.Name = "ButtonSnipits"
        Me.ButtonSnipits.Size = New System.Drawing.Size(62, 33)
        Me.ButtonSnipits.TabIndex = 12
        Me.ButtonSnipits.Text = "Browse"
        Me.ButtonSnipits.UseVisualStyleBackColor = True
        '
        'ButtonTemplates
        '
        Me.ButtonTemplates.Location = New System.Drawing.Point(394, 269)
        Me.ButtonTemplates.Name = "ButtonTemplates"
        Me.ButtonTemplates.Size = New System.Drawing.Size(62, 33)
        Me.ButtonTemplates.TabIndex = 13
        Me.ButtonTemplates.Text = "Browse"
        Me.ButtonTemplates.UseVisualStyleBackColor = True
        '
        'TextBoxSnipitsNew
        '
        Me.TextBoxSnipitsNew.Location = New System.Drawing.Point(84, 249)
        Me.TextBoxSnipitsNew.Multiline = True
        Me.TextBoxSnipitsNew.Name = "TextBoxSnipitsNew"
        Me.TextBoxSnipitsNew.ReadOnly = True
        Me.TextBoxSnipitsNew.Size = New System.Drawing.Size(287, 53)
        Me.TextBoxSnipitsNew.TabIndex = 14
        '
        'LabelSNew
        '
        Me.LabelSNew.AutoSize = True
        Me.LabelSNew.Location = New System.Drawing.Point(39, 250)
        Me.LabelSNew.Name = "LabelSNew"
        Me.LabelSNew.Size = New System.Drawing.Size(39, 17)
        Me.LabelSNew.TabIndex = 15
        Me.LabelSNew.Text = "New:"
        '
        'TextBoxTemplatesNew
        '
        Me.TextBoxTemplatesNew.Location = New System.Drawing.Point(465, 249)
        Me.TextBoxTemplatesNew.Multiline = True
        Me.TextBoxTemplatesNew.Name = "TextBoxTemplatesNew"
        Me.TextBoxTemplatesNew.ReadOnly = True
        Me.TextBoxTemplatesNew.Size = New System.Drawing.Size(287, 53)
        Me.TextBoxTemplatesNew.TabIndex = 16
        '
        'LabelTNew
        '
        Me.LabelTNew.AutoSize = True
        Me.LabelTNew.Location = New System.Drawing.Point(417, 250)
        Me.LabelTNew.Name = "LabelTNew"
        Me.LabelTNew.Size = New System.Drawing.Size(39, 17)
        Me.LabelTNew.TabIndex = 17
        Me.LabelTNew.Text = "New:"
        '
        'IntroTextBox
        '
        Me.IntroTextBox.Location = New System.Drawing.Point(12, 12)
        Me.IntroTextBox.Multiline = True
        Me.IntroTextBox.Name = "IntroTextBox"
        Me.IntroTextBox.ReadOnly = True
        Me.IntroTextBox.Size = New System.Drawing.Size(739, 107)
        Me.IntroTextBox.TabIndex = 19
        '
        'cDefaultSF
        '
        Me.cDefaultSF.Location = New System.Drawing.Point(19, 151)
        Me.cDefaultSF.Name = "cDefaultSF"
        Me.cDefaultSF.Size = New System.Drawing.Size(352, 28)
        Me.cDefaultSF.TabIndex = 20
        Me.cDefaultSF.Text = "Create Default Snipits Folder (Recommended)"
        Me.cDefaultSF.UseVisualStyleBackColor = True
        '
        'cDefaultTF
        '
        Me.cDefaultTF.Location = New System.Drawing.Point(391, 151)
        Me.cDefaultTF.Name = "cDefaultTF"
        Me.cDefaultTF.Size = New System.Drawing.Size(360, 28)
        Me.cDefaultTF.TabIndex = 21
        Me.cDefaultTF.Text = "Create Default Templates Folder (Recommended)"
        Me.cDefaultTF.UseVisualStyleBackColor = True
        '
        'Configuration
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(760, 398)
        Me.Controls.Add(Me.cDefaultTF)
        Me.Controls.Add(Me.cDefaultSF)
        Me.Controls.Add(Me.IntroTextBox)
        Me.Controls.Add(Me.LabelTNew)
        Me.Controls.Add(Me.TextBoxTemplatesNew)
        Me.Controls.Add(Me.LabelSNew)
        Me.Controls.Add(Me.TextBoxSnipitsNew)
        Me.Controls.Add(Me.ButtonTemplates)
        Me.Controls.Add(Me.ButtonSnipits)
        Me.Controls.Add(Me.TextBoxTemplates)
        Me.Controls.Add(Me.TextBoxSnipits)
        Me.Controls.Add(Me.LabelTCurrent)
        Me.Controls.Add(Me.LabelSCurrent)
        Me.Controls.Add(Me.ButtonSave)
        Me.Controls.Add(Me.LabelTemplates)
        Me.Controls.Add(Me.LabelSnipits)
        Me.Name = "Configuration"
        Me.Text = "Configuration"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents LabelSnipits As System.Windows.Forms.Label
    Friend WithEvents LabelTemplates As System.Windows.Forms.Label
    Friend WithEvents ButtonSave As System.Windows.Forms.Button
    Friend WithEvents LabelSCurrent As System.Windows.Forms.Label
    Friend WithEvents LabelTCurrent As System.Windows.Forms.Label
    Friend WithEvents TextBoxSnipits As System.Windows.Forms.TextBox
    Friend WithEvents TextBoxTemplates As System.Windows.Forms.TextBox
    Friend WithEvents ButtonSnipits As System.Windows.Forms.Button
    Friend WithEvents ButtonTemplates As System.Windows.Forms.Button
    Friend WithEvents FolderBrowserDialogSnipits As System.Windows.Forms.FolderBrowserDialog
    Friend WithEvents FolderBrowserDialogTemplates As System.Windows.Forms.FolderBrowserDialog
    Friend WithEvents TextBoxSnipitsNew As System.Windows.Forms.TextBox
    Friend WithEvents LabelSNew As System.Windows.Forms.Label
    Friend WithEvents TextBoxTemplatesNew As System.Windows.Forms.TextBox
    Friend WithEvents LabelTNew As System.Windows.Forms.Label
    Friend WithEvents IntroTextBox As System.Windows.Forms.TextBox
    Friend WithEvents cDefaultSF As System.Windows.Forms.Button
    Friend WithEvents cDefaultTF As System.Windows.Forms.Button
End Class
