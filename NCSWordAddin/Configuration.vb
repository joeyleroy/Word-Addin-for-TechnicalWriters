Imports System.ComponentModel
Imports System.Windows.Forms

Public Class Configuration

    Private Sub Configuration_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Dim myDocs As String = IO.Path.Combine(Environment.ExpandEnvironmentVariables("%userprofile%"), "Documents")
        Dim myPath As String = myDocs & "\Snipitz\"
        Dim myConfigPath As String = myPath & "Config\"
        Dim mySnipitzTXT As String = "snipitzFolder.txt"
        Dim myTemplatesTXT As String = "templatesFolder.txt"
        Dim mySconfig As String = myConfigPath & mySnipitzTXT
        Dim myTconfig As String = myConfigPath & myTemplatesTXT

        If Not System.IO.File.Exists(mySconfig) Then
            Dim myResult As Integer = MessageBox.Show(mySconfig & " is Missing! Should I create it for you now?", "Missing Snipitz Configuration File!", MessageBoxButtons.YesNoCancel)
            If myResult = DialogResult.Cancel Then
                MessageBox.Show("You CANNOT change the Snipitz folder without a config file!")
                Me.Close()
                Return
            ElseIf myResult = DialogResult.No Then
                MessageBox.Show("You CANNOT change the Snipitz folder without a config file!")
                Me.Close()
                Return
            ElseIf myResult = DialogResult.Yes Then
                System.IO.Directory.CreateDirectory(myConfigPath)
                System.IO.File.Create(mySconfig).Dispose()
                Dim snipitsWriter As New System.IO.StreamWriter(mySconfig)
                ' snipitsWriter.Write(myPath & "Snipits\")
                snipitsWriter.Write("Browse to select a directory, or set the default.")
                snipitsWriter.Close()
                MessageBox.Show(mySconfig & " created successfully!")
            End If
        End If

        If Not System.IO.File.Exists(myTconfig) Then
            Dim myResult As Integer = MessageBox.Show(myTconfig & " is Missing! Should I create it for you now?", "Missing Snipitz Configuration File!", MessageBoxButtons.YesNoCancel)
            If myResult = DialogResult.Cancel Then
                MessageBox.Show("You CANNOT change the Templates folder without a config file!")
                Me.Close()
                Return
            ElseIf myResult = DialogResult.No Then
                MessageBox.Show("You CANNOT change the Templates folder without a config file!")
                Me.Close()
                Return
            ElseIf myResult = DialogResult.Yes Then
                System.IO.Directory.CreateDirectory(myConfigPath)
                System.IO.File.Create(myTconfig).Dispose()
                Dim templatesWriter As New System.IO.StreamWriter(myTconfig)
                templatesWriter.Write("Browse to select a directory, or set the default.")
                templatesWriter.Close()
                MessageBox.Show(myTconfig & " created successfully!")
            End If
        End If

        Dim txtSnipits As String = My.Computer.FileSystem.ReadAllText(mySconfig)
        Dim txtTemplates As String = My.Computer.FileSystem.ReadAllText(myTconfig)
        Dim myIntroText As String
        If txtSnipits = "Browse to select a directory, or set the default." Or txtTemplates = "Browse to select a directory, or set the default." Then
            myIntroText = "Welcome to Configuration. Click the default buttons below to set the default folders, or browse to select your own."
        Else
            myIntroText = "Welcome back to Configuration. You can browse to change your target folders at any time. Press Save to lock your selection in."
        End If

        Me.IntroTextBox.Text = myIntroText
        Me.TextBoxSnipits.Text = txtSnipits
        Me.TextBoxTemplates.Text = txtTemplates

        If Not System.IO.Directory.Exists(myPath & "Snipits\") Then
            Me.cDefaultSF.Visible = True
        Else
            Me.cDefaultSF.Visible = False
        End If

        If Not System.IO.Directory.Exists(myPath & "Templates\") Then
            Me.cDefaultTF.Visible = True
        Else
            Me.cDefaultTF.Visible = False
        End If


        'If Not System.IO.Directory.Exists(myPath & "Snipits\") Then
        '    cDefaultSF.Visible = True
        '    Me.snipitsDefault.Visible = False
        'Else
        '    cDefaultSF.Visible = False
        '    Me.snipitsDefault.Visible = True
        '    Me.snipitsDefault.Text = myPath & "Snipits\"
        'End If

        'If Not System.IO.Directory.Exists(myPath & "Templates\") Then
        '    cDefaultTF.Visible = True
        '    Me.templatesDefault.Visible = False
        'Else
        '    cDefaultTF.Visible = False
        '    Me.templatesDefault.Visible = True
        '    Me.templatesDefault.Text = myPath & "Templates\"
        'End If
    End Sub

    Private Sub ButtonSnipits_Click(sender As Object, e As EventArgs) Handles ButtonSnipits.Click
        ' Button to browse and change path to Snipits
        If (FolderBrowserDialogSnipits.ShowDialog() = DialogResult.OK) Then
            TextBoxSnipitsNew.Text = FolderBrowserDialogSnipits.SelectedPath
        End If
    End Sub

    Private Sub ButtonTemplates_Click(sender As Object, e As EventArgs) Handles ButtonTemplates.Click
        ' Button to browse and change path to Templates
        If (FolderBrowserDialogTemplates.ShowDialog() = DialogResult.OK) Then
            TextBoxTemplatesNew.Text = FolderBrowserDialogTemplates.SelectedPath
        End If
    End Sub

    Private Sub ButtonSave_Click(sender As Object, e As EventArgs) Handles ButtonSave.Click
        ' Button to write paths to files

        Dim myDocs As String = IO.Path.Combine(Environment.ExpandEnvironmentVariables("%userprofile%"), "Documents")
        Dim myPath As String = myDocs & "\Snipitz\"
        Dim myConfigPath As String = myPath & "Config\"
        Dim mySnipitzTXT As String = "snipitzFolder.txt"
        Dim myTemplatesTXT As String = "templatesFolder.txt"
        Dim mySconfig As String = myConfigPath & mySnipitzTXT
        Dim myTconfig As String = myConfigPath & myTemplatesTXT

        If Not TextBoxSnipitsNew.Text = "" Then
            If Not TextBoxSnipits.Text = TextBoxSnipitsNew.Text Then
                If System.IO.File.Exists(mySconfig) = True Then
                    Dim snipitsWriter As New System.IO.StreamWriter(mySconfig)
                    snipitsWriter.Write(TextBoxSnipitsNew.Text)
                    snipitsWriter.Close()
                Else
                    MessageBox.Show(mySconfig & " Doesnt exist! Re-run Configuration to create it.")
                    Return
                End If
            End If
        End If

        If Not TextBoxTemplatesNew.Text = "" Then
            If Not TextBoxTemplates.Text = TextBoxTemplatesNew.Text Then
                If System.IO.File.Exists(myTconfig) = True Then
                    Dim templatesWriter As New System.IO.StreamWriter(myTconfig)
                    templatesWriter.Write(TextBoxTemplatesNew.Text)
                    templatesWriter.Close()
                Else
                    MessageBox.Show(myTconfig & " Doesnt exist! Re-run Configuration to create it.")
                    Return
                End If
            End If
        End If
        Me.Close()
    End Sub

    Private Sub cDefaultSF_Click(sender As Object, e As EventArgs) Handles cDefaultSF.Click
        Dim myDocsS As String = IO.Path.Combine(Environment.ExpandEnvironmentVariables("%userprofile%"), "Documents")
        Dim myPathS As String = myDocsS & "\Snipitz\"
        Dim mySDefault As String = myPathS & "Snipits\"
        Dim mySconfig As String = myPathS & "Config\snipitzFolder.txt"
        System.IO.Directory.CreateDirectory(mySDefault)
        Dim snipWriter As New System.IO.StreamWriter(mySconfig)
        snipWriter.Write(mySDefault)
        snipWriter.Close()
        Me.cDefaultSF.Visible = False
        Me.TextBoxSnipits.Text = mySDefault
    End Sub

    Private Sub cDefaultTF_Click(sender As Object, e As EventArgs) Handles cDefaultTF.Click
        Dim myDocsT As String = IO.Path.Combine(Environment.ExpandEnvironmentVariables("%userprofile%"), "Documents")
        Dim myPathT As String = myDocsT & "\Snipitz\"
        Dim myTDefault As String = myPathT & "Templates\"
        Dim myTconfig As String = myPathT & "Config\templatesFolder.txt"
        System.IO.Directory.CreateDirectory(myTDefault)
        Dim templatesWriter As New System.IO.StreamWriter(myTconfig)
        templatesWriter.Write(myTDefault)
        templatesWriter.Close()
        Me.cDefaultTF.Visible = False
        Me.TextBoxTemplates.Text = myTDefault
    End Sub
End Class