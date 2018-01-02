
Imports System.ComponentModel
Imports System.Windows.Forms

Public Class FormTemplates
    Private Sub Templates_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim myWord As Word.Application = Globals.ThisAddIn.Application
        Dim myDoc As Word.Document = myWord.ActiveDocument
        Dim myNewTemp As Word.Document = Nothing
        RichTextBox1.SelectAll()
        PopulateRichTextBox1(myNewTemp) 'NewTemp
        Dim myUserName As String = Environ("USERPROFILE")

        Dim myDocs As String = IO.Path.Combine(Environment.ExpandEnvironmentVariables("%userprofile%"), "Documents")
        Dim myPathconfig As String = myDocs & "\Snipitz\"
        Dim myConfigPath As String = myPathconfig & "Config\"
        Dim myTemplatesTXT As String = "templatesFolder.txt"
        Dim myTconfig As String = myConfigPath & myTemplatesTXT

        If Not System.IO.File.Exists(myTconfig) Then
            MessageBox.Show("You MUST run 'Configuration' to configure a Templates directory!")
            Me.Close()
            Return
        End If

        Dim templPath As String = My.Computer.FileSystem.ReadAllText(myTconfig)
        Dim myDirectory As String

        If System.IO.Directory.Exists(templPath) Then
            'user defined directory exists
            myDirectory = templPath
        Else
            MessageBox.Show("Run Configuration to set a Template directory.")
            Me.Close()
            Return
        End If

        Dim fi As New IO.DirectoryInfo(myDirectory)

        Me.TopMost = True

        If Not My.Settings.TempSize.IsEmpty Then
            Me.Size = My.Settings.TempSize
        End If

        If Not My.Settings.TempLoc.IsEmpty Then
            Me.Location = My.Settings.TempLoc
        Else
            With Me
                .Top = (myWord.Top)
                .Left = (myWord.Left)
                .StartPosition = 3
            End With
        End If

        With Me
            .PanelInfo.Width = .Width - 17
            .TextBoxInfo.Left = .PanelInfo.Left + 25
            .TextBoxInfo.Width = .PanelInfo.Width - 50
            .Panel1.Width = (.Width / 2) - 13
            .Panel1.Height = (.Height) - 243
            .RichTextBox1.Left = 26
            .RichTextBox1.Width = .Panel1.Width - 50
            .RichTextBox1.Height = .Panel1.Height - 50
            .RichTextBox1.Top = 22
            .Panel2.Left = (.Panel1.Width) + 10
            .Panel2.Width = (.Width / 2) - 13
            .Panel2.Height = (.Height) - 243
            .Label2.Left = .Panel2.Left + 56
            .TreeView1.Left = 26
            .TreeView1.Width = .Panel2.Width - 50
            .TreeView1.Height = .Panel2.Height - 50
            .TreeView1.Top = 22
            .PictureBox1.Top = .Height - 108
            .PictureBox1.Left = Panel1.Left + ((Panel1.Width / 2) - (PictureBox2.Width / 2))
            .PictureBox2.Top = .Height - 108
            .PictureBox2.Left = Panel2.Left + ((Panel2.Width / 2) - (PictureBox1.Width / 2))
        End With

        'Array to store paths
        Dim path() As String = {}
        Me.TreeView1.ImageList = Me.ImageList1
        'Loop through subfolders
        For Each subfolder As IO.DirectoryInfo In fi.GetDirectories()
            'Add this folders name
            Array.Resize(path, path.Length + 1)
            path(path.Length - 1) = subfolder.FullName
        Next

        'Get a list of drives
        Dim rootDir As String = String.Empty
        'Now loop thru each drive and populate the treeview
        For i As Integer = 0 To path.Count - 1
            rootDir = path(i)
            Dim split As String() = rootDir.Split("\")
            Dim parentFolder As String = split(split.Length - 1)
            'Add this drive as a root node
            Dim root As System.Windows.Forms.TreeNode = TreeView1.Nodes.Add(parentFolder)
            root.Tag = rootDir
            'Populate this root node

            PopulateTreeView(rootDir, TreeView1.Nodes(i))
        Next
        Me.TextBoxInfo.Text = "Select a template from the 'Template Selection' window. " &
                                "A limited preview (not exactly accurate) of the selected " &
                                "template will display in the 'Template Preview' window."
    End Sub

    Private Sub FormTemplates_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        My.Settings.TempSize = Me.Size
        My.Settings.TempLoc = Me.Location
    End Sub

    Private Sub FormTemplates_ResizeEnd(sender As Object, e As EventArgs) Handles Me.ResizeEnd
        With Me
            .PanelInfo.Width = .Width - 17
            .TextBoxInfo.Left = .PanelInfo.Left + 25
            .TextBoxInfo.Width = .PanelInfo.Width - 50
            .Panel1.Width = (.Width / 2) - 13
            .Panel1.Height = (.Height) - 243
            .RichTextBox1.Left = 26
            .RichTextBox1.Width = .Panel1.Width - 50
            .RichTextBox1.Height = .Panel1.Height - 50
            .RichTextBox1.Top = 22
            .Panel2.Left = (.Panel1.Width) + 10
            .Panel2.Width = (.Width / 2) - 13
            .Panel2.Height = (.Height) - 243
            .Label2.Left = .Panel2.Left + 56
            .TreeView1.Left = 26
            .TreeView1.Width = .Panel2.Width - 50
            .TreeView1.Height = .Panel2.Height - 50
            .TreeView1.Top = 22
            .PictureBox1.Top = .Height - 108
            .PictureBox1.Left = Panel1.Left + ((Panel1.Width / 2) - (PictureBox2.Width / 2))
            .PictureBox2.Top = .Height - 108
            .PictureBox2.Left = Panel2.Left + ((Panel2.Width / 2) - (PictureBox1.Width / 2))
        End With
    End Sub

    Private Sub PopulateTreeView(ByVal dir As String, ByVal parentNode As System.Windows.Forms.TreeNode)
        Dim folder As String = String.Empty

        Try
            'Add the files to treeview
            Dim files() As String = IO.Directory.GetFiles(dir)
            If files.Length <> 0 Then
                Dim fileNode As System.Windows.Forms.TreeNode = Nothing
                For Each file As String In files
                    Dim myTesting() As String = Split(IO.Path.GetFileName(file), ".")
                    If IO.Path.GetExtension(file) = ".dotm" Or IO.Path.GetExtension(file) = ".dot" Then
                        If Not Mid(IO.Path.GetFileName(file), 1, 2) = "~$" Then
                            fileNode = parentNode.Nodes.Add(IO.Path.GetFileName(file), myTesting(LBound(myTesting)), 2, 2)
                            fileNode.Tag = file
                        End If
                    End If
                Next
            End If
            'Add folders to treeview
            Dim folders() As String = IO.Directory.GetDirectories(dir)
            If folders.Length <> 0 Then
                Dim folderNode As System.Windows.Forms.TreeNode = Nothing
                Dim folderName As String = String.Empty
                For Each folder In folders
                    folderName = IO.Path.GetFileName(folder)
                    folderNode = parentNode.Nodes.Add(folderName, folderName, 0, 1)
                    folderNode.Tag = folder
                    PopulateTreeView(folder, folderNode)
                Next
            End If
        Catch ex As UnauthorizedAccessException
            parentNode.Nodes.Add("Access Denied")
        End Try
    End Sub

    Sub PopulateRichTextBox1(newTemp As Word.Document)
        Dim myWord As Word.Application = Globals.ThisAddIn.Application
        Dim myDoc As Word.Document = myWord.ActiveDocument
        Dim myFontStyle As Drawing.FontStyle
        Dim myIndent As Single
        Dim myFontName As String
        Dim myListText As String
        Dim myStyle As Word.Style
        Dim myListFormat As String
        Dim myListTrailing As String = ""
        Dim myListLevel As Integer
        Dim myFontSize As Integer
        Dim myArray(8) As String
        Dim curStart As Integer
        Dim curEnd As Integer
        Dim i As Integer

        myArray(0) = "Title"
        myArray(1) = "Heading 1"
        myArray(2) = "Body Text"
        myArray(3) = "Heading 2"
        myArray(4) = "Body Text 2"
        myArray(5) = "Heading 3"
        myArray(6) = "Body Text 3"
        myArray(7) = "Heading 4"

        RichTextBox1.DeselectAll()
        RichTextBox1.Clear()

        'Overrides myDoc (default set to "Active Document") with the template selected in the treeview,
        '   and preview displays the selected templates styles
        'If a template is not selected in the treeview, newTemp is set to nothing,
        '   and myDoc retains its default setting to Active document and preview displays active documents styles.
        If Not newTemp Is Nothing Then
            myDoc = newTemp
        End If

        For i = 0 To UBound(myArray) - 1
            myListText = ""
            myIndent = 0
            myStyle = myDoc.Styles(myArray(i))
            If Not myStyle.ListTemplate Is Nothing Then
                myListLevel = myStyle.ListLevelNumber
                myListFormat = myStyle.ListTemplate.ListLevels(myListLevel).NumberFormat.ToString
                myListText = Replace(myListFormat, "%", "") & "   "
            End If

            curStart = RichTextBox1.TextLength
            curEnd = Len(myListText & myArray(i))

            RichTextBox1.AppendText(myListText & myArray(i) & Chr(13))
            RichTextBox1.Select(curStart, curEnd)
            RichTextBox1.SelectionColor = Drawing.Color.FromArgb(myStyle.Font.Color)
            RichTextBox1.SelectionAlignment = myStyle.ParagraphFormat.Alignment.value__

            Dim mystylesName As String = myArray(i)

            If myStyle.ParagraphFormat.FirstLineIndent <> 0 Then

                myIndent = (myStyle.ParagraphFormat.FirstLineIndent)
                If myIndent < 0 Then
                    myIndent = myIndent + (myStyle.ParagraphFormat.LeftIndent)
                End If
            Else
                myIndent = (myStyle.ParagraphFormat.LeftIndent)
            End If
            myIndent = myWord.PointsToPixels(myIndent)
            RichTextBox1.SelectionIndent = myIndent

            If (myStyle.Font.Bold) = True Then
                myFontStyle = Drawing.FontStyle.Bold
            ElseIf (myStyle.Font.Italic) = True Then
                myFontStyle = Drawing.FontStyle.Italic
            ElseIf (myStyle.Font.Bold) = False And (myStyle.Font.Italic) = False Then
                myFontStyle = Drawing.FontStyle.Regular
            End If
            myFontName = myStyle.Font.Name
            myFontSize = myStyle.Font.Size
            RichTextBox1.SelectionFont = New Drawing.Font(
            myFontName, myFontSize, myFontStyle, Drawing.GraphicsUnit.Point)

            RichTextBox1.SelectionStart = RichTextBox1.TextLength
        Next i
    End Sub

    Private Sub TreeView1_AfterSelect(sender As Object, e As TreeViewEventArgs) Handles TreeView1.AfterSelect
        If Not Me.TreeView1.SelectedNode Is Nothing Then
            If Not Dir(Me.TreeView1.SelectedNode.Tag.ToString()) = "" Then
                Dim mySelectedChild As String = Me.TreeView1.SelectedNode.Tag.ToString()
                Dim myWord As Word.Application = Globals.ThisAddIn.Application
                Dim myDoc As Word.Document = myWord.ActiveDocument
                Dim myCurrentTemplate As Word.Template = myDoc.AttachedTemplate
                Dim myNewTemplatePath As String = mySelectedChild
                Dim myNewTemp As Word.Document = myWord.Documents.Open(FileName:=myNewTemplatePath, Visible:=False)
                ' May need to check if document failed to open, report to user, and abort.
                PopulateRichTextBox1(myNewTemp) 'NewTemp
                myNewTemp.Close(SaveChanges:=False)
                Me.TextBoxInfo.Text = "Press the 'Attach Template' button to " &
                                "attach the selected template to the current document."
            End If
        End If
    End Sub

    Sub myUpdateDoc()
        Dim myWord As Word.Application = Globals.ThisAddIn.Application
        Dim myDoc As Word.Document = myWord.ActiveDocument
        Dim oLT As Word.ListTemplate
        myDoc.UpdateStyles()
        myDoc.UpdateStyles()
        On Error Resume Next
        For Each oLT In myDoc.AttachedTemplate.ListTemplates
            If Not oLT.Name = "" Then
                If Not myDoc.ListTemplates(oLT.Name).ListLevels(1) _
                        .LinkedStyle = oLT.ListLevels(1).LinkedStyle Then
                    myDoc.UpdateStyles()
                    Exit For
                End If
            End If
        Next oLT
    End Sub

    ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' 
    'Button 1
    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click
        Dim myWord As Word.Application = Globals.ThisAddIn.Application
        Dim myDoc As Word.Document = myWord.ActiveDocument
        Dim myCurrentTemplate As Word.Template = myDoc.AttachedTemplate

        If Not Me.TreeView1.SelectedNode Is Nothing Then
            Dim mySelectedChild As String = Me.TreeView1.SelectedNode.Tag.ToString()
            myDoc.AttachedTemplate = mySelectedChild
            myUpdateDoc()
            Me.TextBoxInfo.Text = "The selected template was attached to the document!"
        ElseIf Me.TreeView1.SelectedNode Is Nothing Then
            Me.TextBoxInfo.Text = "You must select a template from the 'Template Selection' window!"
        End If
        Me.PictureBox1.Image = My.Resources.sys_printer1
    End Sub
    Private Sub PictureBox1_MouseEnter(sender As Object, e As EventArgs) Handles PictureBox1.MouseEnter
        Me.PictureBox1.Image = My.Resources.sys_printer
    End Sub

    Private Sub PictureBox1_MouseLeave(sender As Object, e As EventArgs) Handles PictureBox1.MouseLeave
        Me.PictureBox1.Image = My.Resources.sys_printer1
    End Sub
    ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' 
    'Button 2
    Private Sub PictureBox2_Click(sender As Object, e As EventArgs) Handles PictureBox2.Click
        myUpdateDoc()
        Me.PictureBox2.Image = My.Resources.Refresh1
        Me.TextBoxInfo.Text = "The documents template was refreshed!"
    End Sub
    Private Sub PictureBox2_MouseEnter(sender As Object, e As EventArgs) Handles PictureBox2.MouseEnter
        Me.PictureBox2.Image = My.Resources.Refresh1BIGGER
    End Sub
    Private Sub PictureBox2_MouseLeave(sender As Object, e As EventArgs) Handles PictureBox2.MouseLeave
        Me.PictureBox2.Image = My.Resources.Refresh1
    End Sub
End Class