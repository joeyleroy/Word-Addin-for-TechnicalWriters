Imports System.ComponentModel
Imports System.Windows.Forms

Public Class Snipitz

    Private Sub Snipitz_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim myWord As Word.Application = Globals.ThisAddIn.Application

        Me.TopMost = True

        If Not My.Settings.SnipSize.IsEmpty Then
            Me.Size = My.Settings.SnipSize
        End If

        If Not My.Settings.SnipitsLoc.IsEmpty Then
            Me.Location = My.Settings.SnipitsLoc
        Else
            With Me
                .Top = (myWord.Top)
                .Left = (myWord.Left)
                .StartPosition = 3
            End With
        End If

        With Me
            .PanelInfo.Top = 15
            .PanelInfo.Left = 0
            .PanelInfo.Height = 100
            .PanelInfo.Width = .Width - 20 ' correct
            .TextBoxInfo.Left = 25 ' correct
            .TextBoxInfo.Top = 20 ' correct
            .TextBoxInfo.Height = .PanelInfo.Height - 50
            .TextBoxInfo.Width = .PanelInfo.Width - 50 ' correct
            .LabelInformation.Left = PanelInfo.Left + 50
            .LabelInformation.Top = PanelInfo.Top - 12

            .Panel2.Top = PanelInfo.Top + PanelInfo.Height + 20
            .Panel2.Left = PanelInfo.Left
            .Panel2.Width = .Width - 20 ' correct
            .Panel2.Height = .Height - (PanelInfo.Height + Panel1.Height + Button1.Height + 115)
            .TreeView1.Left = 25 ' correct
            .TreeView1.Top = 25 ' correct
            .TreeView1.Width = Panel2.Width - 50 ' correct
            .TreeView1.Height = Panel2.Height - 50 ' correct
            .Label2.Left = LabelInformation.Left
            .Label2.Top = Panel2.Top - 12

            .Panel1.Top = Panel2.Top + Panel2.Height + 20 ' correct
            .Panel1.Left = PanelInfo.Left
            .Panel1.Height = 60
            .Panel1.Width = .Width - 20 ' correct
            .TextBox1.Top = 20
            .TextBox1.Left = 25 ' correct
            .TextBox1.Height = .Panel1.Height - 50
            .TextBox1.Width = Panel1.Width - 50 ' correct
            .Label1.Top = Panel1.Top - 12
            .Label1.Left = LabelInformation.Left

            .Button1.Top = Panel1.Top + Panel1.Height + 10
            .Button1.Left = (.Width / 2) - ((Button1.Width * 2) + 40)
            .Button1.Height = 60
            .Button1.Width = .Button1.Height
            .Button2.Top = .Button1.Top
            .Button2.Left = .Button1.Left + .Button1.Width + 20
            .Button2.Height = .Button1.Height
            .Button2.Width = .Button1.Height
            .Button3.Top = .Button1.Top
            .Button3.Left = .Button2.Left + .Button1.Width + 20
            .Button3.Height = .Button1.Height
            .Button3.Width = .Button1.Height
            .Button4.Top = .Button1.Top
            .Button4.Left = .Button3.Left + .Button1.Width + 20
            .Button4.Height = .Button1.Height
            .Button4.Width = .Button1.Height

            .TreeView1.ImageList = .ImageList1
            .PWTextBox.Text = ""
            .PWPanel.Visible = False
            .PWLabel.Visible = False
            .TextBoxInfo.Text = "Select a Snipit to insert into the document, or select content in the document to create a Snipit."
        End With

        myTree()

    End Sub

    Private Sub Snipitz_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        My.Settings.SnipSize = Me.Size
        My.Settings.SnipitsLoc = Me.Location
    End Sub

    Private Sub Snipitz_ResizeEnd(sender As Object, e As EventArgs) Handles Me.ResizeEnd
        With Me
            .PanelInfo.Width = .Width - 20 ' correct
            .TextBoxInfo.Left = 25 ' correct
            .TextBoxInfo.Top = 20 ' correct
            .TextBoxInfo.Height = .PanelInfo.Height - 50
            .TextBoxInfo.Width = .PanelInfo.Width - 50 ' correct
            .LabelInformation.Left = PanelInfo.Left + 50
            .LabelInformation.Top = PanelInfo.Top - 12

            .Panel2.Top = PanelInfo.Top + PanelInfo.Height + 20
            .Panel2.Left = PanelInfo.Left
            .Panel2.Width = .Width - 20 ' correct
            .Panel2.Height = .Height - (PanelInfo.Height + Panel1.Height + Button1.Height + 115)
            .TreeView1.Left = 25 ' correct
            .TreeView1.Top = 25 ' correct
            .TreeView1.Width = Panel2.Width - 50 ' correct
            .TreeView1.Height = Panel2.Height - 50 ' correct
            .Label2.Left = LabelInformation.Left
            .Label2.Top = Panel2.Top - 12

            .Panel1.Top = Panel2.Top + Panel2.Height + 20 ' correct
            .Panel1.Left = PanelInfo.Left
            .Panel1.Height = 60
            .Panel1.Width = .Width - 20 ' correct
            .TextBox1.Top = 20
            .TextBox1.Left = 25 ' correct
            .TextBox1.Height = .Panel1.Height - 50
            .TextBox1.Width = Panel1.Width - 50 ' correct
            .Label1.Top = Panel1.Top - 12
            .Label1.Left = LabelInformation.Left

            .Button1.Top = Panel1.Top + Panel1.Height + 10
            .Button1.Left = (.Width / 2) - ((Button1.Width * 2) + 40)
            .Button1.Height = 60
            .Button1.Width = .Button1.Height
            .LabelInsert.Top = Panel1.Top + Panel1.Height + Button1.Height
            .LabelInsert.Left = Button1.Left
            .Button2.Top = .Button1.Top
            .Button2.Left = .Button1.Left + .Button1.Width + 20
            .Button2.Height = .Button1.Height
            .Button2.Width = .Button1.Height
            .LabelEdit.Top = Panel1.Top + Panel1.Height + Button2.Height
            .LabelEdit.Left = Button2.Left
            .Button3.Top = .Button1.Top
            .Button3.Left = .Button2.Left + .Button1.Width + 20
            .Button3.Height = .Button1.Height
            .Button3.Width = .Button1.Height
            .LabelCreate.Top = Panel1.Top + Panel1.Height + Button3.Height
            .LabelCreate.Left = Button3.Left
            .Button4.Top = .Button1.Top
            .Button4.Left = .Button3.Left + .Button1.Width + 20
            .Button4.Height = .Button1.Height
            .Button4.Width = .Button1.Height
            .LabelDelete.Top = Panel1.Top + Panel1.Height + Button4.Height
            .LabelDelete.Left = Button4.Left
        End With
    End Sub

    ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' 
    'Treeview1 populate
    Private Sub myTree()

        Dim myDocs As String = IO.Path.Combine(Environment.ExpandEnvironmentVariables("%userprofile%"), "Documents")
        Dim myPathconfig As String = myDocs & "\Snipitz\"
        Dim myConfigPath As String = myPathconfig & "Config\"
        Dim mySnipitzTXT As String = "snipitzFolder.txt"
        Dim mySconfig As String = myConfigPath & mySnipitzTXT

        If Not System.IO.File.Exists(mySconfig) Then
            MessageBox.Show("You MUST run 'Configuration' to configure a Snipitz directory!")
            Me.Close()
            Return
        End If

        Dim myUser As String = Environ("USERPROFILE")
        Dim snipPath As String = My.Computer.FileSystem.ReadAllText(mySconfig)
        Dim myRoot As String
        If System.IO.Directory.Exists(snipPath) Then
            'user defined directory exists
            myRoot = snipPath
        Else
            MessageBox.Show("Run Configuration to set a Snipitz directory.")
            Me.Close()
            Return
        End If

        Dim myFileInfo As New IO.DirectoryInfo(myRoot)
        Dim myPath() As String = {}
        Dim rootDir As String = String.Empty
        'Loop through subfolders
        For Each subfolder As IO.DirectoryInfo In myFileInfo.GetDirectories()
            'Add this folders name
            Array.Resize(myPath, myPath.Length + 1)
            myPath(myPath.Length - 1) = subfolder.FullName
        Next
        Array.Resize(myPath, myPath.Length + 1)
        myPath(myPath.Length - 1) = "Add New Category"

        'Loop thru each directory and populate the treeview
        For i As Integer = 0 To myPath.Count - 1
            rootDir = myPath(i)
            Dim mySplit As String() = rootDir.Split("\")
            Dim parentFolder As String = mySplit(mySplit.Length - 1)
            'Add this directory as a root node
            If Not parentFolder = "Add New Category" Then
                Dim root As System.Windows.Forms.TreeNode = TreeView1.Nodes.Add(parentFolder & (i + 1), parentFolder)
                root.Tag = rootDir
                'Populate this root node
                sPopulateTreeView(rootDir, TreeView1.Nodes(i))
            Else
                Dim root As System.Windows.Forms.TreeNode = TreeView1.Nodes.Add("Add New Category" & (i + 1), "Add New Category", 3, 3)
                root.Tag = myRoot
            End If
        Next
    End Sub

    Private Sub sPopulateTreeView(ByVal myDir As String, ByVal parentNode As System.Windows.Forms.TreeNode)

        Dim folder As String = String.Empty
        Dim ii As Integer = 0

        Try
            'Add the files to treeview
            If Not myDir = "Add New Category" Then
                Dim files() As String = IO.Directory.GetFiles(myDir)
                If Not files.Length = 0 Then
                    Dim fileNode As System.Windows.Forms.TreeNode = Nothing
                    For Each file As String In files
                        Dim mySplit() As String = Split(IO.Path.GetFileName(file), ".")
                        If IO.Path.GetExtension(file) = ".docm" Or IO.Path.GetExtension(file) = ".docx" Then
                            If Not Mid(IO.Path.GetFileName(file), 1, 2) = "~$" Then
                                fileNode = parentNode.Nodes.Add(IO.Path.GetFileName(file), mySplit(LBound(mySplit)), 2, 2)
                                fileNode.Tag = file
                            End If
                        End If
                    Next
                End If
                'Add folders to treeview
                Dim folders() As String = IO.Directory.GetDirectories(myDir)
                If folders.Length <> 0 Then
                    Dim folderNode As System.Windows.Forms.TreeNode = Nothing
                    Dim newfolderNode As System.Windows.Forms.TreeNode = Nothing
                    Dim folderName As String = String.Empty
                    For Each folder In folders
                        folderName = IO.Path.GetFileName(folder)
                        folderNode = parentNode.Nodes.Add(folderName, folderName, 0, 1)
                        folderNode.Tag = folder
                        If folder = folders(UBound(folders)) Then
                            newfolderNode = parentNode.Nodes.Add("Add New Category", "Add New Category", 3, 3)
                            newfolderNode.Tag = folder
                        End If
                        sPopulateTreeView(folder, folderNode)
                    Next
                Else
                    Dim newfolderNode As System.Windows.Forms.TreeNode = Nothing
                    newfolderNode = parentNode.Nodes.Add("Add New Category", "Add New Category", 3, 3)
                    newfolderNode.Tag = folder
                End If
            End If
        Catch ex As UnauthorizedAccessException
            parentNode.Nodes.Add("Access Denied")
        End Try
    End Sub

    Private Sub TreeView1_AfterSelect(sender As Object, e As TreeViewEventArgs) Handles TreeView1.AfterSelect
        'MsgBox("Name: " & Me.TreeView1.SelectedNode.Name & vbNewLine & "Tag: " & Me.TreeView1.SelectedNode.Tag & vbNewLine & "Text: " & Me.TreeView1.SelectedNode.Text)
        Try
            If Me.TreeView1.SelectedNode.Name = "Add New Category" Then
                Me.TextBoxInfo.Text = "Enter a Category name into the Title Box."
                Me.TextBox1.Clear()
                Me.TextBox1.Focus()
            ElseIf Me.TreeView1.SelectedNode.ImageIndex = 2 Then
                Me.TextBox1.Text = Me.TreeView1.SelectedNode.Text
            End If
        Catch
        End Try
    End Sub

    ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' 
    'Button 1 Insert Snipit
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        myInsertSnipit()
    End Sub
    Private Sub Button1_MouseEnter(sender As Object, e As EventArgs) Handles Button1.MouseEnter
        Me.Button1.Image = My.Resources.Insert2
    End Sub
    Private Sub Button1_MouseLeave(sender As Object, e As EventArgs) Handles Button1.MouseLeave
        Me.Button1.Image = My.Resources.Insert1
    End Sub
    'Button 2 Edit/Preview
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        myEditPreview()
    End Sub
    Private Sub Button2_MouseEnter(sender As Object, e As EventArgs) Handles Button2.MouseEnter
        Me.Button2.Image = My.Resources.EditPreview2
    End Sub
    Private Sub PictureBox2_MouseLeave(sender As Object, e As EventArgs) Handles Button2.MouseLeave
        Me.Button2.Image = My.Resources.EditPreview1
    End Sub
    'Button 3 Create Snipit
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        myCreateSnipit()
    End Sub
    Private Sub Button3_MouseEnter(sender As Object, e As EventArgs) Handles Button3.MouseEnter
        Me.Button3.Image = My.Resources.Create2
    End Sub
    Private Sub PictureBox3_MouseLeave(sender As Object, e As EventArgs) Handles Button3.MouseLeave
        Me.Button3.Image = My.Resources.Create1
    End Sub
    'Button 4 Delete Snipit
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        myDeleteSnipit()
    End Sub
    Private Sub Button4_MouseEnter(sender As Object, e As EventArgs) Handles Button4.MouseEnter
        Me.Button4.Image = My.Resources.Trash2
    End Sub
    Private Sub PictureBox4_MouseLeave(sender As Object, e As EventArgs) Handles Button4.MouseLeave
        Me.Button4.Image = My.Resources.Trash1
    End Sub

    ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' 
    ' TEXTBOX1 ' ' ' TITLE ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
    Private Sub TextBox1_GotFocus(sender As Object, e As EventArgs) Handles TextBox1.GotFocus
        Dim myWord As Word.Application = Globals.ThisAddIn.Application
        Dim myDoc As Word.Document = myWord.ActiveDocument
        Dim mySelection As Word.Selection = myWord.Selection
        Dim myRange As Word.Range = mySelection.Range

        Dim mySelectedNode As TreeNode = Me.TreeView1.SelectedNode ' Selected Snipitz Node
        If Not IsNothing(mySelectedNode) Then
            Select Case mySelectedNode.ImageIndex
                Case < 2
                    If Not mySelection.Paragraphs(1).Range.Text = "" Then
                        Me.TextBox1.Text = Mid(Replace(Replace(mySelection.Paragraphs(1).Range.Text, vbCr, ""), ".", ""), 1, 40)
                        Me.TextBox1.SelectAll()
                    End If
                Case = 2
                    TextBox1.Text = mySelectedNode.Name
                Case > 2
                    'Nothing
            End Select
        End If
    End Sub

    ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' 
    ' PWTextBox ' ' ' PASSWORD ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
    Private Sub PWTextBox_KeyDown(sender As Object, e As KeyEventArgs) Handles PWTextBox.KeyDown
        If e.KeyCode = Keys.Enter Then
            myDeleteSnipit()
        ElseIf e.KeyCode = Keys.Escape Then
            Me.PWTextBox.Text = ""
            Me.PWPanel.Visible = False
            Me.PWLabel.Visible = False
        End If
    End Sub

    ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' 
    ' TREEVIEW1 ' ' ' DOUBLE CLICK NODES ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
    Private Sub TreeView1_NodeMouseDoubleClick(sender As Object, e As TreeNodeMouseClickEventArgs) Handles TreeView1.NodeMouseDoubleClick

        Dim mySelectedNode As TreeNode = Me.TreeView1.SelectedNode ' Selected Snipitz Node

        Select Case mySelectedNode.ImageIndex
            Case = 0
                mySelectedNode.Expand()
            Case = 1
                mySelectedNode.Collapse()
            Case = 2
                myEditPreview()
            Case = 3
                If TextBox1.Text = "" Then
                    TextBoxInfo.Text = "Enter a title for the new Category"
                    TextBox1.Focus()
                Else
                    myAddNode()
                End If
        End Select
        mySelectedNode = Nothing
    End Sub

    Sub myAddNode()
        '''''''''''''''''''''''''''''''''''''''
        Dim myDocs As String = IO.Path.Combine(Environment.ExpandEnvironmentVariables("%userprofile%"), "Documents")
        Dim myPathconfig As String = myDocs & "\Snipitz\"
        Dim myConfigPath As String = myPathconfig & "Config\"
        Dim mySnipitzTXT As String = "snipitzFolder.txt"
        Dim mySconfig As String = myConfigPath & mySnipitzTXT

        If Not System.IO.File.Exists(mySconfig) Then
            MessageBox.Show("You MUST run 'Configuration' to configure a Snipitz directory!")
            Me.Close()
            Return
        End If

        Dim myUser As String = Environ("USERPROFILE")
        Dim snipPath As String = My.Computer.FileSystem.ReadAllText(mySconfig)
        Dim myRoot As String
        If System.IO.Directory.Exists(snipPath) Then
            'user defined directory exists
            myRoot = snipPath
        Else
            MessageBox.Show("Run Configuration to set a Snipitz directory.")
            Me.Close()
            Return
        End If
        ''''''''''''''''''''''''''''''''''
        Dim myWord As Word.Application = Globals.ThisAddIn.Application
        Dim myDoc As Word.Document = myWord.ActiveDocument
        Dim myDOCX As String = ".docx"
        Dim myNodeInfo As String = ""
        Dim mySelectedNode As TreeNode = Me.TreeView1.SelectedNode ' Selected Snipitz Node

        If mySelectedNode.Level = 0 Then
            myNodeInfo = myRoot & TextBox1.Text & "\"
            Try
                MkDir(myNodeInfo)
                Dim myNewNode As TreeNode
                myNewNode = Me.TreeView1.Nodes.Add(Me.TextBox1.Text, Me.TextBox1.Text, 0, 1)
                myNewNode.Tag = myNodeInfo
                Me.TextBoxInfo.Text = "New Category created!"
                Me.TextBox1.Clear()
                'Me.TreeView1.Refresh()
                Me.Refresh()
            Catch
                Me.TextBoxInfo.Text = "The new Category could not be created!"
            End Try
        ElseIf mySelectedNode.Level > 0 Then
            myNodeInfo = mySelectedNode.Parent.Tag & "\" & TextBox1.Text & "\"
            Try
                MkDir(myNodeInfo)
                Dim myNewNode As TreeNode
                myNewNode = Me.TreeView1.SelectedNode.Parent.Nodes.Add(Me.TextBox1.Text, Me.TextBox1.Text, 0, 1)
                myNewNode.Tag = myNodeInfo
                Me.TextBoxInfo.Text = "New Category created!"
                Me.TextBox1.Clear()
                Me.TreeView1.Refresh()
            Catch
                Me.TextBoxInfo.Text = "The new Category could not be created!"
            End Try
        End If
        mySelectedNode = Nothing
        myNodeInfo = Nothing
    End Sub

    ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' 
    ' BUTTON ACTIONS ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
    Sub myInsertSnipit()
        Dim myWord As Word.Application = Globals.ThisAddIn.Application
        Dim myDoc As Word.Document = myWord.ActiveDocument
        Dim mySelection As Word.Selection = myWord.Selection
        Dim myRange As Word.Range = mySelection.Range

        Dim mySelectedNode As TreeNode = Me.TreeView1.SelectedNode ' Selected Snipitz Node
        If Not IsNothing(mySelectedNode) Then
            Select Case Me.TreeView1.SelectedNode.ImageIndex
                Case < 2 'Folder Closed = 0, Folder Open 1
                    Me.TextBoxInfo.Text = "You cant insert a Category! Select a Snipit instead!"
                Case = 2 'File = 2
                    With myRange
                        .Start = mySelection.Paragraphs(1).Range.Start
                        .End = mySelection.Paragraphs(mySelection.Paragraphs.Count).Range.End
                    End With
                    Try
                        If Replace(myRange.Text, vbCr, "") = "" Then ' Ensure that the user isnt accidentally over-writing part of the document.
                            myRange.InsertFile(mySelectedNode.Tag) ' File being inserted
                            With mySelection ' This is so the cursor is at the end of the inserted content.
                                .Start = myRange.Start
                                .End = myRange.Paragraphs(.Paragraphs.Count).Range.End 'This may need to be tweeked to leave the users cursor at the end of the snipit
                                .Collapse(Direction:=Word.WdCollapseDirection.wdCollapseEnd)
                                .Select()
                            End With
                            Me.TextBoxInfo.Text = "The Snipit was inserted!"
                        Else
                            Me.TextBoxInfo.Text = "You must select an empty line to insert a Snipitz!"
                        End If
                    Catch
                        Me.TextBoxInfo.Text = "The Snipit cound not be inserted!"
                    End Try
                Case = 3 'New Category 3
                    Me.TextBoxInfo.Text = "Select a Snipit not a 'Add New Category' button!"
            End Select
        Else
            Me.TextBoxInfo.Text = "Select a Folder to save to!"
        End If
    End Sub

    Sub myEditPreview()
        Dim myWord As Word.Application = Globals.ThisAddIn.Application
        Dim myDoc2 As New Word.Document
        Dim mySelectedNode As TreeNode = Me.TreeView1.SelectedNode ' Selected Snipitz Node

        Select Case Me.TreeView1.SelectedNode.ImageIndex
            Case < 2 'Folder Closed = 0, Folder Open 1
                Me.TextBoxInfo.Text = "You cant preview or edit a Category! Select a Snipit instead!"
            Case = 2 'File = 2
                Try
                    myDoc2 = myWord.Documents.Open(FileName:=mySelectedNode.Tag, ReadOnly:=False, Visible:=True)
                    'myDoc2 = myWord.Documents.Open(FileName:=mySelectedNode.Tag, ReadOnly:=False, Visible:=True, Revert:=False)
                    With myDoc2
                        .Activate()
                        'Size the Snipitz preview window
                        With .ActiveWindow
                            .WindowState = Word.WdWindowState.wdWindowStateNormal
                            .Top = 35
                            .Left = 35
                            .Height = 500
                            .Width = 750
                        End With
                    End With
                    Me.TextBoxInfo.Text = "Snipit opened for preview/editing!"
                Catch
                    Me.TextBoxInfo.Text = "Snipit failed to open!"
                End Try
            Case = 3 'New Category 3
                Me.TextBoxInfo.Text = "You cant preview or edit a 'Add New Category' button!  Select a Snipit instead!"
        End Select

        myWord = Nothing
        'myDoc2 = Nothing
        mySelectedNode = Nothing

    End Sub

    Sub myCreateSnipit()
        Dim myWord As Word.Application = Globals.ThisAddIn.Application
        Dim myDoc As Word.Document = myWord.ActiveDocument
        Dim myDoc2 As Word.Document
        Dim mySelection As Word.Selection = myWord.Selection
        Dim myRange As Word.Range = mySelection.Range
        Dim myDOCX As String = ".docx"
        Dim myTemplate As Word.Template = myDoc.AttachedTemplate
        Dim myTemplatePath As String = myTemplate.Path & myWord.PathSeparator & myTemplate.Name
        Dim mySelectedNode As TreeNode = Me.TreeView1.SelectedNode ' Selected Snipitz Node

        Select Case mySelectedNode.ImageIndex
            Case < 2 'Folder Closed = 0, Folder Open 1
                If TextBox1.Text = "" Then
                    TextBoxInfo.Text = "Enter a title for the new Snipit"
                    TextBox1.Focus()
                Else
                    If (Dir(mySelectedNode.Tag & "\" & Me.TextBox1.Text & myDOCX) > "") Then
                        Me.TextBoxInfo.Text = "A Snipit with that name already exists! Try a different name."
                    Else
                        With myRange
                            .Start = .Paragraphs(1).Range.Start
                            .End = .Paragraphs(.Paragraphs.Count).Range.End
                        End With

                        If Not mySelection.Text = "" Then
                            myDoc2 = myWord.Documents.Add(Template:=myTemplatePath, Visible:=False)
                            mySelection.Copy()
                            With myDoc2
                                .Range.Collapse(Direction:=Word.WdCollapseDirection.wdCollapseStart)
                                .Range.Paste()
                                Do Until mySelection.Paragraphs.Count = .Paragraphs.Count
                                    .Paragraphs(.Paragraphs.Count).Range.Delete()
                                Loop
                            End With
                            If Dir(mySelectedNode.Tag, vbDirectory) = "" Then
                                Try
                                    MkDir(mySelectedNode.Tag & "\")
                                Catch
                                    Me.TextBoxInfo.Text = "The Snipit could not be created due to Category restrictions!"
                                End Try
                            End If
                            myDoc2.SaveAs2(FileName:=mySelectedNode.Tag & "\" & Me.TextBox1.Text & myDOCX)
                            myDoc2.Close()
                            'May need to use create node code here.
                            Me.TreeView1.SelectedNode.Nodes.Add(mySelectedNode.Tag & "\" & Me.TextBox1.Text & myDOCX, Me.TextBox1.Text, 2, 2)
                            Me.TreeView1.Refresh()
                            Me.TextBoxInfo.Text = " Snipit created!"
                            Me.TextBox1.Clear()
                        Else
                            Me.TextBoxInfo.Text = "Select something in the Document before creating the Snipit!"
                        End If
                    End If
                End If
            Case = 2 'File = 2
                Me.TextBoxInfo.Text = "Select a Category (not a Snipit) to save the Snipit into!"
            Case = 3 'New Category 3
                Me.TextBoxInfo.Text = "Select a Category (not a 'Add New Category' button) to save the Snipit into!"
        End Select
    End Sub

    Sub myDeleteSnipit()
        Dim myDeletePassword As String = Chr(113) & Chr(119) & Chr(101) & Chr(114) & Chr(116) & Chr(121)
        Dim mySelectedNode As TreeNode = Me.TreeView1.SelectedNode ' Selected Snipitz Node

        Me.PWPanel.Visible = True
        Me.PWLabel.Visible = True
        Me.PWTextBox.Focus()
        Me.Button4.Image = My.Resources.Refresh1

        If LCase$(Me.PWTextBox.Text) = LCase$(myDeletePassword) Then
            Select Case Me.TreeView1.SelectedNode.ImageIndex
                Case < 2 'Folder Closed = 0, Folder Open 1
                    Try
                        Dim myOldKey As String = mySelectedNode.Parent.Name
                        mySelectedNode.Parent.Name = "Zapp"
                        System.IO.Directory.Delete(mySelectedNode.Tag, False)
                        Me.TreeView1.Nodes.Remove(Me.TreeView1.SelectedNode)
                        Dim myFindNode As TreeNode() = Me.TreeView1.Nodes.Find(key:="Zapp", searchAllChildren:=True)
                        Me.TreeView1.SelectedNode = Me.TreeView1.Nodes.Find(key:="Zapp", searchAllChildren:=True)(0)
                        Me.TreeView1.SelectedNode.Name = myOldKey
                        Me.TextBoxInfo.Text = "Selected Category deleted!"
                        Me.TextBox1.Clear()
                    Catch
                        Me.TextBoxInfo.Text = "For file assurance, you cannot delete a Category containing SubCatagories or Snipitz!"
                    End Try
                Case = 2 'File = 2
                    Dim mySNodeIndex As TreeNode = mySelectedNode.Parent
                    If (Dir(mySelectedNode.Tag) > "") Then
                        Try
                            Dim mySNodeIndex2 As TreeNode = mySelectedNode.Parent
                            Kill(mySelectedNode.Tag)
                            Me.TextBoxInfo.Text = "The Snipit was deleted."
                            Me.TreeView1.Nodes.Remove(Me.TreeView1.SelectedNode)
                            Me.TreeView1.Refresh()
                            Me.TextBox1.Clear()
                            Me.TreeView1.SelectedNode = mySNodeIndex2
                        Catch
                            Me.TextBoxInfo.Text = "The Snipit could not be deleted!"
                        End Try
                    Else
                        Me.TextBoxInfo.Text = "Select something to delete first!"
                    End If
                Case = 3 'New Category 3
                    Me.TextBoxInfo.Text = "Select a Snipit not a 'Add New Category' button!"
            End Select
            Me.PWTextBox.Text = ""
            Me.PWPanel.Visible = False
            Me.PWLabel.Visible = False
        ElseIf Not Me.PWTextBox.Text = "" Then ' Not Blank
            Me.TextBoxInfo.Text = "The password is incorrect!"
            Me.PWTextBox.Text = ""
            Me.PWPanel.Visible = False
            Me.PWLabel.Visible = False
        ElseIf Me.PWTextBox.Text = "" Then ' Blank
            If Me.TreeView1.SelectedNode.GetNodeCount(True) > 1 Then
                Me.TextBoxInfo.Text = "For file assurance, you cannot delete a Category containing SubCatagories or Snipitz!"
                Me.PWPanel.Visible = False
                Me.PWLabel.Visible = False
            Else
                Me.TextBoxInfo.Text = "Enter the deletion confirmation password into the password box, and then press the Delete button again."
            End If
        End If
    End Sub

    Sub mySelectedSave()
        Dim myNowSelected As String = Me.TreeView1.SelectedNode.Tag
        Dim myNowSelectedName As String = Me.TreeView1.SelectedNode.Name
        MsgBox(myNowSelectedName)
        'Rebuild Treeview
        'Me.TreeView1.Nodes.Find(myNowSelectedName)
    End Sub

End Class