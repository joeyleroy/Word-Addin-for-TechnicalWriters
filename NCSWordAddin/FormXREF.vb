Imports System.ComponentModel

Public Class FormXREF
    Dim myCheck1, myCheck2, myCheck3, myCheck4, myCheck5, myCheck6 As Boolean
    Public myCheck As Integer

    'Loading and Defaults
    Private Sub FormXREF_Load(sender As Object, e As EventArgs) Handles Me.Load
        'Form Loads
        Dim myWord As Word.Application = Globals.ThisAddIn.Application
        Dim myDoc As Word.Document = myWord.ActiveDocument
        Me.TopMost = True

        If Not My.Settings.XRefSize.IsEmpty Then
            Me.Size = My.Settings.XRefSize
        End If

        If Not My.Settings.XRefLoc.IsEmpty Then
            Me.Location = My.Settings.XRefLoc
        Else
            With Me
                .Top = (myWord.Top)
                .Left = (myWord.Left)
                .StartPosition = 3
            End With
        End If

        With Me
            .PanelInfo.Width = .Width - 20 ' correct
            .TextBoxInfo.Left = 25 ' correct
            .TextBoxInfo.Top = 20 ' correct
            .TextBoxInfo.Height = .PanelInfo.Height - 50
            .TextBoxInfo.Width = .PanelInfo.Width - 50 ' correct
            .Label3.Left = PanelInfo.Left + 50
            .Label3.Top = PanelInfo.Top - 12

            .Button1.Top = PanelInfo.Top + PanelInfo.Height + 10
            .Button1.Left = (.Width / 2) - ((Button1.Width * 2) + 40)
            .Button1.Height = 60
            .Button1.Width = .Button1.Height
            .Button2.Top = .Button1.Top
            .Button2.Left = .Button1.Left + .Button1.Width + 20
            .Button2.Height = .Button1.Height
            .Button2.Width = .Button1.Width
            .Button3.Top = .Button1.Top
            .Button3.Left = .Button2.Left + .Button1.Width + 20
            .Button3.Height = .Button1.Height
            .Button3.Width = .Button1.Width
            .Button4.Top = .Button1.Top
            .Button4.Left = .Button3.Left + .Button1.Width + 20
            .Button4.Height = .Button1.Height
            .Button4.Width = .Button1.Width

            .ListBoxBackground.Top = Button1.Top + Button1.Height + 20
            .ListBoxBackground.Left = PanelInfo.Left
            .ListBoxBackground.Width = .Width - 20 ' correct
            .ListBoxBackground.Height = .Height - (PanelInfo.Top + PanelInfo.Height + 20 + Button1.Height + 20 + 43)
            .ListBox1.Left = .ListBoxBackground.Left + 25 ' correct
            .ListBox1.Top = .ListBoxBackground.Top + 25 ' correct
            .ListBox1.Width = ListBoxBackground.Width - 50 ' correct
            .ListBox1.Height = ListBoxBackground.Height - 50 ' correct
            .Label4.Left = Label3.Left
            .Label4.Top = ListBoxBackground.Top - 12

            .Button1Options.Top = .Button3.Top + .Button3.Height
            .Button1Options.Left = .Button3.Left
        End With


        ' XREF
        Dim myHeadings
        Dim i As Integer
        myHeadings = myDoc.GetCrossReferenceItems(Word.WdReferenceType.wdRefTypeHeading)
        For i = 1 To UBound(myHeadings)
            Me.ListBox1.Items.Add(myHeadings(i))
        Next i

        'Button Loads
        Me.Button1.Image = My.Resources.orbital_folder_downloads1
        Me.Button1Options.Visible = False
        Button1Check1.Image = My.Resources.CheckON
        Button1Check2.Image = My.Resources.CheckOFF
        Button1Check3.Image = My.Resources.CheckOFF
        Button1Check4.Image = My.Resources.CheckOFF
        Button1Check5.Image = My.Resources.CheckOFF
        Button1Check6.Image = My.Resources.CheckOFF
        myCheck = 1
        myCheck1 = True
        Me.Button2.Image = My.Resources.orbital_search
        Me.Button3.Image = My.Resources._Select
        Me.Button4.Image = My.Resources.Refresh1

        Dim myInfo As String
        myInfo = "Welcome to X-Ref! Select a heading from the window and press the Insert button to insert a cross-reference into the document, or press the Selection button perform other tasks."
        Me.TextBoxInfo.Text = myInfo
    End Sub

    Private Sub FormXREF_ResizeEnd(sender As Object, e As EventArgs) Handles Me.ResizeEnd
        With Me
            .PanelInfo.Width = .Width - 20 ' correct
            .TextBoxInfo.Left = 25 ' correct
            .TextBoxInfo.Top = 20 ' correct
            .TextBoxInfo.Height = .PanelInfo.Height - 50
            .TextBoxInfo.Width = .PanelInfo.Width - 50 ' correct
            .Label3.Left = PanelInfo.Left + 50
            .Label3.Top = PanelInfo.Top - 12

            .Button1.Top = PanelInfo.Top + PanelInfo.Height + 10
            .Button1.Left = (.Width / 2) - ((Button1.Width * 2) + 40)
            .Button1.Height = 60
            .Button1.Width = .Button1.Height
            .Button2.Top = .Button1.Top
            .Button2.Left = .Button1.Left + .Button1.Width + 20
            .Button2.Height = .Button1.Height
            .Button2.Width = .Button1.Width
            .Button3.Top = .Button1.Top
            .Button3.Left = .Button2.Left + .Button1.Width + 20
            .Button3.Height = .Button1.Height
            .Button3.Width = .Button1.Width
            .Button4.Top = .Button1.Top
            .Button4.Left = .Button3.Left + .Button1.Width + 20
            .Button4.Height = .Button1.Height
            .Button4.Width = .Button1.Width

            .ListBoxBackground.Top = Button1.Top + Button1.Height + 20
            .ListBoxBackground.Left = PanelInfo.Left
            .ListBoxBackground.Width = .Width - 20 ' correct
            .ListBoxBackground.Height = .Height - (PanelInfo.Top + PanelInfo.Height + 20 + Button1.Height + 20 + 43)
            .ListBox1.Left = .ListBoxBackground.Left + 25 ' correct
            .ListBox1.Top = .ListBoxBackground.Top + 25 ' correct
            .ListBox1.Width = ListBoxBackground.Width - 50 ' correct
            .ListBox1.Height = ListBoxBackground.Height - 50 ' correct
            .Label4.Left = Label3.Left
            .Label4.Top = ListBoxBackground.Top - 12

            .Button1Options.Top = .Button3.Top + .Button3.Height
            .Button1Options.Left = .Button3.Left
        End With
    End Sub

    Private Sub FormXREF_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        My.Settings.XRefSize = Me.Size
        My.Settings.XRefLoc = Me.Location
    End Sub

    Private Sub FormXREF_MouseEnter(sender As Object, e As EventArgs) Handles Me.MouseEnter
        Me.Button1Options.Visible = False
        Me.Button3.Image = My.Resources._Select
    End Sub

    ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' 
    'Button 1
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        InsertXREF()
    End Sub
    Private Sub Button1_MouseEnter(sender As Object, e As EventArgs) Handles Button1.MouseEnter
        Me.Button1.Image = My.Resources.orbital_folder_downloadsBIGGER
    End Sub
    Private Sub Button1_MouseLeave(sender As Object, e As EventArgs) Handles Button1.MouseLeave
        Me.Button1.Image = My.Resources.orbital_folder_downloads1
    End Sub


    ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' 
    'Button 2
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        'Find Next
        Dim myWord As Word.Application = Globals.ThisAddIn.Application
        Dim myDoc As Word.Document = myWord.ActiveDocument
        Dim mySelection As Word.Selection = myWord.Selection
        Dim objFld As Word.Field
        Dim docRange As Word.Range = myDoc.Range
        Dim selRange As Word.Range = mySelection.Range
        Dim newRange As Word.Range = myDoc.Range(Start:=selRange.End, End:=docRange.End)
        Dim boRef1 As Boolean = False, boRef2 As Boolean = False
        Dim myHeadings
        Dim i As Integer
        Dim selArray As Array
        Dim headArray As Array
        Dim evalArray(0 To 1)

        ' Loop through fields in the ActiveDocument
        For Each objFld In newRange.Fields
            ' If the field is a cross-ref, do something to it.
            If objFld.Type = Word.WdFieldType.wdFieldRef Then
                objFld.Select() ' Select the Cross-Reference
                selRange = mySelection.Range ' Redefine selRange to the selected Cross-Reference
                mySelection.Collapse(Direction:=Word.WdCollapseDirection.wdCollapseStart) ' Collapse the selection
                mySelection.MoveLeft(Unit:=Word.WdUnits.wdWord, Count:=1, Extend:=Word.WdMovementType.wdExtend) ' Move the selection left by a word
                selRange.SetRange(Start:=mySelection.Range.Start, End:=selRange.End) ' Redefine the start of selRange to the start of the selection
                selRange.Select() ' Select the selRange
                boRef1 = True
                Exit For
            End If
        Next objFld

        myHeadings = myDoc.GetCrossReferenceItems(Word.WdReferenceType.wdRefTypeHeading)
        selRange = mySelection.Range
        If boRef1 = True Then ' Found a Cross-Reference
            For i = 1 To UBound(myHeadings)
                headArray = Split(Trim(myHeadings(i)), Chr(32))
                selArray = Split(Trim(selRange.Text), Chr(160))
                evalArray(0) = selRange.Text
                evalArray(1) = myHeadings(i)
                If headArray(0) = selArray(UBound(selArray)) Then
                    boRef2 = True
                    Exit For
                End If
            Next i
            If boRef2 = True Then ' Found a Cross-Reference that matches a Heading
                With Me
                    '    snipText = "'" & evalArray(0) & "' links to:" & evalArray(1) & "'"
                    '    SpeakText snipText
                    Dim myInfo As String
                    myInfo = "a cross-reference matching one of the headings has been found."
                    .TextBoxInfo.Text = myInfo
                    .Refresh()
                End With
            Else 'Found a Cross-Reference, but it does not have a matching Heading
                '    snipText = evalArray(0) & "  " & evalArray(1)
                '    SpeakText snipText
                Dim myInfo As String
                myInfo = "A broken cross-reference has been found!"
                With Me
                    .TextBoxInfo.Text = myInfo
                    .Refresh()
                End With
            End If
        Else ' Did NOT find a Cross-Reference
            With Me
                '    snipText = evalArray(0) & "  " & evalArray(1)
                '    SpeakText snipText
                Dim myInfo As String
                myInfo = "Cross-reference not found and reached the end of the document. Moving to the Top of the Document."
                .TextBoxInfo.Text = myInfo
                .Refresh()
            End With
            selRange = myDoc.Range ' Redefine selRange to the entire document
            selRange.Collapse(Direction:=Word.WdCollapseDirection.wdCollapseStart) ' Collapses the selection so that the selection is at the start of the document
            selRange.Select()
        End If
    End Sub

    Private Sub Button2_MouseEnter(sender As Object, e As EventArgs) Handles Button2.MouseEnter
        Me.Button2.Image = My.Resources.orbital_searchBIGGER
    End Sub
    Private Sub Button2_MouseLeave(sender As Object, e As EventArgs) Handles Button2.MouseLeave
        Me.Button2.Image = My.Resources.orbital_search
    End Sub

    ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' 
    'Button 3 Open Select Cross-Reference Type Menu
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Me.Button1Options.Visible = Not Me.Button1Options.Visible
    End Sub
    Private Sub Button3_MouseHover(sender As Object, e As EventArgs) Handles Button3.MouseHover
        Me.Button3.Image = My.Resources.SelectBig
    End Sub
    Private Sub Button1Check1_Click(sender As Object, e As EventArgs) Handles Button1Check1.Click
        'Button 3 (Renamed from Button1) - Cross-Reference Step Check 1
        If myCheck1 = False Then
            Button1Check1.Image = My.Resources.CheckON
            Button1Check2.Image = My.Resources.CheckOFF
            Button1Check3.Image = My.Resources.CheckOFF
            Button1Check4.Image = My.Resources.CheckOFF
            Button1Check5.Image = My.Resources.CheckOFF
            Button1Check6.Image = My.Resources.CheckOFF
            myCheck = 1
            myCheck1 = True
            myCheck2 = False
            myCheck3 = False
            myCheck4 = False
            myCheck5 = False
            myCheck6 = False
            Me.Button1Options.Visible = False
            refreshXREF()
        ElseIf myCheck1 = True Then
            Button1Check1.Image = My.Resources.CheckOFF
            myCheck1 = False
            myCheck = 0
        End If
    End Sub
    Private Sub Button1Check2_Click(sender As Object, e As EventArgs) Handles Button1Check2.Click
        'Button 3 (Renamed from Button1) - Cross-Reference Section Check 2
        If myCheck2 = False Then
            Button1Check1.Image = My.Resources.CheckOFF
            Button1Check2.Image = My.Resources.CheckON
            Button1Check3.Image = My.Resources.CheckOFF
            Button1Check4.Image = My.Resources.CheckOFF
            Button1Check5.Image = My.Resources.CheckOFF
            Button1Check6.Image = My.Resources.CheckOFF
            myCheck = 2
            myCheck1 = False
            myCheck2 = True
            myCheck3 = False
            myCheck4 = False
            myCheck5 = False
            myCheck6 = False
            Me.Button1Options.Visible = False
            refreshXREF()
        ElseIf myCheck2 = True Then
            Button1Check2.Image = My.Resources.CheckOFF
            myCheck2 = False
            myCheck = 0
        End If
    End Sub

    Private Sub Button1Check3_Click(sender As Object, e As EventArgs) Handles Button1Check3.Click
        'Button 3 (Renamed from Button1) - Cross-Reference Figure Check 3
        If myCheck3 = False Then
            Button1Check1.Image = My.Resources.CheckOFF
            Button1Check2.Image = My.Resources.CheckOFF
            Button1Check3.Image = My.Resources.CheckON
            Button1Check4.Image = My.Resources.CheckOFF
            Button1Check5.Image = My.Resources.CheckOFF
            Button1Check6.Image = My.Resources.CheckOFF
            myCheck = 3
            myCheck1 = False
            myCheck2 = False
            myCheck3 = True
            myCheck4 = False
            myCheck5 = False
            myCheck6 = False
            Me.Button1Options.Visible = False
            refreshXREF()
        ElseIf myCheck3 = True Then
            Button1Check3.Image = My.Resources.CheckOFF
            myCheck3 = False
            myCheck = 0
        End If
    End Sub

    Private Sub Button1Check4_Click(sender As Object, e As EventArgs) Handles Button1Check4.Click
        'Button 3 (Renamed from Button1) - Cross-Reference Table Check 4
        If myCheck4 = False Then
            Button1Check1.Image = My.Resources.CheckOFF
            Button1Check2.Image = My.Resources.CheckOFF
            Button1Check3.Image = My.Resources.CheckOFF
            Button1Check4.Image = My.Resources.CheckON
            Button1Check5.Image = My.Resources.CheckOFF
            Button1Check6.Image = My.Resources.CheckOFF
            myCheck = 4
            myCheck1 = False
            myCheck2 = False
            myCheck3 = False
            myCheck4 = True
            myCheck5 = False
            myCheck6 = False
            Me.Button1Options.Visible = False
            refreshXREF()
        ElseIf myCheck4 = True Then
            Button1Check4.Image = My.Resources.CheckOFF
            myCheck4 = False
            myCheck = 0
        End If
    End Sub
    Private Sub Button1Check5_Click(sender As Object, e As EventArgs) Handles Button1Check5.Click
        'Button 3 (Renamed from Button1) - Caption Figure Check 5
        If myCheck5 = False Then
            Button1Check1.Image = My.Resources.CheckOFF
            Button1Check2.Image = My.Resources.CheckOFF
            Button1Check3.Image = My.Resources.CheckOFF
            Button1Check4.Image = My.Resources.CheckOFF
            Button1Check5.Image = My.Resources.CheckON
            Button1Check6.Image = My.Resources.CheckOFF
            myCheck = 5
            myCheck1 = False
            myCheck2 = False
            myCheck3 = False
            myCheck4 = False
            myCheck5 = True
            myCheck6 = False
            Me.Button1Options.Visible = False
            refreshXREF()
        ElseIf myCheck5 = True Then
            Button1Check5.Image = My.Resources.CheckOFF
            myCheck5 = False
            myCheck = 0
        End If
    End Sub
    Private Sub Button1Check6_Click(sender As Object, e As EventArgs) Handles Button1Check6.Click
        'Button 3 (Renamed from Button1) - Caption Table Check 6
        If myCheck6 = False Then
            Button1Check1.Image = My.Resources.CheckOFF
            Button1Check2.Image = My.Resources.CheckOFF
            Button1Check3.Image = My.Resources.CheckOFF
            Button1Check4.Image = My.Resources.CheckOFF
            Button1Check5.Image = My.Resources.CheckOFF
            Button1Check6.Image = My.Resources.CheckON
            myCheck = 6
            myCheck1 = False
            myCheck2 = False
            myCheck3 = False
            myCheck4 = False
            myCheck5 = False
            myCheck6 = True
            Me.Button1Options.Visible = False
            refreshXREF()
        ElseIf myCheck6 = True Then
            Button1Check6.Image = My.Resources.CheckOFF
            myCheck6 = False
            myCheck = 0
        End If
    End Sub

    ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' 
    'Button 4 Update All Cross-References
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        'Update All
        Dim myWord As Word.Application = Globals.ThisAddIn.Application
        Dim myDoc As Word.Document = myWord.ActiveDocument
        Dim mySelection As Word.Selection = myWord.Selection
        Dim rngStory As Word.Range = myDoc.Range
        Dim myTOC As Word.TableOfContents
        Dim lngJunk As Long
        Dim oShp As Word.Shape
        Dim myInfo As String = ""

        If myWord.Documents.Count = 0 Then
            myInfo = "You must first select a document to update!"
            Me.TextBoxInfo.Text = myInfo
            Exit Sub
        End If

        myWord.ScreenUpdating = False
        lngJunk = myDoc.Sections(1).Headers(1).Range.StoryType
        For Each rngStory In myDoc.StoryRanges
            Do
                Try
                    rngStory.Fields.Update()
                    Select Case rngStory.StoryType
                        Case 6, 7, 8, 9, 10, 11
                            If rngStory.ShapeRange.Count > 0 Then
                                For Each oShp In rngStory.ShapeRange
                                    If oShp.TextFrame.HasText Then
                                        oShp.TextFrame.TextRange.Fields.Update()
                                    End If
                                Next
                            End If
                    End Select
                Catch
                    myInfo = "Could not update! Action aborted! Something is weird about your document."
                    Me.TextBoxInfo.Text = myInfo
                    myWord.ScreenUpdating = True
                End Try
                rngStory = rngStory.NextStoryRange
            Loop Until rngStory Is Nothing
        Next

        For Each myTOC In myDoc.TablesOfContents
            myTOC.Update()
        Next myTOC

        myWord.ScreenUpdating = True

        myInfo = "All Document Fields Updated!"
        Me.TextBoxInfo.Text = myInfo
    End Sub
    Private Sub Button4_MouseEnter(sender As Object, e As EventArgs) Handles Button4.MouseEnter
        Me.Button4.Image = My.Resources.Refresh1BIGGER
    End Sub
    Private Sub Button4_MouseLeave(sender As Object, e As EventArgs) Handles Button4.MouseLeave
        Me.Button4.Image = My.Resources.Refresh1
    End Sub

    ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' 
    Private Sub ListBox1_MouseEnter(sender As Object, e As EventArgs) Handles ListBox1.MouseEnter
        Me.Button1Options.Visible = False
        Me.Button3.Image = My.Resources._Select
    End Sub

    Private Sub ListBoxBackground_MouseEnter(sender As Object, e As EventArgs) Handles ListBoxBackground.MouseEnter
        Me.Button1Options.Visible = False
        Me.Button3.Image = My.Resources._Select
    End Sub

    Sub InsertXREF()
        Dim myWord As Word.Application = Globals.ThisAddIn.Application
        Dim myDoc As Word.Document = myWord.ActiveDocument
        Dim mySelection As Word.Selection = myWord.Selection
        Dim myRange As Word.Range = myDoc.Range(Start:=mySelection.Paragraphs(1).Range.Start, End:=mySelection.Paragraphs(1).Range.End)

        'Insert X-Reference
        Select Case myCheck
            Case 1 ' Insert Cross-Reference to a Step
                With mySelection
                    If .End > .Start Then .Delete()
                    .InsertBefore("Step" & Chr(160))
                    .Collapse(Direction:=Word.WdCollapseDirection.wdCollapseEnd)
                    On Error Resume Next
                    .InsertCrossReference(ReferenceType:=Word.WdReferenceType.wdRefTypeHeading, ReferenceKind:=Word.WdReferenceKind.wdNumberNoContext, ReferenceItem:=(Me.ListBox1.SelectedIndex + 1))
                    If Err.Number = 4198 Then
                        With Me
                            .TextBoxInfo.Text = "You CANNOT insert a Cross-Reference To a heading that does Not have any text!"
                            .Refresh()
                        End With
                        Exit Sub
                    End If
                    On Error GoTo 0
                    .InsertAfter("")
                    If .Characters.Last.Text <> "" Then
                        .InsertBefore(" ")
                        .Collapse(Direction:=Word.WdCollapseDirection.wdCollapseEnd)
                        Dim myInfo As String
                        myInfo = "Cross-reference to the selected Step has been inserted."
                        Me.TextBoxInfo.Text = myInfo
                    End If
                End With
            Case 2 ' Insert Cross-Reference to a Section
                With mySelection
                    If .End > .Start Then .Delete()
                    .InsertBefore("Section" & Chr(160))
                    .Collapse(Direction:=Word.WdCollapseDirection.wdCollapseEnd)
                    On Error Resume Next
                    .InsertCrossReference(ReferenceType:=Word.WdReferenceType.wdRefTypeHeading, ReferenceKind:=Word.WdReferenceKind.wdNumberNoContext, ReferenceItem:=(Me.ListBox1.SelectedIndex + 1))
                    If Err.Number = 4198 Then
                        With Me
                            .TextBoxInfo.Text = "You CANNOT insert a Cross-Reference To a heading that does Not have any text!"
                            .Refresh()
                        End With
                        Exit Sub
                    End If
                    On Error GoTo 0
                    .InsertAfter("")
                    If .Characters.Last.Text <> "" Then
                        .InsertBefore(" ")
                        .Collapse(Direction:=Word.WdCollapseDirection.wdCollapseEnd)
                        Dim myInfo As String
                        myInfo = "Cross-reference to the selected Section has been inserted."
                        Me.TextBoxInfo.Text = myInfo
                    End If
                End With
            Case 3 ' Insert Cross-Reference to a Figure
                With mySelection
                    If .End > .Start Then .Delete()
                    .InsertBefore("(")
                    .Collapse(Direction:=Word.WdCollapseDirection.wdCollapseEnd)
                    On Error Resume Next
                    .InsertCrossReference(ReferenceType:=Word.WdCaptionLabelID.wdCaptionFigure, ReferenceKind:=Word.WdReferenceKind.wdOnlyLabelAndNumber, ReferenceItem:=(Me.ListBox1.SelectedIndex + 1))
                    If Err.Number = 4198 Then
                        With Me
                            .TextBoxInfo.Text = "You CANNOT insert a Cross-Reference To a Figure Caption that does Not have any text!"
                            .Refresh()
                        End With
                        Exit Sub
                    End If
                    On Error GoTo 0
                    .InsertAfter(")")
                    .Collapse(Direction:=Word.WdCollapseDirection.wdCollapseEnd)
                    Dim myInfo As String
                    myInfo = "Cross-reference to the selected Figure has been inserted."
                    Me.TextBoxInfo.Text = myInfo
                End With
            Case 4 ' Insert Cross-Reference to a Table
                With mySelection
                    If .End > .Start Then .Delete()
                    .InsertBefore("(")
                    .Collapse(Direction:=Word.WdCollapseDirection.wdCollapseEnd)
                    On Error Resume Next
                    .InsertCrossReference(ReferenceType:=Word.WdCaptionLabelID.wdCaptionTable, ReferenceKind:=Word.WdReferenceKind.wdOnlyLabelAndNumber, ReferenceItem:=(Me.ListBox1.SelectedIndex + 1))
                    If Err.Number = 4198 Then
                        With Me
                            .TextBoxInfo.Text = "You CANNOT insert a Cross-Reference To a Table Caption that does Not have any text!"
                            .Refresh()
                        End With
                        Exit Sub
                    End If
                    On Error GoTo 0
                    .InsertAfter(")")
                    .Collapse(Direction:=Word.WdCollapseDirection.wdCollapseEnd)
                    Dim myInfo As String
                    myInfo = "Cross-reference to the selected Table has been inserted."
                    Me.TextBoxInfo.Text = myInfo
                End With
            Case 5 ' Insert Caption for a Figure
                'Insert Caption
                With mySelection
                    If .End > .Start Then .Delete()
                    .Collapse(Direction:=Word.WdCollapseDirection.wdCollapseEnd)
                    .InsertCaption(Label:=Word.WdCaptionLabelID.wdCaptionFigure, ExcludeLabel:=0)
                    .Collapse(Direction:=Word.WdCollapseDirection.wdCollapseEnd)
                    With myRange.Find
                        .Text = Chr(32)
                        .Replacement.Text = Chr(160)
                        .Execute(Replace:=Word.WdReplace.wdReplaceOne)
                    End With
                    NCSFormat("FigureCaption")
                    Dim myInfo As String
                    myInfo = "A figure caption has been inserted."
                    Me.TextBoxInfo.Text = myInfo
                End With
            Case 6 ' Insert Caption for a Table
                'Insert Caption
                With mySelection
                    If .End > .Start Then .Delete()
                    .Collapse(Direction:=Word.WdCollapseDirection.wdCollapseEnd)
                    .InsertCaption(Label:=Word.WdCaptionLabelID.wdCaptionTable, ExcludeLabel:=0)
                    .Collapse(Direction:=Word.WdCollapseDirection.wdCollapseEnd)
                    With myRange.Find
                        .Text = Chr(32)
                        .Replacement.Text = Chr(160)
                        .Execute(Replace:=Word.WdReplace.wdReplaceOne)
                    End With
                    NCSFormat("TableCaption")
                    Dim myInfo As String
                    myInfo = "A table caption has been inserted."
                    Me.TextBoxInfo.Text = myInfo
                End With
        End Select
    End Sub

    Sub refreshXREF()
        Dim myWord As Word.Application = Globals.ThisAddIn.Application
        Dim myDoc As Word.Document = myWord.ActiveDocument
        Dim myHeadings
        Dim i As Integer

        Select Case myCheck
            Case 1
                myHeadings = myDoc.GetCrossReferenceItems(Word.WdReferenceType.wdRefTypeHeading)
                Me.ListBox1.Items.Clear()
                For i = 1 To UBound(myHeadings)
                    Me.ListBox1.Items.Add(myHeadings(i))
                Next i
                Dim myInfo As String
                If Me.ListBox1.Items.Count < 1 Then
                    myInfo = "Their are not any headings in this document!"
                    Me.TextBoxInfo.Text = myInfo
                Else
                    myInfo = "The Window is now full of document headings for you to choose from."
                    Me.TextBoxInfo.Text = myInfo
                End If
            Case 2
                myHeadings = myDoc.GetCrossReferenceItems(Word.WdReferenceType.wdRefTypeHeading)
                Me.ListBox1.Items.Clear()
                For i = 1 To UBound(myHeadings)
                    Me.ListBox1.Items.Add(myHeadings(i))
                Next i
                Dim myInfo As String
                If Me.ListBox1.Items.Count < 1 Then
                    myInfo = "Their are not any headings in this document!"
                    Me.TextBoxInfo.Text = myInfo
                Else
                    myInfo = "The Window is now full of document headings for you to choose from."
                    Me.TextBoxInfo.Text = myInfo
                End If
            Case 3
                myHeadings = myDoc.GetCrossReferenceItems(Word.WdCaptionLabelID.wdCaptionFigure)
                Me.ListBox1.Items.Clear()
                For i = 1 To UBound(myHeadings)
                    Me.ListBox1.Items.Add(myHeadings(i))
                Next i
                Dim myInfo As String
                If Me.ListBox1.Items.Count < 1 Then
                    myInfo = "Their are not any figures with captions in this document!"
                    Me.TextBoxInfo.Text = myInfo
                Else
                    myInfo = "The Window is now full of figures for you to choose from."
                    Me.TextBoxInfo.Text = myInfo
                End If
            Case 4
                myHeadings = myDoc.GetCrossReferenceItems(Word.WdCaptionLabelID.wdCaptionTable)
                Me.ListBox1.Items.Clear()
                For i = 1 To UBound(myHeadings)
                    Me.ListBox1.Items.Add(myHeadings(i))
                Next i
                Dim myInfo As String
                If Me.ListBox1.Items.Count < 1 Then
                    myInfo = "Their are not any tables with captions in this document!"
                    Me.TextBoxInfo.Text = myInfo
                Else
                    myInfo = "The Window is now full of tables for you to choose from."
                    Me.TextBoxInfo.Text = myInfo
                End If
            Case 5
                Me.ListBox1.Items.Clear()
                Dim myInfo As String
                myInfo = "Figure Caption selected. Press the Insert button to insert a figure caption at the cursor position in the document."
                Me.TextBoxInfo.Text = myInfo
            Case 6
                Me.ListBox1.Items.Clear()
                Dim myInfo As String
                myInfo = "Table Caption selected. Press the Insert button to insert a table caption at the cursor position in the document."
                Me.TextBoxInfo.Text = myInfo
        End Select
        Me.Refresh()
    End Sub
End Class