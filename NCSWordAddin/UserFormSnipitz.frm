'Option Explicit On



'Private Sub ToggleButton1_Click()
'    With Me
'            .Caption = "X-Ref"
'            snipText = "ecks ref activated."
'            SpeakText snipText
'    End With
'End Sub

'Private Sub Label6_Click()
'Dim myBool As Boolean

'    With UserFormSnipitz.CheckBox1
'    If .Value = False Then ' If Enable Audio is unchecked
'        myBool = False
'        .Value = True
'    Else ' If Enable Audio is checked
'        myBool = True
'    End If

'    snipText = "Give me a Jay,... Oh,... Ess,... E,... Pee,... Aitch!.. Whats that spell?.. Jayohesepeaitch or something!"
'    SpeakText snipText

'    If myBool = False Then ' If Enable Audio was initially Unchecked
'        .Value = False ' Uncheck it (return to user setting)
'    End If
'End With

'End Sub

'Private Sub XRefLabel2_Click()

'Dim myBool As Boolean

'With UserFormSnipitz.CheckBox1
'    If .Value = False Then ' If Enable Audio is unchecked
'        myBool = False
'        .Value = True
'    Else ' If Enable Audio is checked
'        myBool = True
'    End If

'    snipText = "Joseph... Joseph... He's our man! If he can't do it... Get someone else!"
'    SpeakText snipText

'    If myBool = False Then ' If Enable Audio was initially Unchecked
'        .Value = False ' Uncheck it (return to user setting)
'    End If
'End With

'End Sub

''X-Ref Buttons

'Private Sub XRefButton2_Click()
'Dim i As Integer

''Insert X-Reference
'With Me
'    If .XRefOption1.Value = True Then ' Step
'        For i = 0 To .XRefListBox1.ListCount
'            If .XRefListBox1.Selected(i) = True Then
'                With Selection
'                    If .End > .Start Then .Delete
'                    .InsertBefore "Step" & VBA.Chr$(160)
'                    .Collapse Direction:=wdCollapseEnd
'                    On Error Resume Next
'                    .InsertCrossReference ReferenceType:=wdRefTypeHeading, _
'                        ReferenceKind:=wdNumberNoContext, ReferenceItem:=(i + 1)
'                    If Err.Number = 4198 Then
'                        With Me
'                            .XRefListBox1.Visible = False
'                            With .XRefTextBox2
'                                .Visible = True
'                                .Value = "You CANNOT insert a Cross-Reference to a heading that does not have any text!"
'                                .Height = 218
'                            End With
'                            .Repaint
'                        End With
'                        Exit Sub
'                    End If
'                    On Error GoTo 0
'                    If .Characters.Last <> "" Then
'                        .InsertBefore " "
'                        .Collapse Direction:=wdCollapseEnd
'                        snipText = "Step Reference Inserted."
'                        SpeakText snipText
'                    End If
'                End With
'                Exit For
'            End If
'        Next i
'    ElseIf .XRefOption2.Value = True Then ' Section
'        For i = 0 To .XRefListBox1.ListCount
'            If .XRefListBox1.Selected(i) = True Then
'                With Selection
'                    If .End > .Start Then .Delete
'                    .InsertBefore "Section" & VBA.Chr$(160)
'                    .Collapse Direction:=wdCollapseEnd
'                    On Error Resume Next
'                    .InsertCrossReference ReferenceType:=wdRefTypeHeading, _
'                        ReferenceKind:=wdNumberNoContext, ReferenceItem:=(i + 1)
'                    If Err.Number = 4198 Then
'                        With Me
'                            .XRefListBox1.Visible = False
'                            With .XRefTextBox2
'                                .Visible = True
'                                .Value = "You CANNOT insert a Cross-Reference to a heading that does not have any text!"
'                                .Height = 218
'                            End With
'                            .Repaint
'                        End With
'                        Exit Sub
'                    End If
'                    On Error GoTo 0
'                    If .Characters.Last <> "" Then
'                        .InsertBefore " "
'                        .Collapse Direction:=wdCollapseEnd
'                        snipText = "Section Reference Inserted."
'                        SpeakText snipText
'                    End If
'                End With
'                Exit For
'            End If
'        Next i
'    ElseIf .XRefOption3.Value = True Then ' Figure
'        For i = 0 To .XRefListBox1.ListCount
'            If XRefListBox1.Selected(i) = True Then
'                With Selection
'                    If .End > .Start Then .Delete
'                    .InsertBefore "("
'                    .Collapse Direction:=wdCollapseEnd
'                    .InsertCrossReference ReferenceType:=wdCaptionFigure, _
'                        ReferenceKind:=wdOnlyLabelAndNumber, ReferenceItem:=(i + 1)
'                    .InsertAfter ")"
'                    .Collapse Direction:=wdCollapseEnd
'                    If .Characters.Last <> "" Then
'                        .InsertBefore " "
'                        .Collapse Direction:=wdCollapseEnd
'                        snipText = "Figure Reference Inserted."
'                        SpeakText snipText
'                    End If
'                End With
'                Exit For
'            End If
'        Next i
'    ElseIf .XRefOption4.Value = True Then ' Table
'        For i = 0 To .XRefListBox1.ListCount
'            If .XRefListBox1.Selected(i) = True Then
'                With Selection
'                    If .End > .Start Then .Delete
'                    .InsertBefore "("
'                    .Collapse Direction:=wdCollapseEnd
'                    .InsertCrossReference ReferenceType:=wdCaptionTable, _
'                        ReferenceKind:=wdOnlyLabelAndNumber, ReferenceItem:=(i + 1)
'                    .InsertAfter ")"
'                    .Collapse Direction:=wdCollapseEnd
'                    If .Characters.Last <> "" Then
'                        .InsertBefore " "
'                        .Collapse Direction:=wdCollapseEnd
'                        snipText = "Table Reference Inserted."
'                        SpeakText snipText
'                    End If
'                End With
'                Exit For
'            End If
'        Next i
'    End If
'End With

'Dim wdApp As Word.Application
'Dim winname As String
'Set wdApp = GetObject(, "Word.Application")
'winname = (wdApp.Windows(1).Caption)
'AppActivate winname

'End Sub

'Private Sub XRefButton4_Click()
'Dim i As Integer, pos, pos2
'Dim myRange As Range
''Insert Caption
'With Me
'    If .XRefOption5.Value = True Then ' Insert Figure Caption
'        With Selection
'            If .End > .Start Then .Delete
'            .Collapse Direction:=wdCollapseEnd
'            .InsertCaption Label:="Figure", Title:=": ", _
'                Position:=wdCaptionPositionBelow, ExcludeLabel:=0
'            If .Characters.Last <> "" Then
'                .InsertBefore " "
'                .Collapse Direction:=wdCollapseEnd
'            End If
'            pos = .Paragraphs(1).Range.Start
'            pos2 = .Paragraphs(1).Range.End
'            Set myRange = ActiveDocument.Range(Start:=pos, End:=pos2)
'            With myRange.Find
'                .Text = VBA.Chr(32)
'                .Replacement.Text = VBA.Chr(160)
'                .Execute Replace:=wdReplaceOne
'            End With
'            .Style = ActiveDocument.Styles("FigureCaption")
'            snipText = "Figure Caption Inserted."
'            SpeakText snipText
'        End With
'    ElseIf .XRefOption6.Value = True Then ' Insert Table Caption
'        With Selection
'            If .End > .Start Then .Delete
'            .Collapse Direction:=wdCollapseEnd
'            .InsertCaption Label:="Table", TitleAutoText:="InsertCaption1", _
'                Title:=": ", Position:=wdCaptionPositionBelow, ExcludeLabel:=0
'            If .Characters.Last <> "" Then
'                .InsertBefore " "
'                .Collapse Direction:=wdCollapseEnd
'            End If
'            pos = .Paragraphs(1).Range.Start
'            pos2 = .Paragraphs(1).Range.End
'            Set myRange = ActiveDocument.Range(Start:=pos, End:=pos2)
'            With myRange.Find
'                .Text = VBA.Chr(32)
'                .Replacement.Text = VBA.Chr(160)
'                .Execute Replace:=wdReplaceOne
'            End With
'            .Style = ActiveDocument.Styles("TableCaption")
'            snipText = "Table Caption Inserted."
'            SpeakText snipText
'        End With
'    End If
'End With

'Dim wdApp As Word.Application
'Dim winname As String
'Set wdApp = GetObject(, "Word.Application")
'winname = (wdApp.Windows(1).Caption)
'AppActivate winname

'End Sub

'Private Sub XRefOption1_Click()
'    Dim myHeadings As Variant
'    myHeadings = ActiveDocument.GetCrossReferenceItems(wdRefTypeHeading)
'    With Me
'        .XRefOption1.Value = True
'        .XRefOption2.Value = False
'        .XRefOption3.Value = False
'        .XRefOption4.Value = False
'        .XRefTextBox2.Visible = False
'        .XRefListBox1.Visible = True
'        .XRefListBox1.List = myHeadings
'        snipText = "Steps Selected."
'        SpeakText snipText
'        If .XRefListBox1.ListCount < 1 Then
'            .XRefListBox1.AddItem ("No Headings in the current document.")
'        End If
'        .XRefListBox1.Height = 218
'        .Repaint
'    End With
'End Sub

'Private Sub XRefOption2_Click()
'    Dim myHeadings As Variant
'    myHeadings = ActiveDocument.GetCrossReferenceItems(wdRefTypeHeading)
'    With Me
'        .XRefOption1.Value = False
'        .XRefOption2.Value = True
'        .XRefOption3.Value = False
'        .XRefOption4.Value = False
'        .XRefTextBox2.Visible = False
'        .XRefListBox1.Visible = True
'        .XRefListBox1.List = myHeadings
'        snipText = "Sections Selected."
'        SpeakText snipText
'        If .XRefListBox1.ListCount < 1 Then
'            .XRefTextBox2.Visible = True
'            .XRefTextBox2.Value = "No Headings in the current document."
'            .XRefListBox1.Visible = False
'        End If
'        .XRefListBox1.Height = 218
'        .Repaint
'    End With
'End Sub

'Private Sub XRefOption3_Click()
'    Dim myHeadings As Variant
'    myHeadings = ActiveDocument.GetCrossReferenceItems(wdCaptionFigure)
'    With Me
'        .XRefOption1.Value = False
'        .XRefOption2.Value = False
'        .XRefOption3.Value = True
'        .XRefOption4.Value = False
'        .XRefTextBox2.Visible = False
'        .XRefListBox1.Visible = True
'        .XRefListBox1.List = myHeadings
'        snipText = "Figures Selected."
'        SpeakText snipText
'        If .XRefListBox1.ListCount < 1 Then
'            .XRefTextBox2.Visible = True
'            .XRefTextBox2.Value = "No Headings in the current document."
'            .XRefListBox1.Visible = False
'        End If
'        .XRefListBox1.Height = 218
'        .Repaint
'    End With
'End Sub

'Private Sub XRefOption4_Click()
'    Dim myHeadings As Variant
'    myHeadings = ActiveDocument.GetCrossReferenceItems(wdCaptionTable)
'    With Me
'        .XRefOption1.Value = False
'        .XRefOption2.Value = False
'        .XRefOption3.Value = False
'        .XRefOption4.Value = True
'        .XRefTextBox2.Visible = False
'        .XRefListBox1.Visible = True
'        .XRefListBox1.List = myHeadings
'        snipText = "Tables Selected."
'        SpeakText snipText
'        If .XRefListBox1.ListCount < 1 Then
'            .XRefTextBox2.Visible = True
'            .XRefTextBox2.Value = "No Headings in the current document."
'            .XRefListBox1.Visible = False
'        End If
'        .XRefListBox1.Height = 218
'        .Repaint
'    End With
'End Sub

'Private Sub XRefOption5_Click()
'    With Me
'        .XRefOption5.Value = True
'        .XRefOption6.Value = False
'        .XRefTextBox2.Visible = False
'        .XRefListBox1.Visible = True
'        snipText = "Figure Selected."
'        SpeakText snipText
'        .Repaint
'    End With
'End Sub

'Private Sub XRefOption6_Click()
'    With Me
'        .XRefOption6.Value = True
'        .XRefOption5.Value = False
'        .XRefTextBox2.Visible = False
'        .XRefListBox1.Visible = True
'        snipText = "Table Selected."
'        SpeakText snipText
'        .Repaint
'    End With
'End Sub

'Private Sub XRefButton5_Click()
''Find Next
'Dim objFld As Field
'Dim docRange As Range, selRange As Range, newRange As Range
'Dim boRef1 As Boolean, boRef2 As Boolean

'boRef1 = False
'boRef2 = False

'Set docRange = ActiveDocument.Range ' Set the whole document to a range
'Set selRange = Selection.Range ' Set the current selection to a range
'Set newRange = docRange ' Create a new range identical to the range of the whole document

'newRange.SetRange Start:=selRange.End, End:=docRange.End ' Redefine the new range to start at the end of the selection range

'' Loop through fields in the ActiveDocument
'For Each objFld In newRange.Fields
'    ' If the field is a cross-ref, do something to it.
'    If objFld.Type = wdFieldRef Then
'        objFld.Select ' Select the Cross-Reference
'        Set selRange = Selection.Range ' Redefine selRange to the selected Cross-Reference
'        Selection.Collapse Direction:=wdCollapseStart ' Collapse the selection
'        Selection.MoveLeft unit:=wdWord, Count:=1, Extend:=wdExtend ' Move the selection left by a word
'        selRange.SetRange Start:=Selection.Range.Start, End:=selRange.End ' Redefine the start of selRange to the start of the selection
'        selRange.Select ' Select the selRange
'        boRef1 = True
'        Exit For
'    End If
'Next objFld

'Dim myHeadings As Variant, i%, selArray, headArray, evalArray(0 To 1)

'myHeadings = ActiveDocument.GetCrossReferenceItems(wdRefTypeHeading)
'Set selRange = Selection.Range
'If boRef1 = True Then ' Found a Cross-Reference
'    For i = 1 To UBound(myHeadings)
'        headArray = Split(Trim(myHeadings(i)), VBA.Chr(32))
'        selArray = Split(Trim(selRange.Text), VBA.Chr(160))
'        evalArray(0) = selRange.Text
'        evalArray(1) = myHeadings(i)
'        If headArray(0) = selArray(UBound(selArray)) Then
'            boRef2 = True
'            Exit For
'        End If
'    Next i
'    If boRef2 = True Then ' Found a Cross-Reference that matches a Heading
'        With Me
'            .XRefOption1.Value = False
'            .XRefOption2.Value = False
'            .XRefOption3.Value = False
'            .XRefOption4.Value = False
'            With .XRefTextBox2
'                .Visible = True
'                .Value = evalArray(0) & "  is linked to " & evalArray(1)
'                snipText = "'" & evalArray(0) & "' links to:" & evalArray(1) & "'"
'                SpeakText snipText
'                .Height = 218
'            End With
'            With .XRefListBox1
'                .Height = 218
'                .Visible = False
'            End With
'            .Repaint
'        End With
'    Else 'Found a Cross-Reference, but it does not have a matching Heading
'        evalArray(0) = selRange.Text
'        evalArray(1) = " is a broken Cross-Reference!"
'        With Me
'            .XRefOption1.Value = False
'            .XRefOption2.Value = False
'            .XRefOption3.Value = False
'            .XRefOption4.Value = False
'            With .XRefTextBox2
'                .Visible = True
'                .Value = evalArray(0) & VBA.Chr(13) & evalArray(1)
'                snipText = "'" & evalArray(0) & VBA.Chr(13) & evalArray(1) & "'"
'                SpeakText snipText
'                .Height = 218
'            End With
'            With .XRefListBox1
'                .Height = 218
'                .Visible = False
'            End With
'            .Repaint
'        End With
'    End If
'Else ' Did NOT find a Cross-Reference
'    evalArray(0) = "Reached the end of the document."
'    evalArray(1) = "Restarting from the beginning."
'    With Me
'        .XRefOption1.Value = False
'        .XRefOption2.Value = False
'        .XRefOption3.Value = False
'        .XRefOption4.Value = False
'        With .XRefTextBox2
'            .Visible = True
'            .Value = evalArray(0) & "  " & evalArray(1)
'            snipText = evalArray(0) & "  " & evalArray(1)
'            SpeakText snipText
'            .Height = 218
'        End With
'        With .XRefListBox1
'            .Height = 218
'            .Visible = False
'        End With
'        .Repaint
'    End With
'    Set selRange = ActiveDocument.Range ' Redefine selRange to the entire document
'    selRange.Collapse Direction:=wdCollapseStart ' Collapses the selection so that the selection is at the start of the document
'    selRange.Select
'End If

'Dim wdApp As Word.Application
'Dim winname As String
'Set wdApp = GetObject(, "Word.Application")
'winname = (wdApp.Windows(1).Caption)
'AppActivate winname

'End Sub


'Private Sub XRefButton7_Click()
''Update All
' Dim rngStory As Word.Range, myTOC As TableOfContents
' Dim lngJunk As Long
' Dim oShp As Shape

'    If Documents.Count = 0 Then Exit Sub
'    Application.ScreenUpdating = False
'    lngJunk = ActiveDocument.Sections(1).Headers(1).Range.StoryType
'    For Each rngStory In ActiveDocument.StoryRanges
'        Do
'            On Error Resume Next
'            rngStory.Fields.Update
'            Select Case rngStory.StoryType
'            Case 6, 7, 8, 9, 10, 11
'                If rngStory.ShapeRange.Count > 0 Then
'                    For Each oShp In rngStory.ShapeRange
'                        If oShp.TextFrame.HasText Then
'                            oShp.TextFrame.TextRange.Fields.Update
'                        End If
'                    Next
'                End If
'            Case Else
'                'do nothing
'            End Select
'            On Error GoTo 0
'            Set rngStory = rngStory.NextStoryRange
'        Loop Until rngStory Is Nothing
'    Next

'    For Each myTOC In ActiveDocument.TablesOfContents
'        myTOC.Update
'    Next myTOC

'    With Me
'        .XRefOption1.Value = False
'        .XRefOption2.Value = False
'        .XRefOption3.Value = False
'        .XRefOption4.Value = False
'        .XRefListBox1.Clear
'        .XRefListBox1.AddItem ("All Document Fields Updated!")
'        .XRefListBox1.Height = 218
'        .Repaint
'    End With

'    StatusBar = "All Document Fields Updated!"
'    Application.ScreenUpdating = True

'Dim wdApp As Word.Application
'Dim winname As String
'Set wdApp = GetObject(, "Word.Application")
'winname = (wdApp.Windows(1).Caption)
'AppActivate winname

'End Sub

'Private Sub XRefButton8_Click()
''Update One
'Selection.Fields.Update
'With Me
'    .XRefOption1.Value = False
'    .XRefOption2.Value = False
'    .XRefOption3.Value = False
'    .XRefOption4.Value = False
'    .XRefListBox1.Clear
'    .XRefListBox1.AddItem ("The Selected Document Field is Updated!")
'    .XRefListBox1.Height = 218
'    .Repaint
'End With

'StatusBar = "The Selected Document Field is Updated!"

'Dim wdApp As Word.Application
'Dim winname As String
'Set wdApp = GetObject(, "Word.Application")
'winname = (wdApp.Windows(1).Caption)
'AppActivate winname

'End Sub
