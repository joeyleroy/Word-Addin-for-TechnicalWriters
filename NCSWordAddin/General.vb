Module General

    Sub NCSFormat(ByVal myStyle As String)
        Dim myWord As Word.Application = Globals.ThisAddIn.Application
        Dim myDoc As Word.Document = myWord.ActiveDocument
        Dim mySelection As Word.Selection = myWord.Selection
        Dim selRange As Word.Range = mySelection.Range

        selRange.SetRange(Start:=selRange.Paragraphs(1).Range.Start,
                          End:=selRange.Paragraphs(selRange.Paragraphs.Count).Range.End)
        selRange.Select()
        mySelection.ClearFormatting()
        mySelection.ClearCharacterDirectFormatting()
        mySelection.ClearParagraphDirectFormatting()
        mySelection.ClearParagraphAllFormatting()
        selRange.HighlightColorIndex = Word.WdColorIndex.wdNoHighlight

        ' IMPROVE Record start and end points of user selection and recall selection after format.
        Select Case myStyle
            Case = "Heading 1"
                selRange.Style = myDoc.Styles(Word.WdBuiltinStyle.wdStyleHeading1)
                myWord.StatusBar = "Heading 1 style applied to the selection."
                mySelection.Collapse(Direction:=Word.WdCollapseDirection.wdCollapseStart)
            Case = "Heading 2"
                selRange.Style = myDoc.Styles(Word.WdBuiltinStyle.wdStyleHeading2)
                myWord.StatusBar = "Heading 2 style applied to the selection."
                mySelection.Collapse(Direction:=Word.WdCollapseDirection.wdCollapseStart)
            Case = "Heading 3"
                selRange.Style = myDoc.Styles(Word.WdBuiltinStyle.wdStyleHeading3)
                myWord.StatusBar = "Heading 3 style applied to the selection."
                mySelection.Collapse(Direction:=Word.WdCollapseDirection.wdCollapseStart)
            Case = "Heading 4"
                selRange.Style = myDoc.Styles(Word.WdBuiltinStyle.wdStyleHeading4)
                myWord.StatusBar = "Heading 4 style applied to the selection."
                mySelection.Collapse(Direction:=Word.WdCollapseDirection.wdCollapseStart)
            Case = "Body Text", "Body Text 1"
                selRange.Style = myDoc.Styles(Word.WdBuiltinStyle.wdStyleBodyText)
                myWord.StatusBar = "Body Text 1 style applied to the selection."
                mySelection.Collapse(Direction:=Word.WdCollapseDirection.wdCollapseStart)
            Case = "Body Text 2"
                selRange.Style = myDoc.Styles(Word.WdBuiltinStyle.wdStyleBodyText2)
                myWord.StatusBar = "Body Text 2 style applied to the selection."
                mySelection.Collapse(Direction:=Word.WdCollapseDirection.wdCollapseStart)
            Case = "Body Text 3"
                selRange.Style = myDoc.Styles(Word.WdBuiltinStyle.wdStyleBodyText3)
                myWord.StatusBar = "Body Text 3 style applied to the selection."
                mySelection.Collapse(Direction:=Word.WdCollapseDirection.wdCollapseStart)
            Case = "Table Spacer"
                Dim myStyleTest As Word.Style
                Try
                    selRange.Style = myDoc.Styles("Table Spacer")
                Catch
                    myStyleTest = myDoc.Styles.Add("Table Spacer")
                    With myStyleTest
                        With .ParagraphFormat
                            .Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft
                            .LeftIndent = myWord.InchesToPoints(0.3)
                            .SpaceAfter = myWord.PointsToInches(0)
                            .SpaceBefore = myWord.PointsToInches(0)
                            .LineSpacingRule = Word.WdLineSpacing.wdLineSpaceSingle
                        End With
                        With .Font
                            .Bold = False
                            .Italic = False
                            .Name = "Calibri"
                            .Size = 6
                        End With
                    End With
                    selRange.Style = myDoc.Styles("Table Spacer")
                End Try
                myWord.StatusBar = "Table Space style applied to the selection."
                mySelection.Collapse(Direction:=Word.WdCollapseDirection.wdCollapseStart)
            Case = "FigureFormat"
                'Verify Style Name
                Dim myStyleTest As Word.Style
                Try
                    selRange.Style = myDoc.Styles("FigureFormat")
                Catch
                    myStyleTest = myDoc.Styles.Add("FigureFormat")
                    '''''''''''''''''''''''''''''''''''''''CONFIGURE for Figure Format
                    With myStyleTest
                        With .ParagraphFormat
                            .Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                            .LeftIndent = myWord.InchesToPoints(0)
                            .SpaceAfter = myWord.PointsToInches(0)
                            .SpaceBefore = myWord.PointsToInches(12)
                            .LineSpacingRule = Word.WdLineSpacing.wdLineSpaceSingle
                        End With
                        With .Font
                            .Bold = False
                            .Italic = False
                            .Name = "Calibri"
                            .Size = 72
                            .Color = Word.WdColor.wdColorRed
                        End With
                    End With
                    selRange.Style = myDoc.Styles("FigureFormat")
                End Try
                myWord.StatusBar = "Figure Format style applied to the selection."
                mySelection.Collapse(Direction:=Word.WdCollapseDirection.wdCollapseStart)
            Case = "FigureCaption"
                Dim myStyleTest As Word.Style
                Try
                    selRange.Style = myDoc.Styles("FigureCaption")
                Catch
                    myStyleTest = myDoc.Styles.Add("FigureCaption")
                    With myStyleTest
                        With .ParagraphFormat
                            .Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                            .LeftIndent = myWord.InchesToPoints(0)
                            .SpaceAfter = myWord.PointsToInches(3)
                            .SpaceBefore = myWord.PointsToInches(3)
                            .LineSpacingRule = Word.WdLineSpacing.wdLineSpaceSingle
                        End With
                        With .Font
                            .Bold = True
                            .Italic = False
                            .Name = "Calibri"
                            .Size = 11
                        End With
                    End With
                    selRange.Style = myDoc.Styles("FigureCaption")
                End Try
                myWord.StatusBar = "Figure Caption style applied to the selection."
                mySelection.Collapse(Direction:=Word.WdCollapseDirection.wdCollapseStart)
            Case = "TableCaption"
                Dim myStyleTest As Word.Style
                Try
                    selRange.Style = myDoc.Styles("TableCaption")
                Catch
                    myStyleTest = myDoc.Styles.Add("TableCaption")
                    With myStyleTest
                        With .ParagraphFormat
                            .Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft
                            .LeftIndent = myWord.InchesToPoints(0.3)
                            .SpaceAfter = myWord.PointsToInches(0)
                            .SpaceBefore = myWord.PointsToInches(12)
                            .LineSpacingRule = Word.WdLineSpacing.wdLineSpaceSingle
                        End With
                        With .Font
                            .Bold = True
                            .Italic = False
                            .Name = "Calibri"
                            .Size = 11
                        End With
                    End With
                    selRange.Style = myDoc.Styles("TableCaption")
                End Try
                myWord.StatusBar = "Table Caption style applied to the selection."
                mySelection.Collapse(Direction:=Word.WdCollapseDirection.wdCollapseStart)
            Case Else
                'Do Nothing
        End Select
    End Sub

    Sub PageBreakBefore()
        Dim myWord As Word.Application = Globals.ThisAddIn.Application
        Dim mySelection As Word.Selection = myWord.Selection

        If mySelection.ParagraphFormat.PageBreakBefore = True Then
            mySelection.ParagraphFormat.PageBreakBefore = False
            myWord.StatusBar = "Page Break Before: Off"
        ElseIf mySelection.ParagraphFormat.PageBreakBefore = False Then
            mySelection.ParagraphFormat.PageBreakBefore = True
            myWord.StatusBar = "Page Break Before: On"
        Else
            'Do Nothing
        End If
    End Sub

    Sub KeepWithNext()
        Dim myWord As Word.Application = Globals.ThisAddIn.Application
        Dim mySelection As Word.Selection = myWord.Selection

        If mySelection.ParagraphFormat.KeepWithNext = True Then
            mySelection.ParagraphFormat.KeepWithNext = False
            myWord.StatusBar = "Keep with Next: Off"
        ElseIf mySelection.ParagraphFormat.KeepWithNext = False Then
            mySelection.ParagraphFormat.KeepWithNext = True
            myWord.StatusBar = "Keep with Next: ON"
        Else
            'Do Nothing
        End If
    End Sub

    Sub ToggleDocProps()
        Dim myWord As Word.Application = Globals.ThisAddIn.Application
        Try
            myWord.DisplayDocumentInformationPanel = Not myWord.DisplayDocumentInformationPanel
        Catch
            MsgBox("Something broke... Action aborted!")
        End Try
    End Sub

    Sub updateHeadingStyles()
        Dim myWord As Word.Application = Globals.ThisAddIn.Application
        Dim myDoc As Word.Document = myWord.ActiveDocument
        Dim mySelection As Word.Selection = myWord.Selection

        Try
            If MsgBox("This command will affect the entire document and could cause a loss of information. Continue anyway?", vbOKCancel) = vbOK Then
                For Each para As Word.Paragraph In myDoc.Paragraphs
                    Dim myHeading As Word.Style = para.Style
                    Select Case myHeading.NameLocal
                        Case = "Heading 1"
                            para.Range.Select()
                            NCSFormat("Heading 1")
                        Case = "Heading 2"
                            para.Range.Select()
                            NCSFormat("Heading 2")
                        Case = "Heading 3"
                            para.Range.Select()
                            NCSFormat("Heading 3")
                        Case = "Heading 4"
                            para.Range.Select()
                            NCSFormat("Heading 4")
                        Case = "Heading 5"
                            para.Range.Select()
                            NCSFormat("Heading 5")
                        Case Else
                    End Select
                    myHeading = Nothing
                Next para
                mySelection.Select()
            End If
        Catch
            MsgBox("Something broke... Action aborted!")
        End Try
    End Sub

    Sub InsertGraphic()
        Dim myWord As Word.Application = Globals.ThisAddIn.Application
        Dim myDoc As Word.Document = myWord.ActiveDocument
        Dim mySelection As Word.Selection = myWord.Selection
        Dim selRange As Word.Range = mySelection.Range
        Dim myChar As Long
        Dim intAnswer As Integer

        Try
            With mySelection
                ''''Test if Selection is within a table
                'If selection is not in a table, continue Sub
                'Otherwise, selection is inside of a table prompt with OK/Cancel
                'If OK, Move selection down until not in a table
                'If Cancel, abort Sub
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                If .Information(Word.WdInformation.wdWithInTable) = True Then
                    intAnswer = MsgBox("The requested Action is not allowed inside of a table." &
                    Chr(13) & "Click OK to Perform the requested Action below the selected table." &
                    Chr(13) & "Click Cancel to abort Action request.", vbOKCancel, "Test")
                    If intAnswer = 2 Then
                        Exit Sub
                    Else
                        myWord.ScreenUpdating = False
                        Do Until .Information(Word.WdInformation.wdWithInTable) = False
                            .MoveDown()
                        Loop
                        .Paragraphs(1).Range.Select()
                        If .Style IsNot "Table Spacer" Then
                            .Collapse(Word.WdCollapseDirection.wdCollapseEnd)
                            .Text = Chr(13)
                            NCSFormat("Table Spacer")
                            .Paragraphs(1).Range.Select()
                        End If
                        If .Characters.Last IsNot "" Then
                            .Collapse(Word.WdCollapseDirection.wdCollapseEnd)
                            .Text = Chr(13)
                            .Collapse(Word.WdCollapseDirection.wdCollapseEnd)
                        End If
                    End If
                End If

                ''''Test if selection is at the end of the document.
                ''''by moving to the end of the selection and counting moved units and stores into myChar
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                myChar = .EndOf(Unit:=Word.WdUnits.wdParagraph, Extend:=Word.WdMovementType.wdMove)
                If myChar = 0 Then
                    If .Range.Bookmarks.Exists("\EndOfDoc") = True Then
                        .TypeParagraph()
                    End If
                Else
                    .Text = Chr(13)
                End If
                selRange = mySelection.Range

                ''''Action Stage
                '''''''''''''''''''
                NCSFormat("FigureFormat")
                .TypeText(Text:="G")

                'Caption
                .TypeParagraph()
                .InsertCaption(Label:="Figure", TitleAutoText:="InsertCaption1",
                           Title:=": ", Position:=Word.WdCaptionPosition.wdCaptionPositionBelow, ExcludeLabel:=0)

                'Select the whole line
                .Expand(Word.WdUnits.wdLine)
                selRange = mySelection.Range
                selRange.Select()
                'Find and replace the space [Chr(32)] with a non-breaking space [Chr(160)]
                With .Find
                    .Text = Chr(32)
                    .Replacement.Text = Chr(160)
                    .Execute(Replace:=Word.WdReplace.wdReplaceAll)
                    .Text = ":"
                    .Replacement.Text = ": "
                    .Execute(Replace:=Word.WdReplace.wdReplaceAll)
                End With

                NCSFormat("FigureCaption")

                selRange.Start = selRange.End - 1
                selRange.End = selRange.End - 1

                selRange.Select()

                selRange.Collapse(Direction:=Word.WdCollapseDirection.wdCollapseEnd)
                selRange.Fields.Add(Range:=mySelection.Range, Type:=Word.WdFieldType.wdFieldEmpty,
                                Text:="MACROBUTTON  NoMacro Click Here to Input Graphic Caption Text.",
                                PreserveFormatting:=False)
                selRange.Select()
            End With
        Catch
            MsgBox("Something broke... Action aborted!")
        End Try
        myWord.ScreenUpdating = True
    End Sub

    Sub NCSUpdateFields()
        Dim myWord As Word.Application = Globals.ThisAddIn.Application
        Dim myDoc As Word.Document = myWord.ActiveDocument
        Dim mySelection As Word.Selection = myWord.Selection

        Dim rngStory As Word.Range
        Dim myTOC As Word.TableOfContents
        Dim lngJunk As Long
        Dim oShp As Word.Shape

        Try
            If myWord.Documents.Count = 0 Then Exit Sub
            myWord.ScreenUpdating = False
            lngJunk = myDoc.Sections(1).Headers(1).Range.StoryType
            For Each rngStory In myDoc.StoryRanges
                Do
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
                        Case Else
                            'do nothing
                    End Select
                    rngStory = rngStory.NextStoryRange
                Loop Until rngStory Is Nothing
            Next
            For Each myTOC In myDoc.TablesOfContents
                myTOC.Update()
            Next myTOC
            myWord.StatusBar = "All Document Fields Updated!"
        Catch
            MsgBox("Something broke... Action aborted!")
        End Try
        myWord.ScreenUpdating = True
    End Sub

    Sub FormatNCSNotes()
        Dim myWord As Word.Application = Globals.ThisAddIn.Application
        Dim mySelection As Word.Selection = myWord.Selection
        Dim tbl As Word.Table

        Try
            If mySelection.Information(Word.WdInformation.wdWithInTable) = True Then
                tbl = mySelection.Tables(1)
                With tbl
                    .Rows.LeftIndent = myWord.InchesToPoints(0.75)
                    .AutoFitBehavior(Word.WdAutoFitBehavior.wdAutoFitFixed)
                    .LeftPadding = myWord.InchesToPoints(0.03)
                    .RightPadding = myWord.InchesToPoints(0.03)
                    .TopPadding = myWord.InchesToPoints(0.03)
                    .BottomPadding = myWord.InchesToPoints(0.03)
                    With .Columns(1)
                        .Width = myWord.InchesToPoints(0.5)
                        .Borders(Word.WdBorderType.wdBorderLeft).LineStyle = Word.WdLineStyle.wdLineStyleNone
                        .Borders(Word.WdBorderType.wdBorderRight).LineStyle = Word.WdLineStyle.wdLineStyleNone
                        .Borders(Word.WdBorderType.wdBorderBottom).LineStyle = Word.WdLineStyle.wdLineStyleNone
                        .Borders(Word.WdBorderType.wdBorderTop).LineStyle = Word.WdLineStyle.wdLineStyleNone
                    End With
                    With .Columns(2)
                        .Width = myWord.InchesToPoints(5.5)
                        .Borders(Word.WdBorderType.wdBorderLeft).LineStyle = Word.WdLineStyle.wdLineStyleNone
                        .Borders(Word.WdBorderType.wdBorderRight).LineStyle = Word.WdLineStyle.wdLineStyleNone
                        With .Borders(Word.WdBorderType.wdBorderBottom)
                            .LineStyle = Word.WdLineStyle.wdLineStyleSingle
                            .LineWidth = Word.WdLineWidth.wdLineWidth025pt
                            .Color = RGB(0, 50, 0)
                        End With
                        With .Borders(Word.WdBorderType.wdBorderTop)
                            .LineStyle = Word.WdLineStyle.wdLineStyleSingle
                            .LineWidth = Word.WdLineWidth.wdLineWidth025pt
                            .Color = RGB(0, 50, 0)
                        End With
                    End With
                End With
            Else
                MsgBox("Select a Note, Caution, or Warning first, then format")
            End If
        Catch
            MsgBox("Something broke... Action Aborted!")
        End Try
    End Sub
End Module
