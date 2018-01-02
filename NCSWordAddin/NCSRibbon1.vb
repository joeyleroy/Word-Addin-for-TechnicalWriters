Imports Microsoft.Office.Tools.Ribbon

Public Class NCSRibbon1

    Private Sub NCSRibbon1_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load
        'NCSRibbon1.Font.Name = 'Comic Sans MS';

    End Sub

    Private Sub ButHeading1_Click(sender As Object, e As RibbonControlEventArgs) Handles ButHeading1.Click
        NCSFormat("Heading 1")
    End Sub

    Private Sub ButHeading2_Click(sender As Object, e As RibbonControlEventArgs) Handles ButHeading2.Click
        NCSFormat("Heading 2")
    End Sub

    Private Sub ButHeading3_Click(sender As Object, e As RibbonControlEventArgs) Handles ButHeading3.Click
        NCSFormat("Heading 3")
    End Sub

    Private Sub ButHeading4_Click(sender As Object, e As RibbonControlEventArgs) Handles ButHeading4.Click
        NCSFormat("Heading 4")
    End Sub

    Private Sub ButBodyText1_Click(sender As Object, e As RibbonControlEventArgs) Handles ButBodyText1.Click
        NCSFormat("Body Text")
    End Sub

    Private Sub ButBodyText2_Click(sender As Object, e As RibbonControlEventArgs) Handles ButBodyText2.Click
        NCSFormat("Body Text 2")
    End Sub

    Private Sub ButBodyText3_Click(sender As Object, e As RibbonControlEventArgs) Handles ButBodyText3.Click
        NCSFormat("Body Text 3")
    End Sub

    Private Sub ButTableSpace1_Click(sender As Object, e As RibbonControlEventArgs) Handles ButTableSpace1.Click
        NCSFormat("Table Spacer")
    End Sub

    Private Sub ButInsertGraphic_Click(sender As Object, e As RibbonControlEventArgs) Handles ButInsertGraphic.Click
        InsertGraphic()
    End Sub

    Private Sub ButKeepWithNext_Click(sender As Object, e As RibbonControlEventArgs) Handles ButKeepWithNext.Click
        KeepWithNext()
    End Sub

    Private Sub ButPageBreak_Click(sender As Object, e As RibbonControlEventArgs) Handles ButPageBreak.Click
        PageBreakBefore()
    End Sub

    Private Sub ButToggleDocProps_Click(sender As Object, e As RibbonControlEventArgs) Handles ButToggleDocProps.Click
        ToggleDocProps()
    End Sub

    Private Sub ButSnipitz_Click(sender As Object, e As RibbonControlEventArgs) Handles ButSnipitz.Click
        Dim Snipitz As New Snipitz
        Dim myWord As Word.Application = Globals.ThisAddIn.Application
        Dim myDoc As Word.Document = myWord.ActiveDocument
        Dim frmCollection = System.Windows.Forms.Application.OpenForms

        If frmCollection.OfType(Of Snipitz).Any Then
            frmCollection.Item("Snipitz").Activate()
        Else
            Snipitz.Show()
        End If

    End Sub

    Private Sub ButUnitConverter_Click(sender As Object, e As RibbonControlEventArgs) Handles ButUnitConverter.Click
        unitConversion()
    End Sub

    Private Sub ButUpdateFields_Click(sender As Object, e As RibbonControlEventArgs) Handles ButUpdateFields.Click
        NCSUpdateFields()
    End Sub

    Private Sub ButFixHeadings_Click(sender As Object, e As RibbonControlEventArgs) Handles ButFixHeadings.Click
        updateHeadingStyles()
    End Sub

    Private Sub ButFormatNotes_Click(sender As Object, e As RibbonControlEventArgs) Handles ButFormatNotes.Click
        FormatNCSNotes()
    End Sub

    Private Sub ButXREF_Click(sender As Object, e As RibbonControlEventArgs) Handles ButXREF.Click
        Dim FormXREF As New FormXREF
        Dim myWord As Word.Application = Globals.ThisAddIn.Application
        Dim myDoc As Word.Document = myWord.ActiveDocument
        Dim frmCollection = System.Windows.Forms.Application.OpenForms

        If frmCollection.OfType(Of FormXREF).Any Then
            frmCollection.Item("FormXREF").Activate()
        Else
            FormXREF.Show()
        End If
    End Sub

    Private Sub ButTemplates_Click(sender As Object, e As RibbonControlEventArgs) Handles ButTemplates.Click
        Dim FormTemplates As New FormTemplates
        Dim myWord As Word.Application = Globals.ThisAddIn.Application
        Dim myDoc As Word.Document = myWord.ActiveDocument
        Dim frmCollection = System.Windows.Forms.Application.OpenForms

        If frmCollection.OfType(Of FormTemplates).Any Then
            frmCollection.Item("FormTemplates").Activate()
        Else
            FormTemplates.Show()
        End If
    End Sub

    Private Sub LMFB1_Click(sender As Object, e As RibbonControlEventArgs) Handles LMFB1.Click
        Dim myWord As Word.Application = Globals.ThisAddIn.Application
        Dim frmCollection = System.Windows.Forms.Application.OpenForms
        Dim myint As Integer = frmCollection.Count
        Dim myi As Integer

        For myi = 0 To (myint - 1)
            If frmCollection(myi).Name = "Snipitz" Then
                With frmCollection(myi)
                    .Visible = True
                    .Top = (myWord.Top)
                    .Left = (myWord.Left)
                    .StartPosition = 3
                    If Not My.Settings.SnipSize.IsEmpty Then
                        ' Need to clear this file: My.Settings.SnipitsLoc
                    End If

                    If Not My.Settings.SnipitsLoc.IsEmpty Then
                        ' Need to clear this file: My.Settings.SnipitsLoc
                    End If
                End With
            End If
        Next myi
    End Sub

    Private Sub LMFB2_Click(sender As Object, e As RibbonControlEventArgs) Handles LMFB2.Click
        Dim myWord As Word.Application = Globals.ThisAddIn.Application
        Dim frmCollection = System.Windows.Forms.Application.OpenForms
        Dim myint As Integer = frmCollection.Count
        Dim myi As Integer

        For myi = 0 To (myint - 1)
            If frmCollection(myi).Name = "FormXREF" Then
                With frmCollection(myi)
                    .Visible = True
                    .Top = (myWord.Top)
                    .Left = (myWord.Left)
                    .StartPosition = 3
                    If Not My.Settings.XRefSize.IsEmpty Then
                        ' Need to clear this file: My.Settings.XRefSize
                    End If

                    If Not My.Settings.XRefLoc.IsEmpty Then
                        ' Need to clear this file: My.Settings.XRefLoc
                    End If
                End With
            End If
        Next myi
    End Sub

    Private Sub LMFB3_Click(sender As Object, e As RibbonControlEventArgs) Handles LMFB3.Click
        Dim myWord As Word.Application = Globals.ThisAddIn.Application
        Dim frmCollection = System.Windows.Forms.Application.OpenForms
        Dim myint As Integer = frmCollection.Count
        Dim myi As Integer

        For myi = 0 To (myint - 1)
            If frmCollection(myi).Name = "FormTemplates" Then
                With frmCollection(myi)
                    .Visible = True
                    .Top = (myWord.Top)
                    .Left = (myWord.Left)
                    .StartPosition = 3
                    If Not My.Settings.TempSize.IsEmpty Then
                        ' Need to clear this file: My.Settings.TempSize
                    End If

                    If Not My.Settings.TempLoc.IsEmpty Then
                        ' Need to clear this file: My.Settings.TempLoc
                    End If
                End With
            End If
        Next myi
    End Sub

    Private Sub RibbonFolder_Click(sender As Object, e As RibbonControlEventArgs) Handles Configuration.Click
        Dim FormConfig As New Configuration
        Dim frmCollection = System.Windows.Forms.Application.OpenForms

        If frmCollection.OfType(Of Configuration).Any Then
            frmCollection.Item("Configuration").Activate()
        Else
            FormConfig.Show()
        End If
    End Sub
End Class
