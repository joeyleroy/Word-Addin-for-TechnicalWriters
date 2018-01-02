Option Compare Text
Option Explicit On

Module UnitConverter
    Sub unitConversion()
        Dim selRange As Word.Range
        Dim selArray() As String
        Dim orgUnit As String, newUnit As String, selText As String, newText As String
        Dim orgDigit As Single, newDigit As Single
        Dim usUnit As Boolean, hasNumber As Boolean
        Dim i As Integer, myCount As Integer, iCnt As Integer

        myCount = 0
        selRange = Globals.ThisAddIn.Application.Selection.Range
        selText = selRange.Text
        selText = Replace(selText, vbCr, "") 'Trim paragraph marks from the user selected text if they exist

        If Not selRange.Characters.Count > 1 Then
            GoTo noSel
        End If

        If Mid(selText, Len(selText), 1) = "." Then 'Trim trailing period from the user selected text if one exists
            selText = StrReverse(Replace(StrReverse(selText), StrReverse("."), StrReverse(""), , 1)) 'WTF!?
            selText = Replace(selText, Chr(13), "")
        End If

        selText = Trim(selText) ' Trim leading or trailing spaces from the user selected text
        selText = Replace(selText, Chr(160), Chr(32)) ' Replace non-breaking Spaces with spaces for unified split

        hasNumber = False
        For iCnt = 1 To Len(selText)
            If IsNumeric(Mid(selText, iCnt, 1)) Then
                hasNumber = True
            ElseIf hasNumber = True Then
                If Mid(selText, iCnt, 1) = "," Then
                    'Do Nothing
                ElseIf Mid(selText, iCnt, 1) = "." Then
                    'Do Nothing
                ElseIf Mid(selText, iCnt, 1) = " " Then
                    selText = Mid(selText, 1, iCnt - 1) & Mid(selText, iCnt, Len(selText) - (iCnt - 1))
                    hasNumber = False
                Else
                    selText = Mid(selText, 1, iCnt - 1) & " " & Mid(selText, iCnt, Len(selText) - (iCnt - 1))
                    hasNumber = False
                End If
            End If
        Next iCnt

        selArray = Split(selText, " ") ' Split the User selected text on spaces (" ") into an array

        For i = 0 To UBound(selArray)
            If selArray(i) <> "" Then myCount = myCount + 1
        Next i

        If myCount = 1 Then
            MsgBox("The selected unit is too simple for conversion. Try selecting a digit and unit instead.")
            GoTo canOps
        ElseIf myCount = 2 Then
            orgDigit = selArray(0)
            If Mid(selArray(1), 1, Len(selArray(1))) = "in" Then
                orgUnit = selArray(1) & "."
            Else
                orgUnit = selArray(1)
            End If
        Else
            MsgBox("The selected unit is too complex for conversion. Try selecting a digit and unit instead.")
            GoTo canOps
        End If

        Select Case selArray(1)
            'US to Canada Units
            Case "in", "in.", "inch", "inches", Chr(34), Chr(147), Chr(148)
                newDigit = Math.Round(selArray(0) * 25.4, 2)
                newUnit = "mm"
                orgUnit = "in."
                usUnit = True
            Case "ft", "ft.", "foot", "feet", Chr(39), Chr(145), Chr(146)
                newDigit = Math.Round(selArray(0) * 0.3048, 1)
                newUnit = "m"
                orgUnit = "ft"
                usUnit = True
            Case "psi"
                newDigit = Math.Round(selArray(0) * 0.0068948, 1)
                newUnit = "MPa"
                usUnit = True
            Case "lb", "lbs", "pounds"
                newDigit = Math.Round(selArray(0) * 0.4535924, 1)
                newUnit = "kg"
                orgUnit = "lb"
                usUnit = True
            Case "lbf"
                newDigit = Math.Round(selArray(0) * 0.4448222)
                newUnit = "daN"
                usUnit = True
            Case "lb/ft", "ft/lb"
                newDigit = Math.Round(selArray(0) * 1.49, 2)
                newUnit = "kg/m"
                orgUnit = "lb/ft"
                usUnit = True
            Case "lb/gal", "lb-gal", "gal/lb", "gal-lb"
                newDigit = Math.Round(selArray(0) * 119.8225188)
                newUnit = "kg/m3" ' Need to Superscript the cube
                orgUnit = "lb/gal"
                usUnit = True
            Case "bbl", "bbls", "bbl's", "bbls'"
                newDigit = Math.Round(selArray(0) * 158.9872949)
                newUnit = "L"
                orgUnit = "bbl"
                usUnit = True
            Case "bbl/min", "bbls/min", "bbl's/min", "bbls'/min"
                newDigit = Math.Round(selArray(0) * 158.9872949)
                newUnit = "L/min"
                orgUnit = "bbl/min"
                usUnit = True
            Case "ft/min", "ft-min", "min/ft", "min-ft"
                newDigit = Math.Round(selArray(0) * 0.3048, 2)
                newUnit = "m/min"
                orgUnit = "ft/min"
                usUnit = True
            Case "ft/sec", "ft-sec", "sec/ft", "sec-ft"
                newDigit = Math.Round(selArray(0) * 0.3048, 1)
                newUnit = "m/sec"
                orgUnit = "ft/sec"
                usUnit = True
            Case "ft-lb", "ft-lbs", "lb-ft", "lbs-ft", "ft" & Chr(30) & "lb", "lb" & Chr(30) & "ft"
                newDigit = Math.Round(selArray(0) * 1.3558182)
                newUnit = "N" & Chr(30) & "m"
                orgUnit = "ft" & Chr(30) & "lb"
                usUnit = True
            Case "degF", Chr(176) & "F"
                newDigit = Math.Round((selArray(0) - 32) * (5 / 9))
                newUnit = Chr(176) & "C"
                usUnit = True
            ' Canada to US Units
            Case "mm"
                newDigit = Math.Round(selArray(0) / 25.4, 2)
                newUnit = "in."
                usUnit = False
            Case "m"
                newDigit = Math.Round(selArray(0) / 0.3048, 1)
                newUnit = "ft"
                usUnit = False
            Case "MPa"
                newDigit = Math.Round(selArray(0) / 0.0068948, 1)
                newUnit = "psi"
                usUnit = False
            Case "kg", "kgs"
                newDigit = Math.Round(selArray(0) / 0.4535924, 1)
                newUnit = "lb"
                usUnit = False
            Case "daN"
                newDigit = Math.Round(selArray(0) / 0.4448222)
                newUnit = "lbf"
                usUnit = False
            Case "kg/m"
                newDigit = Math.Round(selArray(0) / 1.49, 2)
                newUnit = "lb/ft"
                usUnit = False
            Case "kg/m3", "kg/m^3" ' Need to Superscript the cube
                newDigit = Math.Round(selArray(0) / 119.8225188)
                newUnit = "lb/gal"
                usUnit = False
            Case "L"
                newDigit = Math.Round(selArray(0) / 158.9872949)
                newUnit = "bbl"
                usUnit = False
            Case "L/min"
                newDigit = Math.Round(selArray(0) / 158.9872949, 1)
                newUnit = "bbl/min"
                usUnit = False
            Case "ft/min"
                newDigit = Math.Round(selArray(0) / 0.3048, 2)
                newUnit = "m/min"
                usUnit = False
            Case "m/sec"
                newDigit = Math.Round(selArray(0) / 0.3048, 1)
                newUnit = "ft/sec"
                usUnit = False
            Case "N-m", "N" & Chr(30) & "m"
                newDigit = Math.Round(selArray(0) / 1.3558182)
                newUnit = "ft" & Chr(30) & "lb"
                usUnit = False
            Case "degC", Chr(176) & "C"
                newDigit = Math.Round(selArray(0) * (9 / 5) + 32)
                newUnit = Chr(176) & "F"
                usUnit = True
            Case Else
                MsgBox("Unknown or non-unit type selected!")
                GoTo canOps
        End Select

        Dim myNewDigit As String = Nothing
        Dim myOrgDigit As String = Nothing

        If IsNumeric(newDigit) Then
            myNewDigit = Format(newDigit, "###,###,###.#####")
        End If
        If IsNumeric(orgDigit) Then
            myOrgDigit = Format(orgDigit, "###,###,###.#####")
        End If

        If usUnit = True Then
            newText = myOrgDigit & Chr(160) & orgUnit & " [" & myNewDigit & Chr(160) & newUnit & "]"
        Else
            newText = myNewDigit & Chr(160) & newUnit & " [" & myOrgDigit & Chr(160) & orgUnit & "]"
        End If

        'Checks to see if the user selection includes a paragraph mark,
        'and if true, reduce the users selection by 1 to exclude it.
        If Mid(selRange.Text, Len(selRange.Text), 1) = vbCr _
            Or Mid(selRange.Text, Len(selRange.Text), 1) = Chr(13) Then
            selRange.End = selRange.End - 1
        End If

        'Checks to see if the user original selection includes a period,
        'and if true, reduce the users selection by 1 to exclude it.
        'This code also tests to see if the period is part of the "in." unit.
        If Mid(selRange.Text, Len(selRange.Text), 1) = Chr(46) Then
            If Mid(selRange.Text, Len(selRange.Text) - 2, 3) <> "in." Then
                selRange.End = selRange.End - 1
            End If
        End If

        selRange.Text = newText

        GoTo canOps

noSel:
        MsgBox("You must select something! Try selecting a number AND a unit of measure.")

canOps:
        Erase selArray ' Erase the array once complete.

    End Sub

End Module
