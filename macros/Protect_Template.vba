Public Sub Protect_Template()

    Dim currentTable As Table
    Dim currentRow As Row
    Dim firstRowSkipped As Boolean
    Dim currentControl As ContentControl

    If ActiveDocument.ProtectionType <> wdNoProtection Then
        ActiveDocument.Unprotect
    End If

    For Each currentTable In ActiveDocument.Tables

        If isValidTableTitle(currentTable.title) Then

            firstRowSkipped = False
            For Each currentRow In currentTable.Rows
                If currentRow.Cells.Count = 2 Then
                    If firstRowSkipped Then
                        currentRow.Cells(2).Select
                        Selection.Editors.Add wdEditorEveryone
                    End If
                    firstRowSkipped = True
                End If
            Next
        End If

    Next

    For Each currentControl In ActiveDocument.ContentControls

        If isValidControlTag(currentControl.tag) Then
            currentControl.Range.Select
            Selection.Editors.Add wdEditorEveryone

        End If

    Next

    ActiveDocument.Protect wdAllowOnlyReading

    Selection.GoTo What:=wdGoToLine, Which:=wdGoToAbsolute, Count:=1

End Sub

Private Function isValidTableTitle(title As String) As Boolean
    isValidTableTitle = _
        title = "0. Indicator information" _
        Or title = "1. Data reporter" _
        Or title = "2. Definition, concepts, and classifications" _
        Or title = "3. Data source type and data collection method" _
        Or title = "4. Other methodological considerations" _
        Or title = "5. Data availability and disaggregation" _
        Or title = "6. Comparability/deviation from international standards" _
        Or title = "7. References and Documentation"
End Function

Private Function isValidControlTag(tag As String) As Boolean
    isValidControlTag = _
        tag = "ddReportingType" _
        Or tag = "ddSeries" _
        Or tag = "ddRefArea" _
        Or tag = "ddLanguage"
End Function

