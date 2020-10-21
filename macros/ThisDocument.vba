Private Sub ImportSdmxDsd_Click()

Dim fDialog As FileDialog
Dim sFileName As String
Dim xDoc As Object
Dim root As Object
Dim listEntryValue As String
Dim listEntryName As String
Dim dropdown As ContentControl

Dim objFileSystem As Object
Dim objTextFile As Object
Dim sSdmxDsd As String
Dim ccSdmxBox As ContentControl

Dim cRefAreas As Collection
Dim vRefArea As Variant
Set cRefAreas = New Collection

Dim aRefAreasAlphabetical() As String
ReDim aRefAreasAlphabetical(10000)
Dim iRefAreaIndex As Integer
Dim bRefAreaWorldExists As Boolean

Set fDialog = Application.FileDialog(msoFileDialogFilePicker)
Set xDoc = CreateObject("MSXML2.DOMDocument.6.0")

Set objFileSystem = CreateObject("Scripting.FileSystemObject")

xDoc.async = False
xDoc.validateOnParse = False

With fDialog
    .Filters.Clear
    .title = "Select your SDMX DSD file"
    .Filters.Add "XML Files", "*.xml?", 1
    .AllowMultiSelect = False

    If .Show Then
        sFileName = .SelectedItems(1)
        If xDoc.Load(sFileName) = False Then
            MsgBox "Unable to load DSD. Reason: " & xDoc.parseError.reason
        End If

        If ActiveDocument.ProtectionType <> wdNoProtection Then
            ActiveDocument.Unprotect
        End If

        'Also load the DSD as plain text and save it hidden.
        Set objTextFile = objFileSystem.OpenTextFile(sFileName, 1)
        sSdmxDsd = objTextFile.ReadAll
        Set ccSdmxBox = ActiveDocument.SelectContentControlsByTag("boxSdmxDsd").Item(1)
        ccSdmxBox.Appearance = wdContentControlHidden
        ccSdmxBox.Range.Text = sSdmxDsd
        ccSdmxBox.Range.Font.Hidden = 1

        xDoc.SetProperty "SelectionNamespaces", "xmlns:str='http://www.sdmx.org/resources/sdmxml/schemas/v2_1/structure' xmlns:com='http://www.sdmx.org/resources/sdmxml/schemas/v2_1/common'"
        Set root = xDoc.DocumentElement

        'Populate the Series dropdown.
        Set dropdown = ActiveDocument.SelectContentControlsByTag("ddSeries").Item(1)
        dropdown.DropdownListEntries.Clear
        'Always include a national catch-all option.
        listEntryValue = "_"
        listEntryName = fixedListEntryName("0.0.0 National series not in global framework", listEntryValue)
        dropdown.DropdownListEntries.Add listEntryName, listEntryValue
        'Get the rest from the DSD.
        For Each codeNode In root.SelectNodes("//str:Codelist[@id='CL_SERIES']/str:Code")
            listEntryValue = codeNode.Attributes.getNamedItem("id").Text
            listEntryName = ""
            'Check for the "RetiredSeries" annotations.
            For Each annotationNode In codeNode.SelectNodes("com:Annotations/com:Annotation")
                If annotationNode.SelectSingleNode("com:AnnotationTitle").Text = "RetiredSeries" Then
                    If listEntryName <> "" Then
                        listEntryName = listEntryName & ", "
                    End If
                    listEntryName = listEntryName & "RETIRED"
                End If
            Next annotationNode
            'Combine all the "Indicator" annotations.
            For Each annotationNode In codeNode.SelectNodes("com:Annotations/com:Annotation")
                If annotationNode.SelectSingleNode("com:AnnotationTitle").Text = "Indicator" Then
                    If listEntryName <> "" Then
                        listEntryName = listEntryName & ", "
                    End If
                    listEntryName = listEntryName & annotationNode.SelectSingleNode("com:AnnotationText").Text
                End If
            Next annotationNode
            If listEntryName <> "" Then
                listEntryName = listEntryName & " "
            End If
            'In addition to the "Indicator" annotations combined above, use the code's Name.
            listEntryName = listEntryName & codeNode.SelectSingleNode("com:Name").Text
            listEntryName = fixedListEntryName(listEntryName, listEntryValue)
            dropdown.DropdownListEntries.Add listEntryName, listEntryValue
        Next codeNode

        'Populate the Reference Area dropdown.
        Set dropdown = ActiveDocument.SelectContentControlsByTag("ddRefArea").Item(1)
        dropdown.DropdownListEntries.Clear
        iRefAreaIndex = 0
        bRefAreaWorldExists = False
        For Each codeNode In root.SelectNodes("//str:Codelist[@id='CL_AREA']/str:Code")
            listEntryValue = codeNode.Attributes.getNamedItem("id").Text
            listEntryName = codeNode.SelectSingleNode("com:Name").Text
            listEntryName = fixedListEntryName(listEntryName, listEntryValue)
            'Reference area codes are duplicated in the global DSD, so we only use the numeric ones.
            If IsNumeric(listEntryValue) = True Then
                If listEntryName = "World (1)" Then
                    bRefAreaWorldExists = True
                End If
                cRefAreas.Add listEntryValue, listEntryName
                aRefAreasAlphabetical(iRefAreaIndex) = listEntryName
                iRefAreaIndex = iRefAreaIndex + 1
            End If
        Next codeNode

        'Sort alphabetically.
        ReDim Preserve aRefAreasAlphabetical(iRefAreaIndex - 1)
        For i = 0 To UBound(aRefAreasAlphabetical)
            For x = UBound(aRefAreasAlphabetical) To i + 1 Step -1
                If aRefAreasAlphabetical(x) < aRefAreasAlphabetical(i) Then
                    holdInt = aRefAreasAlphabetical(x)
                    aRefAreasAlphabetical(x) = aRefAreasAlphabetical(i)
                    aRefAreasAlphabetical(i) = holdInt
                End If
            Next x
        Next i

        If bRefAreaWorldExists Then
            dropdown.DropdownListEntries.Add "World (1)", "1"
        End If
        For i = 0 To UBound(aRefAreasAlphabetical)
            If aRefAreasAlphabetical(i) <> "World (1)" Then
                dropdown.DropdownListEntries.Add aRefAreasAlphabetical(i), cRefAreas(aRefAreasAlphabetical(i))
            End If
        Next i

        'Populate the Reporting Type dropdown.
        Set dropdown = ActiveDocument.SelectContentControlsByTag("ddReportingType").Item(1)
        dropdown.DropdownListEntries.Clear
        For Each codeNode In root.SelectNodes("//str:Codelist[@id='CL_REPORTING_TYPE']/str:Code")
            listEntryValue = codeNode.Attributes.getNamedItem("id").Text
            listEntryName = codeNode.SelectSingleNode("com:Name").Text
            listEntryName = fixedListEntryName(listEntryName, listEntryValue)
            dropdown.DropdownListEntries.Add listEntryName, listEntryValue
        Next codeNode

        Protect_Template

        MsgBox "Successfully updated dropdowns."

    End If
End With

End Sub

Private Function fixedListEntryName(listEntryName As String, listEntryValue As String) As String

    If Len(listEntryName) > 200 Then
        listEntryName = Left(listEntryName, 200) & "..."
    End If

    'Also add the ID at the end, according to a naming convention.
    listEntryName = listEntryName & " (" & listEntryValue & ")"

    fixedListEntryName = listEntryName

End Function


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

Private Sub AddSeries_Click()
    If ActiveDocument.ProtectionType <> wdNoProtection Then
        ActiveDocument.Unprotect
    End If

    Dim lastSeriesDropdown As ContentControl
    Set lastSeriesDropdown = getLastSeriesButton()

    If lastSeriesDropdown Is Nothing Then
        Set lastSeriesDropdown = ActiveDocument.SelectContentControlsByTag("ddSeries").Item(1)
    End If

    lastSeriesDropdown.Copy
    Selection.Collapse wdCollapseStart
    Selection.Paste

    ActiveDocument.Protect wdAllowOnlyReading
End Sub

Private Sub RemoveSeries_Click()
    If ActiveDocument.ProtectionType <> wdNoProtection Then
        ActiveDocument.Unprotect
    End If

    Dim lastSeriesDropdown As ContentControl
    Set lastSeriesDropdown = getLastSeriesButton()

    If Not lastSeriesDropdown Is Nothing Then
        With lastSeriesDropdown
            ActiveDocument.Range(.Range.Start - 1, .Range.End + 2).Delete
        End With
    End If

    ActiveDocument.Protect wdAllowOnlyReading
End Sub

Private Function getLastSeriesButton() As ContentControl
    Dim seriesDropdowns As ContentControls
    Dim seriesDropdown As ContentControl
    Dim lastSeriesDropdown As ContentControl

    Set seriesDropdowns = ActiveDocument.SelectContentControlsByTag("ddSeries")

    If seriesDropdowns.Count > 1 Then

        For Each seriesDropdown In seriesDropdowns
            If lastSeriesDropdown Is Nothing Then
                Set lastSeriesDropdown = seriesDropdown
            ElseIf seriesDropdown.Range.Start > lastSeriesDropdown.Range.Start Then
                Set lastSeriesDropdown = seriesDropdown
            End If
        Next seriesDropdown

    End If

    Set getLastSeriesButton = lastSeriesDropdown
End Function
