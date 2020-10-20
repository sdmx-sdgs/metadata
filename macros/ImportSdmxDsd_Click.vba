Private Sub ImportSdmxDsd_Click()

Dim fDialog As FileDialog
Dim sFileName As String
Dim xDoc As Object
Dim root As Object
Dim listEntryValue As String
Dim listEntryName As String
Dim dropdown As ContentControl

Dim cRefAreas As Collection
Dim vRefArea As Variant
Set cRefAreas = New Collection

Dim aRefAreasAlphabetical() As String
ReDim aRefAreasAlphabetical(10000)
Dim iRefAreaIndex As Integer
Dim bRefAreaWorldExists As Boolean

Set fDialog = Application.FileDialog(msoFileDialogFilePicker)
Set xDoc = CreateObject("MSXML2.DOMDocument.6.0")

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
