Private Sub ImportSdmxDsd_Click()

Dim fDialog As FileDialog
Dim sFileName As String
Dim xDoc As Object
Dim root As Object
Dim listEntryValue As String
Dim listEntryName As String
Dim dropdown As ContentControl

Set fDialog = Application.FileDialog(msoFileDialogFilePicker)
Set xDoc = CreateObject("MSXML2.DOMDocument.6.0")

xDoc.async = False
xDoc.validateOnParse = False

With fDialog
    .Filters.Clear
    .Title = "Select your SDMX DSD file"
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
        For Each codeNode In root.SelectNodes("//str:Codelist[@id='CL_SERIES']/str:Code")
            listEntryValue = codeNode.Attributes.getNamedItem("id").Text
            listEntryName = ""
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
            listEntryName = fixedListEntryName(listEntryName)
            dropdown.DropdownListEntries.Add listEntryName, listEntryValue
        Next codeNode
        
        'Populate the Reference Area dropdown.
        Set dropdown = ActiveDocument.SelectContentControlsByTag("ddRefArea").Item(1)
        dropdown.DropdownListEntries.Clear
        For Each codeNode In root.SelectNodes("//str:Codelist[@id='CL_AREA']/str:Code")
            listEntryValue = codeNode.Attributes.getNamedItem("id").Text
            listEntryName = codeNode.SelectSingleNode("com:Name").Text
            listEntryName = fixedListEntryName(listEntryName)
            'Reference area codes are duplicated in the global DSD, so we only use the numeric ones.
            If IsNumeric(listEntryValue) = True Then
                dropdown.DropdownListEntries.Add listEntryName, listEntryValue
            End If
        Next codeNode
        
        'Populate the Reporting Type dropdown.
        Set dropdown = ActiveDocument.SelectContentControlsByTag("ddReportingType").Item(1)
        dropdown.DropdownListEntries.Clear
        For Each codeNode In root.SelectNodes("//str:Codelist[@id='CL_REPORTING_TYPE']/str:Code")
            listEntryValue = codeNode.Attributes.getNamedItem("id").Text
            listEntryName = codeNode.SelectSingleNode("com:Name").Text
            listEntryName = fixedListEntryName(listEntryName)
            dropdown.DropdownListEntries.Add listEntryName, listEntryValue
        Next codeNode

        Protect_Template

        MsgBox "Successfully updated dropdowns."

    End If
End With

End Sub

Private Function fixedListEntryName(listEntryName As String) As String

    If Len(listEntryName) > 255 Then
        listEntryName = Left(listEntryName, 250) & "..."
    End If
    
    fixedListEntryName = listEntryName

End Function
