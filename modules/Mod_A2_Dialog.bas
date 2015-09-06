Attribute VB_Name = "Mod_A2_Dialog"
Public WorksheetString As String
Sub dir_to_cell()
    ActiveSheet.Cells(1, 1) = ThisWorkbook.path
End Sub
Function String_to_Range(desc As String) As Range
'desc is supposed to be [workbook\]worksheet/row/column

If CountStrings(desc, "/") = 3 Then
    If CountStringsB(Splitword(1), "\") = 2 Then
        Set String_to_Range = Workbooks(SplitwordB(1)).Worksheets(SplitwordB(2)).Cells(Splitword(2), Splitword(3))
    Else
        Set String_to_Range = ThisWorkbook.Worksheets(Splitword(1)).Cells(Val(Splitword(2)), Val(Splitword(3)))
    End If
End If

End Function
Sub SelectWorksheet()
    WorksheetSelectionForm.Show
    If WorksheetString <> "" Then ActiveCell = WorksheetString
End Sub
Sub UseFileDialogOpen()

    Dim lngCount As Long

    ' Open the file dialog
    With Application.FileDialog(msoFileDialogFilePicker)
        .AllowMultiSelect = True
        .Show

        ' Display paths of each file selected
        For lngCount = 1 To .SelectedItems.Count
            ActiveCell.Offset(lngCount - 1) = .SelectedItems(lngCount)
        Next lngCount

    End With

End Sub
Sub UseFileDialogBrowseFolder()

    Dim lngCount As Long

    ' Open the file dialog
    With Application.FileDialog(msoFileDialogFolderPicker)
        .AllowMultiSelect = False
        .Show

        ' Display paths of each file selected
        For lngCount = 1 To .SelectedItems.Count
            Cells(3, 3).Offset(lngCount - 1) = .SelectedItems(lngCount)
        Next lngCount

    End With

End Sub

