Sub Copy_From_Original_Data()
'Getting data from Original Data sheet.
'First find and match the columns
'Then copy those matched columns to Copied Data sheet.

LastRow = Sheet3.Cells(Sheet3.Rows.Count, "H").End(xlUp).Row
LastColumn_OriginalData = Sheet3.Range("A1:H1").Columns.Count
LastColumn_CopiedData = Sheet2.Range("A1:H1").Columns.Count

Dim row_number As Integer
Dim column_number As Integer
Dim c As Integer

row_number = 0
column_number = 0

Do
DoEvents

    column_number = column_number + 1
    copiedData_cell = Sheet2.Cells(1, column_number).Value
    
    For c = 1 To LastColumn_OriginalData
        
        originalData_cell = Sheet3.Cells(1, c).Value
        
        If copiedData_cell = originalData_cell Then
        
            For row_number = 2 To LastRow
                copied_value = Sheet3.Cells(row_number, c).Value
                Sheet2.Cells(row_number, column_number).Value = copied_value
            Next
            
        End If
    
    Next


Loop Until column_number = LastColumn_CopiedData

End Sub
