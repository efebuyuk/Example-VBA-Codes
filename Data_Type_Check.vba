Sub Data_Type_Check()
'Controls data types (formats).
'We will check Column3.

Current_LastRow = Sheet4.Cells(Sheet4.Rows.Count, "A").End(xlUp).Row
Starting_Row = Current_LastRow + 2

Sheet4.Cells(Starting_Row, 1).Value = "Date Format Control for Column3"

Annex2_LastRow = Sheet2.Cells(Sheet2.Rows.Count, "A").End(xlUp).Row

row_number = 2
cell_index = Starting_Row

Do
DoEvents
    row_number = row_number + 1
    cell_value = Sheet2.Range("I" & row_number)
    
    If IsDate(cell_value) = False Then
        cell_index = cell_index + 1
        Sheet4.Cells(cell_index, 1).Value = "Row " & row_number & _
            " in " & Sheet2.Name & " Sheet has " & cell_value & _
            " value which is not a date type. Please change it." & _
            " As an example, value should be like this: 01/01/2020"
    
    End If
    
    
Loop Until row_number = Annex2_LastRow

End Sub
