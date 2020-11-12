Sub Invalid_Value_Control()
'Controls invalid values.
'We will check Column8.

LastRow = Sheet2.Cells(Sheet2.Rows.Count, "H").End(xlUp).Row

Dim element As Variant
Dim coll As Object
Set coll = CreateObject("System.Collections.ArrayList")
    
' Add items
coll.Add "GRANTED"
coll.Add "ADDED"
coll.Add "OK"
coll.Add "VALID"

Sheet4.Cells(1, 1).Value = "Invalid Value Control for column8."

row_number = 2
cell_index = 1

Do
DoEvents
    row_number = row_number + 1
    cell_value = Sheet2.Range("H" & row_number)
    
    For Each element In coll
        If element = cell_value Then
            'cell_index = cell_index + 1
            'Sheet4.Cells(cell_index, 1).Value = "OK"
            Exit For
        Else
            cell_index = cell_index + 1
            Sheet4.Cells(cell_index, 1).Value = "Row " & row_number & _
            " in " & Sheet2.Name & _
            " Sheet has " & cell_value & " value which is invalid. Please change it." & _
            " Valid values are GRANTED, ADDED, OK and VALID."
            Exit For
        End If
    Next element
    
Loop Until row_number = LastRow

End Sub
