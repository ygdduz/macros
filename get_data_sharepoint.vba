Sub copyData()
    ' Define variables
    Dim sourceSheet As Worksheet
    Dim destinationSheet As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim j As Long
    Dim dateValue As Date
    
    ' Set source and destination sheets
    Set sourceSheet = ThisWorkbook.Sheets("Earthquake Reponse (2)")
    Set destinationSheet = ThisWorkbook.Sheets("Sheet1")
    
    ' Find last row of source data
    lastRow = sourceSheet.Cells(Rows.Count, "A").End(xlUp).Row
    
    ' Loop through data and copy only if M column is a number
    j = 2 ' start at row 2 in destination sheet
    For i = 3 To lastRow
        If IsNumeric(sourceSheet.Range("M" & i).Value) Then
            ' Copy row i to destination sheet, starting at row j
            sourceSheet.Range("A" & i & ":J" & i).Copy Destination:=destinationSheet.Range("A" & j)
            j = j + 1 ' increment destination row counter
        End If
    Next i
    
    ' Delete unwanted columns from destination sheet
    destinationSheet.Range("E:E,F:F,G:G,I:I,J:J,P:P,S:S,T:T,U:U,V:V,W:W,X:X,Y:Y").Delete Shift:=xlToLeft
    
    ' Change column A to date format
    destinationSheet.Range("A2:A" & j - 1).NumberFormat = "mm/dd/yyyy"
    
    ' Fill empty cells in destination sheet with values above in columns A to J only
    With destinationSheet.Range("A2:J" & j - 1)
        .SpecialCells(xlCellTypeBlanks).FormulaR1C1 = "=R[-1]C"
        .Value = .Value
    End With
End Sub

