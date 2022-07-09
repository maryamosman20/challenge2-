Attribute VB_Name = "Module1"
Sub Challenge2()

    For Each ws In Worksheets
    ws.Activate
    MsgBox (ws.Name)
    Range("I1").Value = "ticker"
    Range("J1").Value = "yearly change"
    Range("K1").Value = "percent change"
    Range("L1").Value = "total volume change"
    
    
    
    Dim totalvolume As Double
    Dim startrow As Long
    Dim table2row As Integer
    table2row = 0
    totalvolume = 0
    
    
    
    
    startrow = 2
    
RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    For i = 2 To RowCount
    
    
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

      
    totalvolume = totalvolume + Cells(i, 7).Value
    startopenprice = Cells(startrow, 3).Value
    endcloseprice = Cells(i, 6).Value
    yearlychange = endcloseprice - startopenprice
    percentchange = yearlychange / startopenprice * 100
    startrow = i + 1
    Range("I" & 2 + table2row).Value = Cells(i, 1).Value
    Range("J" & 2 + table2row).Value = yearlychange
    Range("K" & 2 + table2row).Value = percentchange
    Range("L" & 2 + table2row).Value = totalvolume
    If yearlychange < 0 Then
    Range("J" & 2 + table2row).Interior.ColorIndex = 3
    Else: Range("J" & 2 + table2row).Interior.ColorIndex = 4
    End If
    totalvolume = 0
    yearlychange = 0
    table2row = table2row + 1

    
    
    Else
    totalvolume = totalvolume + Cells(i, 7).Value
    
    
    End If

    
    
    Next i


    
    
    

Next ws


End Sub
