Attribute VB_Name = "Module1"
Sub Loops()
    
Dim i As Long

    For i = 2 To 1000
    
    Cells(i, 9).Value = Cells(i, 1).Value
    
    Next i
    

Dim j As Long

    For j = 2 To 1000
    
    Cells(j, 10).Value = Cells(j, 3).Value - Cells(j, 6).Value
    
    Next j
    

Dim k As Long
    
    For k = 2 To 1000
    
    Cells(k, 11).Value = (Cells(k, 3).Value - Cells(k, 6)) / Cells(k, 6).Value
    
    Next k
    
    
Dim m As Long

    For m = 2 To 1000
    
    Cells(m, 12).Value = Cells(m, 4).Value * Cells(m, 7).Value
    
    Next m
    
Dim z As Long

For z = 2 To 1000

     If Cells(z, 10).Value > 0 Then
     
     Cells(z, 10).Interior.ColorIndex = 4
     
     Else
     
     Cells(z, 10).Interior.ColorIndex = 3
     
     End If
     
    Next z

End Sub





