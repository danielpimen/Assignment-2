'1)Check Value of First Cell

Sub Counter2()

  Dim Ticker As String

  
  Dim Volume As Double
  Volume = 0
    Dim lastRow As Long
    lastRow = Sheet1.Range("A" & Rows.Count).End(xlUp).Row
  
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2

 
  For i = 2 To lastRow

  
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

  
      Ticker = Cells(i, 1).Value

      
      Volume = Volume + Cells(i, 7).Value

     
      Range("L" & Summary_Table_Row).Value = Ticker


      Range("M" & Summary_Table_Row).Value = Volume


      Summary_Table_Row = Summary_Table_Row + 1
      

      Volume = 0

    
    Else

      Volume = Volume + Cells(i, 7).Value

    End If

  Next i

End Sub

