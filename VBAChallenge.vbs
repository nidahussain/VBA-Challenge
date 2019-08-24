Sub easy()
    Dim ws As Worksheet
    headers = Array("Ticker", "Total Stock Volume")
      For Each ws In Worksheets
        
        ws.Cells(1, 9).Value = "Ticket"
        ws.Cells(1, 10).Value = "Total Stock Volume"
      
        Dim Ticker As String
        Dim Volume As Double
        Volume = 0
        Dim totvol As Long
       
        Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2
       
        totvol = ws.Cells(Rows.Count, 1).End(xlUp).Row
  
  For i = 2 To totvol

    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        Ticker = Cells(i, 1).Value
        Volume = Volume + Cells(i, 7).Value

        
        Range("I" & Summary_Table_Row).Value = Ticker
        Range("J" & Summary_Table_Row).Value = Volume
        

        Summary_Table_Row = Summary_Table_Row + 1
        
        Volume = 0
    Else:
        Volume = Volume + Cells(i, 7).Value
    
    
        
    End If
      
    
  Next i
Next ws
End Sub

