Attribute VB_Name = "Module2"
Sub Stock():

' Loop through all sheets

    For Each WS In Worksheets
    
' Determine the last row

    LastRow = WS.Cells(Rows.Count, 1).End(xlUp).Row
    
' Set Variables

    Dim Ticker As String
    
    Dim Volume As Double
    Volume = 0
    
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
    
' Set Headers

    Cells(1, "H").Value = "Ticker"
    
    Cells(1, "I").Value = "Volume"
    
' Loop through all stock volume

        For i = 2 To LastRow
    
            If Cells(i + 1, "A").Value <> Cells(i, "A").Value Then
    
            Ticker = Cells(i, "A").Value
    
            Volume = Volume + Cells(i, "G").Value
            
            ' Print the Ticker in the Summary Table
        
            Range("H" & Summary_Table_Row).Value = Ticker
            
            ' Print the Volume to the Summary Table
      
            Range("I" & Summary_Table_Row).Value = Volume
    
            Summary_Table_Row = Summary_Table_Row + 1
        
            Volume = 0
        
        Else
        
            Volume = Volume + Cells(i, "G").Value
    
            End If
    
        Next i
    
    Next WS
    
End Sub

