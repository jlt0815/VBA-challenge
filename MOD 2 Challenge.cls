VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub YearData()

   For Each ws In Worksheets
   
lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
   
        Dim tickerSymbol As String
        Dim closingPrice As Double
        Dim yearlyChange As Double
        Dim percentageChange As Double
        
        Dim totalVolume  As Double
        totalVolume = 0
        
        Dim sumRow As Integer
        sumRow = 2
        
        Dim openingPrice As Double
        openingPrice = ws.Cells(2, 3).Value
        
     

    
        ws.Cells(1, 9).Value = "Ticker"
       ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"



        For i = 2 To lastRow

    
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
              tickerSymbol = ws.Cells(i, 1).Value
              ws.Range("I" & sumRow).Value = tickerSymbol

            totalVolume = totalVolume + ws.Cells(i, 7).Value
            ws.Range("L" & sumRow).Value = totalVolume


              closingPrice = ws.Cells(i, 6).Value
              yearlyChange = (closingPrice - openingPrice)
              ws.Range("J" & sumRow).Value = yearlyChange

            If (openingPrice = 0) Then

                    percentageChange = 0

                Else
                    
                    percentageChange = yearlyChange / openingPrice
                
                End If


              ws.Range("K" & sumRow).Value = percentageChange
              ws.Range("K" & sumRow).NumberFormat = "0.00%"
   
    
              sumRow = sumRow + 1
              totalVolume = 0
              openingPrice = ws.Cells(i + 1, 3)
            
            Else
              
     
              totalVolume = totalVolume + ws.Cells(i, 7).Value

            
            End If
        
        Next i
        
    lastrowTable = ws.Cells(Rows.Count, 9).End(xlUp).Row

    For i = 2 To lastrowTable
            If ws.Cells(i, 10).Value > 0 Then
                ws.Cells(i, 10).Interior.ColorIndex = 10
            Else
                ws.Cells(i, 10).Interior.ColorIndex = 3
            
            End If
    
    Next i
    
    
Dim greatestIncrease As Double
Dim greatestDecrease As Double
Dim greatestTotal As Double



For i = 2 To lastRow
    
    If ws.Cells(i, 11).Value > greatestIncrease Then
    greatestIncrease = ws.Cells(i, 11).Value
    
    ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
    ws.Cells(2, 17).Value = ws.Cells(i, 11).Value
    ws.Cells(2, 17).Style = "Percent"
    ws.Cells(2, 17).NumberFormat = "0.00%"
    
    ElseIf ws.Cells(i, 11).Value < greatestDecrease Then
    greatestDecrease = ws.Cells(i, 11).Value
    
    ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
    ws.Cells(3, 17).Value = ws.Cells(i, 11).Value
    ws.Cells(3, 17).Style = "Percent"
    ws.Cells(3, 17).NumberFormat = "0.00%"

    
    End If

    
    If ws.Cells(i, 12).Value > greatestTotal Then
    greatestTotal = ws.Cells(i, 12).Value
    
    ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
    ws.Cells(4, 17).Value = ws.Cells(i, 12).Value
    
    
    End If
Next i
    
    Next ws
    

End Sub


