VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Homework_2()
    ' Create a variable to hold the ticker
    Dim ws As Worksheet
    Dim i As Long
    Dim ticker As String
    Dim tickerCounter As Long
    Dim openingPrice As Double
    Dim closingPrice As Double
    Dim quarterlyChange As Double
    Dim percentageChange As Double
    Dim dateVolume As Double
    Dim totalVolume As Double
    Dim lastRow As Long
 
    ' Variables for finding the greatest values
    Dim greatestIncrease As Double
    Dim greatestDecrease As Double
    Dim greatestVolume As Double
    Dim greatestIncreaseTicker As String
    Dim greatestDecreaseTicker As String
    Dim greatestVolumeTicker As String
    

    
    ' Loop through all sheets in the workbook
    For Each ws In ThisWorkbook.Worksheets

        ' Skip empty sheets or sheets without data
        If WorksheetFunction.CountA(ws.Cells) > 0 Then
        
        ' Activate the sheet
            ws.Activate
    
        ' Get the last row dynamically
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        
        ' Define variable values
        totalVolume = 0
        tickerCounter = 2
        greatestIncrease = 0
        greatestDecrease = 0
        greatestVolume = 0
        
        'Create columns in each ws
        ws.Cells(1, 9).Value = "Ticker Symbol"
        ws.Cells(1, 10).Value = "Quarterly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
    
            
            ' Loop through each row from row 2 to the last row
            For i = 2 To lastRow
            
                ' Check if this is a new ticker
                If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
                    ' Print the ticker symbol
                    ws.Cells(tickerCounter, 9).Value = ws.Cells(i, 1).Value

                    ' Set the opening price for the new ticker
                    openingPrice = ws.Cells(i, 3).Value

                    ' Reset total volume for the new ticker
                    totalVolume = 0
                End If

                'Add to total volume
                totalVolume = totalVolume + ws.Cells(i, 7).Value
        
                ' Check if this is the last row for this ticker
                If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Or i = lastRow Then
                
                    'Set the closing price for the ticker
                    closingPrice = ws.Cells(i, 6).Value
                    
                    'Calculate quarterly change and percentage change
                    quarterlyChange = closingPrice - openingPrice
                    
                    'Avoid division by zero for percentage change
                    If openingPrice <> 0 Then
                        percentageChange = (quarterlyChange / openingPrice)
                    Else
                        percentageChange = 0
                    End If
            
                    'Print the total volume, quarterly change, and percentage change
                    ws.Cells(tickerCounter, 12).Value = totalVolume
                    ws.Cells(tickerCounter, 10).Value = quarterlyChange
                    ws.Cells(tickerCounter, 11).Value = percentageChange
                
        
                    ' Track the greatest percentage increase
                    If percentageChange > greatestIncrease Then
                        greatestIncrease = percentageChange
                        greatestIncreaseTicker = ws.Cells(i, 1).Value
                    End If
                    
                    ' Track the greatest percentage decrease
                    If percentageChange < greatestDecrease Then
                        greatestDecrease = percentageChange
                        greatestDecreaseTicker = ws.Cells(i, 1).Value
                    End If
                    
                    ' Track the greatest total volume
                    If totalVolume > greatestVolume Then
                        greatestVolume = totalVolume
                        greatestVolumeTicker = ws.Cells(i, 1).Value
                    End If
                                
                    ' Format quarterly change with conditional formatting
                    If quarterlyChange > 0 Then
                        ws.Cells(tickerCounter, 10).Interior.ColorIndex = 4 ' Green
                    ElseIf quarterlyChange < 0 Then
                        ws.Cells(tickerCounter, 10).Interior.ColorIndex = 3 ' Red
                    End If

                    ' Format percentage change with conditional formatting
                    If percentageChange > 0 Then
                        ws.Cells(tickerCounter, 11).Interior.ColorIndex = 4 ' Green
                    ElseIf percentageChange < 0 Then
                        ws.Cells(tickerCounter, 11).Interior.ColorIndex = 3 ' Red
                    End If
                    
                    'Format percentage change to percentage format
                    ws.Cells(tickerCounter, 11).NumberFormat = "0.00%"
                    
                ' Increment tickerCounter to move to the next row
                tickerCounter = tickerCounter + 1
            
                End If
            
       

    
        Next i
    End If

  
                  ' Output the results for the greatest values
                  
                    ws.Cells(2, 15).Value = greatestIncreaseTicker
                    ws.Cells(2, 16).Value = Format(greatestIncrease, "0.00%")
            
                    ws.Cells(3, 15).Value = greatestDecreaseTicker
                    ws.Cells(3, 16).Value = Format(greatestDecrease, "0.00%")
                
                    ws.Cells(4, 15).Value = greatestVolumeTicker
                    ws.Cells(4, 16).Value = greatestVolume
     Next ws
        
End Sub
                
            

    

