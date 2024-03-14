Attribute VB_Name = "Module1"
Sub StockCheck()
        
    'First will establish variables
    Dim finalRow As Double
    Dim openValue As Double
    Dim closeValue As Double
    Dim summaryPosition As Integer
    Dim firstTickerRow As Integer
    Dim greatestIncreaseValue As Double
    Dim greatestDecreaseValue As Double
    Dim greatestIncreaseTicker As String
    Dim greatestDecreaseTicker As String
    Dim greatestVolumeTicker As String
    Dim greatestVolume As Double
    Dim totalVolume As Double
    Dim ws As Worksheet
    
    For Each ws In Worksheets
    
        'Initiating these values in Worksheet loop so previous year data does not carry into next year
        finalRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        openValue = ws.Cells(2, 3).Value
        closeValue = 0
        summaryPosition = 2
        firstTickerRow = 2
        greatestIncreaseValue = 0
        greatestDecreaseValue = 0
        greatestIncreaseTicker = "0"
        greatestDecreaseTicker = "0"
        greatestVolumeTicker = "0"
        greatestVolume = 0
        totalVolume = 0
        i = 2
    
        'Setting up headers
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
    
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"

        'Summarizer / main logic
        For i = 2 To finalRow
        totalVolume = totalVolume + ws.Cells(i, 7).Value
        
        If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
            ws.Cells(summaryPosition, 9).Value = ws.Cells(i, 1).Value 'puts the Ticker in our summary chart if the ticker is not the same as the next one
            closeValue = ws.Cells(i, 6).Value 'sets the closeValue variable to the close value of the row it identifies to be the last one belonging to that ticker
            ws.Cells(summaryPosition, 10).Value = closeValue - openValue 'puts the difference in the price at open against the price at close
            ws.Cells(summaryPosition, 11).Value = (closeValue - openValue) / openValue 'puts the difference in the price at open against the price at close as a percentage
            ws.Cells(summaryPosition, 11).NumberFormat = "0.00%"
            ws.Cells(summaryPosition, 12).Value = totalVolume
            
            If ws.Cells(summaryPosition, 10).Value > 0 Then
                ws.Cells(summaryPosition, 10).Interior.ColorIndex = 4
            
                ElseIf ws.Cells(summaryPosition, 10).Value < 0 Then
                ws.Cells(summaryPosition, 10).Interior.ColorIndex = 3
            End If
            
            If ws.Cells(summaryPosition, 11).Value > greatestIncreaseValue Then
                greatestIncreaseValue = ws.Cells(summaryPosition, 11).Value 'Checks if the increase for this particular ticker is greater than the previous greatest percentage, and if so replaces it
                greatestIncreaseTicker = ws.Cells(i, 1).Value
            End If
            
            If ws.Cells(summaryPosition, 11).Value < greatestDecreaseValue Then
                greatestDecreaseValue = ws.Cells(summaryPosition, 11).Value 'Checks if the decrease for this particular ticker is less than the previous low; if so replaces it
                greatestDecreaseTicker = ws.Cells(i, 1).Value
            End If
        
            If ws.Cells(summaryPosition, 12).Value > greatestVolume Then
                greatestVolume = ws.Cells(summaryPosition, 12).Value 'Checks if the increase for this particular ticker is greater than the previous greatest volume; if so replaces previous value
                greatestVolumeTicker = ws.Cells(i, 1).Value
            End If
                      
            openValue = ws.Cells(i + 1, 3).Value 'sets a new openValue for the next unique ticker
            summaryPosition = summaryPosition + 1 'Increases summaryPosition by 1 so we are not overwriting data
            totalVolume = 0 'Resets totalVolume for next unique ticker
        
        End If
        Next i
    
        'Inputting summary table values
        ws.Cells(2, 16).Value = greatestIncreaseTicker
        ws.Cells(2, 17).Value = greatestIncreaseValue
        ws.Cells(3, 16).Value = greatestDecreaseTicker
        ws.Cells(3, 17).Value = greatestDecreaseValue
        ws.Cells(4, 16).Value = greatestVolumeTicker
        ws.Cells(4, 17).Value = greatestVolume
        
    Next ws
End Sub
