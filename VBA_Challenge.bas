Attribute VB_Name = "Module2"
Sub StockMarket()
'Loop through all worksheets
Dim ws As Worksheet

For Each ws In Worksheets
    
        'Make Table headers
        ws.Cells(1, 9) = "Ticker"
        ws.Cells(1, 10) = "Yearly Change"
        ws.Cells(1, 11) = "Percent Change"
        ws.Cells(1, 12) = "Total Stock Volume"

        ws.Cells(1, 15) = "Ticker"
        ws.Cells(1, 16) = "Value"
        ws.Cells(2, 14) = "Greatest % Increase"
        ws.Cells(3, 14) = "Greatest % Decrease"
        ws.Cells(4, 14) = "Greatest Total Volume"

'Define variables

LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
Dim Ticker As String

Dim Table_Row As Integer
Table_Row = 2

Dim Vol As Double
Volume = 0

   'Loop data to pull out ticker and volume
    For x = 2 To LastRow

        'Loop to see if ticker matches and if so add volume to table
        If ws.Cells(x, 1).Value = ws.Cells(x + 1, 1).Value Then
            Volume = Volume + ws.Cells(x, 7).Value
            
       ElseIf ws.Cells(x + 1, 1).Value <> ws.Cells(x, 1).Value Then
            'Add up volume for ticker
            Volume = Volume + ws.Cells(x, 7).Value
            
            'Print Ticker in Table
            ws.Cells(Table_Row, 9).Value = ws.Cells(x, 1).Value
            
            'Print volume into table
            ws.Cells(Table_Row, 12).Value = Volume
            
            'Next row for table
            Table_Row = Table_Row + 1
            
            'Reset volume
            Volume = 0
        
        End If
    Next x
    
'Loop to find yearly change and percent change and insert it in table
    Dim beg_stock As Double
    Dim end_stock As Double
    Dim change As Double
    Dim percent_change
    Dim k As Integer
    k = 2
    
    For x = 2 To LastRow
        'Loop to find the start date of the stock
        If Right(ws.Cells(x, 2), 4) = "0101" Then
            beg_stock = ws.Cells(x, 3)
        'loop to find end date of the stock
        ElseIf Right(ws.Cells(x, 2), 4) = "1231" Then
            end_stock = ws.Cells(x, 6)
            'Create formula to get the yearly change
            change = end_stock - beg_stock
            ws.Cells(k, 10).Value = change
            'Create formula to get the percent change
            percent_change = change / beg_stock
                On Error Resume Next
                
    
        'Repeat above step for if the stock date ends on the 30th and not 31st
        ElseIf Right(ws.Cells(x, 2), 4) = "1230" Then
            end_stock = ws.Cells(x, 6)
            change = end_stock - beg_stock
            ws.Cells(k, 10).Value = change
            percent_change = change / beg_stock
                On Error Resume Next
            
        'Conditional format so that positive changes are green and negative are red and display numbers in the correct format
        If ws.Cells(k, 10) > 0 Then
            ws.Cells(k, 10).Interior.ColorIndex = 4
        ElseIf ws.Cells(k, 10) < 0 Then
            ws.Cells(k, 10).Interior.ColorIndex = 3
        End If
        ws.Cells(k, 11).Value = percent_change
        ws.Cells(k, 11).NumberFormat = "0.00%"
        k = k + 1
        End If
    Next x
        
'Greatest % increase and decreaded and Greatest Total Volume
'Define variables to find min and max
Dim Max As Double
Dim Min As Double
Dim VolumeMax As LongLong
Dim rng1 As Range
Dim rng2 As Range

LastRowTable = ws.Cells(Rows.Count, 11).End(xlUp).Row
LastRowVolume = ws.Cells(Rows.Count, 12).End(xlUp).Row
 
'Set range for the data we will be finding the min and max for
    Set rng1 = ws.Range("K2:K" & LastRowTable)
    Set rng2 = ws.Range("L2:L" & LastRowVolume)

    'Calculate the min and the max
    Max = WorksheetFunction.Max(rng1)
    Min = WorksheetFunction.Min(rng1)
    VolumeMax = WorksheetFunction.Max(rng2)
    
    'report min and max in table
    ws.Cells(2, 16).Value = Max
    ws.Cells(2, 16).NumberFormat = "0.00%"
    ws.Cells(3, 16).Value = Min
    ws.Cells(3, 16).NumberFormat = "0.00%"
    ws.Cells(4, 16).Value = VolumeMax
    
    'Get ticker for min and max table
    For x = 2 To LastRowTable
        If ws.Cells(x, 11).Value = Max Then
            ws.Cells(2, 15) = ws.Cells(x, 9).Value
        End If
        If ws.Cells(x, 11).Value = Min Then
            ws.Cells(3, 15) = ws.Cells(x, 9).Value
        End If
    Next x
    
    'Get ticker max for volume
     For x = 2 To LastRowVolume
        If ws.Cells(x, 12).Value = VolumeMax Then
            ws.Cells(4, 15) = ws.Cells(x, 9).Value
        End If
    Next x
    
    'Format columns to autofit
    ws.Columns("I:L").AutoFit
    ws.Columns("N:P").AutoFit

  
        
Next ws

End Sub
