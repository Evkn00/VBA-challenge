Attribute VB_Name = "AllSheets"
Sub ticker_fun_all_sheets()
'Define all variables
    Dim Ticker As String
    Dim LastRow As Long
    Dim TotalVolume As Double
    Dim OpenPrice As Double
    Dim ClosePrice As Double
    Dim YearlyChange As Double
    Dim PercentChange As Double
    Dim SummaryRow As Integer
    Dim MaxIncreaseTicker As String
    Dim MaxIncreaseValue As Double
    Dim MaxDecreaseTicker As String
    Dim MaxDecreaseValue As Double
    Dim MaxVolumeTicker As String
    Dim MaxVolumeValue As Double


    TotalVolume = 0
    OpenPrice = 0
    ClosePrice = 0
    YearlyChange = 0
    PercentChange = 0

'iterate throguh all sheets
For Each ws In Worksheets

    'set initial values
    Ticker = ws.Range("A2")
    OpenPrice = ws.Range("C2")
    SummaryRow = 2 'does not overwrite header row
    MaxIncreaseValue = 0
    MaxDecreaseValue = 0
    MaxVolumeValue = 0
    
    
'Set up the summary table headers
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
'Set up the Bonus Table
    ws.Range("p1").Value = "Ticker"
    ws.Range("q1").Value = "Value"
    ws.Range("o2").Value = "Greatest % Increase"
    ws.Range("o3").Value = "Greatest % Decrease"
    ws.Range("o4").Value = "Greatest Total Volume"
        
'loop through all rows of data in sheet
    LastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row
    For i = 2 To LastRow
    
        'if a new ticker is active in Ticker
        If ws.Cells(i, 1).Value <> Ticker Then
            'Calculate Close price, YearlyChange, PercentChange
            ClosePrice = ws.Cells(i - 1, "F")
            YearlyChange = ClosePrice - OpenPrice
            PercentChange = YearlyChange / OpenPrice

            
            'Output the summary stats for the previous ticker
            ws.Cells(SummaryRow, "I").Value = Ticker
            ws.Cells(SummaryRow, "J").Value = YearlyChange
            ws.Cells(SummaryRow, "K").Value = PercentChange
            ws.Cells(SummaryRow, "L").Value = TotalVolume
            
            'Check if current ticker is max decrease and update variables
            If PercentChange < MaxDecreaseValue Then
                MaxDecreaseValue = PercentChange
                MaxDecreaseTicker = Ticker
                End If
                
            'Check if current ticker is max increase and update variables
            If PercentChange > MaxIncreaseValue Then
                MaxIncreaseValue = PercentChange
                MaxIncreaseTicker = Ticker
                End If
                
            'Check if current ticker is max volume and update variables
            If TotalVolume > MaxVolumeValue Then
                MaxVolumeValue = TotalVolume
                MaxVolumeTicker = Ticker
                End If
                
            
           'Color the YearlyChange output
            If YearlyChange < 0 Then
                ws.Cells(SummaryRow, "J").Interior.ColorIndex = 3
            ElseIf YearlyChange > 0 Then
                ws.Cells(SummaryRow, "J").Interior.ColorIndex = 4
            End If
 
            'Reset the stats for the new ticker
            Ticker = ws.Cells(i, "A").Value
            TotalVolume = 0
            OpenPrice = ws.Cells(i, "C").Value
            'Iterate the summary row
            SummaryRow = SummaryRow + 1
         End If
         'Add to the total volume for the current ticker
         TotalVolume = TotalVolume + ws.Cells(i, "G").Value
         
    Next i
    
    'Assign values to cells in Bonus table
    ws.Range("p2").Value = MaxIncreaseTicker
    ws.Range("q2").Value = MaxIncreaseValue
    ws.Range("p3").Value = MaxDecreaseTicker
    ws.Range("q3").Value = MaxDecreaseValue
    ws.Range("p4").Value = MaxVolumeTicker
    ws.Range("q4").Value = MaxVolumeValue
    
    'Format percentage cell
    ws.Range("k2:K" & LastRow).NumberFormat = "0.00%"
    ws.Range("q2:q3").NumberFormat = "0.00%"
    'Format decimal places on Yearly Change column
    ws.Range("j2:j" & LastRow).NumberFormat = "0.00"
    'Expand all cells
    ws.Range("A1:q" & LastRow).Columns.AutoFit
    
Next ws

End Sub
