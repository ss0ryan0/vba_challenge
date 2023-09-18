Attribute VB_Name = "Module1"
Option Explicit
Sub columns()

'Given excel columns
Const c_Ticker = 1
Const c_Date = 2
Const c_Open = 3
Const c_Close = 6
Const c_Volume = 7


'New excel columns
Const c_UniqueTicker = 9
Const c_AnnualChange = 10
Const c_Percent = 11
Const c_TotVolume = 12
Const c_oTicker = 16
Const c_Value = 17

'Added column names
Cells(1, c_UniqueTicker).Value = "Ticker"
Cells(1, c_AnnualChange).Value = "Annual Change"
Cells(1, c_Percent).Value = "% Change"
Cells(1, c_TotVolume).Value = "Total Stock Volume"
Cells(1, c_oTicker).Value = "Ticker"
Cells(1, c_Value).Value = "Value"
Cells(2, 15).Value = "Greatest % Increase"
Cells(3, 15).Value = "Greatest % Decrease"
Cells(4, 15).Value = "Greatest Total Volume"

Dim iRow, iFixRow As Integer
Dim sum_Volume, largest_volume As Double
Dim prev_Ticker, highestTicker, lowestTicker, largestTicker As String
Dim first_Open, last_Close, highest_change, lowest_change As Double


    iRow = 2
    iFixRow = 2
    sum_Volume = 0
    prev_Ticker = Cells(iRow, c_Ticker).Value
    first_Open = Cells(iRow, c_Open).Value
    
    highest_change = 0
    lowest_change = 0
    largest_volume = 0
    
    'While on the same stock, accumulate the open and the close values.
    'Then move/calculate the required info at the end for that stock before beginning to process the next stock.
    
    While Not (IsEmpty(Cells(iRow, c_Ticker).Value))  'loop through the entire spreadsheet, assumed to be sorted by ticker symbol
        sum_Volume = sum_Volume + Cells(iRow, c_Volume).Value
        
        iRow = iRow + 1
        If Not (prev_Ticker = Cells(iRow, c_Ticker).Value) Then  'when the ticker changes
            last_Close = Cells(iRow - 1, c_Close).Value
            Cells(iFixRow, c_UniqueTicker).Value = prev_Ticker   'assign the previous ticker to the unique one
            Cells(iFixRow, c_AnnualChange).Value = (last_Close - first_Open)
            Cells(iFixRow, c_Percent).Value = FormatPercent(((last_Close - first_Open) / first_Open))
            Cells(iFixRow, c_TotVolume).Value = sum_Volume
            
            If highest_change < Cells(iFixRow, c_Percent).Value Then
                highest_change = Cells(iFixRow, c_Percent).Value
                highestTicker = Cells(iFixRow, c_UniqueTicker).Value
            End If
            
            If lowest_change > Cells(iFixRow, c_Percent).Value Then
                lowest_change = Cells(iFixRow, c_Percent).Value
                lowestTicker = Cells(iFixRow, c_UniqueTicker).Value
            End If
            
            If largest_volume < Cells(iFixRow, c_TotVolume).Value Then
                largest_volume = Cells(iFixRow, c_TotVolume).Value
                largestTicker = Cells(iFixRow, c_UniqueTicker).Value
            End If
            
           'conditional formating
            If Cells(iFixRow, c_Percent).Value > 0 Then
                Cells(iFixRow, c_Percent).Interior.ColorIndex = 4
            Else
                Cells(iFixRow, c_Percent).Interior.ColorIndex = 3
            End If
            
            Cells(2, 16).Value = highestTicker
            Cells(2, 17).Value = FormatPercent(highest_change)
            Cells(3, 16).Value = lowestTicker
            Cells(3, 17).Value = FormatPercent(lowest_change)
            Cells(4, 16).Value = largestTicker
            Cells(4, 17).Value = largest_volume
    
            first_Open = Cells(iRow, c_Open)
            sum_Volume = 0
            prev_Ticker = Cells(iRow, c_Ticker).Value
            iFixRow = iFixRow + 1
End If
Wend
End Sub

