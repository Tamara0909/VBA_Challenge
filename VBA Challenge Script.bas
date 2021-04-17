Attribute VB_Name = "GreatestPercentVolume"
Sub WallStreet_Stock()
Dim LastRow As Long
Dim Counter As Long
Dim WB As Workbook
Dim WS As Worksheet
Dim Ticker As String
Dim TotalStockVolume As Double
Dim Open_Price As Double
Dim Close_Price As Double
Dim Counter_Result As Long
Dim GreatestPercentIncrease As Double
Dim GreatestPercentDecrease As Double
Dim GreatestVolume As Double
Dim RowIndex As Long
Dim rng1 As String
Dim rng2 As String
Dim Range1 As Range
Dim Range2 As Range
Set WB = ActiveWorkbook
Set WS = WB.ActiveSheet
LastRow = WS.Cells(Rows.Count, 1).End(xlUp).Row
Counter = 2
Ticker = WS.Cells(Counter, 1)
Open_Price = WS.Cells(Counter, 3)
WS.Range("I2:L1000000").ClearContents
WS.Range("I2:L1000000").ClearFormats
Counter_Result = 2

Do While Counter <= LastRow
   If WS.Cells(Counter, 1) <> Ticker Then
        WS.Cells(Counter_Result, 9) = Ticker
        WS.Cells(Counter_Result, 12) = TotalStockVolume
        Close_Price = WS.Cells(Counter - 1, 6)
        WS.Cells(Counter_Result, 10) = Close_Price - Open_Price
        If WS.Cells(Counter_Result, 10) < 0 Then
            WS.Cells(Counter_Result, 10).Interior.Color = vbRed
        Else
            WS.Cells(Counter_Result, 10).Interior.Color = vbGreen
        End If
        If Close_Price = 0 Then
            If Open_Price = 0 Then
                WS.Cells(Counter_Result, 11) = FormatPercent(0, 2)
            Else
                WS.Cells(Counter_Result, 11) = FormatPercent(-1, 2)
            End If
        Else
            WS.Cells(Counter_Result, 11) = FormatPercent((Close_Price - Open_Price) / Close_Price, 2)
        End If
        Ticker = WS.Cells(Counter, 1)
        TotalStockVolume = WS.Cells(Counter, 7)
        Open_Price = WS.Cells(Counter, 3)
        Counter_Result = Counter_Result + 1
   Else
        TotalStockVolume = TotalStockVolume + WS.Cells(Counter, 7)
   End If

Counter = Counter + 1
Loop
    Close_Price = WS.Cells(Counter - 1, 6)
    WS.Cells(Counter_Result, 10) = Close_Price - Open_Price
    
    If WS.Cells(Counter_Result, 10) < 0 Then
        WS.Cells(Counter_Result, 10).Interior.Color = vbRed
    Else
        WS.Cells(Counter_Result, 10).Interior.Color = vbGreen
    End If
        If Close_Price = 0 Then
            If Open_Price = 0 Then
                WS.Cells(Counter_Result, 11) = FormatPercent(0, 2)
            Else
                WS.Cells(Counter_Result, 11) = FormatPercent(-1, 2)
            End If
        Else
            WS.Cells(Counter_Result, 11) = FormatPercent((Close_Price - Open_Price) / Close_Price, 2)
        End If
    WS.Cells(Counter_Result, 9) = WS.Cells(Counter - 1, 1)
    WS.Cells(Counter_Result, 12) = TotalStockVolume
    rng1 = "K2:K" & Counter_Result
    rng2 = "L2:L" & Counter_Result
    Set Range1 = Range(rng1)
    Set Range2 = Range(rng2)
    GreatestPercentDecrease = Application.Min(Range1)
    GreatestPercentIncrease = Application.WorksheetFunction.Max(Range1)
    GreatestVolume = Application.WorksheetFunction.Max(Range2)
    RowIndex = Application.WorksheetFunction.Match(GreatestPercentDecrease, Range1, 0)
    WS.Range("P2") = WS.Cells(RowIndex, 9)
    WS.Range("Q2") = GreatestPercentDecrease
    
    
    RowIndex = Application.WorksheetFunction.Match(GreatestPercentIncrease, Range1, 0)
    WS.Range("P3") = WS.Cells(RowIndex, 9)
    WS.Range("Q3") = GreatestPercentIncrease
    
    
    RowIndex = Application.WorksheetFunction.Match(GreatestVolume, Range2, 0)
    WS.Range("P4") = WS.Cells(RowIndex, 9)
    WS.Range("Q4") = GreatestVolume
    
End Sub

