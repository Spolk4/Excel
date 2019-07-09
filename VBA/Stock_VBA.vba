Part I Easy
Sub WallStreetVBA()

'Set initial variable for holding the unique ticker symbols

Dim TickerGroup As String
Ticker = " "
'Set variable for total stock volume and start the counter at zero
Dim TotlalStockVolume As Double
TotalStockVolume = 0
'Keep Track of location for each ticker group and stock value
Dim SummaryTableRow As Integer
SummaryTableRow = 2

Cells(1, 11).Value = "Ticker Group"
Cells(1, 12).Value = "Total Stock Volume"
'Define last row
LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    'loop for each iteration to last row
    For i = 2 To LastRow
        'If we are still with the same ticker symbol then
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        'Set a ticker group
        TickerGroup = Cells(i, 1).Value
        'Set total stock volume
        TotalStockVolume = TotalStockVolume + Cells(i, 7).Value
        'Print to the Ticker group in the summary table
        Range("K" & SummaryTableRow).Value = TickerGroup
        'Print the Total Stock Volume
        Range("L" & SummaryTableRow).Value = TotalStockVolume
        'Add one to the summary table row
        SummaryTableRow = SummaryTableRow + 1
        'Reset the Total Stock Volume
        TotalStockVolume = 0
        Else
        'If the ticker cell immediately following is the same
        TotalStockVolume = TotalStockVolume + Cells(i, 7).Value
        End If
    Next i
End Sub