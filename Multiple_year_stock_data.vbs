Attribute VB_Name = "Module1"
Sub overall():
    'Declare variables
    Dim j As Integer
    Dim sheetCount As Integer
   
    'count sheets
    sheetCount = Sheets.Count
    'iterate through sheets
    For j = 1 To sheetCount
        stock_volume (j)
    Next j
End Sub
' ----------------------------------------------------
Sub stock_volume(s As Integer):
  ' Set an initial variable for holding the ticker
  Dim tickerSym As String
 
  ' Set an initial variable for volume
  Dim volumeTotal As Double
  ' set variable for open and final price & differences
  Dim openPrice As Double
  Dim closePrice As Double
  ' Keep track of the ticker
  Dim tableRow As Integer
  Dim i As Long
  Dim ws As Worksheet
 
 
  'loop worksheets
  For Each ws In Worksheets
 
        volumeTotal = 0
        tableRow = 2
 
        'determine last row
        Dim lastRow As Long
        lastRow = Cells(Rows.Count, 1).End(xlUp).Row
       
        'print new column headers
        ws.Range("I1") = "Ticker"
        ws.Range("J1") = "Yearly Change"
        ws.Range("K1") = "Percent Change"
        ws.Range("L1") = "Volume"
       
        'Find the distinct value for Ticker and calculate yearly change and Percent change
        dum1 = 2
        dum2 = 2
        amountDiff = 0
 
        ' Loop through all tickers
        For i = 2 To lastRow
          If dum1 = dum2 Then
            openPrice = ws.Cells(i, 3).Value
            dum1 = dum1 + 1
          End If
   
          If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
   
             'store closing price
              closePrice = ws.Cells(i, 6).Value
             'Set the ticker
              tickerSym = ws.Cells(i, 1).Value
   
          ' print ticker
          ws.Range("I" & tableRow).Value = tickerSym
   
          ' print total volume for ticker
          ws.Range("L" & tableRow).Value = volumeTotal
   
          'print amount change
          ws.Range("J" & tableRow).Value = closePrice - openPrice
   
                If ws.Range("J" & tableRow).Value > 0 Then
                    ws.Range("J" & tableRow).Interior.ColorIndex = 4 'Green
                   
                ElseIf ws.Range("J" & tableRow).Value < 0 Then
                    ws.Range("J" & tableRow).Interior.ColorIndex = 3  'Red
                Else
                   
                End If
   
                If openPrice <> 0 Then
                    ws.Range("K" & tableRow).Value = (closePrice - openPrice) / openPrice
                Else
                    ws.Range("K" & tableRow).Value = 0
                End If
                    ws.Range("K" & tableRow).NumberFormat = "0.00%"
   
          ' Add one to the table row
          tableRow = tableRow + 1
   
          ' Reset the volume total
          volumeTotal = 0
   
          'reset amount difference
          amountDiff = 0
   
        ' If the cell immediately following a row is the same ticker
        Else
          ' Add to the volume Total
          volumeTotal = volumeTotal + ws.Cells(i + 1, 7).Value
   
        End If
   
      Next i
 
  Next ws
End Sub

