Attribute VB_Name = "stocktickerworks1"
Sub stocktickerworks():
'
'***data clensing
'loop through all worksheets
For Each ws In Worksheets

'assign value
lastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row
i = 2

    'Go through all the rows
    For i = lastRow To 1 Step -1

        'delete rows where opening and closing are 0
        If ws.Cells(i, "C").Value = 0 Then
            ws.Rows(i & ":" & i).EntireRow.delete
        End If

    Next i
Next ws

'***return summary table values
'loop through all worksheets
For Each ws In Worksheets

'declare
Dim ticker_name As String
Dim volume_total As Double
Dim summary_table_row As Integer
Dim opening As Double
Dim closing As Double
Dim percent_change As Double
Dim percent_formatted As String

'assign value
volume_total = 0
summary_table_row = 2
lastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row
i = 2

'title summary table
ws.Cells(1, "I").Value = "Ticker"
ws.Cells(1, "J").Value = "Yearly Change"
ws.Cells(1, "K").Value = "Percent Change"
ws.Cells(1, "L").Value = "Total Stock Volume"
ws.Range("I1:L1").Font.Bold = True
ws.Range("I1:L1").Columns.AutoFit

    'Assign opening value
    opening = ws.Cells(i, "C").Value
        
        'Go through all the rows
        For i = 2 To lastRow
           
    
            'check if still has same ticker value
            If ws.Cells(i + 1, "A").Value <> ws.Cells(i, "A").Value Then
    
                'set the ticker symbol
                ticker_name = ws.Cells(i, "A").Value
    
                'print ticker symbol
                ws.Range("I" & summary_table_row).Value = ticker_name
    
                'add to the volume total
                volume_total = volume_total + ws.Cells(i, "G").Value
    
                'print the volume total
                ws.Range("L" & summary_table_row) = volume_total
    
                'assign closing
                closing = ws.Cells(i, "F").Value
    
                'calculate yearly change and color
                yearly_change = closing - opening
    
                    'format colors for yearly change
                    If yearly_change >= 0 Then
                        ws.Range("J" & summary_table_row).Interior.ColorIndex = 50
    
                    Else
                        ws.Range("J" & summary_table_row).Interior.ColorIndex = 30
    
                    End If
    
                'caluclate percent change and format
                percent_change = (yearly_change / opening)
                percent_formatted = Format(percent_change, "Percent")
    
                'printing values
                ws.Range("J" & summary_table_row).Value = yearly_change
    
                ws.Range("K" & summary_table_row).Value = percent_formatted
    
                'resassign opening value
                opening = ws.Cells(i + 1, "C").Value
    
                'add one to the summary table row
                summary_table_row = summary_table_row + 1
    
                'reset volume_total
                volume_total = 0
    
            'if the cell immediately following a row is the same as ticker symbol
            Else
                'add to the volume total
                volume_total = volume_total + ws.Cells(i, "G").Value
            End If
                    
        Next i
        
Next ws

'***Additional summary table
'min, max, vol max w/ticker & value
For Each ws In Worksheets
'declare
Dim max As Double
Dim min As Double
Dim vol As Double
Dim k As Integer
Dim ticker As String
Dim max_format As String
Dim min_format As String
Dim vol_format As String

'assign value
lastRow = ws.Cells(Rows.Count, "I").End(xlUp).Row
k = lastRow

'second title summary table
ws.Cells(1, "P").Value = "Ticker"
ws.Cells(1, "Q").Value = "Value"
ws.Cells(2, "O").Value = "Greatest % Increase"
ws.Cells(3, "O").Value = "Greatest % Decrease"
ws.Cells(4, "O").Value = "Greatest Total Volume"
ws.Range("P1:Q1").Font.Bold = True
ws.Range("O2:O4").Font.Bold = True

'find values
max = WorksheetFunction.max(ws.Range("K2:K" & k))
min = WorksheetFunction.min(ws.Range("K2:K" & k))
vol = WorksheetFunction.max(ws.Range("L2:L" & k))

'index and match
'ws.Range("P2") = WorksheetFunction.Index(ws.Range("I2:K" & k), WorksheetFunction.Match(ws.Cells(2, "Q").Value, ws.Range("K2:K" & k), 0))
'ws.Range("P3") = WorksheetFunction.Index(ws.Range("I2:K" & k), WorksheetFunction.Match(ws.Cells(3, "Q").Value, ws.Range("K2:K" & k), 0))
'ws.Range("P4") = WorksheetFunction.Index(ws.Range("I2:L" & k), WorksheetFunction.Match(ws.Cells(4, "Q").Value, ws.Range("L2:L" & k), 0))

'format and print values
max_format = Format(max, "Percent")
ws.Cells(2, "Q").Value = max_format
min_format = Format(min, "Percent")
ws.Cells(3, "Q").Value = min_format
vol_format = Format(vol, "General Number")
ws.Cells(4, "Q").Value = vol_format
ws.Range("O1:Q4").Columns.AutoFit

Next ws
End Sub

'Sources:
'https://www.wallstreetmojo.com/vba-index-match/
'https://docs.microsoft.com/en-us/office/vba/
'https://www.ozgrid.com/forum/index.php?thread/79151-index-and-match-functions-in-macro-code/
'https://stackoverflow.com/questions/37011632/deleting-rows-conditional-on-the-content-of-a-column-in-vba



