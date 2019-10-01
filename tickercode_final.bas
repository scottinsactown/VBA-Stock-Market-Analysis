Attribute VB_Name = "Module21"
Option Explicit

Sub tickercode()

' Prepare for worksheet loops
Dim ws As Worksheet
Dim starting_ws As Worksheet
Set starting_ws = ActiveSheet

For Each ws In ThisWorkbook.Worksheets
ws.Activate

'insert column headers into first available columns to left
Dim lastcol As Double
lastcol = Cells(1, Columns.Count).End(xlToLeft).Column

Cells(1, lastcol + 1).Value = "Ticker  "
Cells(1, lastcol + 2).Value = "Yearly Change"
Cells(1, lastcol + 3).Value = "Percent Change"
Cells(1, lastcol + 4).Value = "Total Stock Volume"

'ID first ticker
Dim ticker_ID As String
ticker_ID = Range("a2").Value

'create other variables
Dim open_price As Double
open_price = Range("c2").Value

Dim end_price As Double

Dim yearly_change As Double

Dim total_vol As Double
total_vol = 0

Dim lastrow As Double
lastrow = Cells(Rows.Count, 1).End(xlUp).row

'sort column A, which must be alpha for rest of code to work
Range("A2" & lastcol).Sort key1:=Range("A2:A" & lastrow), _
   order1:=xlAscending, Header:=xlNo

'this variable is the summary tables location
Dim row_summary As Double
row_summary = 2

'this variable is the current row in the main table
Dim row_search As Double

For row_search = 2 To lastrow + 1

'populate summary stats

    'current ticker ID volume updated
    If Cells(row_search, 1).Value = ticker_ID Then
        Cells(row_summary, lastcol + 1).Value = ticker_ID
        total_vol = total_vol + Cells(row_search, 7).Value
        
    Else
        'update ticker ID
        ticker_ID = Cells(row_search, 1).Value
        row_summary = row_summary + 1
        Cells(row_summary, lastcol + 1).Value = ticker_ID

        'find end price for previous ticker ID
        end_price = Cells(row_search - 1, 6).Value
         
        'calculate and insert yearly change for previous ticker ID
        Cells(row_summary - 1, lastcol + 2).Value = end_price - open_price
        
        yearly_change = Cells(row_summary - 1, lastcol + 2).Value
        
            'calculate and insert percent change for previous ticker ID
            
            'avoid 0/0 overflow issue
            If open_price = 0 Then
            Cells(row_summary - 1, lastcol + 3).Value = 0
            
    '******add to above to if? AND if open_price = 0
            
            Else: Cells(row_summary - 1, lastcol + 3).Value = FormatPercent(yearly_change / open_price)
              
            End If
              
        'insert total volume for previous ticker ID
        Cells(row_summary - 1, lastcol + 4).Value = total_vol
    
        'update open_price for curent ticker ID
        open_price = Cells(row_search, 3).Value
         
        'reset total volume for current ticker ID
        total_vol = Cells(row_search, 7).Value
   
    End If

Next row_search

'add conditional formatting
'update lastrow and row search for summary tables
lastrow = Cells(Rows.Count, lastcol + 2).End(xlUp).row

    For row_search = 2 To lastrow

    'red = negative change
    If Cells(row_search, lastcol + 2).Value < 0 Then
    Cells(row_search, lastcol + 2).Interior.ColorIndex = 3

    'green = positive change (zero change does not get formatting)
    ElseIf Cells(row_search, lastcol + 2).Value > 0 Then
    Cells(row_search, lastcol + 2).Interior.ColorIndex = 4

    End If

Next row_search

'format column width to match headers
Cells(1, lastcol + 1).Columns.AutoFit
Cells(1, lastcol + 2).Columns.AutoFit
Cells(1, lastcol + 3).Columns.AutoFit
Cells(1, lastcol + 4).Columns.AutoFit

'add greatest headers
Cells(1, lastcol + 8).Value = "Ticker"
Cells(1, lastcol + 9).Value = "Value"
Cells(2, lastcol + 7).Value = "Greatest % Increase"
Cells(3, lastcol + 7).Value = "Greatest % Decrease"
Cells(4, lastcol + 7).Value = "Greatest Total Volume"

'insert greatest values
lastrow = row_summary

Dim greatest_inc
greatest_inc = Cells(2, lastcol + 3)
Dim greatest_dec
greatest_inc = Cells(2, lastcol + 3)
Dim greatest_vol
greatest_vol = Cells(2, lastcol + 4)

Dim greatest_inc_ID
Dim greatest_dec_ID
Dim greatest_vol_ID

    'review summary table for greatest values
    For row_summary = 3 To lastrow
        If Cells(row_summary, lastcol + 3) > greatest_inc Then
        greatest_inc = Cells(row_summary, lastcol + 3)
        greatest_inc_ID = Cells(row_summary, lastcol + 1)
        
    
        ElseIf Cells(row_summary, lastcol + 3) < greatest_dec Then
        greatest_dec = Cells(row_summary, lastcol + 3)
        greatest_dec_ID = Cells(row_summary, lastcol + 1)
        
        End If
        
        If Cells(row_summary, lastcol + 4) > greatest_vol Then
        greatest_vol = Cells(row_summary, lastcol + 4)
        greatest_vol_ID = Cells(row_summary, lastcol + 1)
        
        End If
    
    Next row_summary
    
'report greatest values
Cells(2, lastcol + 9).Value = FormatPercent(greatest_inc)
Cells(2, lastcol + 8).Value = greatest_inc_ID
Cells(3, lastcol + 9).Value = FormatPercent(greatest_dec)
Cells(3, lastcol + 8).Value = greatest_dec_ID
Cells(4, lastcol + 9).Value = greatest_vol
Cells(4, lastcol + 8).Value = greatest_vol_ID
        
'format greatest table column widths
Range("N:N").Columns.AutoFit
Range("O:O").Columns.AutoFit
Range("P:P").Columns.AutoFit

'Open next worksheet and repeat
Next ws

'All done, so return to the starting sheet
starting_ws.Activate


End Sub

