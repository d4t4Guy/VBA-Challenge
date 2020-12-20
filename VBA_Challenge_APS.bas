Attribute VB_Name = "VBA_Challenge_APS"
Option Explicit
'To do: add timer - not needed, script runs on big file in under 60 seconds


Sub ProcessWorkbook()
'Assumes every sheet contains input data
'Assumes input data are in cols A:G and output summary is in cols I:L
Dim wsX As Worksheet

On Error GoTo eh
For Each wsX In ActiveWorkbook.Worksheets
    wsX.Select
    Call PrepSheet 'not strictly necessary, but allows user to not have to assume input data are already sorted
    Call Process_Sheet 'calculates metrics for each symbol
    Call MaxMin  'Bonus: Outputs summary of top and bottom performers, and top volume
    
Next wsX
Exit Sub

eh:
    MsgBox ("Ya broke me. I can't go on like this!" & vbNewLine & Err.Number & " " & Err.Description)
    Exit Sub
End Sub





Sub PrepSheet()

Dim Last_Row As Long
Last_Row = Cells(Rows.Count, 1).End(xlUp).Row 'find end of input data
    
    'check if auto filter is on; turn it on if not (source: https://www.mrexcel.com/board/threads/quick-vba-question-checking-if-a-filter-is-applied.258782/   ''visited 2020-12-16)
    If Not ActiveSheet.AutoFilterMode Then
        Range("A1").AutoFilter
    End If
    
    'sort data
    ThisWorkbook.ActiveSheet.AutoFilter.Sort.SortFields.Clear
    ThisWorkbook.ActiveSheet.AutoFilter.Sort.SortFields.Add2 Key:=Range( _
        "A2:A" & Last_Row), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    ThisWorkbook.ActiveSheet.AutoFilter.Sort.SortFields.Add2 Key:=Range( _
        "B2:B" & Last_Row), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ThisWorkbook.ActiveSheet.AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

End Sub





Sub Process_Sheet()  'Makes assumptions about input and output data layout

Dim InputStart_Row, OutputStart_Row, Current_Row, Last_Row As Long

'initialize
InputStart_Row = 2 'Assumes input data has headers in row 1, and is continuous
Last_Row = Cells(Rows.Count, 1).End(xlUp).Row 'find end of input data
OutputStart_Row = 2

'Build headers for output
'Assumes input data are in cols A:G and output is in cols I:L

    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Volume"
    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Value"
    
    Cells(2, 15).Value = "Greatest % Increase"
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(4, 15).Value = "Greatest Total Volume"


'Process individual stocks
    Do Until Current_Row > Last_Row
        Current_Row = End_Row(InputStart_Row, OutputStart_Row) 'calls fn that processes individual rows for a single symbol
        OutputStart_Row = OutputStart_Row + 1
        InputStart_Row = Current_Row
    Loop

'Format numeric output
Range(ActiveSheet.Cells(2, 12), ActiveSheet.Cells(OutputStart_Row, 12)).Select
Selection.NumberFormat = "#,##0"
Range(ActiveSheet.Cells(2, 11), ActiveSheet.Cells(OutputStart_Row, 11)).Select
Selection.NumberFormat = "#,##0.00%;(#,##0.00)%"

'Add conditional formatting
Call RedGreen
Range("A1").Select

End Sub

Function End_Row(ByVal First_Row As Long, ByVal Output_Row As Long) As Long 'Contains assumptions about input data
'Takes output row, and First Row as inputs. Processes one stock at a time. Outputs summary stats for one stock and returns row number where next symbol starts


'Move first_row, output_Row to take as arg - DONE
'Dim First_Row, Output_Row As Long

'return end_row to row_index in process_Sheet() Fn
'Dim End_Row As Long

Dim Symbol As String

Dim Open_Col, Close_Col, Symbol_Col, Volume_Col, SymbolOut_Col, Delta_Col, Pct_Col, VolumeOut_Col As Integer
Dim Current_Row As Long
Dim Total_Volume As LongLong
Dim Open_Price, Close_Price, Price_Change As Currency
Dim Pct_Change As Single

'-----------ASSUMPTIONS--------------------------------------------------------
    Open_Col = 3 'assumes open prices are in column C
    Close_Col = 6 'assumes close prices are in column F
    Symbol_Col = 1 'assumes symbol is in column A
    Volume_Col = 7 'assumes volume is in column G
    
    'Output_Row = 3 ' move to args after testing
    
    SymbolOut_Col = Volume_Col + 2 'Assumes end of input columns is Volume_Col
    Delta_Col = SymbolOut_Col + 1
    Pct_Col = SymbolOut_Col + 2
    VolumeOut_Col = SymbolOut_Col + 3
'------------------------------------------------------------------------------

'init vars
    Current_Row = First_Row
    Total_Volume = 0
    Symbol = Cells(First_Row, Symbol_Col).Value

'Find end_row, and add up volume per day
    Do Until Cells(Current_Row, Symbol_Col).Value <> Cells(Current_Row + 1, Symbol_Col).Value 'repeat loop until symbol changes
        Total_Volume = Total_Volume + Cells(Current_Row, Volume_Col).Value 'Running total of volume
        Current_Row = Current_Row + 1
    Loop
    Total_Volume = Total_Volume + Cells(Current_Row, Volume_Col).Value
    
    
    'Get open and close price
    Open_Price = Cells(First_Row, Open_Col).Value
    Close_Price = Cells(Current_Row, Close_Col).Value
    
    'calc change
    Price_Change = Close_Price - Open_Price
    
    'calc pct change
    If Open_Price <= 0 Then
        Pct_Change = 0
        Else: Pct_Change = (Price_Change / Open_Price)
    End If
'outputs - to do: consider using array for speed
    Cells(Output_Row, SymbolOut_Col).Value = Symbol
    Cells(Output_Row, Delta_Col).Value = Price_Change
    Cells(Output_Row, Pct_Col).Value = Pct_Change
    Cells(Output_Row, VolumeOut_Col).Value = Total_Volume
    
    End_Row = Current_Row + 1
    

End Function

Sub RedGreen()
    Range("J2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlLessEqual, _
        Formula1:="=0"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 255
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
        Formula1:="=0"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 5287936
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False

End Sub
Sub MaxMin() 'BONUS: loads values from columns I:L into array and iterates over values to get max, min

Dim CompareData As Variant '2 dimensional array for values

Dim OutputEnd As Long 'numeric location of last row in output data
Dim strOutputRange As String 'not strictly required, but I find it easier to concat string separately, and then pass to Range() fn
Dim i As Integer 'counter

Dim MaxPct As Double 'holds max Pct change found in array
Dim MaxPctTkr As String 'for ticker symbol related to max pct change found in array

Dim MinPct As Double 'holds min pct change found in array
Dim MinPctTkr As String

Dim MaxVol As LongLong 'holds max volume found in array
Dim MaxVolTkr As String



    OutputEnd = Cells(Rows.Count, 11).End(xlUp).Row 'Goes to column 11 (i.e., column K) and finds bottom row populated.
    
    strOutputRange = "I2:L" & OutputEnd 'Dunno why this works, but it does. str(OutputEnd) puts a space in that I don't want.    input_data = Range(StrRangeOfK)
    
    CompareData = Range(strOutputRange).Value
    
'    Debug.Print strOutputRange
'    Debug.Print CompareData(289)
'    Debug.Print CompareData(LBound(CompareData), 1)
  'assign first value to max and min values
    MaxPct = CompareData(LBound(CompareData), 3)
    MinPct = CompareData(LBound(CompareData), 3)
    MaxVol = CompareData(LBound(CompareData), 4)
    
    For i = LBound(CompareData) To UBound(CompareData)
        If CompareData(i, 3) > MaxPct Then
            MaxPct = CompareData(i, 3)
            MaxPctTkr = CompareData(i, 1)
            
         ElseIf CompareData(i, 3) < MinPct Then
            MinPct = CompareData(i, 3)
            MinPctTkr = CompareData(i, 1)
            
        End If
        
        If CompareData(i, 4) > MaxVol Then
            MaxVol = CompareData(i, 4)
            MaxVolTkr = CompareData(i, 1)
        End If
        
    Next i

'Send output to cells
Range("P2").Value = MaxPctTkr
Range("P3").Value = MinPctTkr
Range("P4").Value = MaxVolTkr

Range("Q2").Value = MaxPct
Range("Q3").Value = MinPct
Range("Q4").Value = MaxVol

'format cells
Range("Q2:Q3").NumberFormat = "#,##0.00%;(#,##0.00)%"
Range("Q4").NumberFormat = "#,##0"
Columns("O:Q").EntireColumn.AutoFit


End Sub



