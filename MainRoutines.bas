Attribute VB_Name = "MainRoutines"
Option Explicit
''!!!!!Execute Main Macro via 'CRTL + W' keyboard shortcut!!!!''''''''

'Main Support Functions''''''============================================================================================================================================
'''=================================================================================================================================================================
Private Function CalculateOpenClose(MinElements, MaxElements)
    ''''Used to calculate Year Change and % Change metrics from extracted data
    Dim CalcArr(1 To 1, 1 To 2)
    Dim xlFunction As WorksheetFunction
    
    Set xlFunction = Application.WorksheetFunction
    
    'Statement in place to prevent run-time error Divide by 0
    If MinElements(3) = 0 Or MaxElements(3) = 0 Then
        CalcArr(1, 1) = 0
        CalcArr(1, 2) = 0
    Else
        CalcArr(1, 1) = MaxElements(3) - MinElements(3)
        CalcArr(1, 2) = (MaxElements(3) / MinElements(3)) - 1
    End If
    
    CalculateOpenClose = CalcArr

End Function

Public Function GetMin(ItemID As String, StartPoint, EndPoint, fTable As Variant)
     
     Dim xlFunction As WorksheetFunction
     Dim i As Long
     Dim j As Long
     Dim Min As Double
     Dim TempArr()
     
     Set xlFunction = Application.WorksheetFunction
     
     'Dynamiclly resize temp array to handle Divide and Conquer process
     ReDim TempArr(StartPoint - StartPoint + 1 To EndPoint - StartPoint + 1, 1 To 3)
     
     'Extract and consolidate particular TickerID data points from master table
     For i = StartPoint To EndPoint
        For j = 1 To 3
                TempArr(i - StartPoint + 1, j) = fTable(i, j)
        Next j
     Next i
     'Identify oldest/Max date value within a particular TickerIDs mini array
    Min = xlFunction.Min(xlFunction.Index(TempArr, 0, 2))
    
    'Slice 1x3 array of data elements related to Max date in master table
    GetMin = xlFunction.Index(TempArr, xlFunction.Match(Min, xlFunction.Index(TempArr, 0, 2), 0), 0)
    
 
End Function

Public Function GetMax(ItemID As String, StartPoint, EndPoint, fTable As Variant)
     
     Dim xlFunction As WorksheetFunction
     Dim i As Long
     Dim j As Long
     Dim k As Long
     Dim IndexID
     Dim Max As Double
     Dim IndexDate
     Dim TempArr()
     
     Set xlFunction = Application.WorksheetFunction
     
     'Dynamiclly resize temp array to handle Divide and Conquer process
     ReDim TempArr(StartPoint - StartPoint + 1 To EndPoint - StartPoint + 1, 1 To 3)
     
     
     i = 0
     j = 0
     ' k used to offset noncontiguous Close price data
     k = 0
     
     'Extract and consolidate particular TickerID data points from master table
     For i = StartPoint To EndPoint
        For j = 1 To 3
            If j = 3 Then k = 3
                TempArr(i - StartPoint + 1, j) = fTable(i, j + k)
                k = 0
        Next j
     Next i
     
    'Identify oldest/Max date value within a particular TickerIDs mini array
    Max = xlFunction.Max(xlFunction.Index(TempArr, 0, 2))
    
    'Slice 1x3 array of data elements related to Max date in master table
    GetMax = xlFunction.Index(TempArr, xlFunction.Match(Max, xlFunction.Index(TempArr, 0, 2), 0), 0)
 
End Function

Function GetSheetName(Rng As Range) As String

    '''''''''''Support Routine
    
    'Capture property of sheet object
    GetSheetName = Rng.Parent.Name

End Function
'Main Sub-Routines''''''============================================================================================================================================
'''=================================================================================================================================================================
Sub PrepData()
    ''''''''''1st Routine
    Dim i As Integer
    Dim WS As Worksheet
    
    'Sort entire data set by TickerID  ascending and then by Volume descending for Divide and Conquer process
    Set WS = ActiveSheet
    Cells(1, 1).CurrentRegion.Select
    WS.Sort.SortFields.Clear
    WS.Sort.SortFields.Add2 Key:=Range( _
        Cells(2, 1), Cells(2, 1).End(xlDown)), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    WS.Sort.SortFields.Add2 Key:=Range( _
        Cells(2, 7), Cells(2, 7).End(xlDown)) _
        , SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:= _
        xlSortNormal
    With WS.Sort
        .SetRange Range(Cells(1, 1), Cells(1, 7).End(xlDown))
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    'Replace column headers with unique friendly name for Named Range based sheet name
    Range(Cells(1, 1), Cells(1, 7)).Value = Array("Ticker_" & GetSheetName(Cells(1, 1)), _
    "Date_" & GetSheetName(Cells(1, 1)), "Open_" & GetSheetName(Cells(1, 1)), "High_" & _
    GetSheetName(Cells(1, 1)), "Low_" & GetSheetName(Cells(1, 1)), "Close_" & _
    GetSheetName(Cells(1, 1)), "Vol_" & GetSheetName(Cells(1, 1)))

    'Suppress Excel prompts to overwrite existing named ranges and use default response
    Application.DisplayAlerts = False
    
    'Add data columns to unique named ranges based on column headers
    Cells(1, 1).CurrentRegion.CreateNames _
    Top:=True, Left:=False, Bottom:=False, Right:=False
    
End Sub

Sub PopulateHeaders()

    '''''''''''''''2nd Routine
    
    'Populate Headers for table calculations
    Range("I1:L1").Value = Array("Ticker", "Yearly Change", "Percent Change", "Total Stock Volume")
    Range("O1:P1").Value = Array("Ticker", "Value")
    Range("N2:N4").Value = Application.WorksheetFunction.Transpose(Array("Greatest % Increase", "Greatest % Decrease", "Greatest Total Volume"))

End Sub

Sub CaptureUniqueTickerID()

    ''''''''''''''''''''3rd Routine
    Dim ID_Range As Range
    Dim Row As Long

    'Define range area of Ticker ID column and set to range object variable
    Row = Cells(1, 1).CurrentRegion.Rows.Count
    Set ID_Range = Range(Cells(1, 1), Cells(Row, 1))

    'Dynamically copy range object variable and paste outside main table
    ID_Range.Copy Range("I1")

    'Assign temp range name to paste destination area and remove duplicate Ticker IDs
    Range("I1").CurrentRegion.Name = "TickerList"
    Range("Tickerlist").RemoveDuplicates Columns:=1, Header:= _
        xlYes
    'Clear out temp name range
    Range("TickerList").Name.Delete

    'Clear out object from memory
    Set ID_Range = Nothing

End Sub

Sub AggregationTransferArray()

    '''''''''''''''''''4th Routine
    Dim Cell As Range
    Dim FullRange As Range
    Dim UniqueRange As Range
    Dim RowCount As Long
    Dim StartPoint
    Dim EndPoint
    Dim TempArray()
    Dim TempSum As Double
    Dim i As Long
    Dim j As Long

    'Dimension range objects needed for array transfer
    Set FullRange = Range(Cells(2, 1), Cells(2, 1).End(xlDown))
    Set UniqueRange = Range(Cells(2, 9), Cells(2, 9).End(xlDown))
    RowCount = UniqueRange.Rows.Count
    
    'Compartmentalize(Divide and Conquer) TrickerID transactions for processing
    With UniqueRange.Range(Cells(1, 1), Cells(1, 1))
        .Offset(0, 1).Value = "=MATCH(I2, Ticker_" & GetSheetName(Cells(1, 1)) & ",0)"
        .Offset(0, 1).AutoFill UniqueRange.Range(Cells(1, 1), _
         Cells(RowCount, 1)).Offset(0, 1), xlFillDefault
        .Offset(0, 2).Value = "=IF(J3<>"""",J3-1,COUNTA(A:A)-1)"
        .Offset(0, 2).AutoFill UniqueRange.Range(Cells(1, 1), Cells(RowCount, 1)) _
        .Offset(0, 2), xlFillDefault
    End With
    
    'Switch to manual calculation to improve processing speed
    Application.Calculate
    Application.Calculation = xlCalculationManual
    
    'initialize loop iterators
    TempSum = 0
    i = 0
    j = 0
    
    'initialize Process Range point variables for aggregation criteria
    StartPoint = UniqueRange.Cells(1, 1).Offset(0, 1).Value
    EndPoint = UniqueRange.Cells(1, 1).Offset(0, 2).Value
    
    'Adjust container array size to fit aggregated Volumes
    ReDim TempArray(1 To UniqueRange.Rows.Count + 1, 1 To 1)
    
    'loop and sum volumes per TickerID
    For i = 2 To UniqueRange.Rows.Count + 1
        StartPoint = UniqueRange.Cells(i - 1, 1).Offset(0, 1).Value
        EndPoint = UniqueRange.Cells(i - 1, 1).Offset(0, 2).Value
        For j = StartPoint To EndPoint
                TempSum = TempSum + FullRange.Cells(j, 1).Offset(0, 6).Value
         Next j
         TempArray(i - 1, 1) = TempSum
         TempSum = 0
    Next i
    
    'Move processed Volume values back to sheet
    UniqueRange.Offset(0, 3) = TempArray
    
    Range(Cells(2, 9), Cells(2, 11).End(xlDown)).Offset(0, 10) _
    .Value = Range(Cells(2, 9), Cells(2, 11).End(xlDown)).Value
    Range(Cells(2, 10), Cells(2, 11).End(xlDown)).ClearContents
    
    'Clear object level variables from memory
    Set FullRange = Nothing
    Set UniqueRange = Nothing
    Set Cell = Nothing
    Erase TempArray
    
    'Restore calculation
    Application.Calculation = xlCalculationAutomatic
   
End Sub

 Sub PopulateYearlyPriceChange()
 ''''''''''''''''''''''''''5th Routine
    Dim xlFunction As WorksheetFunction
    Dim TrID As String
    Dim TableRng As Range
    Dim TableArr
    Dim UniqueArr
    Dim UniqueCount As Long
    Dim Item
    Dim MinElements
    Dim MaxElements
    Dim LocaterArr
    Dim StartPoint
    Dim EndPoint
    Dim OutPutRng As Range
    Dim Counter1 As Double
    Dim Counter2 As Double
    Dim TempArr()
    Dim x
    Dim ArrMarker As Double
    Dim tempMarker As Double
    Dim i As Long
    Dim j As Long
    
    Set xlFunction = Application.WorksheetFunction
    
    'Define data set for variant array
    Set TableRng = Cells(1, 1).CurrentRegion.Offset(1, 0)
    
    'Create variant arrays for data processing in memory
    TableArr = TableRng.Value
    UniqueArr = Range(Cells(2, 9), Cells(2, 9).End(xlDown)).Value
    LocaterArr = Range(Cells(2, 19), Cells(2, 21).End(xlDown)).Value
    UniqueCount = Range(Cells(2, 9), Cells(2, 9).End(xlDown)).Rows.Count + 1
    Set OutPutRng = Range(Cells(2, 10), Cells(UniqueCount, 11))
    ReDim TempArr(1 To UBound(UniqueArr, 1), 1 To 2)
    
    'Intialize Loop counter variables
    i = 1
    j = 0
    
    'Marker used capture next available dimension in array
    ArrMarker = 0
    
    ' For Each loop used to process all unique Ticker Ids to create mini-array tables
    For Each Item In UniqueArr
    
            'Test for counter variable values that exceed unique Ticker count
            If i > UBound(UniqueArr, 1) Then Exit For
            TrID = UniqueArr(i, 1)
            
            'Establish position to BEGIN processing like TickerIds
            StartPoint = xlFunction.Index(LocaterArr, xlFunction.Match _
            (TrID, xlFunction.Index(LocaterArr, 0, 1), 0), 2)
            
            'Establish position to STOP processing TickerIds to prevent unnecesary iterations
            EndPoint = xlFunction.Index(LocaterArr, xlFunction.Match _
            (TrID, xlFunction.Index(LocaterArr, 0, 1), 0), 3)
            
            'Process mini-array to capture raw data for Open(MinDate) and Close(Maxdate) price information
            MinElements = GetMin(TrID, StartPoint, EndPoint, TableArr)
            MaxElements = GetMax(TrID, StartPoint, EndPoint, TableArr)
            
            '1x3 horizontal array slice generated in function call after MinMax elements extracted
            'Elements further procesed to calculate requested metrics and temp assigned to variant
            x = CalculateOpenClose(MinElements, MaxElements)
            
            'Append processed metrics to final output table array
            For Counter1 = LBound(x, 1) To UBound(x, 1)
                If ArrMarker > UBound(UniqueArr, 1) Then Exit For
                tempMarker = 0
                For Counter2 = LBound(x, 2) To UBound(x, 2)
                    TempArr(Counter1 + ArrMarker, Counter2) = x(Counter1, Counter2)
                Next Counter2
                tempMarker = tempMarker + (Counter1)
                ArrMarker = ArrMarker + tempMarker
             Next Counter1
             Erase x
            i = i + 1
    Next Item
    ' Transfer final data array to sheet
    OutPutRng = TempArr
    
    'Run clean-up routines and format conditions
    Call rgPerformanceAndClean
    Call ClearNames
    
End Sub

Public Sub rgPerformanceAndClean()

    Dim xlFunction As WorksheetFunction
    Dim x
    Dim gPerInc As Double
    Dim gPerDecr As Double
    Dim gVolume As Double
    Dim gPerIncTRid As String
    Dim gPerDecrTRid As String
    Dim gVolumeTRid As String
    
    Set xlFunction = Application.WorksheetFunction
    
    'Variant array holding summarized values
    x = Range(Cells(2, 9), Cells(2, 12).End(xlDown)).Value
    

    
    'Find and extract Top Metrics
    gPerInc = xlFunction.Max(xlFunction.Index(x, 0, 3))
    gPerIncTRid = xlFunction.Index(x, xlFunction.Match(gPerInc, xlFunction.Index(x, 0, 3), 0), 1)
    gPerDecr = xlFunction.Min(xlFunction.Index(x, 0, 3))
    gPerDecrTRid = xlFunction.Index(x, xlFunction.Match(gPerDecr, xlFunction.Index(x, 0, 3), 0), 1)
    gVolume = xlFunction.Max(xlFunction.Index(x, 0, 4))
    gVolumeTRid = xlFunction.Index(x, xlFunction.Match(gVolume, xlFunction.Index(x, 0, 4), 0), 1)
            
    'Assign Top Metric values to worksheet and format summary data tables
    Range(Cells(2, 16), Cells(4, 16)).Value = xlFunction.Transpose(Array(gPerInc, gPerDecr, gVolume))
    Range(Cells(2, 15), Cells(4, 15)).Value = xlFunction.Transpose(Array(gPerIncTRid, gPerDecrTRid, gVolumeTRid))
    Range(Cells(1, 1), Cells(1, 16)).Font.Bold = True
    
    With Range(Cells(2, 9), Cells(2, 12).End(xlDown)).Borders
        .LineStyle = xlContinuous
        .Color = RGB(157, 157, 157)
    End With
    
    'Tidy basic sheet formatiing
    Range(Cells(2, 10), Cells(2, 10).End(xlDown)).NumberFormat = "0.00"
    Range(Cells(2, 11), Cells(2, 11).End(xlDown)).NumberFormat = "0.0%"
    Range(Cells(2, 12), Cells(2, 12).End(xlDown)).NumberFormat = "#,##0"
    Cells.Columns.AutoFit
    
    Range(Cells(2, 16), Cells(4, 16)).NumberFormat = "0.0%"
    Range(Cells(4, 16), Cells(4, 16)).NumberFormat = "#,##0"
    
    With Range(Cells(1, 14), Cells(4, 16))
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(157, 157, 157)
        .Font.Size = 10
        .Font.Bold = True
        .Font.Color = vbBlack
        .Interior.ThemeColor = xlThemeColorAccent4
        .Interior.TintAndShade = 0.599993896298105
    End With

    'Condtional formatting
    'Script for Positive Values
     With Range(Cells(2, 10), Cells(2, 11) _
        .End(xlDown)).FormatConditions _
        .Add(xlCellValue, xlGreater, 0)
        With .Interior
             .Color = RGB(27, 237, 107)
        End With
        With .Borders
            .LineStyle = xlContinuous
            .Color = RGB(157, 157, 157)
        End With
        With .Font
            .Bold = True
            .ColorIndex = vbBlack
        End With
    End With

    'Script for Negative Values
    With Range(Cells(2, 10), Cells(2, 11) _
        .End(xlDown)).FormatConditions _
        .Add(xlCellValue, xlLess, 0)
        With .Interior
             .Color = RGB(252, 54, 74)
        End With
        With .Borders
            .LineStyle = xlContinuous
            .Color = RGB(157, 157, 157)
        End With
        With .Font
            .Bold = True
            .Color = RGB(255, 255, 255)
        End With
    End With
    
    'Clear look-up table
    Range(Cells(2, 19), Cells(2, 21).End(xlDown)).ClearContents
    Sheet1.Activate
    Sheet1.Range(Cells(1, 1), Cells(1, 1)).Select
    
    
End Sub

Sub ClearNames()

Dim Name As Name

'Clear Named Ranges used in processing data
On Error Resume Next

For Each Name In Application.Names
    Name.Delete
Next Name


End Sub

Sub Main(FrmSelection As Boolean)

    Dim WS As Worksheet
    Dim ProcessAll As Integer
    Dim Ans As Integer
    Dim Cycle As Integer
    '''Master routine containing all sub-routines and directs flow based on user response'''
    
    'initialize decision boolean variable for prompted userform response
    ProcessAll = FrmSelection
    
    'Apply all sub-routines to data based on scope selected by user
    Select Case ProcessAll
    Case True
        Cycle = 1
        Application.StatusBar = "0%"
        
        For Each WS In ActiveWorkbook.Worksheets
            If WS.Name = "2014" Or WS.Name = "2015" Or WS.Name = "2016" Then
                Application.ScreenUpdating = False
                WS.Activate
                Call PrepData
                Application.StatusBar = "Completion Status: " & Format(Str(Cycle / 15), "0%")
                Cycle = Cycle + 1
                Call PopulateHeaders
                Application.StatusBar = "Completion Status: " & Format(Str(Cycle / 15), "0%")
                Cycle = Cycle + 1
                Call CaptureUniqueTickerID
                Application.StatusBar = "Completion Status: " & Format(Str(Cycle / 15), "0%")
                Cycle = Cycle + 1
                Call AggregationTransferArray
                Application.StatusBar = "Completion Status: " & Format(Str(Cycle / 15), "0%")
                Cycle = Cycle + 1
                Call PopulateYearlyPriceChange
                Cycle = Cycle + 1
                Application.StatusBar = "Completion Status: " & Format(Str(Cycle / 15), "0%")
            End If
        Next WS
        Application.ScreenUpdating = True
        
    Case Else
        Set WS = ThisWorkbook.ActiveSheet
        If WS.Name = "2014" Or WS.Name = "2015" Or WS.Name = "2016" Then
                Application.ScreenUpdating = False
                WS.Activate
                Call PrepData
                Application.StatusBar = "Completion Status: 20%"
                Call PopulateHeaders
                Application.StatusBar = "Completion Status: 40%"
                Call CaptureUniqueTickerID
                Application.StatusBar = "Completion Status: 60%"
                Call AggregationTransferArray
                Application.StatusBar = "Completion Status: 80%"
                Call PopulateYearlyPriceChange
                Application.StatusBar = "Completion Status: 100%"
        End If
        
        Application.ScreenUpdating = True
        Set WS = Nothing
    End Select
    
    Application.StatusBar = ""
    MsgBox "All Requested Summaries Hvae been Processed.", vbInformation
    
End Sub


