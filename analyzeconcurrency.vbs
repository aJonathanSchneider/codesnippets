'
' Developed by and copyright by Jonathan Schneider (except otherwise noted)
'
' with this code snippet you can analyse data sets that overlap in a certain period of time and count the
' number of overlaps. You can do this in two ways:
' 1) SI:  get the maximum number of overlaps at one point in time (using Marzullo's algorithm)
' 2) SI4: get the maximum number of overlaps that lasted a given period of time, demo data is ~4 years: 4*364 days
'
'
' note: the calculation result will be placed in column_targetSI or column_targetSItimePeriod respectively.
'       SI stems from the name of the original data set; simultaneous investments. SI4; simultaneous investments
'		for a period of 4 years

Rem Attribute VBA_ModuleType=VBAModule
Option VBASupport 1
'SI settings: adjust to your column headings
Public Const sourceTableName = "Table1"
Public Const iteratorColumn = "id"
Public Const column_iterator = "id"
Public Const column_timeFr_start = "startDate"
Public Const column_timeFr_end = "endDate"
Public Const column_groupBy = "group"
Public Const column_targetSI = "SI"
Public Const column_targetSItimePeriod = "SI4" 'SI_period

'SI_period settings (4 in this case)
Public Const timeFrameLength = 4 'in years


'working variables
Dim colId_iterator As Integer
Dim colId_timeFr_start As Integer
Dim colId_timeFr_end As Integer
Dim colId_groupBy As Integer
Dim colId_targetSI As Integer
Dim colId_targetSItimePeriod As Integer
Dim dsLen As Integer 'height of the data set
Dim dsFirstCol As Integer
Dim dsLastCol As Integer 'width of the data set
Dim shSource As Worksheet
Dim analyzedRange As Range
Dim timeFrameLengthInternal


Sub calculateSI()
    'meta data generation
    Set shSource = ThisWorkbook.Sheets(sourceTableName)
    Call getColumnIndices
    Call getDataSetLength
    timeFrameLengthInternal = timeFrameLength * 364
    'get relevant range
    With shSource
        Set analyzedRange = .Range(.Cells(2, dsFirstCol), .Cells(dsLen, dsLastCol))
    End With
    'calculate SI4
    Dim i As Integer
    For i = 1 To dsLen
        Dim si4Return As Integer
        si4Return = getSI4forRow(i)
        shSource.Cells(i + 1, colId_targetSItimePeriod).Value = si4Return
    Next i
    'calculate SI
    For i = 1 To dsLen
        Dim siReturn As Integer
        siReturn = getSIforRow(i)
        shSource.Cells(i + 1, colId_targetSI).Value = siReturn
    Next i

    On Error GoTo ErrorHandler
    MsgBox column_targetSI & " und " & column_targetSItimePeriod & " erfolgreich neu gefüllt"
    Exit Sub
    
ErrorHandler:
    MsgBox "Error! While creating " & column_targetSI & " and " & column_targetSItimePeriod & " an error has occured. Details: " + Err.Description
End Sub

Private Function getColumnIndices()
    Dim colHeadi As Integer
    Dim curCol As String
    colHeadi = 1
    Do
        curCol = ThisWorkbook.Sheets(sourceTableName).Cells(1, colHeadi)
        If curCol = "" Then
            Exit Do
        End If
        Select Case curCol
            Case column_iterator
                colId_iterator = colHeadi
            Case column_timeFr_start
                colId_timeFr_start = colHeadi
            Case column_timeFr_end
                colId_timeFr_end = colHeadi
            Case column_groupBy
                colId_groupBy = colHeadi
            Case column_targetSI
                colId_targetSI = colHeadi
            Case column_targetSItimePeriod
                colId_targetSItimePeriod = colHeadi
        End Select
        colHeadi = colHeadi + 1
    Loop
    dsFirstCol = Application.WorksheetFunction.Min(colId_iterator, colId_timeFr_start, colId_timeFr_end, colId_groupBy)
    dsLastCol = Application.WorksheetFunction.Max(colId_iterator, colId_timeFr_start, colId_timeFr_end, colId_groupBy)
    colId_iterator = colId_iterator - dsFirstCol + 1
    colId_timeFr_start = colId_timeFr_start - dsFirstCol + 1
    colId_timeFr_end = colId_timeFr_end - dsFirstCol + 1
    colId_groupBy = colId_groupBy - dsFirstCol + 1
End Function

Private Function getDataSetLength()
    dsLen = 2
    While (ThisWorkbook.Sheets(sourceTableName).Cells(dsLen, 1) <> "")
        dsLen = dsLen + 1
    Wend
    dsLen = dsLen - 2
End Function

Private Function getSI4forRow(pRow As Integer) As Integer
    Dim baseIterator As String, baseGroup As String
    Dim baseTFstart As Date, baseTFend As Date
    Dim curIterator As String, curGroup As String
    Dim curTFstart As Date, curTFend As Date
    Dim curRow As Range
    Dim validRange As Range
    'data for the base row
    baseIterator = analyzedRange.Cells(pRow, colId_iterator)
    baseTFstart = analyzedRange.Cells(pRow, colId_timeFr_start)
    baseTFend = analyzedRange.Cells(pRow, colId_timeFr_end)
    baseGroup = analyzedRange.Cells(pRow, colId_groupBy)
    'exit if base row doesn't have a sufficient timeframe
    If baseTFend - baseTFstart < timeFrameLengthInternal Then
        getSI4forRow = 0
        Exit Function
    End If
    'create clean arrays with relevant data
    Dim startDates() As Date
    Dim endDates() As Date
    Dim datesArrLen As Integer
    datesArrLen = -1
    Dim i As Integer
    For i = 1 To dsLen
        With analyzedRange
            Set curRow = .Range(.Cells(i, 1), .Cells(i, .Columns.Count))
        End With
        curIterator = curRow(0, colId_iterator)
        curTFstart = curRow(0, colId_timeFr_start)
        curTFend = curRow(0, colId_timeFr_end)
        curGroup = curRow(0, colId_groupBy)
        'simple exit criteria
        Dim isValidRow As Boolean
        isValidRow = True
        If pRow = i Then
            isValidRow = False
        End If
        If isValidRow And baseGroup <> curGroup Then
            isValidRow = False
        End If
        If isValidRow And curTFend < baseTFstart And curTFstart > baseTFend Then
            isValidRow = False
        End If
        Dim testtt
        testtt = curTFend - curTFstart
        If isValidRow And (curTFend - curTFstart) < timeFrameLengthInternal Then
            isValidRow = False
        End If
        If isValidRow Then
            datesArrLen = datesArrLen + 1
            ReDim Preserve startDates(datesArrLen)
            ReDim Preserve endDates(datesArrLen)
            startDates(datesArrLen) = curTFstart
            endDates(datesArrLen) = curTFend
        End If
    Next i
    'analyze the cleaned array
    Dim j As Integer
    Dim validTFcounter
    validTFcounter = 0
    For j = 0 To datesArrLen
        If startDates(j) < baseTFstart And (endDates(j) - baseTFstart) > timeFrameLengthInternal Then
            validTFcounter = validTFcounter + 1
        Else
            If endDates(j) > baseTFend And (baseTFend - startDates(j)) > timeFrameLengthInternal Then
                validTFcounter = validTFcounter + 1
            End If
            If endDates(j) <= baseTFend And startDates(j) >= baseTFstart And endDates(j) - startDates(j) > timeFrameLengthInternal Then
                validTFcounter = validTFcounter + 1
            End If
        End If
    Next j

    getSI4forRow = validTFcounter
End Function

'pRow is the row inside analyzedRange
Private Function getSIforRow(pRow As Integer) As Integer
    Dim baseIterator As String, baseGroup As String
    Dim baseTFstart As Date, baseTFend As Date
    Dim curIterator As String, curGroup As String
    Dim curTFstart As Date, curTFend As Date
    Dim curRow As Range
    Dim validRange As Range
    'data for the base row
    baseIterator = analyzedRange.Cells(pRow, colId_iterator)
    baseTFstart = analyzedRange.Cells(pRow, colId_timeFr_start)
    baseTFend = analyzedRange.Cells(pRow, colId_timeFr_end)
    baseGroup = analyzedRange.Cells(pRow, colId_groupBy)
    'create clean arrays with relevant data
    Dim startDates() As Date
    Dim endDates() As Date
    Dim datesArrLen As Integer
    datesArrLen = -1
    Dim i As Integer
    For i = 1 To dsLen
        With analyzedRange
            Set curRow = .Range(.Cells(i, 1), .Cells(i, .Columns.Count))
        End With
        curIterator = curRow(0, colId_iterator)
        curTFstart = curRow(0, colId_timeFr_start)
        curTFend = curRow(0, colId_timeFr_end)
        curGroup = curRow(0, colId_groupBy)
        'simple exit criteria
        Dim isValidRow As Boolean
        isValidRow = True
        If pRow = i Then
            isValidRow = False
        End If
        If isValidRow And baseGroup <> curGroup Then
            isValidRow = False
        End If
        If isValidRow And curTFend < baseTFstart And curTFstart > baseTFend Then
            isValidRow = False
        End If
        If isValidRow Then
            datesArrLen = datesArrLen + 1
            ReDim Preserve startDates(datesArrLen)
            ReDim Preserve endDates(datesArrLen)
            startDates(datesArrLen) = curTFstart
            endDates(datesArrLen) = curTFend
        End If
    Next i
    If datesArrLen = -1 Then
        getSIforRow = 0
        Exit Function
    End If
    'prepare data for Marzullo's algorithm
    'quick introduction to the algorithm: https://en.wikipedia.org/wiki/Marzullo%27s_algorithm
    Dim j As Integer
    Dim endDateIt As Integer
    Dim combinedArrLen As Integer
    combinedArrLen = (datesArrLen + 1) * 2 - 1
    Dim combinedArray()
    ReDim combinedArray(combinedArrLen, 1)
    For i = 0 To datesArrLen
        combinedArray(i, 0) = startDates(i)
        combinedArray(i, 1) = 1
    Next i
    For i = datesArrLen + 1 To combinedArrLen
        endDateIt = combinedArrLen - i
        combinedArray(i, 0) = endDates(endDateIt)
        combinedArray(i, 1) = -1
    Next i
    QuickSortArray combinedArray, , , 0
    'main part of Marzullo's algorithm
    Dim best As Integer
    Dim cnt As Integer
    Dim mType As Integer
    Dim bestStart
    Dim bestEnd
    best = 0
    cnt = 0
    For i = 0 To combinedArrLen
        mType = combinedArray(i, 1)
        cnt = cnt + mType
        If (best < cnt) Then
            best = cnt
            bestStart = combinedArray(i, 0)
            If (i < combinedArrLen) Then
                bestEnd = combinedArray(i + 1, 0)
            End If
        End If
    Next i
    getSIforRow = best
End Function


'VBA quicksort implementation source: http://stackoverflow.com/questions/4873182/sorting-a-multidimensionnal-array-in-vba
Public Sub QuickSortArray(ByRef SortArray As Variant, Optional lngMin As Long = -1, Optional lngMax As Long = -1, Optional lngColumn As Long = 0)
    On Error Resume Next

    'Sort a 2-Dimensional array

    ' SampleUsage: sort arrData by the contents of column 3
    '
    '   QuickSortArray arrData, , , 3

    '
    'Posted by Jim Rech 10/20/98 Excel.Programming

    'Modifications, Nigel Heffernan:

    '       ' Escape failed comparison with empty variant
    '       ' Defensive coding: check inputs

    Dim i As Long
    Dim j As Long
    Dim varMid As Variant
    Dim arrRowTemp As Variant
    Dim lngColTemp As Long

    If IsEmpty(SortArray) Then
        Exit Sub
    End If
    If InStr(TypeName(SortArray), "()") < 1 Then  'IsArray() is somewhat broken: Look for brackets in the type name
        Exit Sub
    End If
    If lngMin = -1 Then
        lngMin = LBound(SortArray, 1)
    End If
    If lngMax = -1 Then
        lngMax = UBound(SortArray, 1)
    End If
    If lngMin >= lngMax Then    ' no sorting required
        Exit Sub
    End If

    i = lngMin
    j = lngMax

    varMid = Empty
    varMid = SortArray((lngMin + lngMax) \ 2, lngColumn)

    ' We  send 'Empty' and invalid data items to the end of the list:
    If IsObject(varMid) Then  ' note that we don't check isObject(SortArray(n)) - varMid *might* pick up a valid default member or property
        i = lngMax
        j = lngMin
    ElseIf IsEmpty(varMid) Then
        i = lngMax
        j = lngMin
    ElseIf IsNull(varMid) Then
        i = lngMax
        j = lngMin
    ElseIf varMid = "" Then
        i = lngMax
        j = lngMin
    ElseIf VarType(varMid) = vbError Then
        i = lngMax
        j = lngMin
    ElseIf VarType(varMid) > 17 Then
        i = lngMax
        j = lngMin
    End If

    While i <= j
        While SortArray(i, lngColumn) < varMid And i < lngMax
            i = i + 1
        Wend
        While varMid < SortArray(j, lngColumn) And j > lngMin
            j = j - 1
        Wend

        If i <= j Then
            ' Swap the rows
            ReDim arrRowTemp(LBound(SortArray, 2) To UBound(SortArray, 2))
            For lngColTemp = LBound(SortArray, 2) To UBound(SortArray, 2)
                arrRowTemp(lngColTemp) = SortArray(i, lngColTemp)
                SortArray(i, lngColTemp) = SortArray(j, lngColTemp)
                SortArray(j, lngColTemp) = arrRowTemp(lngColTemp)
            Next lngColTemp
            Erase arrRowTemp

            i = i + 1
            j = j - 1
        End If
    Wend

    If (lngMin < j) Then Call QuickSortArray(SortArray, lngMin, j, lngColumn)
    If (i < lngMax) Then Call QuickSortArray(SortArray, i, lngMax, lngColumn)

End Sub
