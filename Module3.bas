Attribute VB_Name = "Module3"
Function MarkAsShipped(wB As Workbook, coNum As Variant) As String

    Dim coRow As Integer, rowNum As Integer
    
    On Error GoTo errhandler
    
    MarkAsShipped = "FALSE"
    
    wB.Activate
    wB.Worksheets(1).Activate
    If WorksheetFunction.CountIf(Range("B:B"), coNum) = 0 Then
        MarkAsShipped = MarkAsShipped & "[This CO not found]"
        Exit Function
    ElseIf WorksheetFunction.CountIf(Range("B:B"), coNum) > 1 Then
        'check the months for each time the CO appears
        Dim i As Integer, j As Integer, monthNames() As String, coRows() As Integer, x As Integer, varIt As Variant
        rowNum = 1
        x = 0
        ReDim monthNames(0)
        For i = 1 To WorksheetFunction.CountIf(Range("B:B"), coNum)
            rowNum = Range("B:B").Find(what:=coNum, lookat:=xlWhole, after:=Range("B" & rowNum + 1)).Row
            ReDim Preserve coRows(x)
            coRows(x) = rowNum
            Do While UCase(Range("C" & rowNum)) <> "OPPORTUNITIES"
                rowNum = rowNum + 1
                If rowNum > 5000 Then
                    MarkAsShipped = MarkAsShipped & "[Issue locating month for an instance of this CO number]"
                    Exit Function
                End If
            Loop
            ReDim Preserve monthNames(x)
            monthNames(x) = Right(Range("C" & rowNum + 1).Value, Len(Range("C" & rowNum + 1).Value) - 6)
            x = x + 1
        Next
        'see if there's only one left for this month
        i = 0
        j = -1
        For Each varIt In monthNames
            j = j + 1
            If UCase(varIt) = UCase(ThisWorkbook.Worksheets(2).Range("M" & gMonthNum).Value) Then 'this month
                x = j
                i = i + 1
            End If
        Next
        If i <> 1 Then '# of times the CO is under this month is <> 1
            If i = 0 Then
                MarkAsShipped = MarkAsShipped & "[This CO doesn't appear under this month]"
            Else
                MarkAsShipped = MarkAsShipped & "[This CO appears multiple times for this month]"
            End If
            Exit Function
        End If
        coRow = coRows(x) 'this is the only row with this CO for this month
    Else
        coRow = Range("B:B").Find(what:=coNum, lookat:=xlWhole).Row
        rowNum = coRow
        Do While UCase(Range("C" & rowNum)) <> "OPPORTUNITIES" 'check that it's in this month
            rowNum = rowNum + 1
            If rowNum > 5000 Then
                MarkAsShipped = MarkAsShipped & "[Issue locating month for an instance of this CO number]"
                Exit Function
            End If
        Loop
        If UCase(Right(Range("C" & rowNum + 1).Value, Len(Range("C" & rowNum + 1).Value) - 6)) <> _
                            UCase(ThisWorkbook.Worksheets(2).Range("M" & gMonthNum).Value) Then
            MarkAsShipped = MarkAsShipped & "[This CO doesn't appear under this month]"
            Exit Function
        End If
    End If
    
    If Range("G" & coRow).Interior.ColorIndex <> 6 Then
        MarkAsShipped = MarkAsShipped & "[Order value not highlighted yellow]"
        Exit Function
    ElseIf UCase(Range("L" & coRow).Value) = "SHIPPED" Then
        MarkAsShipped = MarkAsShipped & "[This order already marked shipped]"
        Exit Function
    End If
    
    Range("G" & coRow).Interior.ColorIndex = 0
    With Range("L" & coRow)
        .Value = "SHIPPED"
        .Interior.Color = 5287936
        .Font.Bold = True
        .Font.Underline = False
        .Font.Italic = False
    End With
    
    MarkAsShipped = "TRUE"
    Exit Function
errhandler:
    MarkAsShipped = MarkAsShipped & "[Unspecified issue]"
    
End Function

Function MoveShipMonth(wB As Workbook, coNum As Long, oldMonth As String, newMonth As String) As String
    Dim markShipped As Boolean
    markShipped = False
    If UCase(newMonth) = UCase(ThisWorkbook.Worksheets(2).Range("N" & gMonthNum)) Then
        ans = MsgBox("Do you want CO number " & coNum & " (moving to this month) to be marked as shipped?", vbYesNo)
        If ans = vbYes Then markShipped = True
    End If
    'find CO num in old month
    'determine whether move jumps quarters
    'find new home (same prod line)
    'move
End Function

Function PartialShipment(wB As Workbook, coNum As Long, Month1 As String, value1 As Single, Month2 As String) As String
    MsgBox coNum & ": not able to do partial shipments yet"
    'determine whether new month jumps quarters
    'find new month (same prod line)
    'change orig line value
    'create new line
End Function

