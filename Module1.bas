Attribute VB_Name = "Module1"
Option Explicit
Sub BuildList()
    Dim shipSchedule As Workbook, wB As Workbook
    Dim shipSheet As Worksheet
    Dim bSchedIsOpen As Boolean, bContinue As Boolean
    Dim sWbPath As String, sWbFilename As String, attWbPath As String
    Dim sPassword As String
    Dim qStartRow As Integer, qEndRow As Integer
    Dim prodLineRows() As Integer, prodLines() As String, outInfo() As String
    Dim mStartRow As Integer, mEndRow As Integer
    Dim x As Integer, i As Integer, j As Integer
    Dim varStep As Variant, varAns As Variant
    Dim outProdLines() As String, outCustNames() As String, outCoNums() As String
    Dim outDescriptions() As String, outPrices() As String, outComments() As String
    
    ''''Hardcoded values''''''''''''''
    sPassword = "T"
    ''''''''''''''''''''''''''''''''''
    
    'On Error GoTo errhandler
    
    sWbPath = GetShipSchedulePath
    
    If sWbPath = "" Then Exit Sub
    
    If Month(Date) = 1 Then
        varAns = MsgBox("Run for the month of " & Sheet2.Range("M" & Month(Date)).Value & "?" & vbCrLf & vbCrLf & _
                "(choose 'No' to run the program for " & Sheet2.Range("M12").Value & ")", vbYesNo)
    Else
        varAns = MsgBox("Run for the month of " & Sheet2.Range("M" & Month(Date)).Value & "?" & vbCrLf & vbCrLf & _
                "(choose 'No' to run the program for " & Sheet2.Range("M" & Month(Date) - 1).Value & ")", vbYesNo)
    End If
    If varAns = vbYes Then
        gMonthNum = Month(Date)
    Else
        gMonthNum = Month(Date) - 1
    End If
    
    sWbFilename = Right(sWbPath, Len(sWbPath) - InStrRev(sWbPath, "\")) 'get filename
    sWbPath = Left(sWbPath, InStr(sWbPath, sWbFilename) - 1) 'get path directory w/o filename
    
    If Right(sWbPath, 1) <> "\" Then sWbPath = sWbPath & "\"
    
    bSchedIsOpen = False
    
    For Each wB In Application.Workbooks
        If wB.Name = sWbFilename Then 'assuming only one period
            Set shipSchedule = wB
            bSchedIsOpen = True
            Exit For
        End If
    Next

    If shipSchedule Is Nothing Then 'open it
        Application.DisplayAlerts = False
        Set shipSchedule = Workbooks.Open(sWbPath & sWbFilename)
    End If
    
    shipSchedule.Activate
    shipSchedule.Worksheets(1).Activate
    
    'find start & end rows for current quarter in ship schedule
    qStartRow = QuarterStartRow(shipSchedule)
    If qStartRow = 0 Then Exit Sub
    qEndRow = QuarterEndRow(shipSchedule)
    If qEndRow = 0 Then Exit Sub
    
    x = 0 'get product lines & corresponding rows in ship schedule
    ReDim prodLineRows(x)
    ReDim prodLines(x)
    For i = qStartRow To qEndRow
        If Cells(i, 1).Font.Bold And Cells(i, 1).Font.Underline <> 2 And Cells(i, 1).Value > 0 Then
            ReDim Preserve prodLineRows(x + 1)
            ReDim Preserve prodLines(x)
            prodLineRows(x) = i
            prodLines(x) = Cells(i, 1).Value
            x = x + 1
        End If
    Next
    prodLineRows(x) = qEndRow 'required because last product line needs an "end row" for the find month functions
    
    x = 0
    
    On Error GoTo 0 'forerr 'sometimes array length is wrong -> index error
    For i = 0 To UBound(prodLines) - LBound(prodLines) 'for each starting row of a product line
        mStartRow = MonthStartRow(shipSchedule, prodLineRows(i), prodLineRows(i + 1))
        mEndRow = MonthEndRow(shipSchedule, prodLineRows(i), prodLineRows(i + 1))
        Debug.Print prodLines(i) & "    strt:"; mStartRow & "    end:" & mEndRow
        If mStartRow = 0 Or mEndRow = 0 Then GoTo forerr
        For j = mStartRow To mEndRow
            ReDim outInfo(5) 'one of these (#5) is an indicator of failure/success
            outInfo = GetDetails(shipSchedule, j)
            If outInfo(5) <> "OK" Then Exit Sub
            If outInfo(1) > "" Then 'CO num was found
                ReDim Preserve outProdLines(x)
                ReDim Preserve outCustNames(x)
                ReDim Preserve outCoNums(x)
                ReDim Preserve outDescriptions(x)
                ReDim Preserve outPrices(x)
                ReDim Preserve outComments(x)
                outProdLines(x) = UCase(prodLines(i))
                outCustNames(x) = outInfo(0)
                outCoNums(x) = outInfo(1)
                outDescriptions(x) = outInfo(2)
                outPrices(x) = outInfo(3)
                outComments(x) = outInfo(4)
                x = x + 1
            End If
        Next
    Next
forerr:
    Err.Clear
    On Error GoTo errhandler

    ThisWorkbook.Activate
    Sheet1.Activate
    bContinue = UpdateWbInfo(shipSchedule, sPassword)
    If Not bSchedIsOpen Then 'schedule wasn't open initially
        'shipSchedule.Close savechanges:=False
    End If
    
    If x = 0 Then 'no results found
        MsgBox "No results found"
        Exit Sub
    End If
    Debug.Print "1"
    attWbPath = MakeWorkbook(outProdLines, outCustNames, outCoNums, outDescriptions, outPrices, outComments)
    Debug.Print "2"
    bContinue = DraftEmail(attWbPath)
    Exit Sub
errhandler:
    MsgBox "Error in BuildList sub"

End Sub

Sub UpdateSchedule()

    Dim shipSchedule As Workbook, wB As Workbook
    Dim shipSheet As Worksheet
    Dim bSchedIsOpen As Boolean, bContinue As Boolean, validMonth As Boolean
    Dim sWbPath As String, sWbFilename As String
    Dim sPassword As String, sContinue As String
    Dim qStartRow As Integer, qEndRow As Integer
    Dim prodLineRows() As Integer, prodLines() As String
    Dim a As Integer, b As Integer 'shipped
    Dim i As Integer, j As Integer 'moved
    Dim x As Integer, y As Integer 'partial
    Dim foo As Integer, bar As Integer 'counters/iterables
    Dim cosShipped() As Long, cosToMove() As Long, cosPartial() As Long
    Dim failedShipCOs() As String, failedMoveCOs() As String, failedParCos() As String
    Dim oldMonths() As String, newMonths() As String
    Dim firMonths() As String, secMonths() As String, parAmts() As Single
    Dim varCO As Variant, varAns As Variant
    Dim markShipped As Boolean, newMonth As Boolean, partialShip As Boolean
    
    ''''Hardcoded values''''''''''''''
    sPassword = "T"
    ''''''''''''''''''''''''''''''''''
    
    On Error GoTo errhandler
    
    'build CO lists
    ThisWorkbook.Activate
    Sheet2.Activate
    a = -1 'shipped
    i = -1 'moved
    x = -1 'partial
    ReDim cosShipped(0)
    ReDim cosToMove(0)
    ReDim newMonths(0)
    ReDim oldMonths(0)
    ReDim cosPartial(0)
    ReDim oldParMonths(0)
    ReDim newParMonths(0)
    ReDim parAmts(0)
    
    If WorksheetFunction.CountA(Range("F:F")) < 2 Then Exit Sub
    
    If Month(Date) = 1 Then
        varAns = MsgBox("Run for the month of " & Sheet2.Range("M" & Month(Date)).Value & "?" & vbCrLf & vbCrLf & _
                "(choose 'No' to run the program for " & Sheet2.Range("M12").Value & ")", vbYesNo)
    Else
        varAns = MsgBox("Run for the month of " & Sheet2.Range("M" & Month(Date)).Value & "?" & vbCrLf & vbCrLf & _
                "(choose 'No' to run the program for " & Sheet2.Range("M" & Month(Date) - 1).Value & ")", vbYesNo)
    End If
    If varAns = vbYes Then
        gMonthNum = Month(Date)
    Else
        gMonthNum = Month(Date) - 1
    End If
    
    If Range("F3").Value > 0 Then '
        foo = 3
        Do While Range("F" & foo).Value > 0
            If Not IsNumeric(Range("F" & foo).Value) Then 'can't use
                If Len(Range("F" & foo).Value) > 6 Then 'maybe it's "CO" followed by #
                    Range("F" & foo).Value = Right(Range("F" & foo).Value, 6)
                End If
            End If
            If Not IsNumeric(Range("F" & foo).Value) Then 'can't use
                MsgBox "Only enter CO numbers in column F"
                Exit Sub
            End If
            
            markShipped = False 'see if shipped or moved or partial
            newMonth = False
            partialShip = False
            
            If Range("G" & foo).Value > 0 Then 'shipped
                markShipped = True
            End If
            
            If Range("H" & foo).Value > 0 Or Range("I" & foo).Value > 0 Then 'new month
                If markShipped Then 'too many sections filled out
                    MsgBox "Too many columns filled out for row " & foo & ", CO number " & Range("F" & foo).Value
                    Exit Sub
                End If
                If Range("H" & foo).Value > 0 And Range("I" & foo).Value > 0 Then
                    newMonth = True
                Else 'only one filled out
                    MsgBox "Both an old month and a new month must be filled out" & vbCrLf & vbCrLf & _
                            "(check row " & foo & ", CO number " & Range("F" & foo).Value & ")"
                    Exit Sub
                End If
            End If
            
            If Range("J" & foo).Value > 0 Or Range("K" & foo).Value > 0 Or Range("L" & foo).Value > 0 Then
                If markShipped Or newMonth Then 'too many sections filled out
                    MsgBox "Too many columns filled out for row " & foo & ", CO number " & Range("F" & foo).Value
                    Exit Sub
                End If
                If Range("J" & foo).Value > 0 And Range("K" & foo).Value > 0 And Range("L" & foo).Value > 0 Then
                    newMonth = True
                Else 'only one filled out
                    MsgBox "A first month, a value, and a second month must be filled out" & vbCrLf & vbCrLf & _
                            "(check row " & foo & ", CO number " & Range("F" & foo).Value & ")"
                    Exit Sub
                End If
            End If
            
            If markShipped Then
                a = a + 1
                ReDim Preserve cosShipped(a)
                cosShipped(a) = Range("F" & foo).Value
            ElseIf newMonth Then
                validMonth = False 'check first month
                For bar = 1 To 12
                    If UCase(Range("H" & foo).Value) = UCase(Range("N" & bar).Value) Then
                        validMonth = True
                        Exit For
                    End If
                Next bar
                If Not validMonth Then
                    MsgBox "Issue with old month in row " & foo
                    Exit Sub
                End If
                validMonth = False 'check second month
                For bar = 1 To 12
                    If UCase(Range("I" & foo).Value) = UCase(Range("N" & bar).Value) Then
                        validMonth = True
                        Exit For
                    End If
                Next bar
                If Not validMonth Then
                    MsgBox "Issue with new month in row " & foo
                    Exit Sub
                End If
                i = i + 1
                ReDim Preserve cosToMove(i)
                ReDim Preserve newMonths(i)
                ReDim Preserve oldMonths(i)
                cosToMove(i) = Range("F" & foo).Value
                oldMonths(i) = Range("H" & foo).Value
                newMonths(i) = Range("I" & foo).Value
            ElseIf partialShip Then
                x = x + 1
                ReDim Preserve cosPartial(x)
                ReDim Preserve firMonths(x)
                ReDim Preserve secMonths(x)
                ReDim Preserve parAmts(x)
                cosPartial(x) = Range("F" & foo).Value
                firMonths(x) = Range("J" & foo).Value
                parAmts(x) = Range("K" & foo).Value
                secMonths(x) = Range("L" & foo).Value
            Else
                MsgBox "Each row must have an indication of what to do with the CO" & vbCrLf & vbCrLf & _
                        "(Check row " & foo & ", CO number " & Range("F" & foo).Value & ")"
                Exit Sub
            End If
            
            foo = foo + 1
        Loop
    Else
        MsgBox "No CO listed in F3. Enter the first CO number in F3 and don't skip any rows."
        Exit Sub
    End If
    
    If a < 0 And i < 0 And x < 0 Then Exit Sub 'no nada -- shouldn't be reachable
    
    'sWbPath = GetShipSchedulePath
    sWbPath = "C:\users\englandt\desktop\FY20 PSA Clearwater Ship Schedule.xlsx"
    
    If sWbPath = "" Then Exit Sub
    
    sWbFilename = Right(sWbPath, Len(sWbPath) - InStrRev(sWbPath, "\")) 'get filename
    sWbPath = Left(sWbPath, InStr(sWbPath, sWbFilename) - 1) 'get path directory w/o filename
    
    If Right(sWbPath, 1) <> "\" Then sWbPath = sWbPath & "\"
    
    bSchedIsOpen = False
    
    For Each wB In Application.Workbooks
        If wB.Name = sWbFilename Then
            Set shipSchedule = wB
            bSchedIsOpen = True
            Exit For
        End If
    Next
    
    If shipSchedule Is Nothing Then 'open it
        Application.DisplayAlerts = False
        Set shipSchedule = Workbooks.Open(Filename:=sWbPath & sWbFilename, IgnoreReadOnlyRecommended:=True)
    End If
    
    If shipSchedule.ReadOnly Then
        MsgBox "The Shipping schedule is read only"
        Exit Sub
    End If
    
    shipSchedule.Activate
    shipSchedule.Worksheets(1).Activate
    
    'find start & end rows for current quarter in ship schedule
    qStartRow = QuarterStartRow(shipSchedule)
    If qStartRow = 0 Then Exit Sub
    qEndRow = QuarterEndRow(shipSchedule)
    If qEndRow = 0 Then Exit Sub
    
    foo = 0 'get product lines & corresponding rows in ship schedule
    ReDim prodLineRows(foo)
    ReDim prodLines(foo)
    For bar = qStartRow To qEndRow
        If Cells(bar, 1).Font.Bold Then
            ReDim Preserve prodLineRows(foo + 1)
            ReDim Preserve prodLines(foo)
            prodLineRows(foo) = bar
            prodLines(foo) = Cells(bar, 1).Value
            foo = foo + 1
        End If
    Next
    prodLineRows(foo) = qEndRow 'required because last product line needs an "end row" for the find month functions
    
    'for shipped co's -> mark as shipped
    ReDim failedShipCOs(0)
    b = 0 'for ship errors
    If a >= 0 Then
        For Each varCO In cosShipped
            sContinue = MarkAsShipped(shipSchedule, varCO)
            If InStr(UCase(sContinue), "FALSE") > 0 Then
                ReDim Preserve failedShipCOs(b)
                failedShipCOs(b) = varCO & " " & Right(sContinue, Len(sContinue) - 5) '"FALSE" is 5 char
                b = b + 1
            End If
        Next varCO
    End If
'
'    'for moving co's ->
'    ReDim failedMoveCOs(0)
'    j = 0 'for moving errors
'    If i >= 0 Then
'        For foo = 0 To UBound(cosToMove)
'            sContinue = MoveShipMonth(shipSchedule, cosToMove(foo), oldMonths(foo), newMonths(foo))
'            If InStr(UCase(sContinue), "FAIL") > 0 Then
'                ReDim Preserve failedMoveCOs(j)
'                failedMoveCOs(j) = cosToMove(foo) & " " & Right(sContinue, Len(sContinue) - 5) '"FALSE" is 5 char
'                j = j + 1
'            End If
'        Next foo
'    End If
'
'    'for partial ship co's ->
'    ReDim failedParCos(0)
'    y = 0 'for moving errors
'    If x >= 0 Then
'        For foo = 0 To UBound(cosPartial)
'            sContinue = PartialShipment(shipSchedule, cosPartial(foo), firMonths(foo), parAmts(foo), secMonths(foo))
'            If InStr(UCase(sContinue), "FAIL") > 0 Then
'                ReDim Preserve failedParCos(y)
'                failedParCos(y) = cosPartial(foo) & " " & Right(sContinue, Len(sContinue) - 5) '"FALSE" is 5 char
'                y = y + 1
'            End If
'        Next foo
'    End If
    
    bContinue = UpdateWbInfo(shipSchedule, sPassword)
    Sheet2.Activate
    
    If b > 0 Or j > 0 Or y > 0 Then 'some errors exist
        If b > 0 Then 'marking as shipped error
            MsgBox "The following CO's could not be marked as shipped:" & vbCrLf & vbCrLf & _
                    Join(failedShipCOs, vbCrLf)
        End If
        If j > 0 Then 'moving months error
            MsgBox "The following CO's could not be moved to the new month:" & vbCrLf & vbCrLf & _
                    Join(failedMoveCOs, vbCrLf)
        End If
        If y > 0 Then 'partial ship error
            MsgBox "The following CO's could not be moved to the new month:" & vbCrLf & vbCrLf & _
                    Join(failedParCos, vbCrLf)
        End If
    Else 'no errors
        Sheet2.Unprotect sPassword
        Range("F3:L500").ClearContents
        Sheet2.Protect sPassword
        MsgBox "All specified updates successful"
    End If
    
    shipSchedule.Activate
    Exit Sub
    
errhandler:
    Resume ending
ending:
    On Error Resume Next
    shipSchedule.Close savechanges:=False
    MsgBox "Error in UpdateSchedule sub"

End Sub
