Attribute VB_Name = "Module2"
Public gMonthNum As Integer
Function GetShipSchedulePath() As String
    Dim shipSchedule As Workbook, wB As Workbook
    Dim sWbPath As String, sWbFilename As String
    Dim bAskUser As Boolean, i As Integer
    
    On Error GoTo errhandler
    
    ThisWorkbook.Activate
    Sheet1.Activate
    i = 1
    Do While Cells(1, i).Value = 0
        i = i + 1
        If i > 100 Then
            Exit Do
        End If
    Loop
    sWbPath = Cells(1, i).Value 'where the workbook was last found
    sWbFilename = Cells(2, i).Value 'what workbook was last called
    
    If sWbPath = "" Then 'make user choose it
        bAskUser = True
    Else
        If Right(sWbPath, 1) <> "\" Then sWbPath = sWbPath & "\"
        If Dir$(sWbPath & sWbFilename) > "" Then 'schedule is where it's expected to be
            bAskUser = False
            sWbFilename = Dir$(sWbPath & sWbFilename)
        Else
            bAskUser = True
            If Month(Date) < 10 Then 'current year = fiscal year
                sWbFilename = "FY" & Right(Str(Year(Date)), 2) & "*C*L*W*" 'try this filename
            Else 'next year = fiscal year
                sWbFilename = "FY" & Right(Str(Year(Date) + 1), 2) & "*C*L*W*" 'try this filename
            End If
            If Dir$(sWbPath & sWbFilename) > "" Then 'retry finding workbook
                bAskUser = False
                sWbFilename = Dir$(sWbPath & sWbFilename)
            End If
        End If
    End If
    
    If Not bAskUser Then
        GetShipSchedulePath = sWbPath & sWbFilename
        Exit Function
    End If
    
    MsgBox "You will have to choose which file is the shipping schedule"
    
    With Application.FileDialog(msoFileDialogFilePicker)
        .AllowMultiSelect = False
        .Filters.Add "Excel Files", "*.xlsx; *.xlsm; *.xls; *.xlsb", 1
        .Show
        
        GetShipSchedulePath = .SelectedItems.Item(1)
    End With
    
    Exit Function
    
errhandler:
    MsgBox "Error in GetShipSchedulePath function"
    
End Function
Function QuarterStartRow(wB As Workbook) As Integer
    
    Dim iQuarter As Integer, i As Integer, j As Integer
    Dim prodLineOne As String
    Dim rngSearchResult As Range
    
    If gMonthNum > 9 Then 'changes may happen after month ends?
        iQuarter = 1
    ElseIf gMonthNum < 4 Then
        iQuarter = 2
    ElseIf gMonthNum > 3 And gMonthNum < 7 Then
        iQuarter = 3
    Else
        iQuarter = 4
    End If
    
    QuarterStartRow = 0
    On Error GoTo errhandler
    wB.Activate
    wB.Worksheets(1).Activate
    i = 3 'don't include doc title
    Do While Not Cells(i, 1).Font.Bold
        i = i + 1
    Loop
    
    If iQuarter = 1 Then 'use top value
        QuarterStartRow = i
        Exit Function
    End If
    
    prodLineOne = Cells(i, 1).Value
    For j = 1 To iQuarter - 1 'jump to next "first product line" j times
        i = i + 1
        i = Range("A:C").Find(prodLineOne, after:=Range("A" & i), searchorder:=xlByRows, lookat:=xlWhole).Row
    Next

    QuarterStartRow = i
    
    Exit Function
    
errhandler:
    MsgBox "Error in QuarterStartRow function"
End Function
Function QuarterEndRow(wB As Workbook) As Integer
    
    Dim iQuarter As Integer, i As Integer, j As Integer
    Dim prodLineOne As String
    Dim rngSearchResult As Range
    
    If gMonthNum > 9 Then
        iQuarter = 1
    ElseIf gMonthNum < 4 Then
        iQuarter = 2
    ElseIf gMonthNum > 3 And gMonthNum < 7 Then
        iQuarter = 3
    Else
        iQuarter = 4
    End If
    
    QuarterEndRow = 0
    On Error GoTo errhandler
    wB.Activate
    wB.Worksheets(1).Activate
    i = 3 'don't include doc title
    Do While Not Cells(i, 1).Font.Bold
        i = i + 1
    Loop
    
    If iQuarter = 4 Then
        Range("B10000").Select
        Selection.End(xlUp).Select
        QuarterEndRow = ActiveCell.Row
        Exit Function
    End If
    
    prodLineOne = Cells(i, 1).Value
    For j = 0 To iQuarter - 1 'jump to next "first product line" j times
        i = i + 1
        i = Range("A:C").Find(prodLineOne, after:=Range("A" & i), searchorder:=xlByRows, lookat:=xlWhole).Row
    Next

    QuarterEndRow = i
    
    Exit Function
    
errhandler:
    MsgBox "Error in QuarterEndRow function"
End Function
Function MonthStartRow(wB As Workbook, curStartRow As Integer, nextStartRow As Integer) As Integer
    
    Dim monthName As String, lastMonth As String, i As Integer, searchStartRow As Integer
    
    MonthStartRow = 0
    'On Error GoTo errhandler
    ThisWorkbook.Activate
    Sheet2.Activate
    monthName = Range("M" & gMonthNum).Value
    If gMonthNum > 1 Then
        lastMonth = Range("M" & gMonthNum - 1).Value
    Else
        lastMonth = Range("M12").Value
    End If
    
    wB.Activate
    wB.Worksheets(1).Activate
    searchStartRow = curStartRow + 1
    Do While searchStartRow < nextStartRow
        i = Range("C:C").Find("OPPORTUNITIES", after:=Range("C" & searchStartRow), lookat:=xlPart).Row
        If InStr(UCase(Range("C" & i + 1)), UCase(monthName)) > 0 Or InStr(UCase(Range("C" & i + 2)), UCase(monthName)) > 0 Then
            MonthStartRow = searchStartRow
            Exit Do
        Else
            searchStartRow = i + 1
        End If
    Loop
    
    Exit Function
    
errhandler:
    MsgBox "Error in MonthStartRow function"
    Sheet1.Activate
End Function
Function MonthEndRow(wB As Workbook, curStartRow As Integer, nextStartRow As Integer) As Integer
    
    Dim monthName As String, i As Integer, searchStartRow As Integer
    
    MonthEndRow = 0
    On Error GoTo errhandler
    ThisWorkbook.Activate
    Sheet2.Activate
    monthName = Range("M" & gMonthNum).Value
    
    wB.Activate
    wB.Worksheets(1).Activate
    searchStartRow = curStartRow
    Do While searchStartRow < nextStartRow
        i = Range("C:C").Find("OPPORTUNITIES", after:=Range("C" & searchStartRow), lookat:=xlWhole).Row
        If InStr(UCase(Range("C" & i + 1)), UCase(monthName)) > 0 Or InStr(UCase(Range("C" & i + 2)), UCase(monthName)) > 0 Then
            MonthEndRow = i
            Exit Do
        End If
        searchStartRow = i + 1
    Loop
    Exit Function
    
errhandler:
    MsgBox "Error in MonthEndRow function"
    Sheet1.Activate
End Function

Function GetDetails(wB As Workbook, rowNum As Integer) As String()
    
    Dim outputVals(5) As String
    
    On Error GoTo errhandler
    outputVals(5) = "ERROR"
    GetDetails = outputVals
    wB.Activate
    wB.Worksheets(1).Activate
    If Range("B" & rowNum) > 0 Then
        If Range("G" & rowNum).Interior.ColorIndex = 6 Then
            outputVals(0) = Range("A" & rowNum).Value
            outputVals(1) = Range("B" & rowNum).Value
            outputVals(2) = Range("C" & rowNum).Value
            outputVals(3) = Range("G" & rowNum).Value
            outputVals(4) = Range("L" & rowNum).Value
        End If
    End If
    outputVals(5) = "OK"
    GetDetails = outputVals
    
    Exit Function
    
errhandler:
    MsgBox "Error in GetDetails function"
End Function

Function UpdateWbInfo(wB As Workbook, sPass As String) As Boolean
    
    UpdateWbInfo = False
    ThisWorkbook.Activate
    Sheet1.Activate
    Sheet1.Unprotect sPass
    Range("T1").Value = LocalToUNC(wB.Path) & Right(wB.Path, Len(wB.Path) - 2) & "\"
    If Left(Range("T1").Value, 2) <> "\\" Then Range("T1").Value = "\\" & Range("T1").Value
    Range("T2").Value = wB.Name
    Sheet1.Protect sPass
    UpdateWbInfo = True
    Exit Function
errhrandler:
    MsgBox "Error in UpdateWbInfo function"
End Function
Function MakeWorkbook(prodLines() As String, custNames() As String, coNums() As String, _
                    sDescriptions() As String, sPrices() As String, sComments() As String) As String

    Dim rowNum As Integer, i As Integer, x As Integer
    Dim wbNew As Workbook
    Dim sSavePath As String, sSavePathBackup As String, sMonth As String, wbName As String
    Dim iFY As Integer
    Dim lastF As String
    
    On Error GoTo errhandler
    
    lastF = Left(Application.UserName, InStr(Application.UserName, ",") - 1)
    lastF = lastF & Mid(Application.UserName, InStr(Application.UserName, ",") + 2, 1)
    
    '''''''hardcoded values'''''
    If Dir$("C:\users\" & lastF & "\OneDrive - Barry-Wehmiller\", vbDirectory) <> "" Then 'onedrive exists
        sSavePath = "C:\Users\" & lastF & "\OneDrive - Barry-Wehmiller\"
        If Dir$(sSavePath & "SHIP SCHED*", vbDirectory) <> "" Then
            sSavePath = sSavePath & Dir$(sSavePath & "SHIP SCHED*", vbDirectory)
        End If
    ElseIf Dir$("C:\users\" & lastF & "\desktop\", vbDirectory) <> "" Then
        sSavePath = "C:\users\" & lastF & "\desktop\Remaining Shipment List\"
    Else 'no one drive and/or wrong username
        sSavePath = "C:\Remaining Shipment List\"
    End If
    sSavePathBackup = "C:\Users\" & lastF & "\Desktop\SHIP SCHEDULE SENDOUTS\"
    ''''''''''''''''''''''''''''
    
    MakeWorkbook = ""
    
    If Dir$(sSavePath, vbDirectory) = "" Then
        MkDir sSavePath
    End If
    
    If Dir$(sSavePathBackup, vbDirectory) = "" Then
        MkDir sSavePathBackup
    End If
    
    sMonth = ThisWorkbook.Worksheets(2).Range("M" & gMonthNum).Value
    
    Set wbNew = Workbooks.Add 'make workbook
    
    wbNew.Activate 'fill it with data
    
    Range("E1").Value = "Y/N"
    
    
    Range("A:G").Font.Name = "Arial"
    Range("A:G").Font.Size = 12
    Range("B:C").HorizontalAlignment = xlHAlignCenter
    Range("E:E").HorizontalAlignment = xlHAlignCenter
    'Range("G:G").HorizontalAlignment = xlHAlignCenter ''removed initials
    Range("A1").Value = "REMAINING SHIPMENTS"
    Range("A1").Font.Size = 18
    Range("D:D").NumberFormat = "$#,###"
    Range("D1").NumberFormat = "mm/dd/yyyy"
    Range("D1").Value = Date
    
    rowNum = 3
    Range("A" & rowNum).Value = prodLines(0)
    Range("A" & rowNum).Font.Size = 15
    Range("1:" & rowNum).Font.Bold = True
    'Range("E1").Value = "Y/N" ''moved up (for column sizing)
    Range("E:E").Interior.ColorIndex = 15
    With Range("E:E")
        With .Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        With .Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlMedium
        End With
    End With
    Range("F1").Value = "COMMENTS"
    'Range("G1").Value = "INITIALS" ''removed initials
    Range("E1:F1").Font.Underline = True
    Range("E:F").Font.ColorIndex = 3
    'Range("G:G").Font.ColorIndex = 23 ''removed initials
    Range("F1").HorizontalAlignment = xlHAlignCenter
    With ActiveWindow
        If .FreezePanes Then .FreezePanes = False
        .SplitColumn = 0
        .SplitRow = 1
        .FreezePanes = True
    End With
    Range("A:G").Locked = False
    
    x = rowNum
    i = 0
    Do While i <= UBound(prodLines)
        rowNum = rowNum + 1
        If prodLines(i) <> Range("A" & x).Value Then
            Range("C" & rowNum).Value = "TOTAL"
            Range("D" & rowNum).Formula = "=SUM(D" & x + 1 & ":D" & rowNum - 1 & ")"
            Range(rowNum & ":" & rowNum + 2).Font.Bold = True
            x = rowNum + 2
            Range("A" & x).Value = prodLines(i)
            Range("A" & x).Font.Size = 15
            rowNum = rowNum + 3
        End If
        Range("A" & rowNum).Value = Replace(custNames(i), "=", "")
        Range("B" & rowNum).Value = Replace(coNums(i), "=", "")
        Range("C" & rowNum).Value = Replace(sDescriptions(i), "=", "")
        Range("D" & rowNum).Value = Replace(sPrices(i), "=", "")
        Range("D" & rowNum).Interior.ColorIndex = 6
        Range("F" & rowNum).Value = Replace(sComments(i), "=", "")
        i = i + 1
    Loop
    
    Columns("A").ColumnWidth = 50
    Columns("B:F").AutoFit
    Columns("E").ColumnWidth = 7
    'ActiveSheet.Protect ''no longer want this protected
    With ActiveSheet.PageSetup
        .Orientation = xlLandscape
        .LeftMargin = Application.InchesToPoints(0.25)
        .RightMargin = Application.InchesToPoints(0.25)
        .TopMargin = Application.InchesToPoints(0.75)
        .BottomMargin = Application.InchesToPoints(0.75)
        .HeaderMargin = Application.InchesToPoints(0.3)
        .FooterMargin = Application.InchesToPoints(0.3)
        .Zoom = False
        .FitToPagesTall = False
        .FitToPagesWide = 1
    End With
    
    If Right(sSavePathBackup, 1) <> "\" Then sSavePathBackup = sSavePathBackup & "\"
    If Right(sSavePath, 1) <> "\" Then sSavePath = sSavePath & "\"
    
    If gMonthNum > 9 Then
        iFY = Year(Date) + 1
        If gMonthNum = 12 And Month(Date) = 1 Then
            iFY = Year(Date)
        End If
    Else
        iFY = Year(Date)
    End If
    
    wbName = "Remaining_Shipments_" & UCase(Left(sMonth, 3)) & "_FY" & Right(Str(iFY), 2) & "_" & Format(Now, "yyyymmdd-hhmmss") & ".xlsx"
    sSavePathBackup = sSavePathBackup & wbName
    sSavePath = sSavePath & wbName
    wbNew.SaveAs sSavePath
    wbNew.SaveAs sSavePathBackup
    
    MakeWorkbook = sSavePathBackup

    Exit Function
errhandler:
    MsgBox "Error in MakeWorkbook function"

End Function
Sub test()
    Dim sName As String
    sName = Application.UserName
    sName = Trim(Mid(sName, InStr(sName, ",") + 1, InStr(sName, "(") - InStr(sName, ",") - 1))
    Debug.Print sName
End Sub

Function DraftEmail(wbPath As String) As Boolean

    Dim sEmailText As String, sSubject As String, sName As String
    Dim oOutlook As Object, oMail As Object
    Dim i As Integer, x As Integer
    Dim sEmails() As String, vEmail As Variant
    Dim newWB As Workbook, wbURL As String
    Dim oneDrive As Boolean
    Dim wbWorkbook As Workbook
    
    DraftEmail = False
    
    On Error Resume Next
    sName = Application.UserName
    sName = Trim(Mid(sName, InStr(sName, ",") + 1, InStr(sName, "(") - InStr(sName, ",") - 1))
    If sName = "" Then sName = "Michelle"
    
    ThisWorkbook.Activate
    Sheet1.Activate
    i = 3
    ReDim sEmails(0)
    Do While Range("F" & i).Value > 0
        ReDim Preserve sEmails(i - 3)
        sEmails(i - 3) = Range("F" & i).Value
        i = i + 1
    Loop
    
    For Each wbWorkbook In Application.Workbooks
        If wbWorkbook.Path = wbPath Then
            Set newWB = wbWorkbook
        End If
    Next wbWorkbook
    
    If newWB Is Nothing Then
        Set newWB = Workbooks.Open(wbPath)
    End If
    
    wbURL = "<" & newWB.FullName & ">" 'chevrons required for outlook if spaces exist in the link
    
    sEmailText = "Good Morning, All!" & vbCrLf & vbCrLf & _
                "Below, I have included a link to"
    
    sEmailText = sEmailText & "the updated remaining shipment list. Please " & _
                "let Mike M. and me know of any changes/updates." & vbCrLf & vbCrLf
                
    sEmailText = sEmailText & wbURL & vbCrLf & vbCrLf ''figure out how to get onedrive URL
                
    If Weekday(Date) = 6 Then
        sEmailText = sEmailText & "Happy Friday and hope everyone enjoys their weekend!" & vbCrLf & vbCrLf
    End If
    
    sEmailText = sEmailText & "Thank you," & vbCrLf & _
                                sName
                                
    sSubject = Right(wbPath, Len(wbPath) - InStrRev(wbPath, "\"))
    sSubject = Left(sSubject, InStr(LCase(sSubject), ".xls") - 1)
    
    Set oOutlook = CreateObject("Outlook.Application")
    oOutlook.Session.Logon
    Set oMail = oOutlook.CreateItem(olMailItem)
    
    With oMail
        For Each vEmail In sEmails
            .Recipients.Add vEmail
        Next
        .Subject = sSubject
        .Body = sEmailText
    End With
    
    oMail.Display
    
    DraftEmail = True
    
End Function

Function LocalToUNC(ByVal localPath As String) As String
    
Dim objNetwork As Object, objDrives As Object, indivDrive As Long

On Error GoTo errhandler

localPath = Left(localPath, 2)

Set objNetwork = CreateObject("WScript.Network")
Set objDrives = objNetwork.enumnetworkdrives

For indivDrive = 0 To objDrives.Count - 1 Step 2
    If UCase(objDrives.Item(indivDrive)) = UCase(localPath) Then
        LocalToUNC = objDrives.Item(indivDrive + 1)
        Exit For
    End If
Next

Exit Function

errhandler:
MsgBox "Error in LocalToUNC function"

End Function
