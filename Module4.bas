Attribute VB_Name = "Module4"
Dim iNumMsgs As Integer, arrErrorEmails() As String
Sub TransferComments()

    Dim wbSPoint As Workbook, wbShipSched As Workbook
    Dim wsSPoint As Worksheet, wsShipSched As Worksheet
    Dim i As Integer, j As Integer, iRow As Integer, iEmpty As Integer, x As Integer, y As Integer
    Dim sWbPath As String, sWbName As String, sComment As String, sLastF As String
    Dim sCO As String
    Dim arrCOs() As String, arrComms() As String, arrFailed() As String
    Dim arrAmts() As String, arrDups() As String, arrInits() As String
    Dim varVar As Variant
    Dim rngSearch As Range, rngResult As Range
    Dim bSPopen As Boolean, bSSopen As Boolean, bDup As Boolean

    On Error GoTo errhandler

    'find sharepoint workbook
'    sLastF = Left(Application.UserName, InStr(Application.UserName, ",") - 1)
'    sLastF = sLastF & Mid(Application.UserName, InStr(Application.UserName, ",") + 2, 1)
'    If Dir$("C:\users\" & sLastF & "\OneDrive - Barry-Wehmiller\", vbDirectory) <> "" Then 'onedrive exists
'        If Dir$("C:\Users\" & lastF & "\OneDrive - Barry-Wehmiller\" & "SHIP SCHED*", vbDirectory) <> "" Then
'            sWbPath = "C:\Users\" & sLastF & "\OneDrive - Barry-Wehmiller\" & Dir$("C:\Users\" & lastF & "\OneDrive - Barry-Wehmiller\" & "SHIP SCHED*", vbDirectory)
'        End If
'    End If
'
'    If sWbPath = "" Then
        MsgBox "Please choose which file is the newly commented shipping schedule" 'Michelle requested to choose location each time
        On Error Resume Next
        With Application.FileDialog(msoFileDialogFilePicker)
            .AllowMultiSelect = False
            .Filters.Add "Excel Files", "*.xlsx; *.xlsm; *.xls; *.xlsb", 1
            .Show
            sWbPath = .SelectedItems.Item(1)
        End With
        Err.Clear
'    Else
'        If Right(sWbPath, 1) <> "\" Then sWbPath = sWbPath & "\"
'        sWbPath = sWbPath & Dir(sWbPath & "*.xls*")
'        If Dir$(sWbPath & "*.xls*") <> "" Then sWbPath = "" 'multiple excel files in that directory
'    End If
    If sWbPath = "" Then Exit Sub
    On Error GoTo errhandler
    sWbName = Right(sWbPath, Len(sWbPath) - InStrRev(sWbPath, "\"))

    For Each varVar In Workbooks
        If UCase(varVar.Name) = UCase(sWbName) Then
            Set wbSPoint = varVar
            bSPopen = True
            Exit For
        End If
    Next
    If wbSPoint Is Nothing Then Set wbSPoint = Workbooks.Open(sWbPath)

    'find official ship schedule
    sWbPath = GetShipSchedulePath
    Sheet2.Activate
    If sWbPath = "" Then Exit Sub

    sWbName = Right(sWbPath, Len(sWbPath) - InStrRev(sWbPath, "\"))

    For Each varVar In Workbooks
        If UCase(varVar.Name) = UCase(sWbName) Then
            Set wbShipSched = varVar
            bSSopen = True
            Exit For
        End If
    Next
    If wbShipSched Is Nothing Then
        Application.DisplayAlerts = False
        Set wbShipSched = Workbooks.Open(sWbPath, IgnoreReadOnlyRecommended:=True)
    End If

    If wbShipSched.ReadOnly Then
        MsgBox "Ship schedule is read only"
        'Exit Sub
    End If

    'vvvvvvvv old/current way vvvvvvvvvvvvv
    
    'copy the comments over
    Set wsSPoint = wbSPoint.Worksheets(1)
    Set wsShipSched = wbShipSched.Worksheets(1)
    ReDim arrFailed(x)
    i = 1
    Do While iEmpty < 5 '5 empty rows indicates end of file
        If wsSPoint.Range("B" & i).Value > 0 Then
            iEmpty = 0
            sCO = wsSPoint.Range("B" & i).Value
            sComment = wsSPoint.Range("F" & i).Value
            If sComment <> "" Then
                If wsSPoint.Range("G" & i).Value > 0 Then sComment = sComment & " (" & wsSPoint.Range("G" & i).Value & ")"
            End If
            On Error Resume Next
                Set rngSearch = wsShipSched.Range("B:B").Find(what:=sCO, LookIn:=xlValues)
                Err.Clear
            On Error GoTo errhandler
            If rngSearch Is Nothing Then 'add to list of failed
                ReDim Preserve arrFailed(x)
                arrFailed(x) = sCO
                x = x + 1
            ElseIf wsShipSched.Range("L" & rngSearch.Row).Value <> "SHIPPED" Then 'update comments
                wsShipSched.Range("L" & rngSearch.Row).Value = sComment
                Set rngSearch = Nothing
            End If
        Else
            iEmpty = iEmpty + 1
        End If
        Debug.Print i & ":  " & sCO & " - " & sComment
        i = i + 1
    Loop
    '^^^^^^^^^ old/current way ^^^^^^^^^^^

    
'' vvvvvvvvvvvv  new way (should work?)  vvvvvvvvvvvvvv
'    x = 0
'    y = 0
'    ReDim arrCOs(x)
'    ReDim arrAmts(x) 'only used when duplicate CO's
'    ReDim arrComms(x)
'    ReDim arrInits(x)
'    ReDim arrFailed(y)
'
'    Do While iEmpty < 5 'get CO's, Amounts, Comments, Initials
'        If wsSPoint.Range("B" & i).Value > 0 Then
'            iEmpty = 0
'            ReDim Preserve arrCOs(x)
'            ReDim Preserve arrAmts(x)
'            ReDim Preserve arrComms(x)
'            ReDim Preserve arrInits(x)
'            arrCOs(x) = wsSPoint.Range("B" & i).Value
'            arrAmts(x) = wsSPoint.Range("D" & i).Value
'            arrComms(x) = wsSPoint.Range("F" & i).Value
'            arrInits(x) = wsSPoint.Range("G" & i).Value
'            x = x + 1
'        Else
'            iEmpty = iEmpty + 1
'        End If
'    Loop
'
'    arrDups = GetDuplicates(arrCOs)
'
'    On Error Resume Next 'the "FIND" function needs this
'    For i = 0 To UBound(arrCOs) 'update each on ship schedule
'        For Each varVar In arrDups 'see if duplicate
'            If varVar = sCO(i) Then
'                bDup = True
'                Exit For
'            End If
'        Next
'        If bDup Then 'duplicate -> check amount
'            j = 0
'            Set rngSearch = wsShipSched.Range("B:B").Find(what:=arrCOs(i), LookIn:=xlValues) 'find CO
'            Do While j < 10
'                If wsShipSched.Range("G" & rngSearch.Row).Value = arrAmts(i) Then Exit Do
'                Set rngSearch = wsShipSched.Range("B:B").FindNext(after:=rngSearch)
'                j = j + 1
'            Loop
'            If j = 10 Then Set rngSearch = Nothing 'none have the proper amounts
'        Else 'update regardless of amount
'            Set rngSearch = wsShipSched.Range("B:B").Find(what:=arrCOs(i), LookIn:=xlValues) 'find CO
'        End If
'        If rngSearch Is Nothing Then 'add to list of failed
'            ReDim Preserve arrFailed(y)
'            arrFailed(y) = arrCOs(i)
'            y = y + 1
'        Else 'CO (& amount?) was found
'            If wsShipSched.Range("L" & rngSearch.Row).Value <> "SHIPPED" Then 'update comments
'                wsShipSched.Range("L" & rngSearch.Row).Value = arrComms(i)
'                wsShipSched.Range("M" & rngSearch.Row).Value = arrInits(i)
'            End If
'        End If
'    Next
'
''^^^^^^^ new way (should work?) ^^^^^^^^^^^^^^^




    'save ship sched?

    If Not bSPopen Then wbSPoint.Close savechanges:=False
    'If Not bSSopen Then wbShipSched.Close savechanges:=False

    If x > 0 Then
        MsgBox "The following CO numbers could not have comments updated:" & _
                vbCrLf & vbCrLf & Join(arrFailed, vbCrLf)
    Else
        MsgBox "Comments transferred successfully"
    End If

    Exit Sub
errhandler:
    MsgBox "Error copying the comments over"
    Call ErrorRep("TransferComments", "Sub", "N/A", Err.Number, Err.Description, "")
End Sub
Function GetDuplicatez(arrCOs() As String) As String()
    Dim x As Integer, arrOutput() As String, varVar As Variant, i As Integer, j As Integer
    ReDim arrOutput(0)
    For i = 0 To UBound(arrCOs)
        j = 0
        For Each varVar In arrCOs
            If varVar = arrCOs(i) Then j = j + 1
        Next
        If j > 1 Then 'duplicates exist
            j = 0
            For Each varVar In arrOutput
                If varVar = arrCOs(i) Then
                    j = 1
                    Exit For
                End If
            Next
            If j = 0 Then 'add to arrOutput
                ReDim Preserve arrOutput(x)
                arrOutput(x) = arrCOs(i)
                x = x + 1
            End If
        End If
    Next
    GetDuplicates = arrOutput
End Function
Public Sub ErrorRep(rouName, rouType, curVal, errNum, errDesc, miscInfo)
    Exit Sub
    Dim oApp As Object, oEmail As MailItem, arrEmailTxt(10) As String
    Dim outlookOpen As Boolean, emailTxt As String, varMsg As Variant
    
    Application.ScreenUpdating = False
    arrEmailTxt(2) = "--Issue finding Workbook"
    arrEmailTxt(3) = "--Issue finding User"
    arrEmailTxt(4) = "--Issue finding Workbook path"
    arrEmailTxt(5) = "--Issue finding Routine name"
    arrEmailTxt(6) = "--Issue finding Routine type"
    arrEmailTxt(7) = "--Issue finding Current value"
    arrEmailTxt(8) = "--Issue finding Error number"
    arrEmailTxt(9) = "--Issue finding Error description"
    arrEmailTxt(10) = "--Issue finding Misc. add'l info"
    
    On Error Resume Next
        Set oApp = GetObject(, "Outlook.Application")
        outlookOpen = True
        
        ''''''can't use error handler because these varTypes might be problematic
        If Not VarType(curVal) = vbString Then 'make into string
            If VarType(curVal) > 8000 Then 'array of some sort
                curVal = Join(curVal, ";")
            Else 'hopefully this will make it a string
                curVal = Str(curVal)
            End If
        End If
        
        If Not VarType(miscInfo) = vbString Then 'make into string
            If VarType(miscInfo) > 8000 Then 'array of some sort
                curVal = Join(miscInfo, ";")
            Else 'hopefully this will make it a string
                curVal = Str(miscInfo)
            End If
        End If
        
    On Error Resume Next 'types might cause errors
        arrEmailTxt(0) = "REPORT"
        arrEmailTxt(1) = "Error occurred in VBA program. Details are listed below." & vbCrLf
        arrEmailTxt(2) = Right(arrEmailTxt(2), Len(arrEmailTxt(2)) - 16) & ": " & ThisWorkbook.Name
        arrEmailTxt(3) = Right(arrEmailTxt(3), Len(arrEmailTxt(3)) - 16) & ": " & Application.UserName & vbCrLf
        arrEmailTxt(4) = Right(arrEmailTxt(4), Len(arrEmailTxt(4)) - 16) & ": " & ThisWorkbook.Path
        arrEmailTxt(5) = Right(arrEmailTxt(5), Len(arrEmailTxt(5)) - 16) & ": " & rouName
        arrEmailTxt(6) = Right(arrEmailTxt(6), Len(arrEmailTxt(6)) - 16) & ": " & rouType
        arrEmailTxt(7) = Right(arrEmailTxt(7), Len(arrEmailTxt(7)) - 16) & ": " & curVal & vbCrLf
        arrEmailTxt(8) = Right(arrEmailTxt(8), Len(arrEmailTxt(8)) - 16) & ": " & errNum
        arrEmailTxt(9) = Right(arrEmailTxt(9), Len(arrEmailTxt(9)) - 16) & ": " & errDesc & vbCrLf
        arrEmailTxt(10) = Right(arrEmailTxt(10), Len(arrEmailTxt(10)) - 16) & ": " & vbCrLf & miscInfo
    On Error GoTo errhandler
    
    emailTxt = Join(arrEmailTxt, vbCrLf)
    
    'see if emailTxt has been sent already this session
    bNewMsg = True 'default value
    If iNumMsgs > 0 Then 'at least one email has been generated already
        For Each varMsg In arrErrorEmails 'see if there were any matches
            If UCase(varMsg) = UCase(emailTxt) Then 'this was already sent this session
                bNewMsg = False
                Exit For
            End If
        Next
    End If
    
    If bNewMsg Then 'new message -> add to array for next time
        iNumMsgs = iNumMsgs + 1
        ReDim Preserve arrErrorEmails(iNumMsgs)
        arrErrorEmails(iNumMsgs) = emailTxt
    Else 'repeat message
        Exit Sub
    End If
    
    If oApp Is Nothing Then
        Set oApp = CreateObject("Outlook.Application")
        outlookOpen = False
    End If
    
    Set oEmail = oApp.CreateItem(0)

    With oEmail
        .To = "tyler.england@bwpackagingsystems.com"
        .Subject = "VBA Program Error Report"
        .Body = emailTxt
        If InStr(UCase(Application.UserName), "ENGLAND, TYLER") > 0 Then
            .Display 'it me
        Else:
            .Send
        End If
    End With
    
    If Not outlookOpen Then oApp.Close
errhandler:
End Sub

