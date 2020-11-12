Attribute VB_Name = "Module5"
Dim iNumMsgs As Integer, arrErrorEmails() As String, bNewMsg As Boolean

Function ExportModules() As Boolean
    Dim wbMacro As Workbook, varVar As Variant, bOpen As Boolean
    For Each varVar In Application.Workbooks
        If UCase(varVar.Name) = "MACROBOOK.XLSM" Then
            bOpen = True
            Set wbMacro = varVar
            Exit For
        End If
    Next
    If Not bOpen Then Set wbMacro = Workbooks.Open("\\PSACLW02\HOME\SHARED\MacroBook.xlsm")
    Application.Run "'" & wbMacro.Name & "'!ExportModules", ThisWorkbook
    If Not bOpen Then wbMacro.Close savechanges:=False
    ExportModules = True
End Function

Public Sub ErrorRep(rouName, rouType, curVal, errNum, errDesc, miscInfo)
    Dim wbMacro As Workbook, varVar As Variant, bOpen As Boolean
    bNewMsg = True 'default value
    If iNumMsgs > 0 Then 'at least one email has been generated already
        For Each varVar In arrErrorEmails 'see if there were any matches
            If UCase(varVar) = UCase(ThisWorkbook.Name & "-" & errNum) Then Exit Sub 'repeat message (this was already sent this session)
        Next
    End If
    For Each varVar In Application.Workbooks
        If UCase(varVar.Name) = "MACROBOOK.XLSM" Then
            bOpen = True
            Set wbMacro = varVar
            Exit For
        End If
    Next
    If Not bOpen Then Set wbMacro = Workbooks.Open("\\PSACLW02\HOME\SHARED\MacroBook.xlsm")
    Application.Run "'MacroBook.xlsm'!ErrorReport", rouName, rouType, curVal, errNum, errDesc, miscInfo
    If Not bOpen Then wbMacro.Close savechanges:=False
    iNumMsgs = iNumMsgs + 1
    ReDim Preserve arrErrorEmails(iNumMsgs)
    arrErrorEmails(iNumMsgs) = ThisWorkbook.Name & "-" & errNum
End Sub




