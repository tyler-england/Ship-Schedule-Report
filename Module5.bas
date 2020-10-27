Attribute VB_Name = "Module5"
Option Explicit

Function ExportModules() As Boolean
    Dim s1DPath As String, sFolderPath As String, sSubFolder As String, sFileFolder As String
    Dim varVar As Variant, bNewFolder As Boolean, sExt As String
    Dim sFailed() As String, x As Integer
    Const Module = 1
    Const ClassModule = 2
    Const Form = 3
    Const Document = 100
    '''''''hardcoded'''''''
    s1DPath = "C:\Users\englandt\*"
    '''''''''''''''''''''''
    'On Error GoTo errhandler
    sSubFolder = Dir(Replace(s1DPath, "*", "OneDrive*"), vbDirectory)
    If sSubFolder = "" Then
        Exit Function 'no OneDrive
    ElseIf UCase(sSubFolder) = "ONEDRIVE" Then
        MsgBox "May be using wrong OneDrive folder (not BW directory)"
    End If
    sFolderPath = Replace(s1DPath, "*", sSubFolder) & "\"
    sSubFolder = Dir(sFolderPath & "scripts*", vbDirectory)
    If sSubFolder = "" Then
        MkDir sFolderPath & "Scripts" 'make directory
        sSubFolder = "Scripts"
    End If
    sFolderPath = sFolderPath & sSubFolder & "\"
    sSubFolder = Dir(sFolderPath & "VBA*", vbDirectory)
    If sSubFolder = "" Then
        MkDir sFolderPath & "VBA_Modules" 'make directory
        sSubFolder = sFolderPath & "VBA_Modules"
    End If
    sFolderPath = sFolderPath & sSubFolder & "\" 'vba modules folder
    sFileFolder = Replace(Replace(Replace(Replace(ThisWorkbook.Path & "\" & ThisWorkbook.Name, "\", "-"), ".", ""), ":", "+"), " ", "")
    sSubFolder = Dir(sFolderPath & sFileFolder, vbDirectory)
    If sSubFolder = "" Then 'folder doesn't exist
        bNewFolder = True
        sSubFolder = Dir(sFolderPath & "*" & Replace(ThisWorkbook.Name, ".", "") & "*", vbDirectory)
        Do While sSubFolder <> "" 'check for any partial matches (diff path, etc)
            varVar = MsgBox("No folder exists with the following name..." & vbCrLf & sFileFolder & _
                    vbCrLf & vbCrLf & "However this folder does exist..." & vbCrLf & sSubFolder & _
                    vbCrLf & vbCrLf & "Do you want to use this one instead?", vbYesNo, "VBA Modules")
            If varVar = vbYes Then 'use this folder -> don't make a new one
                Name sFolderPath & sSubFolder As sFolderPath & sFileFolder
                bNewFolder = False
                Exit Do
            End If
            sSubFolder = Dir()
        Loop
        If bNewFolder Then 'make new folder
            MkDir sFolderPath & sFileFolder
        End If
        sFolderPath = sFolderPath & sFileFolder
    Else
        sFolderPath = sFolderPath & sSubFolder
    End If
    If Right(sFolderPath, 1) <> "\" Then sFolderPath = sFolderPath & "\"
    x = 0
    ReDim sFailed(x)
    For Each varVar In ThisWorkbook.VBProject.VBComponents
        On Error GoTo errhandler
        Select Case varVar.Type
            Case ClassModule, Document
                sExt = ".cls"
            Case Form
                sExt = ".frm"
            Case Module
                sExt = ".bas"
            Case Else
                sExt = ".txt"
        End Select
        If sExt = ".bas" Or sExt = ".cls" Then 'only care about modules/sheets
            On Error Resume Next
            Debug.Print varVar.Name
            Err.Clear
            Call varVar.Export(sFolderPath & varVar.Name & sExt)
            If Err.Number <> 0 Then
                ReDim Preserve sFailed(x)
                sFailed(x) = varVar.Name
                x = x + 1
            End If
        End If
    Next
    If x > 0 Then
        MsgBox "Failed to export the following modules:" & vbCrLf & vbCrLf & _
                Join(sFailed, vbCrLf)
        ExportModules = True 'cancel close
    End If
errhandler:
    If Err.Number <> 0 Then
        MsgBox "Error " & Err.Number & ": " & Err.Description
        ExportModules = True 'cancel close
    End If
End Function



