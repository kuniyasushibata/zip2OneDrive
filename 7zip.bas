Attribute VB_Name = "Module1"
Option Explicit

Const pathOf7zip As String = "C:\Program Files\7-Zip\7z.exe"
Const wordOfSuccess As String = "Everything is Ok"
Const fileNameOfResult As String = "7zip_result.txt"
Const maxCount As Long = 10

Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As Long)
Private Declare PtrSafe Function PathIsDirectoryEmpty Lib "SHLWAPI.DLL" Alias "PathIsDirectoryEmptyA" (ByVal pszPath As String) As Boolean

Function Is7zipInstalled() As Boolean
    Dim ret As Boolean
    If Dir(pathOf7zip) <> "" Then
        ret = True
    Else
        ret = False
    End If
    Is7zipInstalled = ret
End Function

Function Compress7zip(folderPath As String, zipFilePath As String, pass As String) As Boolean
    Dim result, ret As Boolean
    If IsFolderExistsAndNotEmpty(folderPath) = True Then
        If Dir(zipFilePath) = "" Then
            result = Execute7zip(CreateCmdOf7zip(True, folderPath, zipFilePath, pass))
            If result = True Then
                ret = True
            Else
                If Dir(zipFilePath) <> "" Then
                    Kill zipFilePath
                End If
                ret = False
            End If
        Else
            ret = False
        End If
    Else
        ret = False
    End If
    Compress7zip = ret
End Function

Function Extract7zip(zipFilePath As String, folderPath As String, pass As String) As Boolean
    Dim ret As Boolean
    If Dir(zipFilePath) <> "" Then
        If PathIsDirectoryEmpty(folderPath) = 1 Then
            ret = Execute7zip(CreateCmdOf7zip(False, folderPath, zipFilePath, pass))
            If ret = False Then
                DeleteFileAndFolder folderPath
            End If
        Else
            ret = False
        End If
    Else
        ret = False
    End If
    Extract7zip = ret
End Function

Private Function CreateCmdOf7zip(isCompress As Boolean, folderPath As String, zipFilePath As String, pass As String) As String
    Dim cmd As String
    If isCompress = True Then
        cmd = """" & pathOf7zip & """" & " " & "a" & " " & """" & zipFilePath & """" & " " & """" & folderPath & """"
        If pass <> "" Then
            cmd = cmd & " " & "-p" & """" & pass & """"
        End If
    Else
        cmd = "echo q | " & """" & pathOf7zip & """" & " " & "x" & " " & """" & zipFilePath & """" & " " & "-o" & """" & folderPath & """"
        If pass <> "" Then
            cmd = cmd & " " & "-p" & """" & pass & """"
        End If
    End If
    CreateCmdOf7zip = cmd
End Function


Private Function Execute7zip(cmd As String) As Boolean
    Dim WSH As Object
    Dim filePathOfResult As String
    filePathOfResult = Environ("TMP")
    filePathOfResult = filePathOfResult & "\" & fileNameOfResult
    If Dir(filePathOfResult) <> "" Then
        Kill filePathOfResult
    End If
    cmd = cmd & " > " & filePathOfResult
    cmd = """" & cmd & """"
    Debug.Print cmd
    Set WSH = CreateObject("WScript.Shell")
    WSH.Run "%ComSpec% /c " & cmd
    Set WSH = Nothing
    Execute7zip = CheckResultOf7zip(filePathOfResult)
End Function

Private Function CheckResultOf7zip(filePath As String) As Boolean
    Dim buf As String
    Dim ret, isFound As Boolean
    ret = False
    isFound = False
    Dim i As Long
    For i = 1 To maxCount
        If Dir(filePath) <> "" Then
            isFound = True
            Exit For
        End If
        Sleep 500
    Next i
    If isFound = True Then
        Open filePath For Input As #1
            Do Until EOF(1)
                Line Input #1, buf
                If InStr(buf, wordOfSuccess) > 0 Then
                    ret = True
                End If
            Loop
        Close #1
    End If
    CheckResultOf7zip = ret
End Function

Private Sub DeleteFileAndFolder(folder As String)
    Dim fso As Object
    Dim all As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    all = fso.BuildPath(folder, "*")
    Debug.Print all
    fso.DeleteFile all, True
    fso.DeleteFolder all, True
    Set fso = Nothing
End Sub

Private Function IsFolderExistsAndNotEmpty(folderPath As String) As Boolean
    Dim ret As Boolean
    ret = False
    If Dir(folderPath, vbDirectory) <> "" Then
        If PathIsDirectoryEmpty(folderPath) = 0 Then
            ret = True
        End If
    Else
        ret = False
    End If
    IsFolderExistsAndNotEmpty = ret
End Function

Private Sub Test()
    Dim installed As Boolean
    installed = Is7zipInstalled()
    'MsgBox Compress7zip("D:\github\UiPathLogToExcel", "D:\github\UiPathLogToExcel.zip", "test")
    'MsgBox Compress7zip("D:\github\UiPathLogToExcela", "D:\github\UiPathLogToExcel.zip", "test")
    'MsgBox Extract7zip("D:\github\UiPathLogToExcel.zip", "D:\github\a", "test")
    'MsgBox Extract7zip("D:\github\UiPathLogToExcel2.zip", "D:\github\a", "test")
    'MsgBox Extract7zip("D:\github\UiPathLogToExcel.zip", "D:\github\a2", "test")
    MsgBox Extract7zip("D:\github\UiPathLogToExcel.zip", "D:\github\a", "test2")
End Sub

Private Sub Test2()
    MsgBox PathIsDirectoryEmpty("D:\github\a")
End Sub

Private Sub Test3()
    DeleteFileAndFolder "D:\github\a"
End Sub

Private Sub Test4()
    MsgBox IsFolderExistsAndNotEmpty("D:\github\a")
End Sub
