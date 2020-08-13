Attribute VB_Name = "FolderFile"
Option Explicit

Dim fso As Object

Private Declare PtrSafe Function PathIsDirectoryEmpty Lib "SHLWAPI.DLL" Alias "PathIsDirectoryEmptyA" (ByVal pszPath As String) As Boolean


Function FFPathIsDirectoryEmpty(folderPath As String) As Boolean
    Dim ret As Boolean
    ret = False
    If PathIsDirectoryEmpty(folderPath) = 1 Then
        ret = True
    End If
    FFPathIsDirectoryEmpty = ret
End Function

Function FFGetFso() As Object
    If fso Is Nothing Then
        Set fso = CreateObject("Scripting.FileSystemObject")
    End If
    Set FFGetFso = fso
End Function

Sub FFDeleteFileAndFolder(folder As String)
    Dim all As String
    all = FFGetFso().BuildPath(folder, "*")
    Debug.Print all
    fso.DeleteFile all, True
    fso.DeleteFolder all, True
End Sub

Function FFIsFolderExistsAndNotEmpty(folderPath As String) As Boolean
    Dim ret As Boolean
    ret = False
    If Dir(folderPath, vbDirectory) <> "" Then
        If FFPathIsDirectoryEmpty(folderPath) = False Then
            ret = True
        End If
    Else
        ret = False
    End If
    FFIsFolderExistsAndNotEmpty = ret
End Function

Function FFIsIncludeProhibitionWin(name As String) As Boolean
    Dim ret As Boolean
    CheckProhibitionStringWin name, ret, False, "", ""
    FFIsIncludeProhibitionWin = ret
End Function

Function FFIsIncludeProhibitionExcel(name As String) As Boolean
    Dim ret As Boolean
    CheckProhibitionStringExcel name, ret, False, "", ""
    FFIsIncludeProhibitionExcel = ret
End Function

Function FFIsIncludeProhibitionSharePoint(name As String) As Boolean
    Dim ret As Boolean
    CheckProhibitionStringSharePoint name, ret, False, "", ""
    FFIsIncludeProhibitionSharePoint = ret
End Function

Function FFIsIncludeProhibitionAll(name As String) As Boolean
    Dim ret As Boolean
    CheckProhibitionStringAll name, ret, False, "", ""
    FFIsIncludeProhibitionAll = ret
End Function

Function FFReplaceProhibitionWin(name As String, replace As String) As String
    Dim ret As String
    Dim isInclude As Boolean
    CheckProhibitionStringWin name, isInclude, True, replace, ret
    FFReplaceProhibitionWin = ret
End Function

Function FFReplaceProhibitionExcel(name As String, replace As String) As String
    Dim ret As String
    Dim isInclude As Boolean
    CheckProhibitionStringExcel name, isInclude, True, replace, ret
    FFReplaceProhibitionExcel = ret
End Function

Function FFReplaceProhibitionSharePoint(name As String, replace As String) As String
    Dim ret As String
    Dim isInclude As Boolean
    CheckProhibitionStringSharePoint name, isInclude, True, replace, ret
    FFReplaceProhibitionSharePoint = ret
End Function

Function FFReplaceProhibitionAll(name As String, replace As String) As String
    Dim ret As String
    Dim isInclude As Boolean
    CheckProhibitionStringAll name, isInclude, True, replace, ret
    FFReplaceProhibitionAll = ret
End Function

Private Sub CheckProhibitionStringWin(name As String, isInclude As Boolean, isReplace As Boolean, replaced As String, result As String)
    Dim strsOfProhibition As Variant
    Dim resultOfWin As String
    Dim isIncludeOfWin As Boolean
    If isReplace = True Then
        resultOfWin = name
    End If
    strsOfProhibition = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    CheckProhibitionString strsOfProhibition, name, isIncludeOfWin, isReplace, replaced, resultOfWin
    If isIncludeOfWin = True And isReplace = False Then
        isInclude = isIncludeOfWin
    Else
        Dim isSpaceFound
        isSpaceFound = False
        If Left(name, 1) = " " Or Right(name, 1) = " " Then
            isInclude = True
            isSpaceFound = True
        End If
        If isSpaceFound = True And isReplace = True Then
            result = ReplaceStartEndSpace(resultOfWin, replaced)
        End If
    End If
End Sub

Private Function ReplaceStartEndSpace(name As String, replaced As String) As String
    Dim isFound As Boolean
    isFound = False
    If Left(name, 1) = " " Then
        name = replace(name, " ", replaced, 1, 1)
        isFound = True
    End If
    If Right(name, 1) = " " Then
        name = Mid(name, 1, Len(name) - 1)
        name = name & replaced
        isFound = True
    End If
    If isFound = True Then
        name = ReplaceStartEndSpace(name, replaced)
    End If
    ReplaceStartEndSpace = name
End Function

Private Sub CheckProhibitionStringExcelOnly(name As String, isInclude As Boolean, isReplace As Boolean, replaced As String, result As String)
    If isReplace = True Then
        result = name
    End If
    Dim strsOfProhibition As Variant
    strsOfProhibition = Array("[", "]")
    CheckProhibitionString strsOfProhibition, name, isInclude, isReplace, replaced, result
End Sub

Private Sub CheckProhibitionStringSharePointOnly(name As String, isInclude As Boolean, isReplace As Boolean, replaced As String, result As String)
    If isReplace = True Then
        result = name
    End If
    Dim strsOfProhibition As Variant
    strsOfProhibition = Array("~", "%", "&", "{", "}")
    CheckProhibitionString strsOfProhibition, name, isInclude, isReplace, replaced, result
End Sub

Private Sub CheckProhibitionStringExcel(name As String, isInclude As Boolean, isReplace As Boolean, replaced As String, result As String)
    Dim isIncludeWin As Boolean
    Dim resultOfWin As String
    If isReplace = True Then
        result = name
    End If
    CheckProhibitionStringWin name, isIncludeWin, isReplace, replaced, resultOfWin
    If isIncludeWin = True And isReplace = False Then
        isInclude = isIncludeWin
    Else
        Dim isIncludeExcel As Boolean
        If isReplace = False Then
            resultOfWin = name
        End If
        CheckProhibitionStringExcelOnly resultOfWin, isIncludeExcel, isReplace, replaced, result
        If isIncludeWin = True Or isIncludeExcel = True Then
            isInclude = True
        End If
    End If
End Sub

Private Sub CheckProhibitionStringSharePoint(name As String, isInclude As Boolean, isReplace As Boolean, replaced As String, result As String)
    Dim isIncludeWin As Boolean
    Dim resultOfWin As String
    If isReplace = True Then
        result = name
    End If
    CheckProhibitionStringWin name, isIncludeWin, isReplace, replaced, resultOfWin
    If isIncludeWin = True And isReplace = False Then
        isInclude = isIncludeWin
    Else
        Dim isIncludeSharePoint As Boolean
        If isReplace = False Then
            resultOfWin = name
        End If
        CheckProhibitionStringSharePointOnly resultOfWin, isIncludeSharePoint, isReplace, replaced, result
        If isIncludeWin = True Or isIncludeSharePoint = True Then
            isInclude = True
        End If
    End If
End Sub

Private Sub CheckProhibitionStringAll(name As String, isInclude As Boolean, isReplace As Boolean, replaced As String, result As String)
    Dim isIncludeWin As Boolean
    Dim resultOfWin As String
    If isReplace = True Then
        result = name
    End If
    CheckProhibitionStringWin name, isIncludeWin, isReplace, replaced, resultOfWin
    If isIncludeWin = True And isReplace = False Then
        isInclude = isIncludeWin
    Else
        Dim isIncludeExcel As Boolean
        Dim resultOfExcel As String
        If isReplace = False Then
            resultOfWin = name
        End If
        CheckProhibitionStringExcelOnly resultOfWin, isIncludeExcel, isReplace, replaced, resultOfExcel
        If isIncludeExcel = True And isReplace = False Then
            isInclude = True
        Else
            Dim isIncludeSharePoint As Boolean
            If isReplace = False Then
                resultOfExcel = name
            End If
            CheckProhibitionStringSharePointOnly resultOfExcel, isIncludeSharePoint, isReplace, replaced, result
            If isIncludeWin = True Or isIncludeSharePoint = True Or isIncludeSharePoint = True Then
                isInclude = True
            End If
        End If
    End If
End Sub

Private Sub CheckProhibitionString(strsOfProhibition As Variant, name As String, isInclude As Boolean, isReplace As Boolean, replaced As String, result As String)
    Dim i As Long
    isInclude = False
    For i = LBound(strsOfProhibition) To UBound(strsOfProhibition)
        If InStr(name, strsOfProhibition(i)) > 0 Then
            isInclude = True
            If isReplace = False Then
                Exit For
            Else
                result = replace(result, strsOfProhibition(i), replaced)
            End If
        End If
    Next i
End Sub

Private Sub Test2()
    Debug.Print FFPathIsDirectoryEmpty("D:\github\a")
End Sub

Private Sub Test3()
    FFDeleteFileAndFolder "D:\github\a"
End Sub

Private Sub Test4()
    Debug.Print FFIsFolderExistsAndNotEmpty("D:\github\a")
End Sub

Private Sub Test5()
    'Debug.Print FFIsIncludeProhibitionWin("abc")
    Debug.Print FFIsIncludeProhibitionWin("a|b")
End Sub

Private Sub Test6()
    Dim isInclude As Boolean
    Dim result As String
    'CheckProhibitionStringWin "abc", isInclude, False, "", ""
    'CheckProhibitionStringWin "a/bc", isInclude, False, "", ""
    'CheckProhibitionStringExcel "abc", isInclude, False, "", ""
    'CheckProhibitionStringExcel "ab/c", isInclude, False, "", ""
    'CheckProhibitionStringExcel "ab[c", isInclude, False, "", ""
    'CheckProhibitionStringExcel "a/b[c", isInclude, False, "", ""
    'CheckProhibitionStringSharePoint "ab~c", isInclude, False, "", ""
    'CheckProhibitionStringAll "abcd", isInclude, False, "", ""
    'CheckProhibitionStringAll "a/b~c[d", isInclude, False, "", ""
    'CheckProhibitionStringAll "ab~c[d", isInclude, False, "", ""
    'CheckProhibitionStringAll "ab~cd", isInclude, False, "", ""
    'Debug.Print isInclude
    'CheckProhibitionStringWin "a/b\c|d", isInclude, True, "", result
    'CheckProhibitionStringExcel "a/b\c|d", isInclude, True, "", result
    'CheckProhibitionStringExcel "a/b\c[d", isInclude, True, "_", result
    'CheckProhibitionStringSharePoint "a/b\c~d", isInclude, True, "_", result
    'CheckProhibitionStringAll "a/b[c~d", isInclude, True, "_", result
    'Debug.Print result
End Sub

Private Sub Test7()
    'Debug.Print FFIsIncludeProhibitionWin("abc")
    'Debug.Print FFIsIncludeProhibitionWin("ab/c")
    'Debug.Print FFIsIncludeProhibitionWin(" ab c ")
    'Debug.Print FFIsIncludeProhibitionWin("  ")
    'Debug.Print FFIsIncludeProhibitionExcel("abc")
    'Debug.Print FFIsIncludeProhibitionExcel("ab[c")
    'Debug.Print FFIsIncludeProhibitionSharePoint("abc")
    'Debug.Print FFIsIncludeProhibitionSharePoint("ab%c")
    'Debug.Print FFIsIncludeProhibitionAll("abc")
    'Debug.Print FFIsIncludeProhibitionAll("ab%c")
End Sub

Private Sub Test8()
    'Debug.Print FFReplaceProhibitionWin("abc", "_")
    'Debug.Print FFReplaceProhibitionWin("ab/c", "_")
    'Debug.Print FFReplaceProhibitionWin("  ab/c  ", "_")
    Debug.Print FFReplaceProhibitionWin("  ab/c  ", "")
    'Debug.Print FFReplaceProhibitionExcel("abc", "_")
    'Debug.Print FFReplaceProhibitionExcel("ab[c", "_")
    'Debug.Print FFReplaceProhibitionSharePoint("abc", "_")
    'Debug.Print FFReplaceProhibitionSharePoint("ab%c", "_")
    'Debug.Print FFReplaceProhibitionAll("abc", "_")
    'Debug.Print FFReplaceProhibitionAll("ab%c", "_")
End Sub

Private Sub Test9()
    'Debug.Print ReplaceStartEndSpace(" abc ", "_")
    'Debug.Print ReplaceStartEndSpace("", "_")
    'Debug.Print ReplaceStartEndSpace("  ", "_")
    Debug.Print ReplaceStartEndSpace("  ", "")
End Sub

