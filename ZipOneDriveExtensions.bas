Attribute VB_Name = "ZipOneDriveExtensions"
Option Explicit

Const nameOfZODE As String = "ZODE"
Const keywordOfExtract As String = "展開せよ"
Const keywordOfCompress As String = "圧縮せよ"
Const folderNameOfExtract As String = "展開"
Const folderNameOfCompress As String = "圧縮"
Const folderPathOfOneDrive As String = "D:\OneDrive"
Const maxCounterOfUniqueFile As Integer = 100
Const extensionsOfTarget As String = "zip,zi_,zi_p"

Dim senderName As String
Dim folderPathOfOneDriveZODE As String
Dim folderPathOfOneDriveExtract As String
Dim folderPathOfOneDriveCompress As String
Dim folderPathOfTmpZODE As String
Dim folderPathOfTmpExtract As String
Dim folderPathOfTmpCompress As String

Sub ZODErecieved(ByVal EntryIDCollection As String)
    If Dir(folderPathOfOneDrive, vbDirectory) = "" Then
        Exit Sub
    End If
    If Is7zipInstalled() = False Then
        Exit Sub
    End If
    
    If CreateFolders() = True Then
        Debug.Print EntryIDCollection
        'Recieved EntryIDCollection
    End If
End Sub

Private Sub Recieved(ByVal EntryIDCollection As String)
    Dim namespace, mailItem As Object
    Set namespace = GetNamespace("MAPI")
    Set mailItem = namespace.GetItemFromID(EntryIDCollection)
    If mailItem.To = mailItem.senderName And mailItem.To = GetSenderFromSentFolder() Then
        If InStr(mailItem.subject, keywordOfExtract) > 0 Then
            Extract mailItem
        ElseIf InStr(mailItem.subject, keywordOfCompress) > 0 Then
            Compress mailItem
        End If
    End If
    Set mailItem = Nothing
    Set namespace = Nothing
End Sub

Private Sub Extract(mailItem As Object)
    Dim currentFolderName, currentFolderPath As String
    Dim f As Object
    currentFolderName = Format(Date, "mmdd") & Format(Time, "hhMMss") & "_" & Left(FFReplaceProhibitionAll(RemoveKeywordInSubject(mailItem.subject), "_"), 8)
    Set f = GetFso()
    currentFolderPath = f.BuildPath(GetFolderPathOfTmpExtract(), currentFolderName)
    If Dir(currentFolderPath, vbDirectory) <> "" Then
        Exit Sub
    End If
    MkDir currentFolderPath
    Dim pass As String
    pass = GetPasswordInBody(mailItem.body)
    Dim attachment As Object
    For Each attachment In mailItem.attachments
        If IsTargetExtension(f.GetExtensionName(attachment.FileName)) = True Then
            Dim currentFilePath As String
            currentFilePath = f.BuildPath(currentFolderPath, attachment.FileName)
            attachment.SaveAsFile CreateUniqueFilePath(currentFilePath)
        End If
    Next attachment
    Set attachment = Nothing
End Sub

Private Sub Compress(objId As Object)

End Sub

Private Function RemoveKeywordInSubject(subject As String) As String
    Dim ret As String
    Dim keywords As Variant
    Dim i As Long
    keywords = Array("RE: ", "Re: ", "FW: ", "Fw: ", keywordOfExtract)
    ret = subject
    For i = LBound(keywords) To UBound(keywords)
        ret = replace(ret, keywords(i), "")
    Next i
    RemoveKeywordInSubject = ret
End Function

Private Function CreateUniqueFilePath(filePath As String) As String
    Dim ret As String
    ret = ""
    If Dir(filePath) <> "" Then
        Dim folder, name, ext As String
        Dim f As Object
        Set f = GetFso()
        folder = f.GetParentFolderName(filePath)
        name = f.GetBaseName(filePath)
        ext = f.GetExtensionName(filePath)
        Dim i As Integer
        For i = 1 To maxCounterOfUniqueFile
            Dim uniqueName As String
            uniqueName = name & Format(i, "000")
            Dim uniquePath As String
            uniquePath = f.BuildPath(folder, uniqueName & "." & ext)
            If Dir(uniquePath) = "" Then
                ret = uniquePath
                Exit For
            End If
        Next i
    Else
        ret = filePath
    End If
    CreateUniqueFilePath = ret
End Function

Private Function CreateFolders() As Boolean
    Dim ret As Boolean
    ret = False
    If Dir(folderPathOfOneDrive, vbDirectory) <> "" Then
        If Dir(GetFolderPathOfOneDriveZODE(), vbDirectory) = "" Then
            MkDir GetFolderPathOfOneDriveZODE()
        End If
        If Dir(GetFolderPathOfOneDriveCompress(), vbDirectory) = "" Then
            MkDir GetFolderPathOfOneDriveCompress()
        End If
        If Dir(GetFolderPathOfOneDriveExtract(), vbDirectory) = "" Then
            MkDir GetFolderPathOfOneDriveExtract()
        End If
        
        If Dir(GetFolderPathOfTmpZODE(), vbDirectory) = "" Then
            MkDir GetFolderPathOfTmpZODE()
        End If
        If Dir(GetFolderPathOfTmpCompress(), vbDirectory) = "" Then
            MkDir GetFolderPathOfTmpCompress()
        End If
        If Dir(GetFolderPathOfTmpExtract(), vbDirectory) = "" Then
            MkDir GetFolderPathOfTmpExtract()
        End If
        ret = True
    Else
        ret = False
    End If
    CreateFolders = ret
End Function

Private Function GetSenderFromSentFolder() As String
    If senderName = vbNullString Then
        Dim namespace, sentFolder As Object
        Set namespace = GetNamespace("MAPI")
        Set sentFolder = namespace.GetDefaultFolder(olFolderSentMail)
        Dim mailItem As Object
        Set mailItem = sentFolder.Items(1)
        senderName = mailItem.senderName
    End If
    GetSenderFromSentFolder = senderName
End Function

Private Function IsTargetExtension(extension As String) As Boolean
    Dim ret As Boolean
    Dim extensions As Variant
    Dim i As Long
    ret = False
    extensions = Split(extensionsOfTarget, ",")
    For i = LBound(extensions) To UBound(extensions)
        If UCase(extension) = UCase(extensions(i)) Then
            ret = True
            Exit For
        End If
    Next i
    IsTargetExtension = ret
End Function

Private Function GetPasswordInBody(body As String) As String
    Dim ret As String
    Dim tmp As Variant
    ret = ""
    tmp = Split(body, vbCrLf)
    If UBound(tmp) <> -1 Then
        ret = tmp(LBound(tmp))
    End If
    GetPasswordInBody = ret
End Function

Private Function GetFso() As Object
    Set GetFso = FFGetFso()
End Function

Private Function GetFolderPathOfOneDriveZODE() As String
    If folderPathOfOneDriveZODE = vbNullString Then
        folderPathOfOneDriveZODE = GetFso().BuildPath(folderPathOfOneDrive, nameOfZODE)
    End If
    GetFolderPathOfOneDriveZODE = folderPathOfOneDriveZODE
End Function

Private Function GetFolderPathOfOneDriveExtract() As String
    If folderPathOfOneDriveExtract = vbNullString Then
        folderPathOfOneDriveExtract = GetFso().BuildPath(GetFolderPathOfOneDriveZODE(), folderNameOfExtract)
    End If
    GetFolderPathOfOneDriveExtract = folderPathOfOneDriveExtract
End Function

Private Function GetFolderPathOfOneDriveCompress() As String
    If folderPathOfOneDriveCompress = vbNullString Then
        folderPathOfOneDriveCompress = GetFso().BuildPath(GetFolderPathOfOneDriveZODE(), folderNameOfCompress)
    End If
    GetFolderPathOfOneDriveCompress = folderPathOfOneDriveCompress
End Function

Private Function GetFolderPathOfTmpZODE() As String
    If folderPathOfTmpZODE = vbNullString Then
        folderPathOfTmpZODE = GetFso().BuildPath(Environ("TMP"), nameOfZODE)
    End If
    GetFolderPathOfTmpZODE = folderPathOfTmpZODE
End Function

Private Function GetFolderPathOfTmpExtract() As String
    If folderPathOfTmpExtract = vbNullString Then
        folderPathOfTmpExtract = GetFso().BuildPath(GetFolderPathOfTmpZODE(), folderNameOfExtract)
    End If
    GetFolderPathOfTmpExtract = folderPathOfTmpExtract
End Function

Private Function GetFolderPathOfTmpCompress() As String
    If folderPathOfTmpCompress = vbNullString Then
        folderPathOfTmpCompress = GetFso().BuildPath(GetFolderPathOfTmpZODE(), folderNameOfCompress)
    End If
    GetFolderPathOfTmpCompress = folderPathOfTmpCompress
End Function

Private Sub Test1()
    'Recieved "00000000DCD046AF784B41439F6B3EE112AD60350700D4E3394F76CE43458D782CFAA00BC4DD0000007649CB000033EEE52455E77F4FA12B877F3493A24700020BBC1C470000"
    'Recieved "00000000DCD046AF784B41439F6B3EE112AD60350700D4E3394F76CE43458D782CFAA00BC4DD0000007649CB000033EEE52455E77F4FA12B877F3493A24700020BBC1C4B0000"
    'Recieved "00000000DCD046AF784B41439F6B3EE112AD60350700D4E3394F76CE43458D782CFAA00BC4DD0000007649CB000033EEE52455E77F4FA12B877F3493A24700020BBC1C4F0000"
    'Recieved "00000000DCD046AF784B41439F6B3EE112AD60350700D4E3394F76CE43458D782CFAA00BC4DD0000007649CB000033EEE52455E77F4FA12B877F3493A24700020BBC1C540000" 'テキスト
    'Recieved "00000000DCD046AF784B41439F6B3EE112AD60350700D4E3394F76CE43458D782CFAA00BC4DD0000007649CB000033EEE52455E77F4FA12B877F3493A24700020BBC1C570000" 'html
    'Recieved "00000000DCD046AF784B41439F6B3EE112AD60350700D4E3394F76CE43458D782CFAA00BC4DD0000007649CB000033EEE52455E77F4FA12B877F3493A24700020BBC1C580000" 'リッチテキスト
    Recieved "00000000DCD046AF784B41439F6B3EE112AD60350700D4E3394F76CE43458D782CFAA00BC4DD0000007649CB000033EEE52455E77F4FA12B877F3493A24700020BBC1C5F0000"
End Sub

Private Sub Test2()
    Debug.Print GetSenderFromSentFolder()
End Sub

Private Sub Test3()
    Dim a As Object
    Set a = GetFso()
    Debug.Print GetFso().BuildPath("c", "d")
End Sub

Private Sub Test4()
   'Debug.Print GetFolderPathOfOneDriveZODE()
   'Debug.Print GetFolderPathOfOneDriveZODE()
   'Debug.Print GetFolderPathOfOneDriveExtract()
   'Debug.Print GetFolderPathOfOneDriveExtract()
   'Debug.Print GetFolderPathOfOneDriveCompress()
   'Debug.Print GetFolderPathOfOneDriveCompress()
   'Debug.Print GetFolderPathOfTmpExtract()
   'Debug.Print GetFolderPathOfTmpExtract()
   'Debug.Print GetFolderPathOfTmpCompress()
   'Debug.Print GetFolderPathOfTmpCompress()
End Sub

Private Sub Test5()
    Debug.Print CreateFolders()
    Debug.Print CreateFolders()
End Sub

Private Sub Test6()
    Debug.Print CreateUniqueFilePath("D:\UiPathLogToExcel.zip")
End Sub

Private Sub Test7()
    'Debug.Print Year(Date) & " " & Month(Date) & " " & Day(Date) & " " & Hour(Time) & " " & Minute(Time) & " " & Second(Time)
    Debug.Print Format(Date, "yyyymmdd") & " " & Format(Time, "hhMMss")
End Sub

Private Sub Test8()
    Debug.Print RemoveKeywordInSubject("Re: テスト")
End Sub

Private Sub Test9()
    'Debug.Print IsTargetExtension("zip")
    Debug.Print IsTargetExtension("Zi_p")
End Sub

Private Sub Test10()
    'Debug.Print GetPasswordInSubject("")
    'Debug.Print GetPasswordInSubject("test")
    'Debug.Print GetPasswordInSubject("test" & vbCrLf & vbCrLf)
    'Debug.Print GetPasswordInSubject("a" & vbCrLf & "b")
End Sub
