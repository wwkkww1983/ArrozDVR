Attribute VB_Name = "File"
Option Explicit

Public Function JoinPathFile(p As String, f As String) As String
    JoinPathFile = AddSlash(p) & f
End Function

Public Function AddSlash(Directory As String) As String
    If Right$(Directory, 1) <> "\" Then Directory = Directory + "\"
    AddSlash = Directory
End Function

Public Sub RemoveSlash(Directory As String)
    If Len(Directory) > 3 And InStrRev(Directory, "\") = Len(Directory) Then Directory = Left$(Directory, Len(Directory) - 1)
End Sub

'分解出文件和目录
Public Function CutPathFile(nStr As String, nPath As String, nFile As String)
    Dim i As Long, s As Long
    
    For i = 1 To Len(nStr)
        If Mid(nStr, i, 1) = "\" Then s = i                                     '查找最后一个目录分隔符
    Next
    If s > 0 Then
        nPath = Left(nStr, s): nFile = Mid(nStr, s + 1)
    Else
        nPath = "": nFile = nStr
    End If
End Function

'逐级建立目录,成功返回 True
Public Function MakePath(ByVal nPath As String) As Boolean
    Dim i As Long, Path1 As String, IsPath As Boolean
    nPath = Trim(nPath)
    If Right(nPath, 1) <> "\" Then nPath = nPath & "\"
    On Error GoTo Exit1
    For i = 1 To Len(nPath)
        If Mid(nPath, i, 1) = "\" Then
            Path1 = Left(nPath, i - 1)
            If Dir(Path1, 23) = "" Then
                MkDir Path1
            Else
                IsPath = GetAttr(Path1) And 16
                If Not IsPath Then Exit Function                                '有一个同名的文件
            End If
        End If
    Next
    MakePath = True: Exit Function
Exit1:
End Function

'检查目录或文件夹，返回值：0不存在，1是文件，2是目录
Public Function CheckDirFile(nDirFile) As Long
    Dim nStr As String, nD As Boolean
    nStr = Dir(nDirFile, 23)
    If nStr = "" Then Exit Function
    nD = GetAttr(nDirFile) And 16
    If nD Then CheckDirFile = 2 Else CheckDirFile = 1
End Function

'查找指定目录下的所有文件，返回路径和文件名，不支持子文件夹递归
'调用示例
'  SearchFiles "C:\Program Files\WinRAR\", "*" '查找所有文件
'  SearchFiles "C:\Program Files\WinRAR\", "*.exe" '查找所有exe文件
'  SearchFiles "C:\Program Files\WinRAR\", "*in*.exe" '查找文件名中包含有 in 的exe文件
Public Function SearchFiles(Path As String, FileType As String) As String()
    Dim sPath As String, numFiles As Long
    Dim saFiles() As String
    
    If Right$(Path, 1) <> "\" Then Path = Path & "\"
    
    sPath = Dir(Path & FileType) '查找第一个文件
    
    numFiles = 0
    Do While Len(sPath) '循环到没有文件为止
        ReDim Preserve saFiles(numFiles) As String
        saFiles(numFiles) = Path & sPath
        numFiles = numFiles + 1
        sPath = Dir '查找下一个文件
        'DoEvents '让出控制权
    Loop
    
    If numFiles Then
        SearchFiles = saFiles
    Else
        SearchFiles = Split("")
    End If
End Function

