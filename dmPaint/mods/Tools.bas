Attribute VB_Name = "Tools"
Option Explicit

Public ButtonPress As VbMsgBoxResult
Public DataDir As String

Public Function FixPath(lPath As String) As String
    If Right$(lPath, 1) = "\" Then
        FixPath = lPath
    Else
        FixPath = lPath & "\"
    End If
End Function

Public Function GetFilePath(ByVal lPath As String) As String
Dim sPos As Integer
    'Return file path
    sPos = InStrRev(lPath, "\", Len(lPath), vbBinaryCompare)
    
    If (sPos > 0) Then
        GetFilePath = Left$(lPath, sPos - 1)
    Else
        GetFilePath = lPath
    End If
    
End Function

Public Function GetFileExt(ByVal lFilename As String) As String
Dim sPos As Integer
    'Return file .ext
    sPos = InStrRev(lFilename, ".", Len(lFilename), vbBinaryCompare)
    
    If (sPos > 0) Then
        GetFileExt = Mid$(lFilename, sPos + 1)
    End If
    
End Function
