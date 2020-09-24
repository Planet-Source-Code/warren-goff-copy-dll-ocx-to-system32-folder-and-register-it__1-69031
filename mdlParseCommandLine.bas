Attribute VB_Name = "mdlParseCommandLine"
Option Explicit
Public Reg As Boolean

Public Function Register(ByVal CommandLine As String) As Boolean
If LCase$(Trim$(Left$(CommandLine, 2))) = "/u" Then
    Register = False
Else
    Register = True
End If
End Function

Public Function GetFileNameFromCommandLine(ByVal CommandLine As String) As String
If Reg = True Then
    GetFileNameFromCommandLine = Trim$(CommandLine)
ElseIf Reg = False Then
    GetFileNameFromCommandLine = Trim$(Mid$(CommandLine, 3, Len(CommandLine) - 2))
End If
End Function
