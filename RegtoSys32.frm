VERSION 5.00
Begin VB.Form RegtoSys32 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Move file to system32 and register"
   ClientHeight    =   765
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4965
   Icon            =   "RegtoSys32.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   765
   ScaleWidth      =   4965
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAddToShell 
      Caption         =   "Add to Shell"
      Height          =   300
      Left            =   30
      TabIndex        =   2
      Top             =   390
      Width           =   2430
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Remove from Shell"
      Height          =   315
      Left            =   2490
      TabIndex        =   1
      Top             =   390
      Width           =   2412
   End
   Begin VB.TextBox txtFilename 
      Height          =   285
      Left            =   -75
      TabIndex        =   0
      Top             =   1410
      Width           =   5295
   End
   Begin VB.Label Label1 
      Caption         =   "You must compile this and drag/drop dll or ocx on the executable!!"
      Height          =   240
      Left            =   75
      TabIndex        =   3
      Top             =   75
      Width           =   4860
   End
End
Attribute VB_Name = "RegtoSys32"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Function GetFileTitle(ByVal sFileName As String) As String
    'Returns FileTitle Without Path
    Dim lPos As Long
    lPos = InStrRev(sFileName, "\")


    If lPos > 0 Then


        If lPos < Len(sFileName) Then
            GetFileTitle = Mid$(sFileName, lPos + 1)
        Else
            GetFileTitle = ""
        End If
    Else
        GetFileTitle = sFileName
    End If
    
End Function


Private Sub cmdAddToShell_Click()
Dim RegSvr32ForWindowsFileName As String
'If Right$(App.Path, 1) = "\" Then
    RegSvr32ForWindowsFileName = App.Path & "\RegtoSys32.exe"  '"c:\windows\system32\RegtoSys32.exe"
'Else
    'RegSvr32ForWindowsFileName = App.Path & "\" & App.EXEName & ".exe"
'End If

CreateNewKey HKEY_CLASSES_ROOT, "dllfile\shell\reg\command"
CreateNewKey HKEY_CLASSES_ROOT, "dllfile\shell\unreg\command"
CreateNewKey HKEY_CLASSES_ROOT, "ocxfile\shell\reg\command"
CreateNewKey HKEY_CLASSES_ROOT, "ocxfile\shell\unreg\command"

SetKeyValue HKEY_CLASSES_ROOT, "dllfile\shell\reg", (Standard), "Register", REG_SZ
SetKeyValue HKEY_CLASSES_ROOT, "dllfile\shell\unreg", (Standard), "Unregister", REG_SZ
SetKeyValue HKEY_CLASSES_ROOT, "dllfile\shell\reg\command", (Standard), Chr(34) & RegSvr32ForWindowsFileName & Chr(34) & " " & Chr(34) & "%1" & Chr(34), REG_SZ
SetKeyValue HKEY_CLASSES_ROOT, "dllfile\shell\unreg\command", (Standard), Chr(34) & RegSvr32ForWindowsFileName & Chr(34) & " /u " & Chr(34) & "%1" & Chr(34), REG_SZ

SetKeyValue HKEY_CLASSES_ROOT, "ocxfile\shell\reg", (Standard), "Register", REG_SZ
SetKeyValue HKEY_CLASSES_ROOT, "ocxfile\shell\unreg", (Standard), "Unregister", REG_SZ
SetKeyValue HKEY_CLASSES_ROOT, "ocxfile\shell\reg\command", (Standard), Chr(34) & RegSvr32ForWindowsFileName & Chr(34) & " " & Chr(34) & "%1" & Chr(34), REG_SZ
SetKeyValue HKEY_CLASSES_ROOT, "ocxfile\shell\unreg\command", (Standard), Chr(34) & RegSvr32ForWindowsFileName & Chr(34) & " /u " & Chr(34) & "%1" & Chr(34), REG_SZ

End Sub

Private Sub Command2_Click()
DeleteKey HKEY_CLASSES_ROOT, "dllfile\shell\reg\command"
DeleteKey HKEY_CLASSES_ROOT, "dllfile\shell\unreg\command"
DeleteKey HKEY_CLASSES_ROOT, "dllfile\shell\reg"
DeleteKey HKEY_CLASSES_ROOT, "dllfile\shell\unreg"

DeleteKey HKEY_CLASSES_ROOT, "ocxfile\shell\reg\command"
DeleteKey HKEY_CLASSES_ROOT, "ocxfile\shell\unreg\command"
DeleteKey HKEY_CLASSES_ROOT, "ocxfile\shell\reg"
DeleteKey HKEY_CLASSES_ROOT, "ocxfile\shell\unreg"

End Sub

Private Sub Form_Load()
Dim System32Path As String, Holden As String
Dim FileToReg As String

If Trim(Command$) = "" Then
        Exit Sub
        MsgBox "You must compile this and " & vbCrLf _
                    & "drag/drop dll or ocx " & vbCrLf & "on the executable!!"
        Unload Me
        Exit Sub
End If
'Holden = Replace(Command$, Chr(34), "")
'MsgBox System32Path
'MsgBox System32Path & "\" & GetFileTitle(Command$)
    Reg = Register(Command$)
    FileToReg = GetFileNameFromCommandLine(Command$)
    Holden = Trim$(Mid$(FileToReg, 2, Len(FileToReg) - 2))
    If Right(LCase(Holden), 4) <> ".dll" And Right(LCase(Holden), 4) <> ".ocx" Then
            MsgBox "You must drag/drop only a dll or ocx!!"
            Unload Me
            Exit Sub
    End If
    System32Path = GetShellFolderPath(&H25)
    FileCopy Holden, System32Path & "\" & GetFileTitle(Holden)
    txtFilename.Text = System32Path & "\" & GetFileTitle(Holden)
    Reg = Register(Command$)
    If Reg = True Then
        If RegServer(txtFilename.Text) = True Then
            MsgBox txtFilename.Text & " was correctly registered.", vbInformation, "Success"
        Else
            MsgBox "Failure registering " & txtFilename.Text & ".", vbCritical, "Failure"
        End If
    ElseIf Reg = False Then
        If UnRegServer(txtFilename.Text) = True Then
            MsgBox txtFilename.Text & " was correctly unregistered.", vbInformation, "Success"
        Else
            MsgBox "Failure unregistering " & txtFilename.Text & ".", vbCritical, "Failure"
        End If
    End If
Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set RegtoSys32 = Nothing
End Sub
