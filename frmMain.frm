VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "SWFLASH.OCX"
Begin VB.Form frmMain 
   Caption         =   "Shock! Install"
   ClientHeight    =   5970
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8190
   ClipControls    =   0   'False
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   398
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   546
   StartUpPosition =   3  'Windows Default
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash ShockwaveFlash1 
      Height          =   6000
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8250
      Movie           =   ""
      Src             =   ""
      WMode           =   "Window"
      Play            =   -1  'True
      Loop            =   -1  'True
      Quality         =   "AutoHigh"
      SAlign          =   ""
      Menu            =   -1  'True
      Base            =   ""
      Scale           =   "ShowAll"
      DeviceFont      =   0   'False
      EmbedMovie      =   0   'False
      BGColor         =   ""
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long


'This program interperets FS commands from Flash movies. If you dont know
'what that means, you will need to learn Flash.


Private Sub Form_Load()

Dim strTmp As String
strTmp = command$



strTmp = Replace(strTmp, "$APPDIR$", GetAppPath)   'Its easier than keep retyping absolute paths (using dollar tags cos windows strips %param% formatted parameters)



If Dir(strTmp) = "" Or strTmp = "" Then
    MsgBox "Flash file not found!", vbCritical, "Error"
    Unload Me
    End
End If

ShockwaveFlash1.Movie = strTmp




'ShockwaveFlash1.Movie = "f:\test.swf"
'ShockwaveFlash1.Movie = "\\KNEPTUNE\HARD_DISK_3\test.swf"
'ShockwaveFlash1.Movie = "http://147.132.18.1/solsupp/flash/test.swf"

End Sub

Private Sub ShockwaveFlash1_FSCommand(ByVal command As String, ByVal args As String)

Select Case UCase$(command)
    Case "QUIT"
        ShockwaveFlash1.Stop
        'NEVER USE UNLOAD EVENT HERE, OCX will crash the kernel
        End
    Case "SHELL"
        
        'Replace reserved words
        
        args = Replace(args, "%WINDIR%", WindowsDirectory)
        args = Replace(args, "%TEMPDIR%", TempDirectory)
        
        
        'usr want to shell
        If Mid(args, 2, 1) = ":" Then 'must be absolute path
            ShellExecute vbNull, "open", args, "", "", 1
        Else 'and this relative path
            ShellExecute vbNull, "open", GetAppPath & args, "", "", 1
        End If
        
        
    Case "WIDTH"
        On Error GoTo bad_fs_val
        If Not IsNumeric(args) Then GoTo bad_fs_val
        ShockwaveFlash1.Width = Val(args)
        Me.ScaleWidth = Val(args) + 10
    
    Case "HEIGHT"
        On Error GoTo bad_fs_val
        If Not IsNumeric(args) Then GoTo bad_fs_val
        ShockwaveFlash1.Height = Val(args)
        Me.ScaleHeight = Val(args) + 10
        
    Case "INFOBOX"
        MsgBox args, vbInformation, "Information"

    Case Else
    
End Select

Exit Sub
bad_fs_val:
    MsgBox "Bad FS value sent to Shock! Install" & vbCrLf & vbCrLf & "Please check the FS Action arguments for " & command, vbExclamation, "Error"
    ShockwaveFlash1.Stop
    End
End Sub



