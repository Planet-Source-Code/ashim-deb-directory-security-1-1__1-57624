VERSION 5.00
Begin VB.Form frmSecurity 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Directory Security"
   ClientHeight    =   2805
   ClientLeft      =   3960
   ClientTop       =   3900
   ClientWidth     =   3795
   Icon            =   "TEST.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "TEST.frx":57E2
   ScaleHeight     =   2805
   ScaleWidth      =   3795
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      Picture         =   "TEST.frx":31CB3
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton cmdUnsecure 
      Caption         =   "&Unsecure"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      Picture         =   "TEST.frx":5E184
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton cmdSecure 
      Caption         =   "&Secure"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      Picture         =   "TEST.frx":8A655
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.DirListBox dirDir 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2130
      Left            =   120
      TabIndex        =   1
      Top             =   570
      Width           =   2295
   End
   Begin VB.DriveListBox drvDrive 
      BackColor       =   &H00EED7CE&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
   Begin VB.Menu ashim 
      Caption         =   "File"
      Index           =   0
      Begin VB.Menu np 
         Caption         =   "Change Password"
         Index           =   1
         Shortcut        =   ^C
      End
      Begin VB.Menu aa 
         Caption         =   "-"
      End
      Begin VB.Menu about 
         Caption         =   "About"
      End
      Begin VB.Menu data 
         Caption         =   "-"
      End
      Begin VB.Menu exit 
         Caption         =   "Exit"
         Index           =   2
         Shortcut        =   ^X
      End
   End
End
Attribute VB_Name = "frmSecurity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetWindowsDirectory Lib "Kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

Dim WindowsDirectory As String

Private Sub about_Click()
Dim MsgText As String
    MsgText = "This Simply Project By Ashim Deb,There Some Project Like -."
    MsgText = MsgText & vbCrLf & "Folder password,I Think U Like That"
    MsgText = MsgText & vbCrLf & "It was Support All Verson Windows 3.1/95/98/2000/Me/Xp "
    MsgText = MsgText & vbCrLf & "E-mail Me:programergroups@yahoo.com & Website: "
    MsgBox MsgText, , "WWW.ashimdeb123.blinkz.com"
End Sub

Private Sub cmdExit_Click()
    End
End Sub

Private Sub cmdSecure_Click()
On Error GoTo Err
    Dim Path As String
    Dim data As String
    Dim File As String
    Dim Ext As String
    Dim FileName As String
    Ext = ".{645FF040-5081-101B-9F08-00AA002F954E}"
    Path = dirDir.Path
    data = Mid$(Path, InStrRev(Path, "\") + 1, Len(Path))
    File = Left$(Path, Len(Path) - Len(data))
    If Not UCase$(Path) = UCase$(WindowsDirectory) _
    And Not UCase$(data) = UCase("desktop") Then
        
        FileName = File & data & Ext
        Name dirDir.Path As FileName
        dirDir.Path = File
        MsgBox "Security applied to the directory.", vbApplicationModal + vbInformation, "Security..."
    Else
        MsgBox "Cann't be protected.", vbApplicationModal + vbInformation, "Security..."
    End If
Err:
    Exit Sub
End Sub

Private Sub cmdUnsecure_Click()
On Error GoTo Err
    Dim Path As String
    Dim Temp As String
    Dim data As String
    Dim File As String
    Dim Ext As String
    Dim FileName As String
    Path = dirDir.Path
    Temp = Mid$(Path, InStrRev(Path, "\") + 1, Len(Path))
    data = Left$(Temp, InStr(Temp, ".{") - 1)
    File = Left$(Path, Len(Path) - Len(Temp))
    FileName = File & data
    Name dirDir.Path As FileName
    dirDir.Path = File
    MsgBox "Security removed to the directory.", vbApplicationModal + vbInformation, "Security..."
Err:
    Exit Sub
End Sub

Private Sub dirDir_Change()
On Error Resume Next
    dirDir.Path = drvDrive.Drive
End Sub

Private Sub drvDrive_Change()
On Error GoTo NotReady
    dirDir.Path = drvDrive.Drive
    Exit Sub
NotReady:
    MsgBox "Drive is not ready.", vbExclamation + vbApplicationModal, "Not Ready..."
End Sub

Private Sub exit_Click(Index As Integer)
End
End Sub

Private Sub Form_Load()


On Error Resume Next
    Dim ret As Long
    Dim buff As String
    buff = Space(255)
    ret = GetWindowsDirectory(buff, 255)
    WindowsDirectory = Left$(buff, InStr(buff, vbNullChar) - 1)
End Sub


Private Sub np_Click(Index As Integer)
Form1.Show
End Sub
