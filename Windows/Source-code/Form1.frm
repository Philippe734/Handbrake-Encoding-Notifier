VERSION 5.00
Begin VB.Form FrmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "HandBrake"
   ClientHeight    =   975
   ClientLeft      =   14220
   ClientTop       =   13215
   ClientWidth     =   4740
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   975
   ScaleWidth      =   4740
   Begin VB.Frame Frame 
      Caption         =   "Encoding is over"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      Begin VB.TextBox txtFile 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   1
         Text            =   "txtFile"
         Top             =   300
         Width           =   4095
      End
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private sFileName As String
Private sFilePath As String

Private Declare Function InitCommonControls Lib "comctl32.dll" () As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal HWnd As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal HWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub Form_Initialize()
    InitCommonControls
End Sub

Private Sub Form_Load()
' windows position
    Me.Top = Screen.Height - 2500
    Me.Left = Screen.Width - Me.Width - 200
End Sub

Private Sub Form_Activate()

' get command line
    sFilePath = Command
    ' sFilePath = "C:\dossier\film jlk fqsmdlkjaze fsqdfq.mp4"
    sFilePath = Replace(sFilePath, Chr(34), vbNullString)
    sFileName = Right$(sFilePath, Len(sFilePath) - InStrRev(sFilePath, "\"))


    txtFile.Text = sFileName

    ' set icon in systray
    SystrayOn Me, "Encoding is over"

    DoEvents

    ' set tooltip
    Call PopupBalloon(Me, "Encoding is over" & vbNewLine & sFileName, "HandBrake")

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim bFlag As Boolean

    If bFlag = True Then Exit Sub
    bFlag = True

    'If Button = vbRightButton Then

    ' this is for exit menu when user clic out
    '    Call SetForegroundWindow(Me.HWnd)

    ' show menu popup
    'PopupMenu mnuPop

    'End If

    bFlag = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SystrayOff Me
End Sub

Private Sub ClicEvent()

' open video encoded
    On Error Resume Next
    If Len(sFileName) > 4 Then
        ShellExecute 0&, vbNullString, sFilePath, vbNullString, vbNullString, vbNormalFocus
        DoEvents
    End If
    ' quit
    Unload Me

End Sub

Private Sub Form_Click()
    Call ClicEvent
End Sub

Private Sub Frame_Click()
    Call ClicEvent
End Sub

Private Sub txtFile_Click()
    Call ClicEvent
End Sub

