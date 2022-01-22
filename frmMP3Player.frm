VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmMP3_Player 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AiE VB5 MP3 Player Example Projects By : Rw Bivins"
   ClientHeight    =   1440
   ClientLeft      =   150
   ClientTop       =   795
   ClientWidth     =   6705
   Icon            =   "frmMP3Player.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1440
   ScaleWidth      =   6705
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1455
      Left            =   2520
      ScaleHeight     =   1455
      ScaleWidth      =   3855
      TabIndex        =   0
      Top             =   0
      Width           =   3855
      Begin VB.VScrollBar VScroll1 
         Height          =   3615
         Left            =   0
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   0
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Timer Timer1 
         Interval        =   55
         Left            =   240
         Top             =   3120
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H00C0C0C0&
         Height          =   6000
         Left            =   240
         MouseIcon       =   "frmMP3Player.frx":030A
         MousePointer    =   99  'Custom
         MultiLine       =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         Text            =   "frmMP3Player.frx":0614
         Top             =   0
         Width           =   3495
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   2040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   $"frmMP3Player.frx":07FD
   End
   Begin VB.Image Image1 
      Height          =   1260
      Left            =   120
      Picture         =   "frmMP3Player.frx":0886
      Top             =   120
      Width           =   2250
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuOpen 
         Caption         =   "Play"
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuStop 
         Caption         =   "Stop"
      End
      Begin VB.Menu sep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAiE 
         Caption         =   "AiE "
         Begin VB.Menu mnuWebsite 
            Caption         =   "WebSite"
         End
         Begin VB.Menu sep3 
            Caption         =   "-"
         End
         Begin VB.Menu mnuWebBoard 
            Caption         =   "WebBoard"
         End
         Begin VB.Menu sep4 
            Caption         =   "-"
         End
         Begin VB.Menu mnudownloads 
            Caption         =   "Downloads"
         End
      End
   End
End
Attribute VB_Name = "frmMP3_Player"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
GradientForm Me
VScroll1.Max = Picture1.Height
VScroll1.Min = 0 - Text1.Height
VScroll1.Value = VScroll1.Max
End Sub
Private Sub mnuOpen_Click()
'Dim MP3File As String
'CommonDialog1.ShowOpen
'      MP3File = CommonDialog1.filename
ActiveMovie1.FileName = "AiE.mp3"
ActiveMovie1.Run
End Sub

Private Sub mnuStop_Click()
ActiveMovie1.Stop
End Sub

Private Sub mnuWebBoard_Click()
Dim ret&
    ret = ShellExecute(Me.hwnd, "Open", _
        "http://disc.server.com/discussion.cgi?id=17429", _
        "", App.Path, 1)
End Sub
Private Sub mnuWebsite_Click()
Dim ret&
    ret = ShellExecute(Me.hwnd, "Open", _
        "http://members.tripod.com/~AceInterActive/", _
        "", App.Path, 1)
End Sub
Private Sub Timer1_Timer()
'Move the textbox here
If VScroll1.Value >= VScroll1.Min + 30 Then
  VScroll1.Value = VScroll1.Value - 20
Else
  VScroll1.Value = VScroll1.Max
  DoEvents
End If
Text1.Top = VScroll1.Value
DoEvents
End Sub
