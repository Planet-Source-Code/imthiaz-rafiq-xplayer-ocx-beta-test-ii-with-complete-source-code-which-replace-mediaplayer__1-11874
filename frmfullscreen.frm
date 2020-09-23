VERSION 5.00
Begin VB.Form frmfullscreen 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   3195
   ClientLeft      =   -330
   ClientTop       =   -345
   ClientWidth     =   7380
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   7380
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Timer EraseString 
      Interval        =   1
      Left            =   2070
      Top             =   2460
   End
   Begin VB.PictureBox Screener 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4500
      Left            =   1350
      ScaleHeight     =   4500
      ScaleWidth      =   6495
      TabIndex        =   0
      Top             =   1620
      Width           =   6500
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmfullscreen.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   585
      Left            =   90
      TabIndex        =   1
      Top             =   150
      Width           =   7095
   End
End
Attribute VB_Name = "frmfullscreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub EraseString_Timer()
CurString = ""
EraseString.Enabled = False
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
  Case 27: Unload Me
  
  Case 99: CurString = "Pause"
  Case 67: CurString = "Pause"
  Case 32: CurString = "Pause"
  
  Case 98: CurString = "Stop"
  Case 66: CurString = "Stop"
  
  Case 118: CurString = "FF"
  Case 86: CurString = "FF"
  
  Case 120: CurString = "Play"
  Case 88: CurString = "Play"
  
  Case 90: CurString = "RR"
  Case 122: CurString = "RR"
  
  Case 48 To 58: CurString = "V" + Str(KeyAscii - 48)
  
  Case Else
    frmfullhelp.Show vbModal
End Select

End Sub

Private Sub Form_Load()
Me.Top = 0
Me.Left = 0
Me.WindowState = 2

If Started = False And PlayMode = "" Then
  Started = True
  PlayMode = "Play"
  Form_KeyPress 98
ElseIf Started = True And PlayMode = "Play" Then
  Form_KeyPress 120
ElseIf Started = True And PlayMode = "Pause" Then
  Form_KeyPress 99
End If

End Sub

Private Sub Timer1_Timer()

End Sub
