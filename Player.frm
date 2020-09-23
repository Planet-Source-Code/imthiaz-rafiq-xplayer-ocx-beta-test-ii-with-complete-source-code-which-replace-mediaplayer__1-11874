VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Frmplayer 
   Caption         =   "PlayerX.ocx Beta test III With Source Code........."
   ClientHeight    =   4380
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   7125
   Icon            =   "Player.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   4380
   ScaleWidth      =   7125
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   1620
      Left            =   540
      TabIndex        =   8
      Top             =   2160
      Visible         =   0   'False
      Width           =   6315
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Height          =   645
      Left            =   150
      TabIndex        =   1
      Top             =   1050
      Width           =   6885
      Begin XPlayer.GoldButton GoldButton2 
         Height          =   345
         Left            =   2304
         TabIndex        =   2
         Top             =   180
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   609
         Caption         =   "<<"
         Alignment       =   2
         ForeColor       =   -2147483630
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         OnDown          =   2
         OnHover         =   5
      End
      Begin XPlayer.GoldButton GoldButton1 
         Height          =   345
         Left            =   60
         TabIndex        =   3
         Top             =   180
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   609
         Caption         =   "Playlist"
         Alignment       =   2
         ForeColor       =   -2147483630
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         OnDown          =   2
         OnHover         =   5
      End
      Begin XPlayer.GoldButton GoldButton3 
         Height          =   345
         Left            =   4638
         TabIndex        =   4
         Top             =   180
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   609
         Caption         =   "Mp3 Info"
         Alignment       =   2
         ForeColor       =   -2147483630
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         OnDown          =   2
         OnHover         =   5
      End
      Begin XPlayer.GoldButton GoldButton4 
         Height          =   345
         Left            =   5760
         TabIndex        =   5
         Top             =   180
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   609
         Caption         =   "Cd O/C"
         Alignment       =   2
         ForeColor       =   -2147483630
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         OnDown          =   2
         OnHover         =   5
      End
      Begin XPlayer.GoldButton GoldButton5 
         Height          =   345
         Left            =   3516
         TabIndex        =   6
         Top             =   180
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   609
         Caption         =   ">>"
         Alignment       =   2
         ForeColor       =   -2147483630
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         OnDown          =   2
         OnHover         =   5
      End
      Begin XPlayer.GoldButton GoldButton6 
         Height          =   345
         Left            =   1182
         TabIndex        =   7
         Top             =   180
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   609
         Alignment       =   2
         ForeColor       =   -2147483630
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         OnDown          =   2
         OnHover         =   5
      End
   End
   Begin XPlayer.PlayerX Player1 
      Height          =   960
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   1693
      CurrentPosition =   0
      ChildNo         =   705.547485351563
      ColorVideoWindow=   8421504
      ColorStatusText =   16744576
      ColorPositionBarForeColor=   255
      ColorPositionBarBackColor=   64
      ColorVolumeBarForeColor=   16711680
      ColorVolumeBarBackColor=   4194304
      Repeat          =   -1  'True
      FastForwardBySec=   5
      FastRewindBySec =   5
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   5040
      Top             =   2190
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Menu mnusimple 
      Caption         =   "Simple"
      Visible         =   0   'False
      Begin VB.Menu mnu1 
         Caption         =   "1"
      End
      Begin VB.Menu mnu2 
         Caption         =   "2"
      End
      Begin VB.Menu mnu3 
         Caption         =   "3"
      End
      Begin VB.Menu mnu4 
         Caption         =   "4"
      End
   End
End
Attribute VB_Name = "Frmplayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim size As Boolean
Dim X As Boolean
Dim PlayCount As Integer




Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 98 Then Player1.ResumePlay
If KeyAscii = 114 Then Player1.Play
If KeyAscii = 112 Then Player1.Pause
Player1.Play
End Sub

Private Sub Form_Load()
If App.PrevInstance = True Then
  Player1.ChildNo = Abs(App.ThreadID)
End If
If Player1.MovieSizer = Original_Size Then
  GoldButton6.Caption = "Org Size"
Else
  GoldButton6.Caption = "Fit to Size"
End If
X = True
End Sub

Private Sub Form_Resize()
'Exit Sub
Player1.Top = 0
Player1.Left = 0
If Frmplayer.WindowState = 0 Then
  If Frmplayer.Height < Player1.Height + Frame1.Height + 450 Then
    Frmplayer.Height = Player1.Height + Frame1.Height + 450
  Else
    Frmplayer.Height = Player1.Height + Frame1.Height + 450
  End If
  If Frmplayer.Width < Player1.Width + 100 Then
    Frmplayer.Width = Player1.Width + 100
  Else
    Frmplayer.Width = Player1.Width + 100
  End If
  Frame1.Top = Player1.Height
  Frame1.Width = Frmplayer.ScaleWidth
  Frame1.Left = 0
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Player1.ClosePlayers
    Player1.ShowAbout
End Sub

Private Sub GoldButton1_Click()
    PlayCount = 0
    Form1.Show vbModal
    If List1.ListCount <> 0 Then
      PlayCount = -1
      List1.ListIndex = 0
      Player1_OnPlayFinish
    End If
End Sub
Sub XX()
On Error GoTo Errhand
If cd1.FileName = "" Then
  cd1.InitDir = App.Path
Else
  For i = Len(cd1.FileName) To 1 Step -1
    If Mid$(cd1.FileName, i, 1) = "\" Then Exit For
  Next i
  cd1.InitDir = Mid$(cd1.FileName, 1, i + 1)
End If
cd1.Filter = "All Supported files (*.dat;*.avi;*.mp3;*.wav;*.mpg;*.wma)|*.dat;*.avi;*.mp3;*.wav;*.mpg;*.wma"
cd1.DialogTitle = "Open a media file to play"
cd1.ShowOpen
If cd1.FileName <> "" Then
  Player1.File = cd1.FileName
End If
Exit Sub
Errhand:

End Sub

Private Sub GoldButton2_Click()
If List1.ListCount <> 0 Then
    If PlayCount = 0 Then
        PlayCount = List1.ListCount - 1
    End If
    PlayCount = PlayCount - 1
    List1.ListIndex = PlayCount
    Player1.ClosePlayers
    Player1.File = List1.Text
    Player1.Play
End If
End Sub

Private Sub GoldButton3_Click()
  MsgBox Player1.Mp3Information, vbInformation + vbOKOnly
End Sub

Private Sub GoldButton4_Click()
If X = True Then
  X = False
  Player1.OpenCdBay
Else
  X = True
  Player1.CloseCdBay
End If
End Sub

Private Sub GoldButton5_Click()
    Player1_OnPlayFinish
End Sub

Private Sub GoldButton6_Click()
If Player1.MovieSizer = Original_Size Then
  GoldButton6.Caption = "Fit to Size"
  Player1.MovieSizer = Resize_to_Fit_Current_Control
Else
  GoldButton6.Caption = "Org Size"
  Player1.MovieSizer = Original_Size
End If
End Sub

Private Sub mnu1_Click()
If Player1.ShowPosBar = True Then
  Player1.ShowPosBar = False
Else
  Player1.ShowPosBar = True
End If
End Sub

Private Sub mnu2_Click()
If Player1.ShowControls = True Then
  Player1.ShowControls = False
Else
  Player1.ShowControls = True
End If
End Sub

Private Sub mnu3_Click()
If Player1.ShowStatus = True Then
  Player1.ShowStatus = False
Else
  Player1.ShowStatus = True
End If

End Sub

Private Sub mnu4_Click()
If Player1.ShowVolume = True Then
  Player1.ShowVolume = False
Else
  Player1.ShowVolume = True
End If
End Sub

Private Sub Player1_OnMouseMove(Button As Integer, X As Single, Y As Single)
If Button = 2 Then
  SetmenuCaption
  PopupMenu mnusimple
End If
End Sub

Private Sub Player1_OnPlayFinish()
If List1.ListCount <> 0 Then
    If PlayCount = List1.ListCount - 1 Then
        PlayCount = -1
    End If
    PlayCount = PlayCount + 1
    List1.ListIndex = PlayCount
    Player1.ClosePlayers
    Player1.File = List1.Text
    Player1.Play
End If
End Sub

Private Sub Player1_OnSelfResize(NewHeight As Long, NewWidth As Long)
size = True
Form_Resize
End Sub

Sub SetmenuCaption()
If Player1.ShowPosBar = True Then
  mnu1.Caption = "Hide Progress Bar"
Else
  mnu1.Caption = "Show Progress Bar"
End If

If Player1.ShowControls = True Then
  mnu2.Caption = "Hide Controls Bar"
Else
  mnu2.Caption = "Show Controls Bar"
End If

If Player1.ShowStatus = True Then
  mnu3.Caption = "Hide Status Bar"
Else
  mnu3.Caption = "Show Status Bar"
End If

If Player1.ShowVolume = True Then
  mnu4.Caption = "Hide Volume Bar"
Else
  mnu4.Caption = "Show Volume Bar"
End If

End Sub
