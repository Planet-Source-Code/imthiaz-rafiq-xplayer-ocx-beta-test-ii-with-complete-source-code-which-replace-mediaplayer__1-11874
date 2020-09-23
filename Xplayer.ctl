VERSION 5.00
Begin VB.UserControl PlayerX 
   BackColor       =   &H00000000&
   ClientHeight    =   3750
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5835
   ScaleHeight     =   3750
   ScaleWidth      =   5835
   ToolboxBitmap   =   "Xplayer.ctx":0000
   Begin VB.Timer Readstring 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   420
      Top             =   300
   End
   Begin VB.PictureBox FrameVideo 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1740
      Left            =   420
      ScaleHeight     =   1740
      ScaleWidth      =   2175
      TabIndex        =   7
      Top             =   210
      Width           =   2175
   End
   Begin VB.PictureBox Control 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1035
      Left            =   -120
      ScaleHeight     =   1035
      ScaleWidth      =   5985
      TabIndex        =   0
      Top             =   2340
      Width           =   5985
      Begin VB.PictureBox Tool 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H00404040&
         Height          =   210
         Index           =   2
         Left            =   60
         ScaleHeight     =   210
         ScaleWidth      =   5040
         TabIndex        =   8
         Top             =   420
         Width           =   5040
         Begin XPlayer.ProgYbar Volume 
            Height          =   105
            Left            =   30
            TabIndex        =   9
            ToolTipText     =   "Volume"
            Top             =   30
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   185
            ForeColor       =   4210752
            BackColor       =   8421504
            Max             =   100
            Mode            =   0
            Border          =   1
            Mark            =   0   'False
            MarkThicness    =   3
            MarkColor       =   65535
         End
         Begin XPlayer.ProgYbar Pan 
            Height          =   105
            Left            =   1650
            TabIndex        =   10
            ToolTipText     =   "Balance"
            Top             =   30
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   185
            ForeColor       =   4210752
            BackColor       =   8421504
            Max             =   100
            Mode            =   0
            Border          =   1
            Mark            =   0   'False
            MarkThicness    =   3
            MarkColor       =   65535
         End
      End
      Begin VB.PictureBox Tool 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   165
         Index           =   0
         Left            =   60
         ScaleHeight     =   165
         ScaleWidth      =   4065
         TabIndex        =   5
         Top             =   0
         Width           =   4065
         Begin XPlayer.ProgYbar PosBar 
            Height          =   105
            Left            =   30
            TabIndex        =   6
            Top             =   30
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   185
            ForeColor       =   8421504
            BackColor       =   4210752
            Max             =   100
            Mode            =   0
            Border          =   1
            Mark            =   0   'False
            MarkThicness    =   3
            MarkColor       =   65535
         End
      End
      Begin VB.PictureBox Tool 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   3
         Left            =   60
         ScaleHeight     =   225
         ScaleWidth      =   3885
         TabIndex        =   2
         Top             =   600
         Width           =   3885
         Begin VB.Label Status 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Ready.."
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   810
            TabIndex        =   3
            Top             =   0
            Width           =   510
         End
         Begin VB.Label statusc 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Status -"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   30
            TabIndex        =   4
            Top             =   0
            Width           =   570
         End
      End
      Begin VB.PictureBox Tool 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   60
         ScaleHeight     =   255
         ScaleWidth      =   3825
         TabIndex        =   1
         Top             =   150
         Width           =   3825
         Begin VB.Image Open 
            Height          =   255
            Left            =   2760
            Picture         =   "Xplayer.ctx":0312
            Top             =   0
            Width           =   555
         End
         Begin VB.Image Pause 
            Height          =   255
            Left            =   1110
            Picture         =   "Xplayer.ctx":09AB
            Top             =   0
            Width           =   555
         End
         Begin VB.Image Ffwd 
            Height          =   255
            Left            =   2220
            Picture         =   "Xplayer.ctx":1044
            Top             =   0
            Width           =   555
         End
         Begin VB.Image Play 
            Height          =   255
            Left            =   555
            Picture         =   "Xplayer.ctx":171A
            Top             =   0
            Width           =   555
         End
         Begin VB.Image Rrwd 
            Height          =   255
            Left            =   0
            Picture         =   "Xplayer.ctx":1DCD
            Top             =   0
            Width           =   555
         End
         Begin VB.Image Stop 
            Height          =   255
            Left            =   1665
            Picture         =   "Xplayer.ctx":24B4
            Top             =   0
            Width           =   555
         End
      End
   End
   Begin VB.Timer TimerAtEndFile 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   840
      Top             =   3270
   End
   Begin VB.Timer TimerMisc 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   330
      Top             =   3300
   End
End
Attribute VB_Name = "PlayerX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Dim Child As Double

Dim OriGinalX As Long
Dim OriGinalY As Long

Dim FileName As String

Dim LbActualCx As Long
Dim LbActualCy As Long
Dim LbFramesPerSecond As Long
Dim LbTotalFrames As Long
Dim LbTotalTime As Long
Dim LbCurrPos As Long

Dim OpenedSucess As Boolean
Dim Autostart As Boolean

Dim CurVolMainPer As Integer
Dim CurVolBalPer As Integer
Dim VolMuteState As Boolean

Dim ToolVis(3) As Boolean
Dim FullScreen As Boolean

Dim FFby As Integer
Dim FBBy As Integer

Public Enum MovSize
  Original_Size
  Resize_to_Fit_Current_Control
End Enum

Dim MSize As MovSize

Dim Reapet As Boolean

Private CdOC As New clsCDTray



Public Event OnTimer(PlayerPlayMode As String, PlayerCurrentTime As Long, PlayerOldTime As Long, PlayerTotalTime As Long)

Public Event OnPlayerError(PlayerError As String, PlayerErrorNo As Integer)

Public Event OnSelfResize(NewHeight As Long, NewWidth As Long)

Public Event OnPlayFinish()

Public Event OnMouseMove(Button As Integer, X As Single, Y As Single)



Sub FullX()
MsgBox "Fulscreen Is still in Beta testing", vbInformation, "Player.ocx Fullscreen Option"
End Sub
Function StartFullScreen() As String
Dim X As VbMsgBoxResult

  Dim cx As Double
  Dim Cy As Double
  
  If FrameVideo.Enabled = False Then
    StartFullScreen = "Error No video Found "
    Exit Function
  End If
  If OpenedSucess = False Then
    StartFullScreen = "File not loaded..":
    Exit Function
  End If

  If GetSize(AliasName, "cx") > 250 And GetSize(AliasName, "cy") > 200 Then
    LbActualCx = GetSize(AliasName, "cx") * (((Screen.Width / Screen.TwipsPerPixelX) - 5) / GetSize(AliasName, "cx"))
    LbActualCy = GetSize(AliasName, "cy") * (((Screen.Height / Screen.TwipsPerPixelY) - 5) / GetSize(AliasName, "cy"))
  Else
    X = MsgBox("Full Screen May hang the system." + vbCrLf + "If you are running a uncompress movie file.." + vbCrLf + " A Small Bug", vbInformation + vbRetryCancel, "Please correct this bug and mail me.")
    
    If X <> vbRetry Then
      StartFullScreen = "Error exit by user."
      Exit Function
    End If
    LbActualCx = GetSize(AliasName, "cx") * (((Screen.Width / Screen.TwipsPerPixelX) - 5) / GetSize(AliasName, "cx"))
    LbActualCy = GetSize(AliasName, "cy") * (((Screen.Height / Screen.TwipsPerPixelY) - 5) / GetSize(AliasName, "cy"))
  End If
  
  FullScreen = True
  Load frmfullscreen
  frmfullscreen.Screener.Top = 0 ' (Screen.Height - frmfullscreen.Screener.Height) / 2
  frmfullscreen.Screener.Left = 0 ' (Screen.Width - frmfullscreen.Screener.Width) / 2
  frmfullscreen.Screener.Width = LbActualCx * Screen.TwipsPerPixelX
  frmfullscreen.Screener.Height = LbActualCy * Screen.TwipsPerPixelY
  frmfullscreen.Screener.Top = (Screen.Height - frmfullscreen.Screener.Height) / 2
  frmfullscreen.Screener.Left = (Screen.Width - frmfullscreen.Screener.Width) / 2

LbCurrPos = GetCurrentMultimediaPos(AliasName)

CloseMultimedia (AliasName)
Result = OpenMultimedia(frmfullscreen.Screener.hwnd, AliasName, FileName, typeDevice)      'call now function OpenMultimedia
Result = PutMultimedia(frmfullscreen.Screener.hwnd, AliasName, 0, 0, 0, 0)           'call now function PutMultimedia

Result = MoveMultimedia(AliasName, LbCurrPos)      'call now function MoveMultimedia
Readstring.Enabled = True


frmfullscreen.Show vbModal


LbCurrPos = GetCurrentMultimediaPos(AliasName)

CloseMultimedia (AliasName)

Result = OpenMultimedia(FrameVideo.hwnd, AliasName, FileName, typeDevice)      'call now function OpenMultimedia

Result = PutMultimedia(FrameVideo.hwnd, AliasName, 0, 0, 0, 0)         'call now function PutMultimedia

Result = MoveMultimedia(AliasName, LbCurrPos)      'call now function MoveMultimedia

If Started = True And PlayMode = "Pause" Then
  PlayMode = "Play"
  Pause_Click
ElseIf Started = True And PlayMode = "Play" Then
  PlayMode = "Pause"
  Pause_Click
Else
  PlayMode = "Play"
  Started = True
  Stop_Click
End If


FullScreen = False
End Function

Sub IncreaseVolumeByOnePercent()
  If CurVolMainPer - 1 <> 0 Then
    CurVolMainPer = CurVolMainPer + 1
    AdjustOutput PercentF(CurVolMainPer), CurVolBalPer
  End If
End Sub

Sub DecreaseVolumeByOnePercent()
  If CurVolMainPer + 1 <> 0 Then
    CurVolMainPer = CurVolMainPer + 1
    AdjustOutput PercentF(CurVolMainPer), CurVolBalPer
  End If
End Sub

Private Sub Ffwd_Click()
Dim pos As Long

If Started = True And PlayMode = "Play" Then
  
  LbCurrPos = GetCurrentMultimediaPos(AliasName)
  
  If (LbTotalTime * 1000) > LbCurrPos + (600 * FFby) Then
    Result = MoveMultimedia(AliasName, LbCurrPos + (600 * FFby))    'call now function MoveMultimedia
    If Result = "Sucess" Then
      PosBar.DrawBar LbCurrPos / LbFramesPerSecond
      Status.Caption = "Fast farward by 10 sec"
    End If
  End If
Else
  Status.Caption = "Idle"
End If

End Sub

Private Sub FrameVideo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  RaiseEvent OnMouseMove(Button, X, Y)

End Sub

Private Sub FrameVideo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  RaiseEvent OnMouseMove(Button, X, Y)

End Sub

Private Sub Open_Click()
X = StartFullScreen()
End Sub

Private Sub Pan_Click(Value As Double)
Dim X As Integer
X = Value
AdjustOutput PercentF(CurVolMainPer), CallBal(X)
PropertyChanged "VolumePercent"
End Sub

Private Sub Readstring_Timer()
Select Case CurString
  Case "Pause": Pause_Click: CurString = "s": frmfullscreen.EraseString.Enabled = True
  Case "Play": Play_Click: CurString = "s": frmfullscreen.EraseString.Enabled = True
  Case "Stop": Stop_Click: CurString = "s": frmfullscreen.EraseString.Enabled = True
  Case "FF": Ffwd_Click: CurString = "s": frmfullscreen.EraseString.Enabled = True
  Case "RR": Rrwd_Click: CurString = "s": frmfullscreen.EraseString.Enabled = True
  Case Else
    If Mid$(CurString, 1, 1) = "V" Then
      CurVolMainPer = 100 - Val(Mid$(CurString, 3, 1)) * 10
      AdjustOutput PercentF(CurVolMainPer), CallBal(CurVolBalPer)
      PropertyChanged "VolumePercent"
    End If
End Select
End Sub

Private Sub Rrwd_Click()
Dim pos As Long

If Started = True And PlayMode = "Play" Then
  
  LbCurrPos = GetCurrentMultimediaPos(AliasName)
  
  If (LbTotalTime * 1000) > LbCurrPos - (600 * FBBy) Then
    Result = MoveMultimedia(AliasName, LbCurrPos - (600 * FBBy))
    If Result = "Sucess" Then
      PosBar.DrawBar LbCurrPos / LbFramesPerSecond
      Status.Caption = "Fast Rewind by 10 sec"
    End If
  End If
Else
  Status.Caption = "Idle"
End If
End Sub

Private Sub Status_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  RaiseEvent OnMouseMove(Button, X, Y)
End Sub

Private Sub Status_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  RaiseEvent OnMouseMove(Button, X, Y)
End Sub

Private Sub Pause_Click()
If OpenedSucess = False Then Status.Caption = "File not loaded..": Exit Sub
If Started = True And PlayMode = "Play" Then
  
  Result = PauseMultimedia(AliasName)
  Status.Caption = "Paused.."
  If Result = "Success" Then
    TimerAtEndFile.Enabled = False
    TimerMisc.Enabled = False
    PlayMode = "Pause"
  End If
ElseIf Started = True And PlayMode = "Pause" Then
  Result = ResumeMultimedia(AliasName)
  Status.Caption = "Playing.."
  If Result = "Success" Then
    TimerAtEndFile.Enabled = True
    TimerMisc.Enabled = True
    PlayMode = "Play"
  End If
Else
  Status.Caption = "Play the file..."
End If
End Sub

Private Sub Play_Click()
If OpenedSucess = False Then Status.Caption = "File not loaded..": Exit Sub
If Started = False And PlayMode = "" Then
  MovieSize
  Result = PlayMultimedia(AliasName, "", "")
  If Result = "Success" Then
    Status.Caption = "Playing.."
    PosBar.DrawBar 0
    TimerAtEndFile.Enabled = True
    TimerMisc.Enabled = True
    PlayMode = "Play"
    Started = True
  End If
ElseIf Started = True And PlayMode = "Pause" Then
  Result = ResumeMultimedia(AliasName)
  Status.Caption = "Playing.."
  If Result = "Success" Then
    TimerAtEndFile.Enabled = True
    TimerMisc.Enabled = True
    PlayMode = "Play"
  End If
End If
End Sub

Private Sub PosBar_Click(Value As Double)
Dim SValue As Long
If Value <> -1 Then
  SValue = Value
  If OpenedSucess = False Then PosBar.DrawBar 0: Status.Caption = "File not loaded..": Exit Sub
  
  If LbFramesPerSecond = 0 And SValue = -1 Then Exit Sub 'if this alias not opened then exit (improtant)
  
  
  Dim pos As Long
  
  pos = SValue * LbFramesPerSecond
  
  If Started = False And PlayMode = "" Then
    Result = MoveMultimedia(AliasName, pos)      'call now function MoveMultimedia
    Result = PauseMultimedia(AliasName)
    Started = True
    PlayMode = "Pause"
  ElseIf Started = True And PlayMode = "Play" Then
    Result = MoveMultimedia(AliasName, pos)
  ElseIf Started = True And PlayMode = "Pause" Then
    Result = MoveMultimedia(AliasName, pos)      'call now function MoveMultimedia
    Result = PauseMultimedia(AliasName)
  End If
  
  If Result = "Success" Then 'this mean MoveMultimedia success
    Status.Caption = "Moved to " + CalTime(SValue) + " of " + CalTime(LbTotalTime)
  Else 'not success
    RaiseEvent OnPlayerError(Result, ErrorNo(Result))
    PosBar.DrawBar 0
  End If
End If
End Sub



Private Sub PosBar_ValueChange(NewVal As Double, Oldval As Double)
  Dim XX As Long
  Dim xr As Long
  Dim xg As Long
  Dim xc As Long
  XX = NewVal
  xr = Oldval
  xg = PosBar.Max
  
  RaiseEvent OnTimer(PlayMode, XX, xr, xg)
End Sub

Private Sub Stop_Click()
StopPlay
End Sub
Private Sub TimerMisc_Timer()

Dim Percent As Double

If Started = True Then
  LbCurrPos = GetCurrentMultimediaPos(AliasName)
  Percent = GetPercent(AliasName)
  PosBar.DrawBar LbCurrPos / LbFramesPerSecond
  Status.Caption = "Playing time -> " + CalTime(LbCurrPos / LbFramesPerSecond) + " of " + CalTime(LbTotalTime) + " <-"
Else
  Status.Caption = "Idle"
End If
Simple = VolAvail()
CurVolMainPer = PercentE(MixerState(1).MxrVol)
ControlVolume
End Sub

Private Sub TimerAtEndFile_Timer()

 'this is the main improtant point to select the file which you want change position for it

If AreMultimediaAtEnd(AliasName, LbTotalFrames) = True Then ' alias name for e.g.:"movie"
  PlayMode = ""
  Started = False
  PosBar.DrawBar PosBar.Max
  TimerMisc.Enabled = False
  TimerAtEndFile.Enabled = False
  RaiseEvent OnPlayFinish
  If Reapet = True Then
    Play_Click
  End If
End If


End Sub



Private Sub Tool_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  RaiseEvent OnMouseMove(Button, X, Y)
End Sub

Private Sub Tool_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  RaiseEvent OnMouseMove(Button, X, Y)

End Sub

Private Sub UserControl_DragDrop(Source As Control, X As Single, Y As Single)
  MsgBox Control.Font
End Sub

Private Sub UserControl_Initialize()
'Init mixer
Dim Zx As Double
Simple = VolAvail()

If Not GetDefaultDevice("MPEGVideo") = "mciqtz.drv" Then
  SetDefaultDevice "MPEGVideo", "mciqtz.drv"
End If

If Not GetDefaultDevice("sequencer") = "mciseq.drv" Then
  SetDefaultDevice "sequencer", "mciseq.drv"
End If

If Not GetDefaultDevice("avivideo") = "mciavi.drv" Then
  SetDefaultDevice "avivideo", "mciavi.drv"
End If

CurVolMainPer = PercentE(MixerState(1).MxrVol)

VolMuteState = MixerState(1).MxrMute

Pan.DrawBar 50
Zx = CurVolMainPer
Volume.DrawBar 100 - Zx

CurVolBalPer = 50

AdjustOutput PercentF(CurVolMainPer), CallBal(CurVolBalPer)
PropertyChanged "VolumePercent"
Started = True
PlayMode = ""
MSize = Original_Size
For i = 0 To 3
  ToolVis(i) = True
Next i
FullScreen = False
Reapet = False
FFby = 1
FBBy = 1
End Sub


Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  RaiseEvent OnMouseMove(Button, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  RaiseEvent OnMouseMove(Button, X, Y)
End Sub
Sub ControlSize()
Dim Xas As Long

    Control.Left = 0
    Xas = 0
    Control.Width = UserControl.Width
    Control.Height = 0
    For i = 0 To 3
      Tool(i).Left = 0
      Tool(i).Width = Control.Width
      Tool(i).Top = Xas
      If ToolVis(i) = True Then
        Tool(i).Visible = True
        Control.Height = Control.Height + Tool(i).Height + 20
        Xas = Xas + Tool(i).Height
      Else
        Tool(i).Visible = False
        Control.Height = Control.Height + Tool(i).Height - 20
      End If
    Next i
    Xas = 0
    For i = 0 To 3
      If Tool(i).Visible = True Then
        Xas = Xas + Tool(i).Height
      End If
    Next i
    Control.Height = Xas
    Control.Height = Control.Height + 100
    Volume.Left = 20
    Volume.Width = (Tool(2).Width / 2) - 60
    Pan.Left = Volume.Left + Volume.Width + 30
    Pan.Width = Volume.Width
    PosBar.Left = 20
    PosBar.Width = Tool(0).Width - 70
    Control.Top = UserControl.Height - Control.Top
End Sub

Sub ControlVolume()
Dim Zx As Double
  
  Zx = CurVolMainPer
  
  Volume.DrawBar 100 - Zx
  
  CurVolBalPer = 50
  
  Zx = CurVolBalPer
    
  Pan.DrawBar Zx
  
  AdjustOutput PercentF(CurVolMainPer), CallBal(CurVolBalPer)
  
End Sub

Private Sub UserControl_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox Data
End Sub

Private Sub UserControl_Resize()
ControlSize
ControlVolume
MovieSize
If FrameVideo.Enabled = False Then
  Control.Top = 0
  FrameVideo.Top = UserControl.Height + 1000
  UserControl.Width = Control.Width
  UserControl.Height = Control.Height
  RaiseEvent OnSelfResize(UserControl.Height, UserControl.Width)
  Exit Sub
End If
If MSize = Original_Size Then
  If UserControl.Height > Control.Height + FrameVideo.Height + 100 Then
    Control.Top = UserControl.Height - Control.Height
    FrameVideo.Top = ((UserControl.Height - Control.Height) - FrameVideo.Height) / 2
  Else
    UserControl.Height = Control.Height + FrameVideo.Height + 101
    Control.Top = UserControl.Height - Control.Height
    FrameVideo.Top = ((UserControl.Height - Control.Height) - FrameVideo.Height) / 2
    RaiseEvent OnSelfResize(UserControl.Height, UserControl.Width)
  End If
  If UserControl.Width > FrameVideo.Width Then
    FrameVideo.Left = (UserControl.Width - FrameVideo.Width) / 2
  Else
    UserControl.Width = FrameVideo.Width
    FrameVideo.Left = (UserControl.Width - FrameVideo.Width) / 2
    RaiseEvent OnSelfResize(UserControl.Height, UserControl.Width)
  End If
  ControlVolume
  MovieSize
End If
If MSize = Resize_to_Fit_Current_Control Then
  If UserControl.Width > 3500 Then
    FrameVideo.Left = 0
    FrameVideo.Width = UserControl.Width
  Else
    UserControl.Width = 3501
    FrameVideo.Left = 0
    FrameVideo.Width = UserControl.Width
    RaiseEvent OnSelfResize(UserControl.Height, UserControl.Width)
  End If
  If UserControl.Height > 2000 + Control.Height Then
    FrameVideo.Top = 0
    FrameVideo.Height = UserControl.Height - Control.Height
  Else
    UserControl.Height = 2001 + Control.Height
    FrameVideo.Height = UserControl.Height - Control.Height
    FrameVideo.Top = 0
    RaiseEvent OnSelfResize(UserControl.Height, UserControl.Width)
  End If
    ControlVolume
    MovieSize
    Control.Top = FrameVideo.Height
End If
End Sub

Private Sub UserControl_Show()
Dim Zx As Double
CurVolMainPer = PercentE(MixerState(1).MxrVol)

VolMuteState = MixerState(1).MxrMute

CurVolBalPer = 50
AdjustOutput PercentF(CurVolMainPer), CallBal(CurVolBalPer)
PropertyChanged "VolumePercent"


Pan.DrawBar 50
Zx = CurVolMainPer
Volume.DrawBar 100 - Zx
CurVolBalPer = 0
UserControl_Resize

End Sub

Private Sub UserControl_Terminate()

TimerMisc.Enabled = False
TimerAtEndFile.Enabled = False

mixerClose hMixer


DoEvents
If OpenedSucess = True Then
  
  AliasName = "movie" & Child
  
  Result = CloseMultimedia(AliasName)
  
End If

If Result = "Success" Then 'this mean CloseAll success
'Write your commands here
Else 'not success
  RaiseEvent OnPlayerError(Result, ErrorNo(Result))
  Result = CloseMultimedia(AliasName)
End If

End Sub

Sub CloseAllPlayers()
  Result = CloseAll
  RaiseEvent OnPlayerError(Result, ErrorNo(Result))
End Sub

Sub ClosePlayers()
TimerMisc.Enabled = False
TimerAtEndFile.Enabled = False

mixerClose hMixer


DoEvents
If OpenedSucess = True Then
  
  AliasName = "movie" & Child
  
  Result = CloseMultimedia(AliasName)
  
End If

If Result = "Success" Then 'this mean CloseAll success
'Write your commands here
Else 'not success
  RaiseEvent OnPlayerError(Result, ErrorNo(Result))
  Result = CloseMultimedia(AliasName)
End If
End Sub


Sub StopPlay()
If Started = True Then
  Started = False
  PlayMode = ""
  TimerMisc.Enabled = False
  Result = StopMultimedia(AliasName)
  PosBar_Click 0
  PosBar.DrawBar 0
  Status.Caption = "Stopped"
End If
End Sub


Sub Play()
If PlayMode = "" And Started = False Then
  Play_Click
End If
End Sub
Sub Pause()
If Started = True And PlayMode = "Play" Then
  Pause_Click
End If
End Sub
Sub ResumePlay()
If Started = True And PlayMode = "Pause" Then
  Pause_Click
End If
End Sub
Sub OpenCdBay()
  CdOC.InitCD
  CdOC.OpenCDTray
End Sub
Sub CloseCdBay()
  CdOC.InitCD
  CdOC.CloseCDTray
End Sub
Sub ToggleVolume()
  If ToolVis(2) = True Then
    ToolVis(2) = False
  Else
    ToolVis(2) = True
  End If
  UserControl_Resize
End Sub
Sub TogglePosBar()
  If ToolVis(0) = True Then
    ToolVis(0) = False
  Else
    ToolVis(0) = True
  End If
  UserControl_Resize
End Sub
Sub ToggleStatus()
  If ToolVis(3) = True Then
    ToolVis(3) = False
  Else
    ToolVis(3) = True
  End If
  UserControl_Resize
End Sub
Sub ToggleControls()
  If ToolVis(1) = True Then
    ToolVis(1) = False
  Else
    ToolVis(1) = True
  End If
  UserControl_Resize
End Sub


Sub MovieSize()
Attribute MovieSize.VB_MemberFlags = "40"
If FullScreen = True Then
  Result = PutMultimedia(frmfullscreen.Screener.hwnd, AliasName, 0, 0, 0, 0)           'call now function PutMultimedia
Else
  Result = PutMultimedia(FrameVideo.hwnd, AliasName, 0, 0, 0, 0)           'call now function PutMultimedia
End If
End Sub

Function CalTime(Timerx As Long) As String
Attribute CalTime.VB_UserMemId = 0
Attribute CalTime.VB_MemberFlags = "40"
On Error Resume Next
Dim X As Long
Dim Y As Long


X = Int(Timerx / 60)

Y = Int(Timerx - (X * 60))

CalTime = LTrim(RTrim(Str(X))) + ":" + LTrim(RTrim(Str(Y)))

End Function

Sub Loadfile()
Attribute Loadfile.VB_MemberFlags = "40"

Started = False
PlayMode = ""

If FileName = "" Then
  OpenedSucess = False
  Status.Caption = ""
  Started = False
  PlayMode = ""
  Exit Sub
End If


Status.Caption = "Loading.."

DoEvents

If Right(FileName, 4) = ".avi" Then
    typeDevice = "AviVideo"
ElseIf Right(FileName, 4) = ".rmi" Or Right(FileName, 4) = ".mid" Then
    typeDevice = "sequencer"
Else
    typeDevice = "MPEGVideo"
End If

OpenedSucess = False

Again:

AliasName = "movie" & Child


Result = OpenMultimedia(FrameVideo.hwnd, AliasName, FileName, typeDevice)      'call now function OpenMultimedia


If ErrorNo(Result) = 0 Then 'this mean OpenMultimedia success
  Dim cx As Double
  Dim Cy As Double

  If MSize = Original_Size Then
    LbActualCx = GetSize(AliasName, "cx")
    LbActualCy = GetSize(AliasName, "cy")
    OriGinalX = LbActualCx
    OriGinalY = LbActualCy
    
    If LbActualCx = -1 Then
      FrameVideo.Width = 0
    Else
      FrameVideo.Width = LbActualCx * Screen.TwipsPerPixelX
    End If
    
    If LbActualCy = -1 Then
      FrameVideo.Height = 0
    Else
      FrameVideo.Height = LbActualCy * Screen.TwipsPerPixelY
    End If
    
    If LbActualCx <> -1 And LbActualCy <> -1 Then
      FrameVideo.Enabled = True
    Else
      FrameVideo.Enabled = False
    End If
  End If
  
  If MSize = Resize_to_Fit_Current_Control Then
    UserControl_Resize
    LbActualCx = GetSize(AliasName, "cx")
    LbActualCy = GetSize(AliasName, "cy")
    If LbActualCx = -1 Then
      FrameVideo.Width = 0
    End If
    If LbActualCy = -1 Then
      FrameVideo.Height = 0
    End If
    If LbActualCx <> -1 And LbActualCy <> -1 Then
      FrameVideo.Enabled = True
    Else
      FrameVideo.Enabled = False
    End If
  End If
  
  PlayMode = ""
  Started = False
  TimerMisc.Enabled = False
  TimerAtEndFile.Enabled = False
  LbFramesPerSecond = GetFramesPerSecond(AliasName)
  LbTotalFrames = GetTotalframes(AliasName)  'Get total frames
  LbTotalTime = GetTotalTimeByMS(AliasName) / 1000   'Get Total Time
  PosBar.Max = LbTotalFrames / LbFramesPerSecond
  Status.Caption = "Loaded Sucessfully"
  OpenedSucess = True
  UserControl_Resize
  If Autostart = True Then Play_Click
Else
  Select Case ErrorNo(Result)
    Case 277
      Status.Caption = "Cannot Initialize Sound.."
    Case 263
      Status.Caption = "Invalid Media File.."
    Case 289
      Result = CloseMultimedia(AliasName)
      GoTo Again
    Case Else
      Status.Caption = "Unknown Error.."
      RaiseEvent OnPlayerError(Result, ErrorNo(Result))
    End Select
  OpenedSucess = False
End If

End Sub
Function ErrorNo(Errstring As String) As Integer
  ErrorNo = Val(Mid$(Errstring, 9, 3))
End Function


'____________________________________________


Public Property Get File() As String
     File = FileName
End Property

Public Property Let File(ByVal N_File As String)
    FileName = N_File
    Loadfile
    PropertyChanged "File"
End Property

Public Property Get CurrentPosition() As Long
Attribute CurrentPosition.VB_MemberFlags = "400"
     CurrentPosition = LbCurrPos
End Property

Public Property Let CurrentPosition(ByVal Npos As Long)
If Npos <> 0 Then
    LbCurrPos = Npos
    PosBar_Click (Npos / LbFramesPerSecond)
    PropertyChanged "CurrentPosition"
End If
End Property


Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("File", File, "")
    Call PropBag.WriteProperty("CurrentPosition", LbCurrPos, "")
    Call PropBag.WriteProperty("ChildNo", Child, Rnd * 1000)
    Call PropBag.WriteProperty("PlayerAutoStart", Autostart, False)
    Call PropBag.WriteProperty("MovieSizer", MSize, 0)
    Call PropBag.WriteProperty("ShowPosBar", ToolVis(0), True)
    Call PropBag.WriteProperty("ShowControls", ToolVis(1), True)
    Call PropBag.WriteProperty("ShowVolume", ToolVis(2), True)
    Call PropBag.WriteProperty("ShowStatus", ToolVis(3), True)
    Call PropBag.WriteProperty("ColorBody", UserControl.BackColor, vbBlack)
    Call PropBag.WriteProperty("ColorVideoWindow", FrameVideo.BackColor, vbBlack)
    Call PropBag.WriteProperty("ColorControlBack", Control.BackColor, vbBlack)
    Call PropBag.WriteProperty("ColorStatusText", Status.ForeColor, vbBlack)
    Call PropBag.WriteProperty("ColorPositionBarForeColor", PosBar.ForeColor, &H404040)
    Call PropBag.WriteProperty("ColorPositionBarBackColor", PosBar.BackColor, &H808080)
    Call PropBag.WriteProperty("ColorVolumeBarForeColor", Volume.ForeColor, &H404040)
    Call PropBag.WriteProperty("ColorVolumeBarBackColor", Volume.BackColor, &H808080)
    Call PropBag.WriteProperty("Repeat", Reapet, False)
    Call PropBag.WriteProperty("FastForwardBySec", FFby, 1)
    Call PropBag.WriteProperty("FastRewindBySec", FBBy, 1)
    
End Sub


Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    File = PropBag.ReadProperty("File", "")
    CurrentPosition = PropBag.ReadProperty("CurrentPosition", 0)
    ChildNo = PropBag.ReadProperty("ChildNo", Rnd * 1000)
    PlayerAutoStart = PropBag.ReadProperty("PlayerAutoStart", False)
    MovieSizer = PropBag.ReadProperty("MovieSizer", 0)
    ShowPosBar = PropBag.ReadProperty("ShowPosBar", True)
    ShowControls = PropBag.ReadProperty("ShowControls", True)
    ShowVolume = PropBag.ReadProperty("ShowVolume", True)
    ShowStatus = PropBag.ReadProperty("ShowStatus", True)
    ColorBody = PropBag.ReadProperty("ColorBody", vbBlack)
    ColorVideoWindow = PropBag.ReadProperty("ColorVideoWindow", vbBlack)
    ColorControlBack = PropBag.ReadProperty("ColorControlBack", vbBlack)
    ColorStatusText = PropBag.ReadProperty("ColorStatusText", vbWhite)
    ColorPositionBarForeColor = PropBag.ReadProperty("ColorPositionBarForeColor", &H404040)
    ColorPositionBarBackColor = PropBag.ReadProperty("ColorPositionBarBackColor", &H808080)
    ColorVolumeBarForeColor = PropBag.ReadProperty("ColorVolumeBarForeColor", &H404040)
    ColorVolumeBarBackColor = PropBag.ReadProperty("ColorVolumeBarBackColor", &H808080)
    Repeat = PropBag.ReadProperty("Repeat", False)
    FastForwardBySec = PropBag.ReadProperty("FastForwardBySec", 1)
    FastRewindBySec = PropBag.ReadProperty("FastRewindBySec", 1)
End Sub
Public Property Get FastForwardBySec() As Integer
    FastForwardBySec = FFby
End Property

Public Property Let FastForwardBySec(ByVal NP As Integer)
    FFby = NP
    PropertyChanged "FastForwardBySec"
End Property

Public Property Get FastRewindBySec() As Integer
    FastRewindBySec = FBBy
End Property

Public Property Let FastRewindBySec(ByVal NP As Integer)
    FBBy = NP
    PropertyChanged "FastRewindBySec"
End Property


Public Property Get Repeat() As Boolean
    Repeat = Reapet
End Property

Public Property Let Repeat(ByVal NP As Boolean)
    Reapet = NP
    PropertyChanged "Repeat"
End Property


Public Property Get ColorVolumeBarForeColor() As OLE_COLOR
    ColorVolumeBarForeColor = Volume.ForeColor
End Property

Public Property Let ColorVolumeBarForeColor(ByVal NP As OLE_COLOR)
    Volume.ForeColor = NP
    Pan.ForeColor = NP
    ControlVolume
    PropertyChanged "ColorVolumeBarForeColor"
End Property

Public Property Get ColorVolumeBarBackColor() As OLE_COLOR
    ColorVolumeBarBackColor = Volume.BackColor
End Property

Public Property Let ColorVolumeBarBackColor(ByVal NP As OLE_COLOR)
    Volume.BackColor = NP
    Pan.BackColor = NP
    ControlVolume
    PropertyChanged "ColorVolumeBarBackColor"
End Property

Public Property Get ColorPositionBarForeColor() As OLE_COLOR
    ColorPositionBarForeColor = PosBar.ForeColor
End Property

Public Property Let ColorPositionBarForeColor(ByVal NP As OLE_COLOR)
    PosBar.ForeColor = NP
    PropertyChanged "ColorPositionBarForeColor"
End Property

Public Property Get ColorPositionBarBackColor() As OLE_COLOR
    ColorPositionBarBackColor = PosBar.BackColor
End Property

Public Property Let ColorPositionBarBackColor(ByVal NP As OLE_COLOR)
    PosBar.BackColor = NP
    PropertyChanged "ColorPositionBarBackColor"
End Property


Public Property Get ColorStatusText() As OLE_COLOR
     ColorStatusText = Status.ForeColor
End Property

Public Property Let ColorStatusText(ByVal NP As OLE_COLOR)
    Status.ForeColor = NP
    statusc.ForeColor = NP
    PropertyChanged "ColorStatusText"
End Property

Public Property Get ColorControlBack() As OLE_COLOR
     ColorControlBack = Control.BackColor
End Property

Public Property Let ColorControlBack(ByVal NP As OLE_COLOR)
    Control.BackColor = NP
    For i = 0 To 3
      Tool(i).BackColor = NP
    Next i
    PropertyChanged "ColorControlBack"
End Property


Public Property Get ColorBody() As OLE_COLOR
     ColorBody = UserControl.BackColor
End Property

Public Property Let ColorBody(ByVal NP As OLE_COLOR)
    UserControl.BackColor = NP
    PropertyChanged "ColorBody"
End Property

Public Property Get ColorVideoWindow() As OLE_COLOR
     ColorVideoWindow = FrameVideo.BackColor
End Property

Public Property Let ColorVideoWindow(ByVal NP As OLE_COLOR)
    FrameVideo.BackColor = NP
    PropertyChanged "ColorVideoWindow"
End Property


Public Property Get ShowPosBar() As Boolean
     ShowPosBar = ToolVis(0)
End Property

Public Property Let ShowPosBar(ByVal NP As Boolean)
    ToolVis(0) = NP
    PropertyChanged "ShowPosBar"
    UserControl_Resize
End Property
Public Property Get ShowControls() As Boolean
     ShowControls = ToolVis(1)
End Property
Public Property Let ShowControls(ByVal NP As Boolean)
    ToolVis(1) = NP
    PropertyChanged "PlayerAutoStart"
    UserControl_Resize
End Property
Public Property Get ShowVolume() As Boolean
     ShowVolume = ToolVis(2)
End Property
Public Property Let ShowVolume(ByVal NP As Boolean)
    ToolVis(2) = NP
    PropertyChanged "ShowVolume"
    UserControl_Resize
End Property

Public Property Get ShowStatus() As Boolean
     ShowStatus = ToolVis(3)
End Property
Public Property Let ShowStatus(ByVal NP As Boolean)
    ToolVis(3) = NP
    PropertyChanged "ShowStatus"
    UserControl_Resize
End Property

Public Property Get VolumePercent() As Integer
     VolumePercent = CurVolMainPer
End Property

Public Property Let VolumePercent(ByVal NCurVolMainPer As Integer)
If NCurVolMainPer > -1 And NCurVolMainPer < 101 Then
    CurVolMainPer = 100 - NCurVolMainPer
    AdjustOutput PercentF(CurVolMainPer), CallBal(CurVolBalPer)
    PropertyChanged "VolumePercent"
End If
End Property


Public Property Get VolumeMute() As Boolean
     VolumeMute = VolMuteState
End Property

Public Property Let VolumeMute(ByVal NVolumeMute As Boolean)
    VolMuteState = NVolumeMute
    MuteMe VolMuteState
    PropertyChanged "VolumeMute"
End Property

Public Property Get BalancePercent() As Integer
     BalancePercent = CurVolBalPer
End Property

Public Property Let BalancePercent(ByVal NBalancePercent As Integer)
If NBalancePercent > -1 And NBalancePercent < 101 Then
    CurVolBalPer = NBalancePercent
    AdjustOutput PercentF(CurVolMainPer), CallBal(CurVolBalPer)
    PropertyChanged "BalancePercent"
End If
End Property


Public Property Get ChildNo() As Double
     ChildNo = Child
End Property

Public Property Let ChildNo(ByVal NChild As Double)
    Child = NChild
    PropertyChanged "Child"
End Property

Public Property Get PlayerAutoStart() As Boolean
     PlayerAutoStart = Autostart
End Property

Public Property Let PlayerAutoStart(ByVal NS As Boolean)
    Autostart = NS
    PropertyChanged "PlayerAutoStart"
End Property

Public Property Get MovieSizer() As MovSize
     MovieSizer = MSize
End Property

Public Property Let MovieSizer(ByVal NS As MovSize)
    MSize = NS
    If MSize = Original_Size Then
    
    If OriGinalX = -1 Then
      FrameVideo.Width = 0
    Else
      FrameVideo.Width = OriGinalX * Screen.TwipsPerPixelX
    End If
    
    If OriGinalY = -1 Then
      FrameVideo.Height = 0
    Else
      FrameVideo.Height = OriGinalY * Screen.TwipsPerPixelY
    End If
    
    If OriGinalX <> -1 And OriGinalY <> -1 And OpenedSucess = True Then
      FrameVideo.Enabled = True
    Else
      FrameVideo.Enabled = False
    End If
    UserControl_Resize
  End If
  
  If MSize = Resize_to_Fit_Current_Control Then
    UserControl_Resize
  End If
    PropertyChanged "MovieSizer"
End Property

Private Sub Volume_Click(Value As Double)
  If Value > -1 And Value < 101 Then
    CurVolMainPer = 100 - Value
    AdjustOutput PercentF(CurVolMainPer), CallBal(CurVolBalPer)
    PropertyChanged "VolumePercent"
  End If
End Sub

Public Function Mp3Information() As String
Dim X As String

If UCase(Right(FileName, 4)) = ".MP3" Then
  Call ReadMP3(FileName, True, True)
  With GetMP3Info
    X = "Bitrate :" + Str(.Bitrate) + vbCrLf
    X = X + "Frequency :" + Str(.Frequency) + vbCrLf
    X = X + "Mode : " + .Mode + vbCrLf
    X = X + "Emphasis : " + .Emphasis + vbCrLf
    X = X + "MpegVersion :" + Str(.MpegVersion) + vbCrLf
    X = X + "MpegLayer :" + Str(.MpegLayer) + vbCrLf
    X = X + "Padding :" + .Padding + vbCrLf
    X = X + "CRC :" + .CRC + vbCrLf
    X = X + "Duration :" + Str(.Duration) + vbCrLf
    X = X + "CopyRight : " + .CopyRight + vbCrLf
    X = X + "Original : " + .Original + vbCrLf
    X = X + "PrivateBit : " + .PrivateBit + vbCrLf
    X = X + "HasTag :" + Str(.HasTag) + vbCrLf
    X = X + "Songname : " + .Songname + vbCrLf
    X = X + "Artist : " + .Artist + vbCrLf
    X = X + "Album : " + .Tag + vbCrLf
    X = X + "Year : " + .Songname + vbCrLf
    X = X + "Comment : " + .Artist + vbCrLf
    X = X + "Genre :" + GenreText(.Genre) + vbCrLf
    X = X + "Track : " + .Artist + vbCrLf
    X = X + "VBR :" + Str(.VBR) + vbCrLf
    X = X + "Frames :" + Str(.Frames) + vbCrLf
  End With
  Mp3Information = X
Else
  Mp3Information = "Not an Mp3 File"
End If
End Function
Sub ShowAbout()
Attribute ShowAbout.VB_UserMemId = -552
  frmabout.Show vbModal
End Sub
