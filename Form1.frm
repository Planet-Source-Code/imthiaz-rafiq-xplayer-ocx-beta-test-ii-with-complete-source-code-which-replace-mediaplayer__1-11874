VERSION 5.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   5475
   ClientLeft      =   -30
   ClientTop       =   -315
   ClientWidth     =   6150
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   5475
   ScaleWidth      =   6150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List2 
      Height          =   2205
      Left            =   6450
      TabIndex        =   4
      Top             =   1620
      Visible         =   0   'False
      Width           =   5985
   End
   Begin VB.DirListBox Dir1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   2250
      Left            =   210
      TabIndex        =   3
      Top             =   990
      Width           =   2175
   End
   Begin VB.FileListBox File1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   2130
      Left            =   2550
      MultiSelect     =   1  'Simple
      Pattern         =   "*.dat;*.avi;*.mp3;*.wav;*.mpg;*.wma"
      TabIndex        =   2
      Top             =   990
      Width           =   3165
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1470
      Left            =   210
      MultiSelect     =   2  'Extended
      TabIndex        =   1
      Top             =   3750
      Width           =   5505
   End
   Begin VB.DriveListBox Drive1 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FFFF&
      Height          =   315
      Left            =   990
      TabIndex        =   0
      Top             =   240
      Width           =   4575
   End
   Begin VB.Menu mnusimple 
      Caption         =   "simp"
      Visible         =   0   'False
      Begin VB.Menu mnuadd 
         Caption         =   "Add To Playlist"
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuDsecall 
         Caption         =   "Select All"
      End
      Begin VB.Menu mnudselnv 
         Caption         =   "Select Inverse"
      End
      Begin VB.Menu mnudselnon 
         Caption         =   "Select None"
      End
   End
   Begin VB.Menu mnus 
      Caption         =   "simp2"
      Visible         =   0   'False
      Begin VB.Menu mnurem 
         Caption         =   "Remove From Playlist"
      End
      Begin VB.Menu sep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuseall 
         Caption         =   "Select All"
      End
      Begin VB.Menu mnuselinv 
         Caption         =   "Select Inverse"
      End
      Begin VB.Menu mnuselnon 
         Caption         =   "Select None"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
On Error GoTo Errhand
Dir1.Path = Drive1.Drive
Exit Sub
Errhand:
If Err = 68 Then
    MsgBox "Device not ready", vbInformation
Else
    MsgBox Error(Err), vbInformation
End If
End Sub

Private Sub File1_DblClick()
Dim File As String
    If Not (Len(Dir1.Path) < 4) Then
        File = Dir1.Path + "\" + File1.FileName
    Else
        File = Dir1.Path + File1.FileName
    End If
    List1.AddItem File1.FileName
    List2.AddItem File
    List1.ListIndex = List1.ListCount - 1
End Sub

Private Sub File1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
    PopupMenu mnusimple
End If
End Sub

Private Sub File1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
    PopupMenu mnusimple
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub Form_Load()
List1.Clear
List2.Clear
For i = 0 To Frmplayer.List1.ListCount - 1
  List2.AddItem Frmplayer.List1.List(i)
  For j = Len(Frmplayer.List1.List(i)) To 1 Step -1
    If Mid$(Frmplayer.List1.List(i), j, 1) = "\" Then
      Exit For
    End If
  Next j
  List1.AddItem Mid$(Frmplayer.List1.List(i), j + 1, Len(Frmplayer.List1.List(i)) - j)
Next i
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Frmplayer.List1.Clear
    For i = 0 To Form1.List2.ListCount - 1
        Frmplayer.List1.AddItem Form1.List2.List(i)
    Next i
End Sub

Private Sub List1_KeyDown(KeyCode As Integer, Shift As Integer)
'MsgBox KeyCode
End Sub


Private Sub List1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
    PopupMenu mnus
End If
End Sub

Private Sub List1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
    PopupMenu mnus
End If
End Sub

Private Sub mnuadd_Click()
For i = 0 To File1.ListCount - 1
    If File1.Selected(i) = True Then
        File1.Selected(i) = True
        If Not (Len(Dir1.Path) < 4) Then
            File = Dir1.Path + "\" + File1.List(i)
        Else
            File = Dir1.Path + File1.List(i)
        End If
        List1.AddItem File1.List(i)
        List2.AddItem File
        List1.ListIndex = List1.ListCount - 1
    End If
Next i

End Sub

Private Sub MnuDsecall_Click()
For i = 0 To File1.ListCount - 1
    File1.Selected(i) = True
Next i
End Sub

Private Sub mnudselnon_Click()
For i = 0 To File1.ListCount - 1
    File1.Selected(i) = False
Next i
End Sub

Private Sub mnudselnv_Click()
For i = 0 To File1.ListCount - 1
    If File1.Selected(i) = True Then
        File1.Selected(i) = False
    Else
        File1.Selected(i) = True
    End If
Next i

End Sub

Private Sub mnurem_Click()
For i = List1.ListCount - 1 To 0 Step -1
    If List1.Selected(i) = True Then
        List1.RemoveItem (i)
        List2.RemoveItem (i)
    End If
Next i
End Sub

Private Sub mnuseall_Click()
For i = 0 To List1.ListCount - 1
    List1.Selected(i) = True
Next i
End Sub

Private Sub mnuselinv_Click()
For i = 0 To List1.ListCount - 1
    If List1.Selected(i) = True Then
        List1.Selected(i) = False
    Else
        List1.Selected(i) = True
    End If
Next i

End Sub

Private Sub mnuselnon_Click()
For i = 0 To List1.ListCount - 1
    List1.Selected(i) = False
Next i

End Sub
