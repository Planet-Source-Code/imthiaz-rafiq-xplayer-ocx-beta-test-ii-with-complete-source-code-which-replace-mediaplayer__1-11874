VERSION 5.00
Begin VB.Form frmabout 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3240
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   7005
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   7005
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image Image2 
      Height          =   1065
      Left            =   180
      Picture         =   "frmabout.frx":0000
      Top             =   1920
      Width           =   2760
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   585
      Left            =   5940
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   2580
      Width           =   1035
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Click this to vote me at www.planetsourcecode.com"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   570
      Left            =   3180
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   2580
      Width           =   2670
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Icq : 24294947"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   285
      Left            =   4110
      TabIndex        =   4
      Top             =   2130
      Width           =   1545
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Email : imthiazrafiq@hotmail.com"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   285
      Left            =   3300
      TabIndex        =   3
      Top             =   1800
      Width           =   3240
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Imthiaz Rafiq.H.M [NightMare]"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   345
      Left            =   3030
      TabIndex        =   2
      Top             =   1410
      Width           =   3720
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Programmed by."
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   405
      Left            =   3690
      TabIndex        =   1
      Top             =   900
      Width           =   2235
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   180
      Picture         =   "frmabout.frx":2799
      Top             =   150
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PlayerX.Ocx Beta III"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   405
      Left            =   750
      TabIndex        =   0
      Top             =   210
      Width           =   3060
   End
End
Attribute VB_Name = "frmabout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Label6.MouseIcon = Image1.Picture
Label7.MouseIcon = Image1.Picture

End Sub

Private Sub Label6_Click()
Shell "Start http://www.planet-source-code.com/xq/ASP/txtCodeId.11874/lngWId.1/qx/vb/scripts/ShowCode.htm", vbHide
End Sub

Private Sub Label7_Click()
  MsgBox "Thanks for downloading the code and voting me at Psc" + vbCrLf + vbCrLf + "Thank you", vbInformation + vbOKOnly, "Thank u"
  Unload Me
End Sub
