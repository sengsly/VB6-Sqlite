VERSION 5.00
Object = "{7935966C-7928-4E31-99B3-E9DB87FD9E30}#2.0#0"; "UniEditControl.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4230
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6105
   LinkTopic       =   "Form1"
   ScaleHeight     =   4230
   ScaleWidth      =   6105
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "User Media"
      Height          =   495
      Left            =   3120
      TabIndex        =   2
      Top             =   2760
      Width           =   1215
   End
   Begin UniEditControl.ControlEdit ControlEdit1 
      Height          =   495
      Left            =   3120
      TabIndex        =   1
      Top             =   600
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   8454143
      FontSize        =   8.25
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Get Label"
      Height          =   495
      Left            =   1680
      TabIndex        =   0
      Top             =   1320
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cYMGuide As New cLanguage
Private Sub Command1_Click()
   'Debug.Print cYMGuide.GetTracksByAlbum(134, True)
   'Debug.Print cYMGuide.GetGenre
   'Debug.Print cYMGuide.GetLabel
   Debug.Print cYMGuide.GetActorName(";1;2;3;", False)
   'Debug.Print cYMGuide.GetDirectorName("1,2", 2)
End Sub

Private Sub Command2_Click()
   'MsgBox cYMGuide.GetLabelName(12324324)
   MsgBox cYMGuide.GetActorName(";14;34;8", True)
End Sub

Private Sub Command3_Click()
   Dim i As New CLocalMedia
   i.Initialize
   'i.addFolder App.Path
   i.addFolder "C:\"
   'i.ScanMedias
   'MsgBox i.getMedias
End Sub

Private Sub Form_Load()
   cYMGuide.Initialize App.Path & "\Database\khmer karaoke.sqlite", 0

End Sub
