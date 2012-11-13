VERSION 5.00
Begin VB.Form frmLions 
   BackColor       =   &H00000000&
   Caption         =   "Lions"
   ClientHeight    =   7545
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10905
   LinkTopic       =   "Form1"
   ScaleHeight     =   7545
   ScaleWidth      =   10905
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return to Teams"
      Height          =   495
      Left            =   840
      TabIndex        =   7
      Top             =   3480
      Width           =   1575
   End
   Begin VB.CommandButton cmdRun2 
      Caption         =   "Run 2"
      Height          =   495
      Left            =   2040
      TabIndex        =   6
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton cmdRun3 
      Caption         =   "Run 3"
      Height          =   495
      Left            =   2040
      TabIndex        =   5
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton cmdRun1 
      Caption         =   "Run 1"
      Height          =   495
      Left            =   2040
      TabIndex        =   4
      Top             =   840
      Width           =   1095
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00FFFFC0&
      Height          =   6015
      Left            =   3600
      ScaleHeight     =   5955
      ScaleWidth      =   5595
      TabIndex        =   3
      Top             =   960
      Width           =   5655
   End
   Begin VB.CommandButton cmdPass3 
      Caption         =   "Pass 3"
      Height          =   495
      Left            =   720
      TabIndex        =   2
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton cmdPass2 
      Caption         =   "Pass 2"
      Height          =   495
      Left            =   720
      TabIndex        =   1
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton cmdPass1 
      Caption         =   "Pass 1"
      Default         =   -1  'True
      Height          =   495
      Left            =   720
      TabIndex        =   0
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label lblChoice 
      Caption         =   "Choose a Play to use"
      Height          =   375
      Left            =   720
      TabIndex        =   8
      Top             =   240
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   7080
      Left            =   480
      Picture         =   "frmLions.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9225
   End
End
Attribute VB_Name = "frmLions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdPass1_Click()
    picResults.Cls 'cleas the picture box
    picResults.Picture = LoadPicture(App.Path & "\calvin.jpg") 'loads the picture
    MsgBox "TouchDown Calvin Johnson!", , "Touchdown"
End Sub

Private Sub cmdPass2_Click()
    picResults.Cls 'cleas the picture box
    picResults.Picture = LoadPicture(App.Path & "\wood.jpg") 'loads the picture
    MsgBox "Interception to Charles Woodson", , "Yikes!"
End Sub

Private Sub cmdPass3_Click()
    picResults.Cls 'cleas the picture box
    picResults.Picture = LoadPicture(App.Path & "\matt.jpg") 'loads the picture
    MsgBox "First Down!", , "Completion"
End Sub

Private Sub cmdReturn_Click()
    frmLions.Hide 'Hides form from user
    frmTeams.Show 'shows form to user
End Sub

Private Sub cmdRun1_Click()
    picResults.Cls 'cleas the picture box
    picResults.Picture = LoadPicture(App.Path & "\Jared.jpg") 'loads the picture
    MsgBox "Jared Allen Just Sacked You!", , "Sack"
End Sub

Private Sub cmdRun2_Click()
    picResults.Cls 'cleas the picture box
    picResults.Picture = LoadPicture(App.Path & "\kevin.jpg ") 'loads the picture"
    MsgBox "TouchDown! Kevin Smith", , "TouchDown!"
End Sub

Private Sub cmdRun3_Click()
    picResults.Cls 'cleas the picture box
    picResults.Picture = LoadPicture(App.Path & "\matt.jpg") 'loads the picture
    MsgBox "QB Sneak for 5 yards.", , "Sneak"
End Sub

