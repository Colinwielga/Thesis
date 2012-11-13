VERSION 5.00
Begin VB.Form frmCover 
   BackColor       =   &H00000040&
   Caption         =   "Cover"
   ClientHeight    =   12810
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   19830
   LinkTopic       =   "Form1"
   ScaleHeight     =   12810
   ScaleWidth      =   19830
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picPicture2 
      Height          =   5055
      Left            =   13920
      ScaleHeight     =   4995
      ScaleWidth      =   4515
      TabIndex        =   3
      Top             =   2280
      Width           =   4575
   End
   Begin VB.PictureBox picPicture 
      Height          =   5175
      Left            =   360
      ScaleHeight     =   5115
      ScaleWidth      =   4395
      TabIndex        =   2
      Top             =   1800
      Width           =   4455
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Lets Add Up Some Points!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   8520
      TabIndex        =   0
      Top             =   7440
      Width           =   3255
   End
   Begin VB.Label Label1 
      Caption         =   $"frmCover.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   6840
      TabIndex        =   1
      Top             =   3240
      Width           =   6615
   End
End
Attribute VB_Name = "frmCover"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdStart_Click()
PointsWanted = InputBox("How many points do you want to play with?")

frmCover.Hide
frmHome.Show

End Sub

Private Sub Form_Load()

picPicture.AutoSize = True
picPicture.Picture = LoadPicture(App.Path & "\codex.JPG")


picPicture2.AutoSize = True
picPicture2.Picture = LoadPicture(App.Path & "\rules.GIF")

End Sub
