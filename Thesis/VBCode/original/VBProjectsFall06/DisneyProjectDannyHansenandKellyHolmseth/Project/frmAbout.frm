VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H008080FF&
   Caption         =   "About Us"
   ClientHeight    =   8040
   ClientLeft      =   2520
   ClientTop       =   1920
   ClientWidth     =   10320
   LinkTopic       =   "Form1"
   ScaleHeight     =   8040
   ScaleWidth      =   10320
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00FF0000&
      Caption         =   "Back"
      Height          =   615
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7200
      Width           =   1935
   End
   Begin VB.PictureBox Picture2 
      Height          =   5055
      Left            =   5280
      Picture         =   "frmAbout.frx":0000
      ScaleHeight     =   4995
      ScaleWidth      =   1995
      TabIndex        =   1
      Top             =   240
      Width           =   2055
   End
   Begin VB.PictureBox Picture1 
      Height          =   3255
      Left            =   1080
      Picture         =   "frmAbout.frx":3BC3
      ScaleHeight     =   3195
      ScaleWidth      =   1875
      TabIndex        =   0
      Top             =   360
      Width           =   1935
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C000&
      Caption         =   "Additional Sources:      www.imdb.com         and http://www.amazon.com/Top-Ten-Disney-Movies"
      Height          =   735
      Left            =   2640
      TabIndex        =   4
      Top             =   7200
      Width           =   3855
   End
   Begin VB.Label Label2 
      BackColor       =   &H000080FF&
      Caption         =   $"frmAbout.frx":65CA
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6855
      Left            =   7560
      TabIndex        =   3
      Top             =   240
      Width           =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      Caption         =   $"frmAbout.frx":6742
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   120
      TabIndex        =   2
      Top             =   3720
      Width           =   3975
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBack_Click()
frmAbout.Hide  'navigates to and from the main page
frmIntro.Show
End Sub

