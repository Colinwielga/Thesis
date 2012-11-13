VERSION 5.00
Begin VB.Form frmBeauty 
   BackColor       =   &H00FF0000&
   Caption         =   "Beauty and the Beast"
   ClientHeight    =   6420
   ClientLeft      =   3540
   ClientTop       =   2760
   ClientWidth     =   9060
   LinkTopic       =   "Form1"
   ScaleHeight     =   6420
   ScaleWidth      =   9060
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H0000C0C0&
      Caption         =   "Back"
      Height          =   1095
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4680
      Width           =   2295
   End
   Begin VB.PictureBox Picture1 
      Height          =   3615
      Left            =   360
      Picture         =   "frmBeauty.frx":0000
      ScaleHeight     =   3555
      ScaleWidth      =   2475
      TabIndex        =   1
      Top             =   480
      Width           =   2535
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   $"frmBeauty.frx":2DBE
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5415
      Left            =   3720
      TabIndex        =   0
      Top             =   480
      Width           =   5055
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmBeauty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Disney Land Trivia
'frmAladdin
'Kelly Holmseth and Danny Hansen
'10/28/06
'Objective: The objective of this form is to provide the user with a summary of the movie "Beauty And The Beast"
Private Sub cmdBack_Click()
frmBeauty.Hide      'Allows user to go from Beauty form to Top form
frmTop.Show

End Sub




