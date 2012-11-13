VERSION 5.00
Begin VB.Form frm165 
   BackColor       =   &H00000000&
   Caption         =   "Form1"
   ClientHeight    =   10065
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11580
   LinkTopic       =   "Form1"
   ScaleHeight     =   10065
   ScaleWidth      =   11580
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      Caption         =   "Return to weight classes"
      BeginProperty Font 
         Name            =   "Viner Hand ITC"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   9480
      TabIndex        =   5
      Top             =   3000
      Width           =   1935
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Exit"
      Height          =   855
      Left            =   9480
      TabIndex        =   4
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "165 lb. Weight Class"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3360
      TabIndex        =   6
      Top             =   360
      Width           =   3615
   End
   Begin VB.Label lblLydon2 
      Caption         =   $"frm165.frx":0000
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   6000
      TabIndex        =   3
      Top             =   4560
      Width           =   3135
   End
   Begin VB.Label lblLydon 
      Caption         =   "Sr. Grant Lydon"
      Height          =   255
      Left            =   6720
      TabIndex        =   2
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label lblBaarson3 
      Caption         =   "Jr. Matt Baarson"
      Height          =   255
      Left            =   1680
      TabIndex        =   1
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label lblBaarson4 
      Caption         =   $"frm165.frx":0117
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4095
      Index           =   0
      Left            =   840
      TabIndex        =   0
      Top             =   4560
      Width           =   3135
   End
   Begin VB.Image Image2 
      Height          =   2460
      Left            =   6360
      Picture         =   "frm165.frx":02D1
      Top             =   1680
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   2565
      Left            =   720
      Picture         =   "frm165.frx":4F95
      Top             =   1680
      Width           =   3405
   End
End
Attribute VB_Name = "frm165"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdback_Click()
frm165.Hide 'hides the 165 form
frmCompetition.Show 'shows the Competition

End Sub

Private Sub cmdQuit_Click()
'ends the program
End
End Sub
