VERSION 5.00
Begin VB.Form frm184 
   BackColor       =   &H00000000&
   Caption         =   "Form1"
   ClientHeight    =   8610
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13800
   LinkTopic       =   "Form1"
   ScaleHeight     =   8610
   ScaleWidth      =   13800
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   855
      Left            =   10080
      TabIndex        =   5
      Top             =   4560
      Width           =   975
   End
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
      Height          =   1575
      Left            =   10200
      TabIndex        =   4
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "184 lb. Weight Class"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3720
      TabIndex        =   6
      Top             =   360
      Width           =   3375
   End
   Begin VB.Label lblRaygor4 
      Caption         =   $"frm184.frx":0000
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   6600
      TabIndex        =   3
      Top             =   5040
      Width           =   2535
   End
   Begin VB.Label lblBaxter2 
      Caption         =   $"frm184.frx":0116
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   1320
      TabIndex        =   2
      Top             =   4560
      Width           =   3135
   End
   Begin VB.Label lblRaygor3 
      Caption         =   "Jr. Dustin Raygor"
      Height          =   255
      Left            =   7080
      TabIndex        =   1
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label lblBaxter 
      Caption         =   "Jr. Dustin Baxter"
      Height          =   255
      Left            =   2280
      TabIndex        =   0
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Image Image2 
      Height          =   3120
      Left            =   6720
      Picture         =   "frm184.frx":0247
      Top             =   1680
      Width           =   2190
   End
   Begin VB.Image Image1 
      Height          =   2535
      Left            =   1920
      Picture         =   "frm184.frx":1D84
      Top             =   1800
      Width           =   1890
   End
End
Attribute VB_Name = "frm184"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdback_Click()
frm184.Hide 'hides the 184 form
frmCompetition.Show 'brings up the Competition form
End Sub

Private Sub cmdQuit_Click()
End 'ends the program
End Sub
