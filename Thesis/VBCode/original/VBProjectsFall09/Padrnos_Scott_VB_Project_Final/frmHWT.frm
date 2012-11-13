VERSION 5.00
Begin VB.Form frmHWT 
   BackColor       =   &H80000001&
   Caption         =   "Form1"
   ClientHeight    =   8745
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12510
   BeginProperty Font 
      Name            =   "Viner Hand ITC"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   8745
   ScaleWidth      =   12510
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   10560
      TabIndex        =   5
      Top             =   5040
      Width           =   975
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Return to weight classes"
      Height          =   1935
      Left            =   10560
      TabIndex        =   4
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Heavyweight Weight Class"
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
      Left            =   2640
      TabIndex        =   6
      Top             =   360
      Width           =   4695
   End
   Begin VB.Label lblSocher2 
      Caption         =   $"frmHWT.frx":0000
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   5880
      TabIndex        =   3
      Top             =   4320
      Width           =   3855
   End
   Begin VB.Label lblEvenson2 
      Caption         =   $"frmHWT.frx":01D4
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   480
      TabIndex        =   2
      Top             =   4320
      Width           =   3615
   End
   Begin VB.Label lblSocher 
      Caption         =   "So. Cody Socher"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6960
      TabIndex        =   1
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label lblEvenson 
      Caption         =   "Sr. Jake Evenson"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1560
      TabIndex        =   0
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Image Image2 
      Height          =   2550
      Left            =   6000
      Picture         =   "frmHWT.frx":03A9
      Top             =   1560
      Width           =   3405
   End
   Begin VB.Image Image1 
      Height          =   2565
      Left            =   600
      Picture         =   "frmHWT.frx":27F3
      Top             =   1560
      Width           =   3405
   End
End
Attribute VB_Name = "frmHWT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdback_Click()
frmCompetition.Show 'hids the heavyweight form
frmHWT.Hide 'brings up the Competition form
End Sub

Private Sub cmdQuit_Click()
End 'ends the program
End Sub
