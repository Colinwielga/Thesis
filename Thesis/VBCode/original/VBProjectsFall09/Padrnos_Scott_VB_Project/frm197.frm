VERSION 5.00
Begin VB.Form frm197 
   BackColor       =   &H000000C0&
   Caption         =   "Form1"
   ClientHeight    =   9690
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15570
   LinkTopic       =   "Form1"
   ScaleHeight     =   9690
   ScaleWidth      =   15570
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Exit"
      Height          =   735
      Left            =   9480
      TabIndex        =   5
      Top             =   8040
      Width           =   855
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
      Height          =   2175
      Left            =   9600
      TabIndex        =   4
      Top             =   600
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "197 lb. Weight Class"
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
      Left            =   4920
      TabIndex        =   6
      Top             =   840
      Width           =   3615
   End
   Begin VB.Label lblJames2 
      Caption         =   $"frm197.frx":0000
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
      Left            =   5160
      TabIndex        =   3
      Top             =   4920
      Width           =   3375
   End
   Begin VB.Label lblWillaert2 
      Caption         =   $"frm197.frx":015E
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4335
      Left            =   840
      TabIndex        =   2
      Top             =   4800
      Width           =   3255
   End
   Begin VB.Label lblJames 
      Caption         =   "Jr. James Carlson"
      Height          =   255
      Left            =   6120
      TabIndex        =   1
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label lblWillaert 
      Caption         =   "Jr. Tony Willaert"
      Height          =   255
      Left            =   1680
      TabIndex        =   0
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Image Image2 
      Height          =   2820
      Left            =   1320
      Picture         =   "frm197.frx":0315
      Top             =   1560
      Width           =   1890
   End
   Begin VB.Image Image1 
      Height          =   2250
      Left            =   5400
      Picture         =   "frm197.frx":1EF1
      Top             =   2280
      Width           =   2820
   End
End
Attribute VB_Name = "frm197"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdback_Click()
frm197.Hide 'hides the 197 form
frmCompetition.Show 'brings up the Competition form
End Sub

Private Sub cmdQuit_Click()
End 'Ends the program
End Sub
