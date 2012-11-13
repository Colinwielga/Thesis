VERSION 5.00
Begin VB.Form frmCrunch 
   Caption         =   "Form1"
   ClientHeight    =   6615
   ClientLeft      =   3795
   ClientTop       =   1680
   ClientWidth     =   8385
   LinkTopic       =   "Form1"
   ScaleHeight     =   6615
   ScaleWidth      =   8385
   Begin VB.CommandButton cmdCrunch 
      Caption         =   "Return to Dancer Page"
      BeginProperty Font 
         Name            =   "Myriad Condensed Web"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   2760
      TabIndex        =   0
      Top             =   4800
      Width           =   2655
   End
   Begin VB.Label lblCrunch 
      BackColor       =   &H80000009&
      Caption         =   "Crunch with his old           girlfriend"
      BeginProperty Font 
         Name            =   "Myriad Condensed Web"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4335
   End
   Begin VB.Image Image1 
      Height          =   7200
      Left            =   -1200
      Picture         =   "frmCrunch.frx":0000
      Top             =   0
      Width           =   9600
   End
End
Attribute VB_Name = "frmCrunch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCrunch_Click()
'This command returns the user to the dancer form by hiding the Crunch form

frmCrunch.Visible = False
frmDancer.Visible = True

End Sub
