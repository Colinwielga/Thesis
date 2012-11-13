VERSION 5.00
Begin VB.Form frmWhat 
   BackColor       =   &H80000012&
   Caption         =   "What Would you Like to Do?"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "This is way too hard for me.  I quit."
      Height          =   855
      Left            =   3960
      TabIndex        =   3
      Top             =   3360
      Width           =   1935
   End
   Begin VB.CommandButton cmdTrivia 
      Caption         =   "Try out some Trivia"
      Height          =   1335
      Left            =   1200
      TabIndex        =   1
      Top             =   1440
      Width           =   2655
   End
   Begin VB.CommandButton cmdAuthors 
      BackColor       =   &H80000005&
      Caption         =   "Test Yourself on your Knowledge of Authors"
      Height          =   1335
      Left            =   5880
      TabIndex        =   0
      Top             =   1440
      Width           =   2535
   End
   Begin VB.Label lblExplain 
      BackColor       =   &H80000012&
      Caption         =   $"frmWhat.frx":0000
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   1095
      Left            =   360
      TabIndex        =   2
      Top             =   120
      Width           =   10215
   End
End
Attribute VB_Name = "frmWhat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdQuit_Click()
    End
End Sub

Private Sub cmdAuthors_Click()
     'go to authors form
     frmWhat.Visible = False
     frmAuthors.Visible = True
   
End Sub

Private Sub cmdTrivia_Click()
    'go to trivia form
    frmTrivia.Visible = True
    frmWhat.Visible = False
    
End Sub
