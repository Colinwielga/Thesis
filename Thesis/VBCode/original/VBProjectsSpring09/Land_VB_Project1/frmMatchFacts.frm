VERSION 5.00
Begin VB.Form frmMatchFacts 
   Caption         =   "Match Facts"
   ClientHeight    =   6405
   ClientLeft      =   4965
   ClientTop       =   3975
   ClientWidth     =   8040
   LinkTopic       =   "Form1"
   ScaleHeight     =   6405
   ScaleWidth      =   8040
   Begin VB.CommandButton cmdMatchBookNumber 
      Caption         =   "Click to match the book numbers"
      Height          =   735
      Left            =   120
      TabIndex        =   4
      Top             =   3000
      Width           =   1455
   End
   Begin VB.PictureBox picResults 
      Height          =   1935
      Left            =   1920
      ScaleHeight     =   1875
      ScaleWidth      =   1995
      TabIndex        =   3
      Top             =   120
      Width           =   2055
   End
   Begin VB.CommandButton cmdReadMatch 
      Caption         =   "Click to see the books in the series"
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   2040
      Width           =   1455
   End
   Begin VB.CommandButton cmdReturnMatch 
      Caption         =   "Click to return to the matching games menu"
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   1455
   End
   Begin VB.CommandButton cmdReturnMain 
      Caption         =   "Click to return to the main menu"
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "frmMatchFacts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub cmdReturnMain_Click()
    'return to main menu
    frmStart.Show
    frmMatchFacts.Hide
    
End Sub

Private Sub cmdReturnMatch_Click()
    'return to matching games menu
    frmMatch.Show
    frmMatchFacts.Hide
    
End Sub
