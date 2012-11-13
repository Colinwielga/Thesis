VERSION 5.00
Begin VB.Form frmQuit 
   Caption         =   "Quit?"
   ClientHeight    =   1785
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3390
   LinkTopic       =   "Form1"
   ScaleHeight     =   1785
   ScaleWidth      =   3390
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdNo 
      Caption         =   "No"
      Height          =   615
      Left            =   1800
      TabIndex        =   2
      Top             =   1080
      Width           =   1335
   End
   Begin VB.CommandButton cmdYes 
      Caption         =   "Yes"
      Height          =   615
      Left            =   240
      TabIndex        =   1
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label lblAreYouSure 
      Alignment       =   2  'Center
      Caption         =   "You have not saved, are you sure you want to quit?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3135
   End
End
Attribute VB_Name = "frmQuit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'opens if the user hasn't saved

Private Sub cmdNo_Click()
    'hides this form
    frmQuit.Hide
End Sub

Private Sub cmdYes_Click()
    'ends the program
    End
End Sub

