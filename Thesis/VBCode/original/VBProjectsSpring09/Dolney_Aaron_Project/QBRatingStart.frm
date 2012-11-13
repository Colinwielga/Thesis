VERSION 5.00
Begin VB.Form frmStart 
   BackColor       =   &H00FF0000&
   Caption         =   "Form1"
   ClientHeight    =   8505
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10920
   LinkTopic       =   "Form1"
   ScaleHeight     =   8505
   ScaleWidth      =   10920
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdNCAA 
      Caption         =   "NCAA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4335
      Left            =   5520
      TabIndex        =   1
      Top             =   1920
      Width           =   4695
   End
   Begin VB.CommandButton cmdNFL 
      Caption         =   "NFL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4335
      Left            =   360
      TabIndex        =   0
      Top             =   1920
      Width           =   4815
   End
End
Attribute VB_Name = "frmStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdNCAA_Click()
frmStart.Hide
frmNFL.Hide
frmNCAA.Show

End Sub

Private Sub cmdNFL_Click()
frmStart.Hide
frmNFL.Show
frmNCAA.Hide

End Sub


