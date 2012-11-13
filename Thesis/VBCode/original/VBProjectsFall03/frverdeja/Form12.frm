VERSION 5.00
Begin VB.Form Form12 
   BackColor       =   &H00000000&
   Caption         =   "Form12"
   ClientHeight    =   8850
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10650
   LinkTopic       =   "Form7"
   ScaleHeight     =   8850
   ScaleWidth      =   10650
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdReturn10 
      Caption         =   "Return to Main Menu"
      BeginProperty Font 
         Name            =   "Onyx"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3840
      TabIndex        =   1
      Top             =   8040
      Width           =   2775
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "Graffiti Glossary"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   2640
      TabIndex        =   2
      Top             =   120
      Width           =   5295
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   $"Form12.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   6855
      Left            =   360
      TabIndex        =   0
      Top             =   1080
      Width           =   9975
   End
End
Attribute VB_Name = "Form12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This form teaches the user the vocabulary of the graffiti underworld, and allows the
'user to return to main menu.

Private Sub cmdReturn10_Click()
Form1.Show
Form12.Hide

End Sub
