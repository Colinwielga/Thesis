VERSION 5.00
Begin VB.Form Form11 
   BackColor       =   &H00000000&
   Caption         =   "Form11"
   ClientHeight    =   7455
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9570
   LinkTopic       =   "Form11"
   Picture         =   "Form11.frx":0000
   ScaleHeight     =   7455
   ScaleWidth      =   9570
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdReturn10 
      Caption         =   "Return to Main Menu"
      BeginProperty Font 
         Name            =   "Onyx"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3240
      TabIndex        =   1
      Top             =   5400
      Width           =   3495
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "~frank verdeja"
      BeginProperty Font 
         Name            =   "Onyx"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   7920
      TabIndex        =   2
      Top             =   6960
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   $"Form11.frx":79A4
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1575
      Left            =   240
      TabIndex        =   0
      Top             =   3480
      Width           =   9015
   End
End
Attribute VB_Name = "Form11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This is the final form in the lesson.  It allows user to return to main menu.

Private Sub cmdReturn10_Click()
Form1.Show
Form11.Hide

End Sub
