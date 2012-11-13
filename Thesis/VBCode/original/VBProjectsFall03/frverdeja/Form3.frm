VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H00000000&
   Caption         =   "Form3"
   ClientHeight    =   5625
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10980
   LinkTopic       =   "Form4"
   Picture         =   "Form3.frx":0000
   ScaleHeight     =   5625
   ScaleWidth      =   10980
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdReturn2 
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
      Height          =   1455
      Left            =   9840
      TabIndex        =   1
      Top             =   2760
      Width           =   855
   End
   Begin VB.CommandButton cmdNext2 
      Caption         =   "Next"
      BeginProperty Font 
         Name            =   "Onyx"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   9840
      TabIndex        =   0
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "~frank verdeja"
      BeginProperty Font 
         Name            =   "Onyx"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   9840
      TabIndex        =   3
      Top             =   5160
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   $"Form3.frx":8E6A
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   120
      TabIndex        =   2
      Top             =   4440
      Width           =   9495
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This form begins the lesson, and allows user to continue to next form, or quit.

Private Sub cmdNext2_Click()
Form4.Show
Form3.Hide

End Sub

Private Sub cmdReturn2_Click()
Form1.Show
Form3.Hide

End Sub

