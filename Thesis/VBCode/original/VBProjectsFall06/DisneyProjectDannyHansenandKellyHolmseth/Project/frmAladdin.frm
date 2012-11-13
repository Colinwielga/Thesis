VERSION 5.00
Begin VB.Form frmAladdin 
   BackColor       =   &H00FF0000&
   Caption         =   "Aladdin"
   ClientHeight    =   8535
   ClientLeft      =   2520
   ClientTop       =   1290
   ClientWidth     =   10410
   LinkTopic       =   "Form1"
   ScaleHeight     =   8535
   ScaleWidth      =   10410
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00FF00FF&
      Caption         =   "Back"
      Height          =   975
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6240
      Width           =   2055
   End
   Begin VB.PictureBox Picture1 
      Height          =   4095
      Left            =   480
      Picture         =   "frmAladdin.frx":0000
      ScaleHeight     =   4035
      ScaleWidth      =   2595
      TabIndex        =   1
      Top             =   840
      Width           =   2655
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FF00FF&
      Caption         =   $"frmAladdin.frx":382D
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7215
      Left            =   4320
      TabIndex        =   0
      Top             =   120
      Width           =   5535
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmAladdin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Disney Land Trivia
'frmAladdin
'Kelly Holmseth and Danny Hansen
'10/28/06
'Objective: The objective of this form is to display to the user a summary of the movie "Aladdin"
Private Sub cmdBack_Click()
frmAladdin.Hide     'Allows user to go from Aladdin form back to Top form
frmTop.Show

End Sub


