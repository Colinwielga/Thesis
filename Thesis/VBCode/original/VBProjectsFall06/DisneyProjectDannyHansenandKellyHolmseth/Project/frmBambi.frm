VERSION 5.00
Begin VB.Form frmBambi 
   BackColor       =   &H00FF0000&
   Caption         =   "Bambi"
   ClientHeight    =   7920
   ClientLeft      =   3135
   ClientTop       =   1920
   ClientWidth     =   9360
   LinkTopic       =   "Form1"
   ScaleHeight     =   7920
   ScaleWidth      =   9360
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H0000C000&
      Caption         =   "Back"
      Height          =   855
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5400
      Width           =   2055
   End
   Begin VB.PictureBox Picture1 
      Height          =   4095
      Left            =   720
      Picture         =   "frmBambi.frx":0000
      ScaleHeight     =   4035
      ScaleWidth      =   2355
      TabIndex        =   1
      Top             =   720
      Width           =   2415
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H0000C000&
      Caption         =   $"frmBambi.frx":3361
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6255
      Left            =   4440
      TabIndex        =   0
      Top             =   240
      Width           =   4215
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmBambi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Disney Land Trivia
'frmAladdin
'Kelly Holmseth and Danny Hansen
'10/28/06
'Objective: The objective of this form is to provide the user with a summary of the movie "Bambi"
Private Sub cmdBack_Click()
frmBambi.Hide   'Allows user to go from Bambi form to Top form
frmTop.Show
End Sub



