VERSION 5.00
Begin VB.Form frmLion 
   BackColor       =   &H00FF0000&
   Caption         =   "Lion King"
   ClientHeight    =   8325
   ClientLeft      =   2715
   ClientTop       =   1500
   ClientWidth     =   10275
   LinkTopic       =   "Form1"
   ScaleHeight     =   8325
   ScaleWidth      =   10275
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00004080&
      Caption         =   "Back"
      Height          =   1215
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6360
      Width           =   2295
   End
   Begin VB.PictureBox Picture1 
      Height          =   3495
      Left            =   840
      Picture         =   "frmLion.frx":0000
      ScaleHeight     =   3435
      ScaleWidth      =   2475
      TabIndex        =   1
      Top             =   960
      Width           =   2535
   End
   Begin VB.Label Label2 
      BackColor       =   &H00004080&
      Height          =   4215
      Left            =   360
      TabIndex        =   3
      Top             =   600
      Width           =   3495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00004080&
      Caption         =   $"frmLion.frx":2AB8
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6855
      Left            =   4560
      TabIndex        =   0
      Top             =   720
      Width           =   5175
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmLion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Disney Land Trivia
'frmAladdin
'Kelly Holmseth and Danny Hansen
'10/28/06
'Objective: The objective of this form is to provide a summary of the movie "The Lion King" to the user
Private Sub cmdBack_Click()
frmLion.Hide    'Allows user to go from Lion form to Top form
frmTop.Show
End Sub

