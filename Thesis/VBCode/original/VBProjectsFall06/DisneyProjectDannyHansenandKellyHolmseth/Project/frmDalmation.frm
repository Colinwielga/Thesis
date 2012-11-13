VERSION 5.00
Begin VB.Form frmDalmation 
   BackColor       =   &H00FF0000&
   Caption         =   "101 Dalmations"
   ClientHeight    =   7515
   ClientLeft      =   2520
   ClientTop       =   1920
   ClientWidth     =   10230
   LinkTopic       =   "Form1"
   ScaleHeight     =   7515
   ScaleWidth      =   10230
   Begin VB.PictureBox Picture2 
      Height          =   2535
      Left            =   7320
      Picture         =   "frmDalmation.frx":0000
      ScaleHeight     =   2475
      ScaleWidth      =   2475
      TabIndex        =   3
      Top             =   1440
      Width           =   2535
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H000000FF&
      Caption         =   "Back"
      Height          =   975
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5760
      Width           =   2535
   End
   Begin VB.PictureBox Picture1 
      Height          =   3615
      Left            =   0
      Picture         =   "frmDalmation.frx":1FF0
      ScaleHeight     =   3555
      ScaleWidth      =   2475
      TabIndex        =   1
      Top             =   1080
      Width           =   2535
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   $"frmDalmation.frx":5042
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4815
      Left            =   2760
      TabIndex        =   0
      Top             =   360
      Width           =   4335
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmDalmation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Load()
'Disney Land Trivia
'frmAladdin
'Kelly Holmseth and Danny Hansen
'10/28/06
'Objective: The objective of this form is to display to the user a summary of the movie "101 Dalmations"
Private Sub cmdBack_Click()
frmDalmation.Hide   'Allows user to go from Dalmation form to Top form
frmTop.Show
End Sub


