VERSION 5.00
Begin VB.Form frmCinderella 
   BackColor       =   &H00FF0000&
   Caption         =   "Cinderella"
   ClientHeight    =   8100
   ClientLeft      =   2520
   ClientTop       =   1920
   ClientWidth     =   10155
   LinkTopic       =   "Form1"
   ScaleHeight     =   8100
   ScaleWidth      =   10155
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H008080FF&
      Caption         =   "Back"
      Height          =   1095
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5640
      Width           =   2295
   End
   Begin VB.PictureBox Picture1 
      Height          =   3495
      Left            =   600
      Picture         =   "frmCinderella.frx":0000
      ScaleHeight     =   3435
      ScaleWidth      =   2475
      TabIndex        =   1
      Top             =   960
      Width           =   2535
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      Caption         =   $"frmCinderella.frx":353A
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6975
      Left            =   3840
      TabIndex        =   0
      Top             =   120
      Width           =   6135
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmCinderella"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Disney Land Trivia
'frmAladdin
'Kelly Holmseth and Danny Hansen
'10/28/06
'Objective: The objective of this form is to display to the user a summary of the movie "Cinderella'
Private Sub cmdBack_Click()
frmCinderella.Hide  'Allows user to go from Cinderella form to Top form
frmTop.Show
End Sub


