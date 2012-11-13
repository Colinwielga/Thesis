VERSION 5.00
Begin VB.Form frmMermaid 
   BackColor       =   &H00FF0000&
   Caption         =   "The Little Mermaid"
   ClientHeight    =   7740
   ClientLeft      =   2520
   ClientTop       =   2130
   ClientWidth     =   10065
   LinkTopic       =   "Form1"
   ScaleHeight     =   7740
   ScaleWidth      =   10065
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H000000FF&
      Caption         =   "Back"
      Height          =   1095
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5880
      Width           =   2535
   End
   Begin VB.PictureBox Picture1 
      Height          =   3615
      Left            =   600
      Picture         =   "frmMermaid.frx":0000
      ScaleHeight     =   3555
      ScaleWidth      =   2475
      TabIndex        =   1
      Top             =   840
      Width           =   2535
   End
   Begin VB.Label Label2 
      BackColor       =   &H000000FF&
      Height          =   4335
      Left            =   240
      TabIndex        =   3
      Top             =   480
      Width           =   3255
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   $"frmMermaid.frx":303C
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
      Left            =   4200
      TabIndex        =   0
      Top             =   480
      Width           =   5055
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmMermaid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Disney Land Trivia
'frmAladdin
'Kelly Holmseth and Danny Hansen
'10/28/06
'Objective: The objective of this form  is to display to the user a summary of the movie "The Little Mermaid"
Private Sub cmdBack_Click()
frmMermaid.Hide     'Allow the user to go back to the Top form
frmTop.Show
End Sub


