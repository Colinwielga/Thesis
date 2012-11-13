VERSION 5.00
Begin VB.Form frmMulan 
   BackColor       =   &H00FF0000&
   Caption         =   "Mulan"
   ClientHeight    =   8625
   ClientLeft      =   2115
   ClientTop       =   1290
   ClientWidth     =   10485
   LinkTopic       =   "Form1"
   ScaleHeight     =   8625
   ScaleWidth      =   10485
   Begin VB.PictureBox Picture2 
      Height          =   4335
      Left            =   1680
      Picture         =   "frmMulan.frx":0000
      ScaleHeight     =   4275
      ScaleWidth      =   3075
      TabIndex        =   3
      Top             =   3000
      Width           =   3135
   End
   Begin VB.PictureBox Picture1 
      Height          =   2895
      Left            =   120
      Picture         =   "frmMulan.frx":2904
      ScaleHeight     =   2835
      ScaleWidth      =   4635
      TabIndex        =   2
      Top             =   0
      Width           =   4695
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H0080C0FF&
      Caption         =   "Back"
      Height          =   1095
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7440
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H0080C0FF&
      Caption         =   $"frmMulan.frx":7DC7
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
      Left            =   5040
      TabIndex        =   0
      Top             =   480
      Width           =   5175
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmMulan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Disney Land Trivia
'frmAladdin
'Kelly Holmseth and Danny Hansen
'10/30/06
'Objective: The objective of this form  is to display to the user a summary of the movie "Mulan"
Private Sub cmdBack_Click()
frmMulan.Hide
frmTop.Show
End Sub


