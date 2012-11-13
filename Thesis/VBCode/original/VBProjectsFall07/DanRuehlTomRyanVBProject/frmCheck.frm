VERSION 5.00
Begin VB.Form frmCheck 
   BackColor       =   &H00800000&
   Caption         =   "Form1"
   ClientHeight    =   4470
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12585
   FillColor       =   &H0000FFFF&
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   18
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   4470
   ScaleWidth      =   12585
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdEnd 
      BackColor       =   &H00800000&
      Caption         =   "Exit"
      Height          =   735
      Left            =   9720
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton cmdShow 
      BackColor       =   &H00FF0000&
      Caption         =   "Print Check!"
      Height          =   1815
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   600
      Width           =   5175
   End
   Begin VB.PictureBox picCheck2 
      BackColor       =   &H00800000&
      FillColor       =   &H0000FFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   735
      Left            =   8520
      ScaleHeight     =   675
      ScaleWidth      =   2475
      TabIndex        =   1
      Top             =   2400
      Width           =   2535
   End
   Begin VB.PictureBox picCheck1 
      BackColor       =   &H00800000&
      FillColor       =   &H0000FFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1095
      Left            =   720
      ScaleHeight     =   1035
      ScaleWidth      =   7395
      TabIndex        =   0
      Top             =   2400
      Width           =   7455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "November 6th 2007"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   735
      Left            =   8280
      TabIndex        =   2
      Top             =   840
      Width           =   2535
   End
   Begin VB.Image Image1 
      Height          =   1230
      Left            =   480
      Picture         =   "frmCheck.frx":0000
      Top             =   600
      Width           =   4815
   End
End
Attribute VB_Name = "frmCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdEnd_Click()
End
End Sub

Private Sub cmdShow_Click()
'This button prints the user's name and final score into a check format.
cmdShow.Visible = False
picCheck1.Print Contestant
picCheck2.Print FormatCurrency(Sum)
End Sub
