VERSION 5.00
Begin VB.Form frmTitle 
   BackColor       =   &H00FF0000&
   Caption         =   "TitlePage"
   ClientHeight    =   8250
   ClientLeft      =   4170
   ClientTop       =   1740
   ClientWidth     =   7725
   LinkTopic       =   "Form1"
   ScaleHeight     =   8250
   ScaleWidth      =   7725
   Begin VB.CommandButton cmdMenu 
      BackColor       =   &H80000010&
      Caption         =   "Menu"
      BeginProperty Font 
         Name            =   "Mathematica5"
         Size            =   18
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7200
      Width           =   1815
   End
   Begin VB.PictureBox Picture1 
      Height          =   4935
      Left            =   1920
      Picture         =   "rugby.frx":0000
      ScaleHeight     =   4875
      ScaleWidth      =   3795
      TabIndex        =   2
      Top             =   2040
      Width           =   3855
   End
   Begin VB.Label lblTitle2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Rugby Football Club"
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Left            =   900
      TabIndex        =   1
      Top             =   1080
      Width           =   5775
   End
   Begin VB.Label lblTitle1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "St. John's"
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2430
      TabIndex        =   0
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "frmTitle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'St. John's Rugby
'Sam Herrmann
'March 2009

'First form viewed by user

Private Sub cmdMenu_Click()

frmTitle.Hide
frmMenu.Show

End Sub
