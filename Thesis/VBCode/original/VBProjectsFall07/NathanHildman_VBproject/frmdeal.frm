VERSION 5.00
Begin VB.Form frmdeal 
   BackColor       =   &H00000000&
   Caption         =   "Form1"
   ClientHeight    =   8160
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12150
   LinkTopic       =   "Form1"
   Picture         =   "frmdeal.frx":0000
   ScaleHeight     =   8160
   ScaleWidth      =   12150
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      BackColor       =   &H0000FFFF&
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4200
      TabIndex        =   2
      Text            =   "You Have Won"
      Top             =   2280
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FFFF&
      Caption         =   "Click here to see how much you've won"
      Height          =   855
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6000
      Width           =   1935
   End
   Begin VB.PictureBox picdeal1 
      BackColor       =   &H0000FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   3000
      ScaleHeight     =   1395
      ScaleWidth      =   5115
      TabIndex        =   0
      Top             =   3600
      Width           =   5175
   End
End
Attribute VB_Name = "frmdeal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()

picdeal1.Print FormatCurrency(bankersoffer)         'displays amount of money won
End Sub

Private Sub Form_Load()

picdeal1.Print FormatCurrency(bankersoffer)         'displays amount of money won

End Sub

