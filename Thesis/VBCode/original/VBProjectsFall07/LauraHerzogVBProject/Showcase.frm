VERSION 5.00
Begin VB.Form Showcase 
   BackColor       =   &H00800080&
   Caption         =   "Form1"
   ClientHeight    =   6210
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9345
   LinkTopic       =   "Form1"
   Picture         =   "Showcase.frx":0000
   ScaleHeight     =   6210
   ScaleWidth      =   9345
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdshow2 
      BackColor       =   &H00404080&
      Caption         =   "Showcase 2!!!"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2520
      Width           =   3495
   End
   Begin VB.CommandButton cmdshow1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Showcase 1!!!"
      BeginProperty Font 
         Name            =   "@Batang"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2520
      Width           =   3375
   End
   Begin VB.Label lblshowcase 
      BackStyle       =   0  'Transparent
      Caption         =   "Choose a showcase of a variety of items that you will be allowed to bid on!!"
      BeginProperty Font 
         Name            =   "@GulimChe"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   1680
      TabIndex        =   2
      Top             =   120
      Width           =   6735
   End
End
Attribute VB_Name = "Showcase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdshow1_Click()
'this allow the user to continue to the showcase that was chosen
cmdshow2.Enabled = False
showcase1.Show
Showcase.Hide
End Sub

Private Sub cmdshow2_Click()
'This button continues to the showcase that was chosen
cmdshow1.Enabled = False
showcase2.Show
Showcase.Hide
End Sub

