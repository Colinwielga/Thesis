VERSION 5.00
Begin VB.Form frmHome 
   Caption         =   "Form1"
   ClientHeight    =   11010
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   Picture         =   "frmHome.frx":0000
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdTrivia 
      BackColor       =   &H000000C0&
      Caption         =   "Trivia"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   240
      Width           =   2175
   End
   Begin VB.CommandButton cmdBuy 
      BackColor       =   &H000000C0&
      Caption         =   "Buy Stuff"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2040
      Width           =   2175
   End
End
Attribute VB_Name = "frmHome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBuy_Click()
'Go from our homepage to our store page.
frmProducts.Show
frmHome.Hide
End Sub

Private Sub cmdTrivia_Click()
'Go from our homepage to our Trivia page.
frmTrivia.Show
frmHome.Hide
End Sub
