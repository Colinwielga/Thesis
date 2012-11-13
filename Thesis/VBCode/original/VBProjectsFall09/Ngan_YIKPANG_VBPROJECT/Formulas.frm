VERSION 5.00
Begin VB.Form Formulas 
   Caption         =   "Form1"
   ClientHeight    =   6255
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8295
   LinkTopic       =   "Form1"
   Picture         =   "Formulas.frx":0000
   ScaleHeight     =   6255
   ScaleWidth      =   8295
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmd4 
      Caption         =   "Go back"
      Height          =   855
      Left            =   720
      TabIndex        =   4
      Top             =   5280
      Width           =   1575
   End
   Begin VB.CommandButton cmd3 
      Caption         =   "To find number of periods requried"
      Height          =   855
      Left            =   480
      TabIndex        =   3
      Top             =   3120
      Width           =   2175
   End
   Begin VB.CommandButton cmd2 
      Caption         =   " Present Value formula"
      Height          =   855
      Left            =   480
      TabIndex        =   2
      Top             =   2160
      Width           =   2175
   End
   Begin VB.PictureBox picresults 
      Height          =   1815
      Left            =   4080
      ScaleHeight     =   1755
      ScaleWidth      =   3675
      TabIndex        =   1
      Top             =   3000
      Width           =   3735
   End
   Begin VB.CommandButton cmd1 
      Caption         =   "Future value formula"
      Height          =   855
      Left            =   480
      TabIndex        =   0
      Top             =   1080
      Width           =   2175
   End
   Begin VB.Label lblFormulas 
      BackColor       =   &H80000012&
      Caption         =   "Formulas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000013&
      Height          =   495
      Left            =   3120
      TabIndex        =   5
      Top             =   360
      Width           =   1695
   End
End
Attribute VB_Name = "Formulas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name:Compound Interest
'Form:formulas
'Author:Yik Pang Ngan (Banny)
'Date Written:Oct 18 2009
Option Explicit
'this form will explain what types of formulas we can use to calculate compound interest

Private Sub cmd1_Click()

picresults.Picture = LoadPicture(App.Path & "\pictures\themostbasicformula.jpg")
'This buttom will show the formula of calulating Future Value
End Sub

Private Sub cmd2_Click()
picresults.Picture = LoadPicture(App.Path & "\pictures\calculatesPV.JPG")
'This buttom will show the formula of calulating Present Value
End Sub

Private Sub cmd3_Click()
picresults.Picture = LoadPicture(App.Path & "\pictures\calculateinterest.JPG")
'This buttom will show the formula of calulating number of periods requried
End Sub

Private Sub cmd4_Click()
Formulas.Hide
CompoundInterest.Show
'This buttom will switch back to the main form
End Sub

