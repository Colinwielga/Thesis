VERSION 5.00
Begin VB.Form frmProduce 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form1"
   ClientHeight    =   8745
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   Picture         =   "frmProduce.frx":0000
   ScaleHeight     =   8745
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdcelery 
      BackColor       =   &H0080FFFF&
      Caption         =   "Celery (one bunch)"
      Height          =   855
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   5280
      Width           =   2055
   End
   Begin VB.CommandButton cmdContinue 
      BackColor       =   &H00FF00FF&
      Caption         =   "Continue Shopping or Check Out"
      Height          =   1215
      Left            =   11280
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   7080
      Width           =   3015
   End
   Begin VB.CommandButton cmdProduceTotal 
      BackColor       =   &H00FF8080&
      Caption         =   "Calculate Produce Subtotal"
      Height          =   1215
      Left            =   8040
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   7080
      Width           =   2895
   End
   Begin VB.CommandButton cmdGrapes 
      BackColor       =   &H0080FFFF&
      Caption         =   "Grapes"
      Height          =   855
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3360
      Width           =   2055
   End
   Begin VB.CommandButton cmdPears 
      BackColor       =   &H0080FFFF&
      Caption         =   "Pear (single)"
      Height          =   855
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2400
      Width           =   1935
   End
   Begin VB.CommandButton cmdBananas 
      BackColor       =   &H0080FFFF&
      Caption         =   "Banana (single)"
      Height          =   855
      Left            =   7440
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1440
      Width           =   2055
   End
   Begin VB.CommandButton cmdApples 
      BackColor       =   &H0080FFFF&
      Caption         =   "Apples (4ct)"
      Height          =   855
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4320
      Width           =   2055
   End
   Begin VB.CommandButton cmdGreenPepper 
      BackColor       =   &H0080FFFF&
      Caption         =   "Green Pepper"
      Height          =   855
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4320
      Width           =   2055
   End
   Begin VB.CommandButton cmdCarrots 
      BackColor       =   &H0080FFFF&
      Caption         =   "Bag of Carrots"
      Height          =   855
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2400
      Width           =   2055
   End
   Begin VB.CommandButton cmdLettuce 
      BackColor       =   &H0080FFFF&
      Caption         =   "Head of Lettuce"
      Height          =   855
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3360
      Width           =   2055
   End
   Begin VB.CommandButton cmdPotatoes 
      BackColor       =   &H0080FFFF&
      Caption         =   "Baking Potato"
      Height          =   855
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1440
      Width           =   2055
   End
   Begin VB.PictureBox picResults 
      Height          =   5535
      Left            =   9960
      ScaleHeight     =   5475
      ScaleWidth      =   4275
      TabIndex        =   1
      Top             =   1320
      Width           =   4335
   End
   Begin VB.Label lblItems 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Items Available: Click on button to add to cart"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   4800
      TabIndex        =   2
      Top             =   960
      Width           =   5055
   End
   Begin VB.Image Image1 
      Height          =   6000
      Left            =   360
      Picture         =   "frmProduce.frx":267C
      Top             =   2160
      Width           =   7500
   End
   Begin VB.Label lblWelcome 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Welcome To The Produce Section"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4440
      TabIndex        =   0
      Top             =   120
      Width           =   7335
   End
End
Attribute VB_Name = "frmProduce"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdApples_Click()
Dim Apples As Single
Apples = 3.29
'if the user clicks the button then the price of the item is added to the produce running total and displayed
ProduceRunningTotal = ProduceRunningTotal + Apples

picResults.Print "Apples"; Tab(25); FormatCurrency(Apples)

End Sub

Private Sub cmdBananas_Click()
Dim Banana As Single
Banana = 0.49
'if the user clicks the button then the price of the item is added to the produce running total and displayed
ProduceRunningTotal = ProduceRunningTotal + Banana

picResults.Print "Banana"; Tab(25); FormatCurrency(Banana)

End Sub

Private Sub cmdCarrots_Click()
Dim Carrots As Single
Carrots = 1.29
'if the user clicks the button then the price of the item is added to the produce running total and displayed
ProduceRunningTotal = ProduceRunningTotal + Carrots

picResults.Print "Bag of Carrots"; Tab(25); FormatCurrency(Carrots)

End Sub

Private Sub cmdcelery_Click()
Dim Celery As Single
Celery = 2.3
'if the user clicks the button then the price of the item is added to the produce running total and displayed
ProduceRunningTotal = ProduceRunningTotal + Celery

picResults.Print "Bunch of Celery"; Tab(25); FormatCurrency(Celery)

End Sub

Private Sub cmdContinue_Click()
'adds the produce total to the running total for all sections
RunningTotal = ProduceRunningTotal + BakeryRunningTotal + FrozenRunningTotal
'displays this amount to the user
MsgBox "Total amount spent so far is " & FormatCurrency(RunningTotal)
'takes user back to the enter form
frmEnter.Show
frmProduce.Hide
frmFrozen.Hide
frmBakery.Hide

End Sub

Private Sub cmdGrapes_Click()
Dim Grapes As Single
Grapes = 2.5
'if the user clicks the button then the price of the item is added to the produce running total and displayed
ProduceRunningTotal = ProduceRunningTotal + Grapes

picResults.Print "Bag of Grapes"; Tab(25); FormatCurrency(Grapes)

End Sub

Private Sub cmdGreenPepper_Click()
Dim GreenPepper As Single
GreenPepper = 0.79
'if the user clicks the button then the price of the item is added to the produce running total and displayed
ProduceRunningTotal = ProduceRunningTotal + GreenPepper

picResults.Print "Green Pepper"; Tab(25); FormatCurrency(GreenPepper)

End Sub

Private Sub cmdLettuce_Click()
Dim Lettuce As Single
Lettuce = 1.25
'if the user clicks the button then the price of the item is added to the produce running total and displayed
ProduceRunningTotal = ProduceRunningTotal + Lettuce

picResults.Print "Head of Lettuce"; Tab(25); FormatCurrency(Lettuce)

End Sub

Private Sub cmdPears_Click()
Dim Pear As Single
Pear = 0.89
'if the user clicks the button then the price of the item is added to the produce running total and displayed
ProduceRunningTotal = ProduceRunningTotal + Pear

picResults.Print "Pear"; Tab(25); FormatCurrency(Pear)

End Sub

Private Sub cmdPotatoes_Click()
Dim Potato As Single
Potato = 0.59
'if the user clicks the button then the price of the item is added to the produce running total and displayed
ProduceRunningTotal = ProduceRunningTotal + Potato

picResults.Print "Potato"; Tab(25); FormatCurrency(Potato)

End Sub

Private Sub cmdProduceTotal_Click()
'displays the produce subtotal in a picture box
picResults.Print "***********************************************************"
picResults.Print "Produce Subtotal: "; FormatCurrency(ProduceRunningTotal)
picResults.Print "***********************************************************"

End Sub

