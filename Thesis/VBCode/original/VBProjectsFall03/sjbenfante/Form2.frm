VERSION 5.00
Begin VB.Form TicketPricing 
   Caption         =   "Form2"
   ClientHeight    =   6660
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9450
   LinkTopic       =   "Form2"
   Picture         =   "Form2.frx":0000
   ScaleHeight     =   6660
   ScaleWidth      =   9450
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   855
      Left            =   240
      TabIndex        =   3
      Top             =   3360
      Width           =   975
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Return To Previous Page"
      Height          =   375
      Left            =   3840
      TabIndex        =   2
      Top             =   6120
      Width           =   4095
   End
   Begin VB.CommandButton cmdDiscounts 
      Caption         =   "To find a group discount, click HERE!"
      Height          =   1335
      Left            =   360
      TabIndex        =   1
      Top             =   4800
      Width           =   1815
   End
   Begin VB.CommandButton cmdPrices 
      Caption         =   "To find the price of tickets, click HERE!"
      Height          =   1335
      Left            =   360
      TabIndex        =   0
      Top             =   1440
      Width           =   1815
   End
End
Attribute VB_Name = "TicketPricing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBack_Click()
TicketPricing.Hide
HomePage.Show
'this hides the second form and shows the first form'
End Sub

Private Sub cmdDiscounts_Click()
TicketPricing.Hide
GiveMeMyDiscount.Show
'this hides the second form and shows the fifth form'
End Sub

Private Sub cmdPrices_Click()
TicketPricing.Hide
WhatIsTheCost.Show
'this hides the second form and shows the third form'
End Sub

Private Sub cmdQuit_Click()
    End
'this automatically end the program'
End Sub

Private Sub Form_Load()
strPath = "n:\CS130\handin\sjbenfante\"
End Sub
