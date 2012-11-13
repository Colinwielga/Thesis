VERSION 5.00
Begin VB.Form frmCheckout 
   BackColor       =   &H00004000&
   Caption         =   "The Campground"
   ClientHeight    =   7800
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7515
   LinkTopic       =   "Form1"
   ScaleHeight     =   7800
   ScaleWidth      =   7515
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCheck 
      Caption         =   "Pay with check."
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Orator Std"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4200
      TabIndex        =   3
      Top             =   6600
      Width           =   2295
   End
   Begin VB.CommandButton cmdCash 
      Caption         =   "Pay with cash."
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Orator Std"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   840
      TabIndex        =   2
      Top             =   6600
      Width           =   2295
   End
   Begin VB.PictureBox picResults 
      BeginProperty Font 
         Name            =   "Orator Std"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4575
      Left            =   840
      ScaleHeight     =   4515
      ScaleWidth      =   5595
      TabIndex        =   1
      Top             =   1680
      Width           =   5655
   End
   Begin VB.CommandButton cmdGrandTotal 
      BackColor       =   &H00004000&
      Caption         =   "Click for your grand total."
      BeginProperty Font 
         Name            =   "Orator Std"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   6375
   End
End
Attribute VB_Name = "frmCheckout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCash_Click()
'takes the user to the cash payment form using the Visible property
frmCheckout.Hide
frmCash.Show
End Sub

Private Sub cmdCheck_Click()
'takes user to the check payment form using the Visible property
frmCheckout.Hide
frmCheck.Show
End Sub


Private Sub cmdGrandTotal_Click()
'computes the subtotals and quantities for each item, subtotal, tax, and grand total
'and prints in a picture box
'allows the user to click the command boxes for each payment after Grand Total button
'has been clicked
picResults.Cls
picResults.Print "Product"; Tab(15); "Quantity"; Tab(35); "Item Subtotal"
picResults.Print "**********************************************************************"
picResults.Print "Sleeping Bag"; Tab(15); SBCTR; Tab(35); FormatCurrency(SleepingBagSub)
picResults.Print "Tent"; Tab(15); TCTR; Tab(35); FormatCurrency(TentSub)
picResults.Print "Jacket"; Tab(15); XJCTR + RJCTR; Tab(35); FormatCurrency(JacketSub)
picResults.Print "Mess Kit"; Tab(15); MKCTR; Tab(35); FormatCurrency(MessKitSub)
picResults.Print
picResults.Print "Subtotal:"; Tab(35); FormatCurrency(Subtotal)
picResults.Print "Tax:"; Tab(15); "6.5%"; Tab(35); FormatCurrency(Subtotal * 0.065, 2)
picResults.Print "**********************************************************************"
picResults.Print "GRAND TOTAL:"; Tab(35); FormatCurrency(Subtotal + Subtotal * 0.065, 2)
GrandTotal = FormatNumber(Subtotal + Subtotal * 0.065, 2)
cmdCash.Enabled = True
cmdCheck.Enabled = True
End Sub
