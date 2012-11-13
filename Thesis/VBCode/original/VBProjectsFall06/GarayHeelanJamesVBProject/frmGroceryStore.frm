VERSION 5.00
Begin VB.Form frmGroceryStore 
   Caption         =   "Central Shopping Menu"
   ClientHeight    =   9705
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13335
   LinkTopic       =   "Form1"
   ScaleHeight     =   9705
   ScaleWidth      =   13335
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCustomerService 
      Caption         =   "Customer Service"
      Height          =   1215
      Left            =   8280
      TabIndex        =   5
      Top             =   6120
      Width           =   2895
   End
   Begin VB.CommandButton cmdLogOut 
      Caption         =   "Log Out"
      Height          =   855
      Left            =   720
      TabIndex        =   4
      Top             =   7800
      Width           =   2895
   End
   Begin VB.CommandButton cmdCheckOut 
      Caption         =   "Check Out "
      Height          =   855
      Left            =   8280
      TabIndex        =   3
      Top             =   7800
      Width           =   2895
   End
   Begin VB.CommandButton cmdGrocery 
      Caption         =   "Grocery"
      Height          =   1335
      Left            =   720
      TabIndex        =   2
      Top             =   6000
      Width           =   2895
   End
   Begin VB.CommandButton cmdProduce 
      Caption         =   "Produce"
      Height          =   1335
      Left            =   720
      TabIndex        =   1
      Top             =   4320
      Width           =   2895
   End
   Begin VB.CommandButton cmdDeli 
      Caption         =   "Deli and Prepared Foods"
      Height          =   1455
      Left            =   720
      TabIndex        =   0
      Top             =   2520
      Width           =   2895
   End
   Begin VB.Image imgStoreFront 
      Height          =   9105
      Left            =   0
      Picture         =   "frmGroceryStore.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12735
   End
End
Attribute VB_Name = "frmGroceryStore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'James Delivers, CSCI 130 Visual Basic Project
'frmGroceryStore
'Written by James Garay Heelan
'on 11-2-06
'The purpose of this page is to provide the user with a central menu by which
'to navigate the online grocery store.  From here, the user may visit any
'department, log out of the program, or proceed to the checkout.

Option Explicit

Private Sub cmdCheckOut_Click()
    frmGroceryStore.Hide 'hides the central menu
    frmCheckOut.Show 'displays the checkout page to the user
End Sub

Private Sub cmdCustomerService_Click()
    frmGroceryStore.Hide 'the central menu is hidden
    frmCustomerService.Show 'the customer service form is displayed for the user
End Sub

Private Sub cmdDeli_Click()
    frmGroceryStore.Hide 'hides the central menu
    frmDeli.Show 'diplays the deli page to the user
End Sub

Private Sub cmdGrocery_Click()
    frmGroceryStore.Hide 'hides the central menu
    frmGrocery.Show 'displays the grocery page to the user
End Sub

Private Sub cmdLogOut_Click()
    End 'exits the program
End Sub

Private Sub cmdProduce_Click()
    frmGroceryStore.Hide 'hides the central menu
    frmProduce.Show 'displays the produce form to the user
End Sub
