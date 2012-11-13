VERSION 5.00
Begin VB.Form frmGrocery 
   Caption         =   "Grocery"
   ClientHeight    =   7485
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10200
   LinkTopic       =   "Form1"
   ScaleHeight     =   7485
   ScaleWidth      =   10200
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCustomerService 
      Caption         =   "Take me to customer service, I want to complain!!!"
      Height          =   1335
      Left            =   3480
      TabIndex        =   2
      Top             =   4080
      Width           =   3135
   End
   Begin VB.CommandButton cmdLogOut 
      Caption         =   "Log Out"
      Height          =   855
      Left            =   5520
      TabIndex        =   1
      Top             =   5880
      Width           =   2775
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Go to a Different Aisle"
      Height          =   855
      Left            =   1680
      TabIndex        =   0
      Top             =   5880
      Width           =   3015
   End
   Begin VB.Image imgGrocery 
      Height          =   7515
      Left            =   0
      Picture         =   "frmGrocery.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10215
   End
End
Attribute VB_Name = "frmGrocery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'James Delivers, CSCI 130 Visual Basic Project
'frmGrocery
'Written by James Garay Heelan
'on 11-2-06
'The purpose of this page is to allow the user to shop for groceries, though
'in its current state, they cannot do so.  Thus, is allows the user to navigate
'to either the customer service page or back to the central shopping menu.

Option Explicit

Private Sub cmdBack_Click()
    frmGrocery.Hide 'hide the current page
    frmGroceryStore.Show 'display the central shopping menu to the user
End Sub

Private Sub cmdCustomerService_Click()
    frmGrocery.Hide 'hide the grocery page
    frmCustomerService.Show 'display the customer service form to the user
End Sub

Private Sub cmdLogOut_Click()
    End 'exit the program
End Sub
