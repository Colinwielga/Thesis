VERSION 5.00
Begin VB.Form GiftCard 
   BackColor       =   &H80000010&
   Caption         =   "Customer Details"
   ClientHeight    =   3090
   ClientLeft      =   5385
   ClientTop       =   5820
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   Begin VB.CommandButton cmdProcess 
      BackColor       =   &H008080FF&
      Caption         =   "Process Transaction"
      Height          =   615
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1920
      Width           =   2415
   End
   Begin VB.TextBox txtOne 
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Top             =   960
      Width           =   2175
   End
   Begin VB.Label lblGiftCard 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Enter Gift Card Number"
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   1815
   End
End
Attribute VB_Name = "GiftCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdProcess_Click()
                
        If txtOne.Text = "" Then
            MsgBox "Please Enter All Required Info", , "Error"
        Else
            MsgBox "Your Payment Has Been Processed.  Thank You!", , "Thank You"
        GiftCard.Visible = False
        TransactionDetails.Visible = False
        PointOfSale.Visible = True
        Status.Visible = False
        End If
                
End Sub



'RetailPOSandInventoryControl program; GiftCard form
'this code was written on Tuesday, October 31, 2006
'this code was edited on Wednesday, November 1, 2006
'this code was revised on Thursday, November 2, 2006
'written by Mark Collette
'the purpose of this form is to take in customer info via giftcard to process a transaction
'the subroutines took in the info, and displayed a message box to note the info had been received and thank the customer for their business

