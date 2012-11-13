VERSION 5.00
Begin VB.Form Check 
   BackColor       =   &H80000010&
   Caption         =   "Customer Details"
   ClientHeight    =   3030
   ClientLeft      =   4920
   ClientTop       =   6135
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4680
   Begin VB.CommandButton cmdProcess 
      BackColor       =   &H008080FF&
      Caption         =   "Process"
      Height          =   495
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2400
      Width           =   2295
   End
   Begin VB.TextBox txtThree 
      Height          =   495
      Left            =   2760
      TabIndex        =   5
      Top             =   1680
      Width           =   1695
   End
   Begin VB.TextBox txtTwo 
      Height          =   495
      Left            =   2760
      TabIndex        =   4
      Top             =   960
      Width           =   1695
   End
   Begin VB.TextBox txtOne 
      Height          =   495
      Left            =   2760
      TabIndex        =   3
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label lblAcctNum 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Enter Account Number:"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   1800
      Width           =   2175
   End
   Begin VB.Label lblRouteNum 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Enter Bank Routing Number:"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   1080
      Width           =   2175
   End
   Begin VB.Label lblCheckNum 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Enter Check Number:"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   1695
   End
End
Attribute VB_Name = "Check"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdProcess_Click()
                
        If txtOne.Text = "" Then
            MsgBox "Please Enter All Required Info", , "Error"
        ElseIf txtTwo.Text = "" Then
            MsgBox "Please Enter All Required Info", , "Error"
        ElseIf txtThree.Text = "" Then
            MsgBox "Please Enter All Required Info", , "Error"
        Else
            MsgBox "Your Payment Has Been Processed.  Thank You!", , "Thank You"
        Check.Visible = False
        TransactionDetails.Visible = False
        PointOfSale.Visible = True
        Status.Visible = False
        End If
                          
End Sub



'RetailPOSandInventoryControl program; Check form
'this code was written on Tuesday, October 31, 2006
'this code was edited on Wednesday, November 1, 2006
'this code was revised on Thursday, November 2, 2006
'written by Mark Collette
'the purpose of this form is to take in customer info via bank check to process a transaction
'the subroutines took in the info, and displayed a message box to note the info had been received and thank the customer for their business
