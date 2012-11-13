VERSION 5.00
Begin VB.Form frmCustomerService 
   Caption         =   "Customer Service"
   ClientHeight    =   7815
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9810
   LinkTopic       =   "Form1"
   ScaleHeight     =   7815
   ScaleWidth      =   9810
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picResults 
      Height          =   4935
      Left            =   4560
      ScaleHeight     =   4875
      ScaleWidth      =   4275
      TabIndex        =   3
      Top             =   1320
      Width           =   4335
   End
   Begin VB.CommandButton cmdLogOut 
      Caption         =   "Log Out"
      Height          =   1215
      Left            =   1080
      TabIndex        =   2
      Top             =   4920
      Width           =   2775
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Continue Shopping"
      Height          =   1215
      Left            =   1080
      TabIndex        =   1
      Top             =   3120
      Width           =   2775
   End
   Begin VB.CommandButton cmdComplain 
      Caption         =   "Complain"
      Height          =   1215
      Left            =   1080
      TabIndex        =   0
      Top             =   1320
      Width           =   2775
   End
   Begin VB.Image imgCustomerService 
      Height          =   7815
      Left            =   0
      Picture         =   "frmCustomerService.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9855
   End
End
Attribute VB_Name = "frmCustomerService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'James Delivers, CSCI 130 Visual Basic Project
'frmCustomerService
'Written by James Garay Heelan
'on 11-2-06
'this page is meant to allow the user to file a complaint with the company
'and then receive a response from a customer service representative.

Option Explicit

Private Sub cmdBack_Click()
    frmCustomerService.Hide 'hides the customer service form and
    frmGroceryStore.Show 'displays the central menu to the user
End Sub

Private Sub cmdComplain_Click()
Dim Complaint As String, Compensation As String, LastWords As String

    picResults.Cls 'clears the picture box of its contents
    
    Complaint = InputBox("What is your complaint? (5 words or less, please)", "Complaint") 'the user enters via inputbox his or her complaint
    Compensation = InputBox("How can we make it up to you?", "Compensation") 'the user enters via inputbox his or her preferred compensation
    LastWords = InputBox("Do you have a final comment?", "Insult?") 'the user enters via inputbox his or her final comment or insult
    
    picResults.Print Names(PurchaseCode); "," 'prints the name of the user
    picResults.Print "we at James Delivers competely understand" 'delivers the response of the customer service representative
    picResults.Print "that you find it frustrating when"
    picResults.Print Complaint 'repeats the customer's complaing
    picResults.Print "and we apologize."
    picResults.Print "However, I spoke with my supervisor and"
    picResults.Print "we are unable to"
    picResults.Print Compensation 'repeats the customer's desired compensation
    picResults.Print "because of company policy."
    picResults.Print " "
    picResults.Print "What was that?"
    picResults.Print "Well, you're a " & LastWords & ", too." 'insults the user with his or her original insult or final comment
    picResults.Print " "
    picResults.Print "Thank you for shopping James Delivers."
    picResults.Print "We hope to deliver to you again, soon!"

End Sub

Private Sub cmdLogOut_Click()
    End 'exits the program
End Sub
