VERSION 5.00
Begin VB.Form frmFinishCheckout 
   BackColor       =   &H000040C0&
   Caption         =   "Finish checkout"
   ClientHeight    =   7860
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10650
   LinkTopic       =   "Form1"
   ScaleHeight     =   7860
   ScaleWidth      =   10650
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdTotal 
      BackColor       =   &H0000C000&
      Caption         =   "View Billing"
      Height          =   735
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2880
      Width           =   2775
   End
   Begin VB.CommandButton cmdEnd 
      BackColor       =   &H0000C000&
      Caption         =   "Goodbye and thanks for shopping with Amazon.com!"
      Height          =   855
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5520
      Width           =   2895
   End
   Begin VB.CommandButton cmdFinish 
      BackColor       =   &H0000C000&
      Caption         =   "Finish checkout"
      Height          =   855
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4200
      Width           =   2895
   End
   Begin VB.CommandButton cmdAverage 
      BackColor       =   &H0000C000&
      Caption         =   "Find your average Amazon.com ranking"
      Height          =   975
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1560
      Width           =   3015
   End
   Begin VB.PictureBox Out 
      BackColor       =   &H0000C000&
      Height          =   6255
      Left            =   480
      ScaleHeight     =   6195
      ScaleWidth      =   5595
      TabIndex        =   0
      Top             =   720
      Width           =   5655
   End
End
Attribute VB_Name = "frmFinishCheckout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : ProjAmazonCDPurchase (Erin Zrust's VB Project.vpb)
'Form Name : frmFinishCheckout (finish_checkout.frm)
'Author: Erin Zrust
'Date Written: October 30, 2003
'Purpose of Form: To purchase various CDS, find the total
                  'including shipping and tax, the Amazon.com
                  'ranking for each CD purchase and the average
                  'of all CDs purchased.

'This forces users to declare variables
Option Explicit
Private Sub cmdAverage_Click()
'Find average of all Amazon.com rankings
Average = 0
'Adds all Amazon.com rankings together
'then divides by 4
Average = DaveMatthewsBandRanking(D) + BenHarperRanking(B) + OARRanking(A) + JackJohnsonRanking(J)
Average = Average / 4
    Out.Print "Your average Amazon ranking is"; Average
'Show "View Billing" button
cmdTotal.Enabled = True
cmdAverage.Enabled = False
End Sub

Private Sub cmdEnd_Click()
End
End Sub

Private Sub cmdFinish_Click()
MsgBox "Your order is completed and will be shipped in the next seven business days", , "Thank You"
'Hide "Finish checkout" button
cmdFinish.Enabled = False

End Sub

Private Sub cmdTotal_Click()
'Clear screen
Out.Cls
'find the combined price without tax or shipping
Subtotal = DaveMatthewsBandPrice(D) + BenHarperPrice(B) + OARPrice(A) + JackJohnsonPrice(J)
'Find the shipping rate
If Subtotal >= 50 Then
    Shipping = 0
If Subtotal >= 25 And Subtotal < 50 Then
    Shipping = 5.99
ElseIf Subtotal < 25 Then
    Shipping = 4.99
End If

'Find the total
Total = Subtotal + Shipping
    Out.Print "Subtotal", FormatCurrency(Subtotal)
    Out.Print "Shipping", FormatCurrency(Shipping)
    Out.Print "Total", FormatCurrency(Total)
'Show "Finish checkout" and
'"Goodbye and thanks for shopping" buttons
cmdFinish.Enabled = True
cmdEnd.Enabled = True
'Hide "Find your average Amazon.com ranking"
'and "View Billing Buttons"
cmdAverage.Enabled = False
cmdTotal.Enabled = False
End If
End Sub

Private Sub Form_Load()
'Hide "Finish checkout", "View Billing" and
'"Goodbye and thanks for shopping" buttons
'The other commands must be completed before
'finishing the program
cmdFinish.Enabled = False
cmdEnd.Enabled = False
cmdTotal.Enabled = False
End Sub
