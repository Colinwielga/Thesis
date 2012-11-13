VERSION 5.00
Begin VB.Form Purchase 
   BackColor       =   &H00FFFFC0&
   Caption         =   "Form1"
   ClientHeight    =   5775
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9825
   LinkTopic       =   "Form1"
   ScaleHeight     =   5775
   ScaleWidth      =   9825
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00FFFF80&
      Caption         =   "Clear Order"
      Height          =   375
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton cmdSubmit 
      BackColor       =   &H00FFFF80&
      Caption         =   "Submit Order"
      Enabled         =   0   'False
      Height          =   375
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   4680
      Width           =   1575
   End
   Begin VB.CommandButton cmdOrder 
      BackColor       =   &H00FFFF80&
      Caption         =   "Create Order"
      Height          =   495
      Left            =   8040
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   240
      Width           =   1575
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00FFFF80&
      Caption         =   "Back to Title Page"
      Height          =   495
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   5160
      Width           =   1575
   End
   Begin VB.PictureBox picBill 
      BackColor       =   &H00FFFF80&
      Height          =   4935
      Left            =   3720
      ScaleHeight     =   4875
      ScaleWidth      =   4155
      TabIndex        =   16
      Top             =   240
      Width           =   4215
   End
   Begin VB.TextBox txtTruth 
      Height          =   285
      Left            =   2760
      TabIndex        =   15
      Text            =   "0"
      Top             =   3960
      Width           =   495
   End
   Begin VB.TextBox txtPirates 
      Height          =   285
      Left            =   2760
      TabIndex        =   13
      Text            =   "0"
      Top             =   3360
      Width           =   495
   End
   Begin VB.TextBox txtGeorge 
      Height          =   285
      Left            =   2760
      TabIndex        =   11
      Text            =   "0"
      Top             =   2880
      Width           =   495
   End
   Begin VB.TextBox txtCode 
      Height          =   285
      Left            =   2760
      TabIndex        =   9
      Text            =   "0"
      Top             =   2400
      Width           =   495
   End
   Begin VB.TextBox txtClick 
      Height          =   285
      Left            =   2760
      TabIndex        =   7
      Text            =   "0"
      Top             =   1920
      Width           =   495
   End
   Begin VB.TextBox txtCars 
      Height          =   285
      Left            =   2760
      TabIndex        =   5
      Text            =   "0"
      Top             =   1440
      Width           =   495
   End
   Begin VB.TextBox txtBreakup 
      Height          =   285
      Left            =   2760
      TabIndex        =   3
      Text            =   "0"
      Top             =   960
      Width           =   495
   End
   Begin VB.TextBox txtLake 
      Height          =   285
      Left            =   2760
      TabIndex        =   1
      Text            =   "0"
      Top             =   480
      Width           =   495
   End
   Begin VB.Label lblP8 
      BackColor       =   &H00FFFFC0&
      Caption         =   "$14.95"
      Height          =   375
      Left            =   1800
      TabIndex        =   28
      Top             =   3960
      Width           =   735
   End
   Begin VB.Label lblP7 
      BackColor       =   &H00FFFFC0&
      Caption         =   "$19.95"
      Height          =   255
      Left            =   1800
      TabIndex        =   27
      Top             =   3360
      Width           =   855
   End
   Begin VB.Label lblP6 
      BackColor       =   &H00FFFFC0&
      Caption         =   "$14.95"
      Height          =   375
      Left            =   1800
      TabIndex        =   26
      Top             =   2880
      Width           =   855
   End
   Begin VB.Label lblP5 
      BackColor       =   &H00FFFFC0&
      Caption         =   "$19.95"
      Height          =   255
      Left            =   1800
      TabIndex        =   25
      Top             =   2400
      Width           =   855
   End
   Begin VB.Label lblP4 
      BackColor       =   &H00FFFFC0&
      Caption         =   "$14.95"
      Height          =   375
      Left            =   1800
      TabIndex        =   24
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label lblP3 
      BackColor       =   &H00FFFFC0&
      Caption         =   "$21.95"
      Height          =   375
      Left            =   1800
      TabIndex        =   23
      Top             =   1440
      Width           =   615
   End
   Begin VB.Label lblP2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "$19.95"
      Height          =   255
      Left            =   1800
      TabIndex        =   22
      Top             =   960
      Width           =   615
   End
   Begin VB.Label lblP1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "$14.95"
      Height          =   255
      Left            =   1800
      TabIndex        =   21
      Top             =   480
      Width           =   495
   End
   Begin VB.Label lblPrice 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Price"
      Height          =   255
      Left            =   1800
      TabIndex        =   20
      Top             =   120
      Width           =   495
   End
   Begin VB.Label lblQuantity 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Quantity"
      Height          =   255
      Left            =   2640
      TabIndex        =   19
      Top             =   120
      Width           =   735
   End
   Begin VB.Label lblTruth 
      BackColor       =   &H00FFFFC0&
      Caption         =   "An inconvienent Truth"
      Height          =   495
      Left            =   240
      TabIndex        =   14
      Top             =   3960
      Width           =   1335
   End
   Begin VB.Label lblPirates 
      BackColor       =   &H00FFFFC0&
      Caption         =   "The Pirates of the Carabbean"
      Height          =   375
      Left            =   240
      TabIndex        =   12
      Top             =   3360
      Width           =   1335
   End
   Begin VB.Label lblGeorge 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Curious George"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   2880
      Width           =   1335
   End
   Begin VB.Label lblCode 
      BackColor       =   &H00FFFFC0&
      Caption         =   "The Davinci Code"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Label lblClick 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Click"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   1920
      Width           =   495
   End
   Begin VB.Label lblCars 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Cars"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label lblbreakup 
      BackColor       =   &H00FFFFC0&
      Caption         =   "The Breakup"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   975
   End
   Begin VB.Label lbllakehouse 
      BackColor       =   &H00FFFFC0&
      Caption         =   "The Lake House"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   1335
   End
End
Attribute VB_Name = "Purchase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Movies
'Form name: Purchase
'Author: Katie Hanson
'Date Written: Nov 1 2006
'Objective: This form allows the user to purchase the movies described on the startup form.
Option Explicit
'Back to startup form
Private Sub cmdBack_Click()
    Title.Show
    Purchase.Hide
End Sub
'clear the order
Private Sub cmdClear_Click()
    picBill.Cls
End Sub
'calculate and print out the users order
Private Sub cmdOrder_Click()
Open App.Path & "\price.txt" For Input As #1 'open all the prices of the movies
Dim Price(1 To 8) As Double
Dim Q1 As Integer, Q2 As Integer, Q3 As Integer, Q4 As Integer, Q5 As Integer, Q6 As Integer, Q7 As Integer, Q8 As Integer
Dim Sum As Double, Tax As Double, Total As Double
Dim ctr As Integer
ctr = 0
Do While Not EOF(1) 'place prices of the movies into an array
     ctr = ctr + 1
     Input #1, Price(ctr)
Loop
Close #1
Q1 = txtLake.Text 'creating variables for the quantity of the movies being purchased
Q2 = txtBreakup.Text
Q3 = txtCars.Text
Q4 = txtClick.Text
Q5 = txtCode.Text
Q6 = txtGeorge.Text
Q7 = txtPirates.Text
Q8 = txtTruth.Text
Sum = (Price(1) * Q1) + (Price(2) * Q2) + (Price(3) * Q3) + (Price(4) * Q4) + (Price(5) * Q5) + (Price(6) * Q6) + (Price(7) * Q7) + (Price(8) * Q8) 'Calculating the total price
Tax = Sum * 0.35
picBill.Print , , "Bill" 'printing out the bill
picBill.Print
If Q1 > 0 Then 'Printing the name, quantity, and price of every movie ordered
    picBill.Print "The Lake House", Q1; " at "; FormatCurrency(Price(1))
End If
If Q2 > 0 Then
    picBill.Print "The Breakup", , Q2; " at "; FormatCurrency(Price(2))
End If
If Q3 > 0 Then
    picBill.Print "Cars", , Q3; " at "; FormatCurrency(Price(3))
End If
If Q4 > 0 Then
    picBill.Print "Click", , Q4; " at "; FormatCurrency(Price(4))
End If
If Q5 > 0 Then
    picBill.Print "The Davinci Code", Q5; " at "; FormatCurrency(Price(5))
End If
If Q6 > 0 Then
    picBill.Print "Curious George", Q6; " at "; FormatCurrency(Price(6))
End If
If Q7 > 0 Then
    picBill.Print "The Pirates of the Carabbean"; Q7; " at "; FormatCurrency(Price(7))
End If
If Q8 > 0 Then
    picBill.Print "An Inconvienent Truth", Q8; " at "; FormatCurrency(Price(8))
End If
picBill.Print
picBill.Print , "Subtotal", FormatCurrency(Sum)
picBill.Print , "Tax", FormatCurrency(Tax)
picBill.Print , "__________________"
picBill.Print , "Total", FormatCurrency(Sum + Tax)
cmdSubmit.Enabled = True 'enabling the Submit button
End Sub
'finalizing the order and recieving the users name and address
Private Sub cmdSubmit_Click()
    InputBox "Please enter the name of the buyer"
    InputBox "Please enter the address to which you would like your purchase to be mailed"
    MsgBox "Thank you for your purchase, it will be mailed in two to three business days"
End Sub
