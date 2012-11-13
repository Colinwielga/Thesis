VERSION 5.00
Begin VB.Form frmSpecial 
   Caption         =   "Form1"
   ClientHeight    =   8475
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11400
   LinkTopic       =   "Form1"
   Picture         =   "frmSpecial.frx":0000
   ScaleHeight     =   8475
   ScaleWidth      =   11400
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   735
      Left            =   9720
      TabIndex        =   6
      Top             =   6480
      Width           =   975
   End
   Begin VB.CommandButton cmdMain 
      Caption         =   "Main"
      Height          =   735
      Left            =   9720
      TabIndex        =   5
      Top             =   5520
      Width           =   975
   End
   Begin VB.CommandButton cmdOrder 
      Caption         =   "Order"
      Height          =   495
      Left            =   9240
      TabIndex        =   4
      Top             =   2760
      Width           =   1455
   End
   Begin VB.PictureBox picResults 
      Height          =   1455
      Left            =   7080
      ScaleHeight     =   1395
      ScaleWidth      =   3555
      TabIndex        =   3
      Top             =   3480
      Width           =   3615
   End
   Begin VB.CommandButton cmdMenu 
      Caption         =   "Check out the specials here!"
      Height          =   495
      Left            =   7080
      TabIndex        =   2
      Top             =   2760
      Width           =   1455
   End
   Begin VB.TextBox txtSpecial 
      Height          =   1095
      Left            =   7080
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   1200
      Width           =   3615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      Caption         =   "What would you like?"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1920
      TabIndex        =   1
      Top             =   120
      Width           =   7335
   End
End
Attribute VB_Name = "frmSpecial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Digital Menu
'Form Name: frmSpecial
'Authors: Gaole Chen
'Date Written: 3/22/09
'Objective: The user can order specials from this form by using a scrollbar.


Option Explicit

Private Sub cmdMain_Click()
frmSpecial.Hide
frmWelcome.Show
End Sub

Private Sub cmdMenu_Click()
'this button reads the special menu to the user, including a scrollbar
'declare the variables
Dim Special As String
Special = "1) Eggrolls, Spicy sauce on beef, White rice--$30" & vbCrLf & "2) Chicken Fingers, Spareribs, Tenyaki Beef--$28" & vbCrLf & "3) Chicken Wings, Crab Rangoon, Peking Ravioli--$35" & vbCrLf & "4) Fried shrimps, Scalion Pancake, Vegetable Eggrolls--$24" & vbCrLf & "5) Beef with Broccoli, Pine Nuts Baby Shrimp, Spicy noodles--$32"
txtSpecial.Text = Special
End Sub

Private Sub cmdOrder_Click()
'The user orders here after checking out the menu
'declare the variables
Dim Order As Integer, Total As Single, Taxrate As Single, Tax As Single, runningTotal As Single
'initialize with total =  0
Total = 0
'Order with an inputbox
Order = InputBox("Please enter the number of special(1-5) you would like. Enter 0 to indicate the end of ordering.")
picResults.Print "You would like:"
Do While Order <> 0
    Select Case Order
        Case 1
        picResults.Print "Special #1--$30."
        Total = Total + 30
        Case 2
        picResults.Print "Special #2--$28."
        Total = Total + 28
        Case 3
        picResults.Print "Special #3--$35."
        Total = Total + 35
        Case 4
        picResults.Print "Special #4--$24."
        Total = Total + 24
        Case 5
        picResults.Print "Special #5--$32."
        Total = Total + 32
        Case Else
        MsgBox "Please enter a valid number."
    End Select
Order = InputBox("Please enter the number of special(1-5) you would like. Enter 0 to indicate the end of ordering.")
Loop
Taxrate = 0.07
Tax = Total * Taxrate
runningTotal = Total + Tax
picResults.Print "-----------------------------------------------------------------------------------------------------------------------"
picResults.Print "Tax: ", FormatNumber(FormatCurrency(Tax), 2)
picResults.Print "Total: ", FormatNumber(FormatCurrency(runningTotal), 2)
End Sub

Private Sub cmdQuit_Click()
End
End Sub
