VERSION 5.00
Begin VB.Form frmstore 
   BackColor       =   &H000000FF&
   Caption         =   "BeerBall Store"
   ClientHeight    =   11520
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15045
   LinkTopic       =   "Form1"
   Picture         =   "frmstore.frx":0000
   ScaleHeight     =   11520
   ScaleWidth      =   15045
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdreturn 
      Caption         =   "Return to Main Menu"
      Height          =   1695
      Left            =   11880
      TabIndex        =   8
      Top             =   840
      Width           =   2655
   End
   Begin VB.CommandButton cmdtotal 
      Caption         =   "Checkout and See Total "
      Height          =   1815
      Left            =   11760
      TabIndex        =   7
      Top             =   3120
      Width           =   2895
   End
   Begin VB.CommandButton cmdcozy 
      Caption         =   "Add A BeerBall Beer Cozy"
      Height          =   975
      Left            =   10800
      TabIndex        =   6
      Top             =   8400
      Width           =   2295
   End
   Begin VB.CommandButton cmdhat 
      Caption         =   "Add a BeerBall Hat"
      Height          =   855
      Left            =   5160
      TabIndex        =   4
      Top             =   9480
      Width           =   1575
   End
   Begin VB.PictureBox picresults 
      Height          =   6615
      Left            =   7320
      ScaleHeight     =   6555
      ScaleWidth      =   4035
      TabIndex        =   1
      Top             =   120
      Width           =   4095
   End
   Begin VB.CommandButton cmdpinpong 
      Caption         =   "Add a BeerBall Ping Pong Ball"
      Height          =   855
      Left            =   5160
      TabIndex        =   0
      Top             =   6120
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackColor       =   &H000000FF&
      Caption         =   "Official BeerBall Cozy                          $ 2.95 each"
      Height          =   495
      Left            =   10920
      TabIndex        =   5
      Top             =   7560
      Width           =   2055
   End
   Begin VB.Image Image3 
      Height          =   4500
      Left            =   7680
      Picture         =   "frmstore.frx":10D2A
      Top             =   6960
      Width           =   3000
   End
   Begin VB.Label Label2 
      BackColor       =   &H000000FF&
      Caption         =   "Official BeerBall Hat              $ 15.99 each"
      Height          =   735
      Left            =   5160
      TabIndex        =   3
      Top             =   8400
      Width           =   2055
   End
   Begin VB.Image Image2 
      Height          =   4500
      Left            =   0
      Picture         =   "frmstore.frx":136C9
      Top             =   7320
      Width           =   4500
   End
   Begin VB.Label Label1 
      BackColor       =   &H000000FF&
      Caption         =   " Offical BeerBall Ping Pong Balls          $2.99/ 6"
      Height          =   495
      Left            =   4920
      TabIndex        =   2
      Top             =   5280
      Width           =   2415
   End
   Begin VB.Image Image1 
      Height          =   4860
      Left            =   0
      Picture         =   "frmstore.frx":16DC8
      Top             =   4800
      Width           =   4860
   End
End
Attribute VB_Name = "frmstore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'this form uses addition and other math functions to allow the user to purchase BeerBall Apparel
Dim sum As Single

Private Sub cmdcozy_Click()
'this subroutine adds a beer cozy to the users bill
Dim cozy As Integer
sum = sum + 2.95
picresults.Print "BeerBall Cozy"; Tab(45); "$2.95"
MsgBox "Thank you a Official BeerBall Cozy has been added to your total"
End Sub

Private Sub cmdhat_Click()
'this subroutine adds a hat to the users total
Dim hat As Integer
sum = sum + 15.99
picresults.Print "BeerBall Hat"; Tab(45); "$15.99"
MsgBox "Thank you a Official BeerBall Hat has been added to your total"
End Sub

Private Sub cmdpinpong_Click()
'this subroutine adds a pack of ping pong balls to the users total
Dim balls As Integer
sum = sum + 2.99
picresults.Print "BeerBall Ping-Pong Ball"; Tab(45); "$2.99"
MsgBox "Thank you a Official Beer Ball has been added to your total"

End Sub

Private Sub cmdreturn_Click()
'this subroutine goes back to the main menu
frmmain.Show
frmstore.Hide

End Sub

Private Sub cmdtotal_Click()
'this subroutine displays the total and tax in a picture box and sends the grand total to the user in a message box
Dim tax As Single
Dim grandtotal As Single

tax = sum * 0.065
grandtotal = sum + tax
picresults.Print "--------------------------------------------------------------"
picresults.Print " Your Subtotal is "; Tab(45); FormatCurrency(sum)
picresults.Print " Your Tax at 6.5% is "; Tab(45); FormatCurrency(tax)
picresults.Print "***************************************************************"
MsgBox " Congratulations " & username & " Your Grand Total is " & FormatCurrency(grandtotal)

End Sub
