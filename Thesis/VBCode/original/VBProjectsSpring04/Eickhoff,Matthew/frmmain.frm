VERSION 5.00
Begin VB.Form frmmain 
   BackColor       =   &H000000FF&
   Caption         =   "Form1"
   ClientHeight    =   6945
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11115
   LinkTopic       =   "Form1"
   ScaleHeight     =   6945
   ScaleWidth      =   11115
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdswitchfrm 
      Caption         =   "Graphs"
      Height          =   975
      Left            =   6600
      TabIndex        =   7
      Top             =   3120
      Width           =   1455
   End
   Begin VB.CommandButton cmdsearch 
      Caption         =   "Search for a Stock"
      Enabled         =   0   'False
      Height          =   975
      Left            =   9480
      TabIndex        =   6
      Top             =   3120
      Width           =   1455
   End
   Begin VB.CommandButton cmdinv 
      Caption         =   "Good Investments"
      Enabled         =   0   'False
      Height          =   975
      Left            =   9480
      TabIndex        =   5
      Top             =   1680
      Width           =   1455
   End
   Begin VB.CommandButton cmdpe 
      Caption         =   "Stocks with Price to Earnings higher than 1.0"
      Enabled         =   0   'False
      Height          =   975
      Left            =   9480
      TabIndex        =   4
      Top             =   240
      Width           =   1455
   End
   Begin VB.CommandButton cmdprintalpha 
      Caption         =   "Print in Alphabetical Order"
      Enabled         =   0   'False
      Height          =   975
      Left            =   6600
      TabIndex        =   3
      Top             =   1680
      Width           =   1455
   End
   Begin VB.CommandButton cmdread 
      BackColor       =   &H80000017&
      Caption         =   "Read"
      Height          =   975
      Left            =   6600
      MaskColor       =   &H8000000F&
      TabIndex        =   2
      Top             =   240
      Width           =   1455
   End
   Begin VB.CommandButton cmdquit 
      Caption         =   "Quit"
      Height          =   975
      Left            =   9480
      TabIndex        =   1
      Top             =   5760
      Width           =   1455
   End
   Begin VB.PictureBox picresults 
      BackColor       =   &H00FFFFFF&
      Height          =   6615
      Left            =   240
      ScaleHeight     =   6555
      ScaleWidth      =   5835
      TabIndex        =   0
      Top             =   120
      Width           =   5895
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Stocks
'written by Mat Eickhoff on 12 March 2004
'this program will read a text file which has information placed in it including:
'the name of the stock, price, sales, earnings and book value
'with that information it can calculate different ratios and sort and display those that
'would make a good investment.
Option Explicit
Dim names(1 To 21) As String, price(1 To 25) As Single, earnings(1 To 25) As Single, book(1 To 25) As Single
Dim sales(1 To 25) As Single, ctr As Integer
Dim path As String
Dim tempnames As String, tempprice As Single, tempearnings As Single, tempbook As Single, tempsales As Single
Dim pass As Integer, listtotal As Integer
Dim peratio(1 To 25) As Single, pbookratio(1 To 25) As Single, score(1 To 25) As Single
Dim found As Boolean, position As Integer, userinput As String

Private Sub cmdinv_Click()
ctr = listtotal
picresults.Cls
picresults.Print "Name of Stock"; Tab(20); "Investment Type"; Tab(40); "Value of Stock"
picresults.Print "*********************************************************************"
For ctr = 1 To listtotal
pbookratio(ctr) = price(ctr) / book(ctr)
picresults.Print names(ctr);
If pbookratio(ctr) > 1.5 Then
score(ctr) = 5.5
picresults.Print Tab(20); "Great Buy";
ElseIf pbookratio(ctr) >= 1.25 Then
score(ctr) = 4.5
picresults.Print Tab(20); "Good Buy";
score(ctr) = 3.5
ElseIf pbookratio(ctr) >= 1 Then
score(ctr) = 2.5
picresults.Print Tab(20); "Average Buy";
ElseIf pbookratio(ctr) >= 0.75 Then
score(ctr) = 1.5
picresults.Print Tab(20); "Poor Buy";
Else
score(ctr) = 0.5
picresults.Print Tab(20); "Awful Buy";
End If
Select Case score(ctr)
Case Is > 4
picresults.Print Tab(40); "Undervalued"
Case 2 To 3
picresults.Print Tab(40); "Valued Correctly"
Case Is < 2
picresults.Print Tab(40); "Overvalued"
Case Else
picresults.Print Tab(40); "Who knows what it's valued at."
End Select
Next ctr
End Sub

Private Sub cmdpe_Click()
picresults.Cls
picresults.Print "Name of Stock"; Tab(20); "Price To Earnings Ratio"
picresults.Print "****************************************************"
For ctr = 1 To listtotal
peratio(ctr) = price(ctr) / earnings(ctr)
If peratio(ctr) > 1 Then
picresults.Print names(ctr); Tab(20); FormatNumber(peratio(ctr), 2)
End If
Next ctr
End Sub

Private Sub cmdprintalpha_Click()
picresults.Cls
ctr = 21
listtotal = ctr
For pass = 1 To listtotal - 1
For ctr = 1 To listtotal - pass
If names(ctr) > names(ctr + 1) Then
tempnames = names(ctr)
names(ctr) = names(ctr + 1)
names(ctr + 1) = tempnames
tempprice = price(ctr)
price(ctr) = price(ctr + 1)
price(ctr + 1) = tempprice
tempearnings = earnings(ctr)
earnings(ctr) = earnings(ctr + 1)
earnings(ctr + 1) = tempearnings
tempbook = book(ctr)
book(ctr) = book(ctr + 1)
book(ctr + 1) = tempbook
tempsales = sales(ctr)
sales(ctr) = sales(ctr + 1)
sales(ctr + 1) = tempsales
End If
Next ctr
Next pass
picresults.Print "Name of Stock"; Tab(20); "Price"; Tab(30); "EPS"; Tab(45); "Book Value"; Tab(60); "Sales Per Share"
picresults.Print "**************************************************************************************************************************************************************************************************************************************************************"
For ctr = 1 To listtotal
picresults.Print names(ctr); Tab(20); price(ctr); Tab(30); earnings(ctr); Tab(45); book(ctr); Tab(60); sales(ctr)
Next ctr
cmdinv.Enabled = True
cmdpe.Enabled = True
cmdsearch.Enabled = True
cmdread.Enabled = False
cmdprintalpha.Enabled = True
End Sub

Private Sub cmdquit_Click()
End
End Sub

Private Sub cmdread_Click()
Open path & "stocks.txt" For Input As #1
Do While Not EOF(1)
ctr = ctr + 1
Input #1, names(ctr), price(ctr), earnings(ctr), book(ctr), sales(ctr)
Loop
cmdprintalpha.Enabled = True
cmdread.Enabled = False
listtotal = ctr
End Sub


Private Sub cmdsearch_Click()
picresults.Cls
ctr = 21
found = False
userinput = InputBox("Please enter the name of the stock you wish to search for...")
If userinput = "" Then
Do While userinput = ""
MsgBox "Sorry, but you must enter a name of a stock", , "Error"
userinput = InputBox("Please enter the name of the stock you wish to search for...")
Loop
End If
position = 0
Do While (Not found) And (position < ctr)
position = position + 1
If names(position) = userinput Then
picresults.Print "The stock is in the list in position"; position
found = True
End If
Loop
If found = False Then
picresults.Print "Sorry the stock was not in the list."
End If
End Sub

Private Sub cmdswitchfrm_Click()
frmmain.Hide
graph.Show
End Sub

Private Sub Form_Load()
path = "N:\CS130\handin\Eickhoff, Matthew\"
End Sub
