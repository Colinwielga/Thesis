VERSION 5.00
Begin VB.Form frmDeli 
   Caption         =   "Deli and Prepared Foods"
   ClientHeight    =   11010
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   14355
   LinkTopic       =   "Form1"
   ScaleHeight     =   11010
   ScaleWidth      =   14355
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdPlaceOrder 
      Caption         =   "Place Order For Everyday Items"
      Height          =   735
      Left            =   600
      TabIndex        =   16
      Top             =   7080
      Width           =   4095
   End
   Begin VB.CommandButton cmdSpecialFour 
      Caption         =   "Order Special #4"
      Height          =   735
      Left            =   6840
      TabIndex        =   15
      Top             =   7440
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton cmdSpecialThree 
      Caption         =   "Order Special #3"
      Height          =   735
      Left            =   6840
      TabIndex        =   14
      Top             =   6360
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton cmdSpecialTwo 
      Caption         =   "Order Special #2"
      Height          =   735
      Left            =   6840
      TabIndex        =   13
      Top             =   5280
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton cmdSpecialOne 
      Caption         =   "Order Special #1"
      Height          =   735
      Left            =   6840
      TabIndex        =   12
      Top             =   4200
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton cmdLogOut 
      Caption         =   "Log Out"
      Height          =   495
      Left            =   1320
      TabIndex        =   11
      Top             =   9960
      Width           =   2535
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Go to Another Aisle"
      Height          =   735
      Left            =   600
      TabIndex        =   10
      Top             =   9000
      Width           =   4095
   End
   Begin VB.CommandButton cmdSpecials 
      Caption         =   "See The Daily Specials"
      Height          =   735
      Left            =   600
      TabIndex        =   9
      Top             =   8040
      Width           =   4095
   End
   Begin VB.PictureBox picResults 
      Height          =   4215
      Left            =   8880
      ScaleHeight     =   4155
      ScaleWidth      =   4515
      TabIndex        =   8
      Top             =   4080
      Width           =   4575
   End
   Begin VB.TextBox txtColeslaw 
      Height          =   615
      Left            =   3240
      TabIndex        =   7
      Text            =   "How many lbs?"
      Top             =   6120
      Width           =   1455
   End
   Begin VB.TextBox txtCod 
      Height          =   615
      Left            =   3240
      TabIndex        =   6
      Text            =   "How many lbs?"
      Top             =   5280
      Width           =   1455
   End
   Begin VB.TextBox txtTurkey 
      Height          =   615
      Left            =   3240
      TabIndex        =   5
      Text            =   "How many lbs?"
      Top             =   4440
      Width           =   1455
   End
   Begin VB.TextBox txtHam 
      Height          =   615
      Left            =   3240
      TabIndex        =   4
      Text            =   "How many lbs?"
      Top             =   3600
      Width           =   1455
   End
   Begin VB.Label lblEveryday 
      Caption         =   "Everyday Items"
      Height          =   375
      Left            =   600
      TabIndex        =   17
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Label lblColeslaw 
      Caption         =   "Coleslaw, $4.00 lb"
      Height          =   615
      Left            =   600
      TabIndex        =   3
      Top             =   6120
      Width           =   2295
   End
   Begin VB.Label lblCod 
      Caption         =   "Cod, $5.00 lb"
      Height          =   615
      Left            =   600
      TabIndex        =   2
      Top             =   5280
      Width           =   2295
   End
   Begin VB.Label lblTurkey 
      Caption         =   "Turkey, $4.75 lb"
      Height          =   615
      Left            =   600
      TabIndex        =   1
      Top             =   4440
      Width           =   2295
   End
   Begin VB.Label lblHam 
      Caption         =   "Ham, $4.50 lb"
      Height          =   615
      Left            =   600
      TabIndex        =   0
      Top             =   3600
      Width           =   2295
   End
   Begin VB.Image Image1 
      Height          =   10905
      Left            =   0
      Picture         =   "frmDeli.frx":0000
      Stretch         =   -1  'True
      Top             =   -120
      Width           =   14220
   End
End
Attribute VB_Name = "frmDeli"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'James Delivers, CSCI 130 Visual Basic Project
'frmDeli
'Written by James Garay Heelan
'on 11-2-06
'The purpose of this page is to allow the user to shop for produce items.
'Also, it is set up so that the company may use a text file to set the
'specials and the prices of the specials every day, without having to
'rewrite code.

Option Explicit
Dim Special(1 To 100) As String, Price(1 To 100) As Single

Private Sub cmdBack_Click()
frmDeli.Hide 'the deli page is hidden
frmGroceryStore.Show 'the central shopping menu is displayed
End Sub

Private Sub cmdLogOut_Click()
End 'exits the program
End Sub


Private Sub cmdPlaceOrder_Click()
If txtHam.Text <> "How many lbs?" Then 'if the value in the Ham text box, which the user will enter how many pounds of ham he or she would like, is not the default text, then
    Sum = Sum + (txtHam.Text * 4.5) ' the amount of ham, in lbs, multiplied by the cost per pound is added to the total amount of the users purchase
    Open App.Path & "/PurchasedItems.txt" For Append As #4 'the shopping cart file is opened for writing in
        Write #4, "Ham " & txtHam.Text & " lbs"; "$4.50/lb" 'the product and its costs are recorded in the shopping cart
    Close #4 'the shopping cart file is closed
End If

If txtTurkey.Text <> "How many lbs?" Then
    Sum = Sum + (txtTurkey.Text * 4.75)
    Open App.Path & "/PurchasedItems.txt" For Append As #4
        Write #4, "Turkey " & txtTurkey.Text & " lbs"; "$4.75/lb"
    Close #4
End If

If txtCod.Text <> "How many lbs?" Then
    Sum = Sum + (txtCod.Text * 5)
    Open App.Path & "/PurchasedItems.txt" For Append As #4
        Write #4, "Cod " & txtCod.Text & " lbs"; "$5.00/lb"
    Close #4
End If

If txtColeslaw.Text <> "How many lbs?" Then
    Sum = Sum + (txtColeslaw.Text * 4)
    Open App.Path & "/PurchasedItems.txt" For Append As #4
        Write #4, "Coleslaw " & txtColeslaw.Text & " lbs"; "$4.00/lb"
    Close #4
End If

End Sub

Private Sub cmdSpecialFour_Click()
    Sum = Sum + Price(4) 'the price of special #4 is added to the total amount of the user's purchase
    Open App.Path & "/PurchasedItems.txt" For Append As #4 'the shopping cart file is opened for writing into
        Write #4, Special(4), Price(4) 'special four and it's cost is entered into the shopping cart file
    Close #4 'the file is closed
End Sub

Private Sub cmdSpecialOne_Click()
    Sum = Sum + Price(1)
    Open App.Path & "/PurchasedItems.txt" For Append As #4
        Write #4, Special(1), Price(1)
    Close #4
End Sub

Private Sub cmdSpecials_Click()
Dim N As Integer, M 'counters are loaded

    N = 0
    Open App.Path & "/DeliDailySpecials.txt" For Input As #3 'the file with the daily specials, pre-entered in by the company, for reading
    Do While Not EOF(3) 'the sort is instructed to end when the file is at its end
        N = N + 1 'increases the counter by 1
        Input #3, Special(N), Price(N) 'sorts the text file into two arrays, the Special array and the Price array
    Loop 'the sort loops till it reaches the end of the file
    Close #3 'the file is closed
    
    For M = 1 To N 'the search is instructed to search through as many items as there are recorded as being in the file
        picResults.Print "Special #"; M 'The number of the special being displayed is shown in the picturebox
        picResults.Print Special(M), Price(M) 'the name of the special and its price are displayed in the picturebox
        picResults.Print " " 'spaces between lines are added
        picResults.Print " "
        picResults.Print " "
        picResults.Print " "
    Next M 'the search continues to the next set of data in the array
    
    cmdSpecialOne.Visible = True 'now that the specials have been listed, the buttons allowing the user to purchase them are shown
    cmdSpecialTwo.Visible = True
    cmdSpecialThree.Visible = True
    cmdSpecialFour.Visible = True
    
End Sub

Private Sub cmdSpecialThree_Click()
    Sum = Sum + Price(3)
    Open App.Path & "/PurchasedItems.txt" For Append As #4
        Write #4, Special(3), Price(3)
    Close #4
End Sub

Private Sub cmdSpecialTwo_Click()
    Sum = Sum + Price(2)
    Open App.Path & "/PurchasedItems.txt" For Append As #4
        Write #4, Special(2), Price(2)
    Close #4
End Sub
