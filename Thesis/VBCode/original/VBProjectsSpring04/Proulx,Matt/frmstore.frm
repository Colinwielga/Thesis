VERSION 5.00
Begin VB.Form frmstore 
   BackColor       =   &H80000003&
   Caption         =   "Counting Crows Store"
   ClientHeight    =   2940
   ClientLeft      =   60
   ClientTop       =   600
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   2940
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdcomment 
      Caption         =   "Purchase Items!"
      Height          =   495
      Left            =   7680
      TabIndex        =   21
      Top             =   6480
      Width           =   1935
   End
   Begin VB.CommandButton cmdclear 
      Caption         =   "Clear your Items"
      Height          =   495
      Left            =   7680
      TabIndex        =   20
      Top             =   7080
      Width           =   1935
   End
   Begin VB.CommandButton cmdbucket 
      Caption         =   "Purchase"
      Height          =   375
      Index           =   8
      Left            =   3120
      TabIndex        =   18
      Top             =   2280
      Width           =   1335
   End
   Begin VB.CommandButton cmdhoodie 
      Caption         =   "Purchase"
      Height          =   375
      Index           =   7
      Left            =   5520
      TabIndex        =   17
      Top             =   2280
      Width           =   1335
   End
   Begin VB.CommandButton cmdtshirt 
      Caption         =   "Purchase"
      Height          =   375
      Index           =   6
      Left            =   7920
      TabIndex        =   16
      Top             =   2280
      Width           =   1335
   End
   Begin VB.CommandButton cmdtank 
      Caption         =   "Purchase"
      Height          =   375
      Index           =   5
      Left            =   840
      TabIndex        =   15
      Top             =   4920
      Width           =   1335
   End
   Begin VB.CommandButton cmdposter 
      Caption         =   "Purchase"
      Height          =   375
      Index           =   4
      Left            =   3240
      TabIndex        =   14
      Top             =   4920
      Width           =   1335
   End
   Begin VB.CommandButton cmdsticker 
      Caption         =   "Purchase"
      Height          =   375
      Index           =   3
      Left            =   5640
      TabIndex        =   13
      Top             =   4920
      Width           =   1335
   End
   Begin VB.CommandButton cmdkeys 
      Caption         =   "Purchase"
      Height          =   375
      Index           =   2
      Left            =   8040
      TabIndex        =   12
      Top             =   4920
      Width           =   1335
   End
   Begin VB.CommandButton cmdhat 
      Caption         =   "Purchase"
      Height          =   375
      Index           =   0
      Left            =   720
      TabIndex        =   11
      Top             =   2280
      Width           =   1335
   End
   Begin VB.CommandButton cmdtotal 
      Caption         =   "Click to see your total"
      Height          =   495
      Left            =   7680
      TabIndex        =   10
      Top             =   5880
      Width           =   1935
   End
   Begin VB.PictureBox Picture2 
      Height          =   1815
      Left            =   7560
      Picture         =   "frmstore.frx":0000
      ScaleHeight     =   1755
      ScaleWidth      =   2115
      TabIndex        =   9
      Top             =   2880
      Width           =   2175
   End
   Begin VB.PictureBox Picture1 
      Height          =   1815
      Index           =   1
      Left            =   2760
      Picture         =   "frmstore.frx":F226
      ScaleHeight     =   1755
      ScaleWidth      =   2115
      TabIndex        =   8
      Top             =   240
      Width           =   2175
   End
   Begin VB.PictureBox Picture1 
      Height          =   1815
      Index           =   0
      Left            =   360
      Picture         =   "frmstore.frx":1FB40
      ScaleHeight     =   1755
      ScaleWidth      =   2115
      TabIndex        =   7
      Top             =   240
      Width           =   2175
   End
   Begin VB.PictureBox Picture1 
      Height          =   1815
      Index           =   7
      Left            =   2760
      Picture         =   "frmstore.frx":2C68E
      ScaleHeight     =   1755
      ScaleWidth      =   2115
      TabIndex        =   6
      Top             =   2880
      Width           =   2175
   End
   Begin VB.PictureBox Picture1 
      Height          =   1815
      Index           =   6
      Left            =   5160
      Picture         =   "frmstore.frx":398EC
      ScaleHeight     =   1755
      ScaleWidth      =   2115
      TabIndex        =   5
      Top             =   2880
      Width           =   2175
   End
   Begin VB.PictureBox Picture1 
      Height          =   1815
      Index           =   4
      Left            =   5160
      Picture         =   "frmstore.frx":4A042
      ScaleHeight     =   1755
      ScaleWidth      =   2115
      TabIndex        =   4
      Top             =   240
      Width           =   2175
   End
   Begin VB.PictureBox Picture1 
      Height          =   1815
      Index           =   3
      Left            =   7560
      Picture         =   "frmstore.frx":5A95C
      ScaleHeight     =   1755
      ScaleWidth      =   2115
      TabIndex        =   3
      Top             =   240
      Width           =   2175
   End
   Begin VB.PictureBox Picture1 
      Height          =   1815
      Index           =   2
      Left            =   360
      Picture         =   "frmstore.frx":6B276
      ScaleHeight     =   1755
      ScaleWidth      =   2115
      TabIndex        =   2
      Top             =   2880
      Width           =   2175
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H0000FFFF&
      Height          =   4095
      Left            =   480
      ScaleHeight     =   4035
      ScaleWidth      =   7035
      TabIndex        =   1
      Top             =   5880
      Width           =   7095
   End
   Begin VB.CommandButton cmdback 
      Caption         =   "Go back to the Main Page"
      Height          =   735
      Left            =   7680
      TabIndex        =   0
      Top             =   8400
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Created by Matt Proulx"
      Height          =   255
      Index           =   1
      Left            =   480
      TabIndex        =   22
      Top             =   10080
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FFFF&
      Caption         =   "What you're Purchasing"
      Height          =   255
      Index           =   0
      Left            =   3120
      TabIndex        =   19
      Top             =   5520
      Width           =   1695
   End
End
Attribute VB_Name = "frmstore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : CountingCrows (Matt Proulx's VB Project.vbp)
'Form Name : frmstore (frmstore.frm)
'Author: Matt Proulx
'Date Written: March 13, 2004
'Purpose of the Form: 'This form will let the user purchase Counting Crows merchandise. It will display the items the
                      'user is purchasing along with the price. It will also calculate the tax and shipping costs of
                      'your total and display a message confirming your purchase.

'Option Explicit is a command to force
'the user to explicitly declare all variables
'before they can be used.
Dim gifts(1 To 8) As String
Dim price(1 To 8) As Single
Dim total As Single
Dim Tax As Single
Dim Shipping As Single
Dim CTR As Integer
Private Sub cmdback_Click()
    frmstore.Hide
    frmtitle.Show
End Sub
Private Sub cmdclear_Click()
picResults.Cls 'Clears your purchases
total = 0
cmdcomment.Enabled = False 'Disable "purchase button" so the user has to click on items and find the total first
                           'before they purchase their items.
cmdtotal.Enabled = True    'Enable "total button" so user can purchase in more items and find the total.
End Sub
Private Sub cmdcomment_Click()
    MsgBox "Thank you! Your order is being processed. Please allow 2 to 3 weeks for shipping and enjoy your purchases."
End Sub
Private Sub cmdhat_Click(Index As Integer)
    picResults.Print gifts(1), FormatCurrency(price(1)) 'Prints the item and price
    total = total + price(1) 'Adds the price of the item to the total
End Sub
Private Sub cmdbucket_Click(Index As Integer)
    picResults.Print gifts(2), FormatCurrency(price(2)) 'Prints the item and price
    total = total + price(2)
End Sub
Private Sub cmdhoodie_Click(Index As Integer)
    picResults.Print gifts(3), FormatCurrency(price(3)) 'Prints the item and price
    total = total + price(3)
End Sub
Private Sub cmdtshirt_Click(Index As Integer)
    picResults.Print gifts(4), FormatCurrency(price(4)) 'Prints the item and price
    total = total + price(4)
End Sub
Private Sub cmdtank_Click(Index As Integer)
    picResults.Print gifts(5), FormatCurrency(price(5)) 'Prints the item and price
    total = total + price(5)
End Sub
Private Sub cmdposter_Click(Index As Integer)
    picResults.Print gifts(6), FormatCurrency(price(6)) 'Prints the item and price
    total = total + price(6)
End Sub
Private Sub cmdsticker_Click(Index As Integer)
    picResults.Print gifts(7), FormatCurrency(price(7)) 'Prints the item and price
    total = total + price(7)
End Sub
Private Sub cmdkeys_Click(Index As Integer)
    picResults.Print gifts(8), FormatCurrency(price(8)) 'Prints the item and price
    total = total + price(8)
End Sub
Private Sub cmdtotal_Click()
    picResults.Print "----------------------"
    picResults.Print "Subtotal ->", FormatCurrency(total)
        Tax = total * 0.09
    picResults.Print "Tax ->", FormatCurrency(Tax)
        Shipping = 2.99 'Defines the cost for shipping
    picResults.Print "Shipping ->", FormatCurrency(Shipping)
        total = total * 1.09 + Shipping
    picResults.Print "Total ->", FormatCurrency(total) 'Prints the final total
    picResults.Print "***********************************************************"
    Select Case total
        Case Is > 200
        picResults.Print "What kind of work did you say you do again?"
        Case Is > 100
        picResults.Print "You're putting us through college!"
        Case Is > 75
        picResults.Print "Someone's got some extra cash lying around."
        Case Is > 50
        picResults.Print "Payday today hu?"
        Case Is > 25
        picResults.Print "Someone got their allowence money this week!"
        Case Else
        picResults.Print "Big spender hu?"
    End Select
    cmdcomment.Enabled = True
    cmdtotal.Enabled = False
End Sub
Private Sub Form_Load()
    Path = "N:\CS130\handin\Proulx, Matt\" 'Loads the prices of the items into the program
    Open Path & "stuff.txt" For Input As #1
    CTR = 0
        Do While Not EOF(1)
            CTR = CTR + 1
            Input #1, gifts(CTR), price(CTR)
        Loop
        Close #1
    cmdcomment.Enabled = False
End Sub

