VERSION 5.00
Begin VB.Form frmOrder 
   Caption         =   "Order your apparel here!"
   ClientHeight    =   10944
   ClientLeft      =   168
   ClientTop       =   552
   ClientWidth     =   15240
   FillColor       =   &H000000FF&
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   Picture         =   "frmOrder.frx":0000
   ScaleHeight     =   10944
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdsubmit 
      Caption         =   "Submit Your Order"
      Height          =   1092
      Left            =   11400
      TabIndex        =   43
      Top             =   7800
      Width           =   3612
   End
   Begin VB.CommandButton cmdTotal 
      Caption         =   "Click Here to See Your Total"
      Height          =   1092
      Left            =   11400
      TabIndex        =   42
      Top             =   1560
      Width           =   3612
   End
   Begin VB.PictureBox picresults 
      Height          =   4692
      Left            =   11400
      ScaleHeight     =   4644
      ScaleWidth      =   3564
      TabIndex        =   31
      Top             =   2880
      Width           =   3612
   End
   Begin VB.CommandButton cmdreturn 
      Caption         =   "Return to main page"
      Height          =   972
      Left            =   11400
      TabIndex        =   30
      Top             =   240
      Width           =   3615
   End
   Begin VB.TextBox txtq7 
      Height          =   375
      Left            =   9960
      TabIndex        =   29
      Text            =   "0"
      Top             =   3960
      Width           =   972
   End
   Begin VB.TextBox txtq8 
      Height          =   375
      Left            =   9960
      TabIndex        =   28
      Text            =   "0"
      Top             =   6120
      Width           =   972
   End
   Begin VB.TextBox txtq9 
      Height          =   375
      Left            =   9960
      TabIndex        =   27
      Text            =   "0"
      Top             =   8880
      Width           =   972
   End
   Begin VB.PictureBox pic9 
      Height          =   2895
      Left            =   7560
      Picture         =   "frmOrder.frx":2AF4C
      ScaleHeight     =   2844
      ScaleWidth      =   2124
      TabIndex        =   25
      Top             =   7800
      Width           =   2175
   End
   Begin VB.PictureBox pic1 
      Height          =   1335
      Left            =   2280
      Picture         =   "frmOrder.frx":36F93
      ScaleHeight     =   1284
      ScaleWidth      =   1284
      TabIndex        =   23
      Top             =   1320
      Width           =   1332
   End
   Begin VB.PictureBox pic8 
      Height          =   2052
      Left            =   7560
      Picture         =   "frmOrder.frx":3C8BD
      ScaleHeight     =   2004
      ScaleWidth      =   1284
      TabIndex        =   22
      Top             =   5640
      Width           =   1332
   End
   Begin VB.PictureBox pic7 
      Height          =   1812
      Left            =   7560
      Picture         =   "frmOrder.frx":4A6C5
      ScaleHeight     =   1764
      ScaleWidth      =   924
      TabIndex        =   19
      Top             =   3720
      Width           =   972
   End
   Begin VB.PictureBox pic6 
      Height          =   2052
      Left            =   7560
      Picture         =   "frmOrder.frx":53AAF
      ScaleHeight     =   2004
      ScaleWidth      =   1164
      TabIndex        =   17
      Top             =   1560
      Width           =   1212
   End
   Begin VB.PictureBox pic5 
      Height          =   1812
      Left            =   2280
      Picture         =   "frmOrder.frx":55AF7
      ScaleHeight     =   1764
      ScaleWidth      =   1764
      TabIndex        =   16
      Top             =   8280
      Width           =   1812
   End
   Begin VB.PictureBox pic4 
      Height          =   2172
      Left            =   2280
      Picture         =   "frmOrder.frx":57EA6
      ScaleHeight     =   2124
      ScaleWidth      =   1884
      TabIndex        =   13
      Top             =   6120
      Width           =   1932
   End
   Begin VB.PictureBox pic3 
      Height          =   2175
      Left            =   2280
      Picture         =   "frmOrder.frx":59EC5
      ScaleHeight     =   2124
      ScaleWidth      =   1524
      TabIndex        =   11
      Top             =   3960
      Width           =   1572
   End
   Begin VB.PictureBox pic2 
      Height          =   1332
      Left            =   2280
      Picture         =   "frmOrder.frx":62133
      ScaleHeight     =   1284
      ScaleWidth      =   1644
      TabIndex        =   10
      Top             =   2640
      Width           =   1692
   End
   Begin VB.TextBox txtq1 
      Height          =   375
      Left            =   4440
      TabIndex        =   6
      Text            =   "0"
      Top             =   1680
      Width           =   972
   End
   Begin VB.TextBox txtq2 
      Height          =   375
      Left            =   4440
      TabIndex        =   5
      Text            =   "0"
      Top             =   2880
      Width           =   972
   End
   Begin VB.TextBox txtq3 
      Height          =   375
      Left            =   4440
      TabIndex        =   4
      Text            =   "0"
      Top             =   4680
      Width           =   972
   End
   Begin VB.TextBox txtq4 
      Height          =   375
      Left            =   4440
      TabIndex        =   3
      Text            =   "0"
      Top             =   6840
      Width           =   972
   End
   Begin VB.TextBox txtq6 
      Height          =   375
      Left            =   9960
      TabIndex        =   2
      Text            =   "0"
      Top             =   2280
      Width           =   972
   End
   Begin VB.TextBox txtq5 
      Height          =   375
      Left            =   4440
      TabIndex        =   1
      Text            =   "0"
      Top             =   8880
      Width           =   972
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      Caption         =   "By Bill Macy"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   44
      Top             =   10200
      Width           =   1935
   End
   Begin VB.Line Line3 
      X1              =   11280
      X2              =   9720
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Line Line2 
      X1              =   4080
      X2              =   5520
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Line Line1 
      X1              =   11280
      X2              =   11280
      Y1              =   1440
      Y2              =   10680
   End
   Begin VB.Label lblcaption 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "(Including two contollers, Super Mario Brothers game and instructions manuals)"
      Height          =   975
      Left            =   5760
      TabIndex        =   41
      Top             =   8880
      Width           =   1695
   End
   Begin VB.Label lblprice4 
      BackStyle       =   0  'Transparent
      Caption         =   "$9.99"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   720
      TabIndex        =   40
      Top             =   7440
      Width           =   1092
   End
   Begin VB.Label lblprice5 
      BackStyle       =   0  'Transparent
      Caption         =   "$12.99"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   39
      Top             =   8760
      Width           =   1095
   End
   Begin VB.Label lblprice3 
      BackStyle       =   0  'Transparent
      Caption         =   "$15.99"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   720
      TabIndex        =   38
      Top             =   5280
      Width           =   1092
   End
   Begin VB.Label lblprice6 
      BackStyle       =   0  'Transparent
      Caption         =   "$7.99"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6360
      TabIndex        =   37
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label lblprice7 
      BackStyle       =   0  'Transparent
      Caption         =   "$19.99"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   6360
      TabIndex        =   36
      Top             =   3960
      Width           =   1092
   End
   Begin VB.Label lblprice8 
      BackStyle       =   0  'Transparent
      Caption         =   "$14.99"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   6360
      TabIndex        =   35
      Top             =   5880
      Width           =   1092
   End
   Begin VB.Label lblprice9 
      BackStyle       =   0  'Transparent
      Caption         =   "$49.99"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   6480
      TabIndex        =   34
      Top             =   9720
      Width           =   1092
   End
   Begin VB.Label lblprice2 
      BackStyle       =   0  'Transparent
      Caption         =   "$9.99"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   33
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Label lblprice1 
      BackStyle       =   0  'Transparent
      Caption         =   "$29.99"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   32
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label lblquantity2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Quantity"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9600
      TabIndex        =   26
      Top             =   960
      Width           =   1575
   End
   Begin VB.Line Line6 
      X1              =   5520
      X2              =   5520
      Y1              =   1440
      Y2              =   10680
   End
   Begin VB.Label lblsystem 
      BackStyle       =   0  'Transparent
      Caption         =   "9)  Nintendo               Entertainment       System  "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5640
      TabIndex        =   24
      Top             =   7800
      Width           =   2055
   End
   Begin VB.Label lblmug 
      BackStyle       =   0  'Transparent
      Caption         =   "8)  Mario Mug"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      TabIndex        =   21
      Top             =   5640
      Width           =   1695
   End
   Begin VB.Label lblbank 
      BackStyle       =   0  'Transparent
      Caption         =   "7)  Mario Bank"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      TabIndex        =   20
      Top             =   3720
      Width           =   1935
   End
   Begin VB.Label lblsoap 
      BackStyle       =   0  'Transparent
      Caption         =   "6)  Mario Bathing      Soap"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5640
      TabIndex        =   18
      Top             =   1560
      Width           =   1935
   End
   Begin VB.Label lblstockcap 
      BackStyle       =   0  'Transparent
      Caption         =   "4)  Mario               Stocking           Cap"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   972
      Left            =   240
      TabIndex        =   15
      Top             =   6600
      Width           =   1692
   End
   Begin VB.Label lblhat 
      BackStyle       =   0  'Transparent
      Caption         =   "5)  Mario                   Baseball Cap"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   14
      Top             =   8280
      Width           =   1935
   End
   Begin VB.Label lblshirt2 
      BackStyle       =   0  'Transparent
      Caption         =   "3)  Mario                 Supersize          me T-shirt    "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   972
      Left            =   240
      TabIndex        =   12
      Top             =   4440
      Width           =   1812
   End
   Begin VB.Label lblshirt1 
      BackStyle       =   0  'Transparent
      Caption         =   "2)  Mario Power      T-shirt         "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   9
      Top             =   2640
      Width           =   1815
   End
   Begin VB.Label lblmouse 
      BackStyle       =   0  'Transparent
      Caption         =   "1)  Mario CPU         Mouse"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   8
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label lblquantity1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Quantity"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      TabIndex        =   7
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label lbltitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Order through us and receive special discounts! "
      BeginProperty Font 
         Name            =   "Pristina"
         Size            =   28.8
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1680
      TabIndex        =   0
      Top             =   120
      Width           =   9735
   End
End
Attribute VB_Name = "frmOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project name: Mario Madness
'Form name: frmOrder
'Author: Bill Macy
'Date Written: Tuesday March 14th, 2006
'Objective of form:  This form allows the user to order special merchandise through my project (okay so not really).  They
                'can select the number of items wanted and then click a submit button to see their total and inform us of
                'where to bill them for the stuff.  You can also return to the main page.  Its a nifty page that allows the
                'user to save some money if they order more than $45 of merchandise but supplies are limited!
                

Option Explicit
    Dim quantityone As Integer      'declares all my variables
    Dim prices(1 To 9) As Single
    Dim quatities(1 To 9) As Integer
    Dim increment As Integer
    Dim discount As Single
    Dim total As Single
    Dim tax As Single
    Dim i As Integer
    Dim q1 As Integer
    Dim q2 As Integer
    Dim q3 As Integer
    Dim q4 As Integer
    Dim q5 As Integer
    Dim q6 As Integer
    Dim q7 As Integer
    Dim q8 As Integer
    Dim q9 As Integer
    Dim total1 As Single
    Dim total2 As Single
    Dim total3 As Single
    Dim total4 As Single
    Dim total5 As Single
    Dim total6 As Single
    Dim total7 As Single
    Dim total8 As Single
    Dim total9 As Single
    Dim overalltotal As Single
    
    

Private Sub cmdreturn_Click()
    frmOrder.Hide       'hides the order page
    frmMain.Show        'shows the main page
End Sub

Private Sub cmdsubmit_Click()
    Dim name As String      'declares variables
    Dim address As String
    name = InputBox("Please input your name", "Shipping Part 1")        'asks the user to input their name for billing purposes
    address = InputBox("Please input your address for shipping purposes  (Commas in between street address, city, and state)", "Shipping part 2")       'asks the user to input their address so the bill can be sent to them
    MsgBox "The total amount of " & FormatCurrency(total) & " will be billed to " & name & " at " & address & ".  Have a nice day!", , "Shipping part 3"        'displays the information they entered as well as a thank you note for ordering stuff through my page
     
End Sub

Private Sub cmdTotal_Click()
    picresults.Cls      'clears the picture box of anything that may have previously been in it
    Open App.Path & "\Mario Prices.txt" For Input As #1     'opens the file that has the prices and names stored
    increment = 0
    Do Until EOF(1)     'looks through the entire file
        increment = increment + 1       'increments increment by one so the array is filled
        Input #1, prices(increment)     'fills the array with the different items and prices
    Loop        'loops back around
    Close #1        'closes the file that was opened
    q1 = txtq1.Text     'stores the amount the user wanted of this particular item
    q2 = txtq2.Text     'stores the amount the user wanted of this particular item
    q3 = txtq3.Text     'stores the amount the user wanted of this particular item
    q4 = txtq4.Text     'stores the amount the user wanted of this particular item
    q5 = txtq5.Text     'stores the amount the user wanted of this particular item
    q6 = txtq6.Text     'stores the amount the user wanted of this particular item
    q7 = txtq7.Text     'stores the amount the user wanted of this particular item
    q8 = txtq8.Text     'stores the amount the user wanted of this particular item
    q9 = txtq9.Text     'stores the amount the user wanted of this particular item
    picresults.Print "  Thank you for purchasing your items through us!"        'tells the user thanks for purchasing through us
    picresults.Print "    ********************************************************"     'prints a bunch of stars
    picresults.Print "  Quantity ", "Price", "        Total"        'prints the titles for each column
    picresults.Print "  (Product #)"        'prints product number under quantity
    picresults.Print "   --------------------------------------------------------------------"      'prints a solid line
    If q1 > 0 Then      'if the quantity they entered is greater than zero, it loops through to print the information
        total1 = q1 * prices(1)     'multiplies the quantity they entered with the appropriate price
        picresults.Print q1; " (1)", prices(1), "          "; total1        'prints the number they ordered, the product number, the price per product, a bunch of spaces, and the total for the product
    End If
    If q2 > 0 Then      'if the quantity they entered is greater than zero, it loops through to print the information
        total2 = q2 * prices(2)     'multiplies the quantity they entered with the appropriate price
        picresults.Print q2; " (2)", prices(2), "          "; total2        'prints the number they ordered, the product number, the price per product, a bunch of spaces, and the total for the product
    End If
    If q3 > 0 Then      'if the quantity they entered is greater than zero, it loops through to print the information
        total3 = q3 * prices(3)     'multiplies the quantity they entered with the appropriate price
        picresults.Print q3; " (3)", prices(3), "          "; total3        'prints the number they ordered, the product number, the price per product, a bunch of spaces, and the total for the product
    End If
    If q4 > 0 Then      'if the quantity they entered is greater than zero, it loops through to print the information
        total4 = q4 * prices(4)     'multiplies the quantity they entered with the appropriate price
        picresults.Print q4; " (4)", prices(4), "          "; total4        'prints the number they ordered, the product number, the price per product, a bunch of spaces, and the total for the product
    End If
    If q5 > 0 Then      'if the quantity they entered is greater than zero, it loops through to print the information
        total5 = q5 * prices(5)     'multiplies the quantity they entered with the appropriate price
        picresults.Print q5; " (5)", prices(5), "          "; total5        'prints the number they ordered, the product number, the price per product, a bunch of spaces, and the total for the product
    End If
    If q6 > 0 Then      'if the quantity they entered is greater than zero, it loops through to print the information
        total6 = q6 * prices(6)     'multiplies the quantity they entered with the appropriate price
        picresults.Print q6; " (6)", prices(6), "          "; total6        'prints the number they ordered, the product number, the price per product, a bunch of spaces, and the total for the product
    End If
    If q7 > 0 Then      'if the quantity they entered is greater than zero, it loops through to print the information
        total7 = q7 * prices(7)     'multiplies the quantity they entered with the appropriate price
        picresults.Print q7; " (7)", prices(7), "          "; total7        'prints the number they ordered, the product number, the price per product, a bunch of spaces, and the total for the product
    End If
    If q8 > 0 Then      'if the quantity they entered is greater than zero, it loops through to print the information
        total8 = q8 * prices(8)     'multiplies the quantity they entered with the appropriate price
        picresults.Print q8; " (8)", prices(8), "          "; total8        'prints the number they ordered, the product number, the price per product, a bunch of spaces, and the total for the product
    End If
    If q9 > 0 Then      'if the quantity they entered is greater than zero, it loops through to print the information
        total9 = q9 * prices(9)     'multiplies the quantity they entered with the appropriate price
        picresults.Print q9; " (9)", prices(9), "          "; total9        'prints the number they ordered, the product number, the price per product, a bunch of spaces, and the total for the product
    End If
    overalltotal = total1 + total2 + total3 + total4 + total5 + total6 + total7 + total8 + total9       'adds up all the totals that were calculated above
    picresults.Print        'prints a blank line
    picresults.Print "  Sub Total"; Tab(33); FormatCurrency(overalltotal)       'prints the words "sub total" and then the total just calculated
    If overalltotal > 45 Then       'if that total is greater than $45 then it loops through to give the discount
        discount = overalltotal * 0.15      'calculates the 15% discount
        picresults.Print "  Discount"; Tab(33); FormatCurrency(discount)        'prints the words "discount" and print the amount they saved by spending over $45
    End If
    tax = overalltotal * 0.065      'calculates the tax by multiplying the total by 6.5 percent
    picresults.Print "  Tax"; Tab(33); FormatCurrency(tax)      'print the value that was derived for tax
    picresults.Print "  --------------------------------------------------------------------"       'prints a line
    total = overalltotal - discount + tax       'calculates the overall total that the user will be billed taking the orignal total - the discount and adding tax
    picresults.Print "  Total"; Tab(33); FormatCurrency(total)      'prints the word "total" and gives the number that the user owes for the merchandise
    
End Sub
