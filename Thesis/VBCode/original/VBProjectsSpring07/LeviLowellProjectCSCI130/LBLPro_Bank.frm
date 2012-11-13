VERSION 5.00
Begin VB.Form frmPurchases 
   BackColor       =   &H00000000&
   Caption         =   "Purchases"
   ClientHeight    =   9735
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   11970
   LinkTopic       =   "Form1"
   ScaleHeight     =   9735
   ScaleWidth      =   11970
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CommandButton cmdCheckOut 
      BackColor       =   &H0000C000&
      Caption         =   "Check Out"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   8520
      Width           =   1575
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H8000000B&
      Height          =   6375
      Left            =   5400
      ScaleHeight     =   6315
      ScaleWidth      =   5595
      TabIndex        =   9
      Top             =   2040
      Width           =   5655
   End
   Begin VB.CommandButton cmdTotal 
      BackColor       =   &H0000C000&
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   8520
      Width           =   2775
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H0000C000&
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Myriad Condensed Web"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   8520
      Width           =   1575
   End
   Begin VB.CommandButton cmdPens 
      BackColor       =   &H0000C000&
      Caption         =   "Lowell Pens"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6360
      Width           =   2775
   End
   Begin VB.CommandButton cmdStationary 
      BackColor       =   &H0000C000&
      Caption         =   "Lowell Stationaries"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7440
      Width           =   2775
   End
   Begin VB.CommandButton cmdCheckbook 
      BackColor       =   &H0000C000&
      Caption         =   "Lowell Checkbooks"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5280
      Width           =   2775
   End
   Begin VB.CommandButton cmdKeychain 
      BackColor       =   &H0000C000&
      Caption         =   "Lowell Keychains"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4200
      Width           =   2775
   End
   Begin VB.CommandButton cmdNotebook 
      BackColor       =   &H0000C000&
      Caption         =   "Lowell Notebooks"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3120
      Width           =   2775
   End
   Begin VB.CommandButton cmdTshirt 
      BackColor       =   &H0000C000&
      Caption         =   "Lowell T-Shirts"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2040
      Width           =   2775
   End
   Begin VB.CommandButton cmdReturn 
      BackColor       =   &H0000C000&
      Caption         =   "Return to Main Page"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   8760
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   8520
      Width           =   2295
   End
   Begin VB.Image Image6 
      BorderStyle     =   1  'Fixed Single
      Height          =   975
      Left            =   840
      Picture         =   "LBLPro_Bank.frx":0000
      Stretch         =   -1  'True
      Top             =   7440
      Width           =   1575
   End
   Begin VB.Image Image5 
      BorderStyle     =   1  'Fixed Single
      Height          =   975
      Left            =   840
      Picture         =   "LBLPro_Bank.frx":1BC85
      Stretch         =   -1  'True
      Top             =   6360
      Width           =   1560
   End
   Begin VB.Image Image4 
      BorderStyle     =   1  'Fixed Single
      Height          =   945
      Left            =   840
      Picture         =   "LBLPro_Bank.frx":2E1A1
      Stretch         =   -1  'True
      Top             =   5280
      Width           =   1575
   End
   Begin VB.Image Image3 
      BorderStyle     =   1  'Fixed Single
      Height          =   1005
      Left            =   840
      Picture         =   "LBLPro_Bank.frx":32361
      Stretch         =   -1  'True
      Top             =   4200
      Width           =   1530
   End
   Begin VB.Image Image2 
      BorderStyle     =   1  'Fixed Single
      Height          =   1020
      Left            =   840
      Picture         =   "LBLPro_Bank.frx":33872
      Stretch         =   -1  'True
      Top             =   3120
      Width           =   1560
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   975
      Left            =   840
      Picture         =   "LBLPro_Bank.frx":399BF
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   1560
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   $"LBLPro_Bank.frx":8BF31
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   1455
      Left            =   2520
      TabIndex        =   10
      Top             =   480
      Width           =   8535
   End
End
Attribute VB_Name = "frmPurchases"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'This program allows the user to "buy" products endorsed by the Lowell Bank!  It's just for fun.
'Also the program asks the user for their information so the Lowell Bank can send them their items.
'Automatically the store adds a tax of 7%.  When the user buys $50.00 worth of items they will
'receive a 15% discount, which is very nice!

Dim sum As Single       'Sets the variables
Dim tax As Single
Dim Discount As Single


Private Sub cmdCheckbook_Click()
Dim Checkbook As Single     'Sets the variable

Checkbook = 5       'Initiates Checkbook variable as 5.00
sum = sum + 5       'Adds 5.00 to the sum
picResults.Print "Checkbook", ; FormatCurrency(Checkbook, 2)        'Displays the item and the amount into the picturebox
End Sub

Private Sub cmdCheckOut_Click()
Dim XName As String     'Sets variables
Dim XAddress As String

XName = InputBox("Please enter your name:", "Shipping Information")     'Asks user for their name
XAddress = InputBox("Please enter your shipping Address:", "Shipping Information")      'Asks user for their shipping address
    MsgBox "Thank you for your purchase " & XName & " your items will be sent to you", , "Thank You!"        'Displays a message letting the user know their items will be sent to them
End Sub

Private Sub cmdKeychain_Click()
Dim Keychain As Single      'Sets the variable

Keychain = 0.75     'initiates Keychain variable as 0.75
sum = sum + 0.75        'Adds 0.75 to the sum
picResults.Print "Keychain", ; FormatCurrency(Keychain, 2)      'Displays the item and the amount into the picturebox
End Sub

Private Sub cmdNotebook_Click()
Dim notebook As Single      'Sets the variable

notebook = 1        'Initiates notebook variable as 1.00
sum = sum + 1       'Adds 1.00 to the sum
picResults.Print "Notebook", ; FormatCurrency(notebook, 2)      'Displays the item and the amount into the picturebox
End Sub

Private Sub cmdPens_Click()
Dim Pens As Single      'Sets the variable

Pens = 2        'Initiates Pens variable as 2.00
sum = sum + 2       'Adds 2.00 to the sum
picResults.Print "Pens", ; FormatCurrency(Pens, 2)      'Displays the item and the amount into the picturebox
End Sub

Private Sub cmdStationary_Click()
Dim Stationary As Single        'Sets the variable

Stationary = 7      'Initiates Stationary variable as 7.00
sum = sum + 7       'Adds 7.00 to the sum
picResults.Print "Stationary", ; FormatCurrency(Stationary, 2)      'Displays the item and the amount into the picturebox
End Sub


Private Sub cmdTshirt_Click()
Dim Tshirt As Single        'Sets the variable

Tshirt = 12     'Initiates tshirts variable as 12.00
sum = sum + 12      'Adds 12.00 to the sum
picResults.Print "T-Shirt", ; FormatCurrency(Tshirt, 2)     'Displays the item and the amount into the picturebox
End Sub

Private Sub cmdTotal_Click()

sum = sum * 1.07        'Adds the sales tax to the sum
    picResults.Print "Your Total is:  ", ; FormatCurrency(sum, 2); " with the sales tax."     'Displays the total if less than $50.00 and without the discount
    Discount = 50       'Sets the discount to $50.00
    
If sum > Discount Then      'If/Then statement that deciphers if the sum is greater than $50.00 so the user gets the discount
        sum = sum - sum * 0.15      'Subtracts the 15% discount from the sum
            picResults.Print "Your Total is over 50.00!  You get the 15% discount! ", ; FormatCurrency(sum, 2)      'Displays the new total after the 15% discount has been given
End If
    sum = 0     'Sets sum equal to zero
    
End Sub

Private Sub cmdClear_Click()

picResults.Cls      'Clears the picturebox
End Sub

Private Sub cmdReturn_Click()

frmPurchases.Hide       'Hides frmPurchases
FrmMain.Show        'Shows frmMain
End Sub


