VERSION 5.00
Begin VB.Form frmGroceryStore 
   BackColor       =   &H00FF8080&
   Caption         =   "Grocery Store"
   ClientHeight    =   9585
   ClientLeft      =   8295
   ClientTop       =   1725
   ClientWidth     =   6840
   LinkTopic       =   "Form1"
   ScaleHeight     =   21407.04
   ScaleMode       =   0  'User
   ScaleWidth      =   4488.189
   Begin VB.CommandButton cmdRange 
      BackColor       =   &H000080FF&
      Caption         =   "List Items in Price Range"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      MaskColor       =   &H000080FF&
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   7440
      Width           =   1695
   End
   Begin VB.PictureBox picFood 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      FillColor       =   &H00C0E0FF&
      ForeColor       =   &H00C0E0FF&
      Height          =   1335
      Left            =   120
      Picture         =   "frmGroceryStore.frx":0000
      ScaleHeight     =   1335
      ScaleWidth      =   1935
      TabIndex        =   6
      Top             =   6000
      Width           =   1935
   End
   Begin VB.CommandButton cmdToMenu 
      BackColor       =   &H000080FF&
      Caption         =   "Return to Menu"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   8280
      Width           =   1695
   End
   Begin VB.CommandButton cmdSortCategory 
      BackColor       =   &H000080FF&
      Caption         =   "Sort Items by Category"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1920
      Width           =   1695
   End
   Begin VB.CommandButton cmdSortPrice 
      BackColor       =   &H000080FF&
      Caption         =   "Sort Items by Price"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1080
      Width           =   1695
   End
   Begin VB.CommandButton cmdShowList 
      BackColor       =   &H000080FF&
      Caption         =   "Show Groceries"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   240
      Width           =   1695
   End
   Begin VB.PictureBox picList 
      BackColor       =   &H00FFFFC0&
      Height          =   9495
      Left            =   2160
      ScaleHeight     =   9435
      ScaleWidth      =   4395
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF8080&
      Caption         =   "by: Natalie Bly"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   8
      Top             =   9120
      Width           =   975
   End
   Begin VB.Label lblStore 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      Caption         =   $"frmGroceryStore.frx":0F60
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   240
      TabIndex        =   5
      Top             =   2760
      Width           =   1695
   End
End
Attribute VB_Name = "frmGroceryStore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Money Manager 2005 (ProjectNataliesMoneyPlanner)
'frmGroceryStore (frmGroceryStore.frm)
'by Natalie Bly
'10/29/05
'The purpose of this form is to give the user a chance to look at a few examples
'of items they might want/need at the grocery store along with their prices to get
'a better idea of how to plan their grocery list.  It's a lot easier to compare prices
'at a computer than to run around at the grocery store because you have no list
'or because you need to compare items.  This form should help that problem a little
'bit at least.  There are buttons that sort items from a file by price or by category.
'There is also a button that allows the user to imput a price range and will display
'the items from the list that are within that range.

Option Explicit                 'makes it easier to debug the code

Private Sub cmdRange_Click()
    picList.Cls                          'clears the output box
    Dim LowerBound As Single, UpperBound As Single
    LowerBound = InputBox("Enter the price range you'd like to search, starting with the lower bound", "Price Range")
                                         'user chooses a lower bound for the price range
    UpperBound = InputBox("And now enter the upper bound of the price range.", "Price Range")
                                         'user chooses an upper bound for the price range
    For K = 1 To I                       'program executes if statement until it has searched the whole file of size I
        If (Price(K) <= UpperBound And Price(K) >= LowerBound) Then     'if item price falls within the bounds then
            picList.Print Item(K); Tab(30); FormatCurrency(Price(K)); Tab(45); Category(K)
                                        'the program prints the item, its price, and its category
        End If
    Next K
End Sub

Private Sub cmdSortCategory_Click()
    picList.Cls                          'clears the output box
    For Pass = 1 To I - 1                'sorts the arrays according to Category
        For K = 1 To I - Pass
            If Category(K) > Category(K + 1) Then
                TempCategory = Category(K) 'stores the "larger" (alphabetically) Category in a temporary variable
                Category(K) = Category(K + 1)   'moves the "smaller" category value back to where the larger value had been
                Category(K + 1) = TempCategory  'replaces the smaller value (in the K+1 slot) with the larger (previously K slot) value, which had been temporarily stored in TempCategory
                TempPrice = Price(K)            'switches the values in the parallel item and price arrays
                Price(K) = Price(K + 1)         'so that they still match up with their appropriate,
                Price(K + 1) = TempPrice        'now sorted, categories.
                TempItem = Item(K)
                Item(K) = Item(K + 1)
                Item(K + 1) = TempItem
            End If
        Next K
    Next Pass
    For L = 1 To I
        picList.Print Item(L); Tab(30); FormatCurrency(Price(L)); Tab(45); Category(L)
                                        'prints the sorted arrays--item, price, and category
    Next L
End Sub
Private Sub cmdSortPrice_Click()
    picList.Cls                         'clears the output box
    For Pass = 1 To I - 1               'Sorts the arrays by Price, from least to greatest
        For K = 1 To I - Pass
            If Price(K) > Price(K + 1) Then    'if the value in the K slot of the price array is greater than the adjacent K+1 slot, then
                TempPrice = Price(K)           'the larger value from slot K is stored in a temporary variable
                Price(K) = Price(K + 1)        'the smaller value from slot K+1 is moved to the K slot
                Price(K + 1) = TempPrice       'and the value in the K+1 slot is replaced with the larger value that was temporarily stored in the TempPrice variable
                TempItem = Item(K)             'switches the values in the parallel Item and Category arrays
                Item(K) = Item(K + 1)          'so that they still match up with the appropriate, now sorted
                Item(K + 1) = TempItem         'prices.
                TempCategory = Category(K)
                Category(K) = Category(K + 1)
                Category(K + 1) = TempCategory
            End If
        Next K
    Next Pass
    For L = 1 To I
        picList.Print Item(L); Tab(30); FormatCurrency(Price(L)); Tab(45); Category(L)
                                        'prints the sorted arrays--item, price, and category
    Next L
End Sub
Private Sub cmdToMenu_Click()
    frmGroceryStore.Hide                'takes the user back to the Menu screen
    frmMenu.Show
End Sub

Private Sub cmdShowList_Click()
    picList.Cls                         'clears the output box
    For J = 1 To I
        picList.Print Item(J); Tab(30); FormatCurrency(Price(J)); Tab(45); Category(J)
                                        'prints the arrays as they appear in the text file
    Next J
End Sub

