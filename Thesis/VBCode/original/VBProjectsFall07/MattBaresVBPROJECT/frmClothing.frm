VERSION 5.00
Begin VB.Form frmClothing 
   Caption         =   "Target"
   ClientHeight    =   8310
   ClientLeft      =   4005
   ClientTop       =   3690
   ClientWidth     =   9735
   LinkTopic       =   "Form1"
   ScaleHeight     =   8310
   ScaleWidth      =   9735
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Leave the Clothing Department"
      Height          =   615
      Left            =   5280
      TabIndex        =   10
      Top             =   7440
      Width           =   4215
   End
   Begin VB.PictureBox picKids 
      Height          =   4215
      Left            =   6600
      ScaleHeight     =   4155
      ScaleWidth      =   2835
      TabIndex        =   2
      Top             =   2280
      Width           =   2895
   End
   Begin VB.PictureBox picWomen 
      Height          =   4215
      Left            =   3360
      ScaleHeight     =   4155
      ScaleWidth      =   2955
      TabIndex        =   1
      Top             =   2280
      Width           =   3015
   End
   Begin VB.PictureBox picMen 
      Height          =   4215
      Left            =   240
      ScaleHeight     =   4155
      ScaleWidth      =   2835
      TabIndex        =   0
      Top             =   2280
      Width           =   2895
   End
   Begin VB.Frame fraClothing 
      Caption         =   "Welcome to the Clothing Department"
      Height          =   8055
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   9495
      Begin VB.CommandButton cmdSortbyName 
         Caption         =   "Sort Clothes by Name"
         Height          =   495
         Left            =   4920
         TabIndex        =   12
         Top             =   6480
         Width           =   2895
      End
      Begin VB.CommandButton cmdSortbyPrice 
         Caption         =   "Sort Clothes by Price"
         Height          =   495
         Left            =   1680
         TabIndex        =   11
         Top             =   6480
         Width           =   2895
      End
      Begin VB.CommandButton cmdAddtoCart 
         Caption         =   "Add Item to Cart"
         Height          =   735
         Left            =   4560
         TabIndex        =   9
         Top             =   360
         Width           =   4695
      End
      Begin VB.CommandButton cmdBegin 
         Caption         =   "Display Clothing"
         Height          =   495
         Left            =   120
         TabIndex        =   8
         Top             =   1320
         Width           =   9255
      End
      Begin VB.PictureBox Picture1 
         Height          =   5895
         Left            =   0
         Picture         =   "frmClothing.frx":0000
         ScaleHeight     =   5835
         ScaleWidth      =   9435
         TabIndex        =   13
         Top             =   1200
         Width           =   9495
      End
      Begin VB.Label lblWhenDone 
         Caption         =   $"frmClothing.frx":123ED
         Height          =   615
         Left            =   240
         TabIndex        =   14
         Top             =   7320
         Width           =   4815
      End
      Begin VB.Label lblKidsClothing 
         Caption         =   "Children's Clothing"
         Height          =   255
         Left            =   7320
         TabIndex        =   7
         Top             =   1920
         Width           =   1455
      End
      Begin VB.Label lblWomensClothing 
         Caption         =   "Women's Clothing"
         Height          =   255
         Left            =   4080
         TabIndex        =   6
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Label lblMensClothing 
         Caption         =   "Men's Clothing"
         Height          =   255
         Left            =   960
         TabIndex        =   5
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label lblSelectItem 
         Caption         =   $"frmClothing.frx":12474
         Height          =   735
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   4335
      End
   End
End
Attribute VB_Name = "frmClothing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim InventoryM(0 To 50) As Integer, InventoryW(0 To 50) As Integer, InventoryK(0 To 50) As Integer
Dim MCName(0 To 50) As String, MCPrice(0 To 50) As Single, WCName(0 To 50) As String, WCPrice(0 To 50) As Single, KCName(0 To 50) As String, KCPrice(0 To 50) As Single
Dim CTR3 As Integer, pos3 As Integer, CTR2 As Integer, pos2 As Integer, CTR As Integer, pos As Integer
Private Sub cmdAddtoCart_Click()
'This subroutine loads users item selections from an inputbox into a 'clothing cart' array and then to a file
'for display during checkout. If the user puts anything in their cart, the program will allow
'the user to now use the "proceed to checkout" button in the department selection form.
Dim Cart(0 To 100) As Integer, AdditiontoCart As Integer, X As Integer, counter As Integer

Open App.Path & "\ClothingCart.txt" For Output As #1

AdditiontoCart = InputBox("Would you like to make a purchase? Input the item's number, type '-1' to quit.", "Add to cart")

Do While AdditiontoCart <> -1
    counter = counter + 1
    'If the user inputs a number that doesn't correspond to an item from this department, they will be warned
    'via a msgbox and asked to try again.
    If AdditiontoCart < -1 Or AdditiontoCart > 37 Then
        MsgBox "That item number does not correspond to an item from this department.  Please try again.", , "Wrong Item Number"
    End If
    
    'If the user puts anything in their cart, the program will allow
    'the user to now use the "proceed to checkout" button in the department selection form.
    If counter > 0 Then
        frmDepartments.cmdCheckout.Enabled = True
    End If
    
    Cart(counter) = AdditiontoCart
    AdditiontoCart = InputBox("Would you like to make another purchase? Input the item's number.", "Add to cart")
Loop

For X = 1 To counter
    Print #1, Cart(X)
Next X

Close #1

End Sub

Private Sub cmdBegin_Click()
'This subroutine opens the file for men's, women's and children's clothing into 9 arrays and
'displays them in their relative picture box.

Open App.Path & "\MensClothing.txt" For Input As #1

CTR = 0

Do While Not EOF(1)
    CTR = CTR + 1
    Input #1, InventoryM(CTR), MCName(CTR), MCPrice(CTR)
Loop
Close #1

picMen.Cls
picMen.Print "#"; Tab(6); "Name"; Tab(31); "Price"
picMen.Print "*********************************************"

For pos = 1 To CTR
    picMen.Print InventoryM(pos); Tab(6); MCName(pos); Tab(31); FormatCurrency(MCPrice(pos))
Next pos

'Next Column

Open App.Path & "\WomensClothing.txt" For Input As #2

CTR2 = 0

Do While Not EOF(2)
    CTR2 = CTR2 + 1
    Input #2, InventoryW(CTR2), WCName(CTR2), WCPrice(CTR2)
Loop
Close #2

picWomen.Cls
picWomen.Print "#"; Tab(6); "Name"; Tab(31); "Price"
picWomen.Print "*********************************************"

For pos2 = 1 To CTR2
    picWomen.Print InventoryW(pos2); Tab(6); WCName(pos2); Tab(31); FormatCurrency(WCPrice(pos2))
Next pos2

'Next Column

Open App.Path & "\KidsClothing.txt" For Input As #3

CTR3 = 0

Do While Not EOF(3)
    CTR3 = CTR3 + 1
    Input #3, InventoryK(CTR3), KCName(CTR3), KCPrice(CTR3)
Loop
Close #3

picKids.Cls
picKids.Print "#"; Tab(6); "Name"; Tab(31); "Price"
picKids.Print "*********************************************"

For pos3 = 1 To CTR3
    picKids.Print InventoryK(pos3); Tab(6); KCName(pos3); Tab(31); FormatCurrency(KCPrice(pos3))
Next pos3

End Sub

Private Sub cmdReturn_Click()
'This subroutine shows the department selection form and hides the clothing department form.
frmClothing.Hide
frmDepartments.Show
End Sub

Private Sub cmdSortbyName_Click()
Dim Pass As Integer, Temp As Integer, TempName As String, TempPrice As Single
'This subroutine sorts each of the clothing types(men, women, and kids) by using bubble sort. It sorts by name.
'It then clears the picture boxes and redisplays the sorted lists.

For Pass = 1 To CTR - 1
    For pos = 1 To CTR - Pass
        If MCName(pos) > MCName(pos + 1) Then
            TempPrice = MCPrice(pos)
            MCPrice(pos) = MCPrice(pos + 1)
            MCPrice(pos + 1) = TempPrice
            TempName = MCName(pos)
            MCName(pos) = MCName(pos + 1)
            MCName(pos + 1) = TempName
            Temp = InventoryM(pos)
            InventoryM(pos) = InventoryM(pos + 1)
            InventoryM(pos + 1) = Temp
        End If
    Next pos
Next Pass

picMen.Cls
picMen.Print "#"; Tab(6); "Name"; Tab(31); "Price"
picMen.Print "*********************************************"

For pos = 1 To CTR
    picMen.Print InventoryM(pos); Tab(6); MCName(pos); Tab(31); FormatCurrency(MCPrice(pos))
Next pos

'Next Column

For Pass = 1 To CTR2 - 1
    For pos2 = 1 To CTR2 - Pass
        If WCName(pos2) > WCName(pos2 + 1) Then
            TempPrice = WCPrice(pos2)
            WCPrice(pos2) = WCPrice(pos2 + 1)
            WCPrice(pos2 + 1) = TempPrice
            TempName = WCName(pos2)
            WCName(pos2) = WCName(pos2 + 1)
            WCName(pos2 + 1) = TempName
            Temp = InventoryW(pos2)
            InventoryW(pos2) = InventoryW(pos2 + 1)
            InventoryW(pos2 + 1) = Temp
        End If
    Next pos2
Next Pass

picWomen.Cls
picWomen.Print "#"; Tab(6); "Name"; Tab(31); "Price"
picWomen.Print "*********************************************"

For pos2 = 1 To CTR2
    picWomen.Print InventoryW(pos2); Tab(6); WCName(pos2); Tab(31); FormatCurrency(WCPrice(pos2))
Next pos2

'Next Column

For Pass = 1 To CTR3 - 1
    For pos3 = 1 To CTR3 - Pass
        If KCName(pos3) > KCName(pos3 + 1) Then
            TempPrice = KCPrice(pos3)
            KCPrice(pos3) = KCPrice(pos3 + 1)
            KCPrice(pos3 + 1) = TempPrice
            TempName = KCName(pos3)
            KCName(pos3) = KCName(pos3 + 1)
            KCName(pos3 + 1) = TempName
            Temp = InventoryK(pos3)
            InventoryK(pos3) = InventoryK(pos3 + 1)
            InventoryK(pos3 + 1) = Temp
        End If
    Next pos3
Next Pass

picKids.Cls
picKids.Print "#"; Tab(6); "Name"; Tab(31); "Price"
picKids.Print "*********************************************"

For pos3 = 1 To CTR3
    picKids.Print InventoryK(pos3); Tab(6); KCName(pos3); Tab(31); FormatCurrency(KCPrice(pos3))
Next pos3

End Sub

Private Sub cmdSortbyPrice_Click()
Dim Pass As Integer, Temp As Integer, TempName As String, TempPrice As Single
'This subroutine sorts each of the clothing types(men, women, and kids) by using bubble sort. It sorts by price.
'It then clears the picture boxes and redisplays the sorted lists.

For Pass = 1 To CTR - 1
    For pos = 1 To CTR - Pass
        If MCPrice(pos) > MCPrice(pos + 1) Then
            TempPrice = MCPrice(pos)
            MCPrice(pos) = MCPrice(pos + 1)
            MCPrice(pos + 1) = TempPrice
            TempName = MCName(pos)
            MCName(pos) = MCName(pos + 1)
            MCName(pos + 1) = TempName
            Temp = InventoryM(pos)
            InventoryM(pos) = InventoryM(pos + 1)
            InventoryM(pos + 1) = Temp
        End If
    Next pos
Next Pass

picMen.Cls
picMen.Print "#"; Tab(6); "Name"; Tab(31); "Price"
picMen.Print "*********************************************"

For pos = 1 To CTR
    picMen.Print InventoryM(pos); Tab(6); MCName(pos); Tab(31); FormatCurrency(MCPrice(pos))
Next pos

'Next Column

For Pass = 1 To CTR2 - 1
    For pos2 = 1 To CTR2 - Pass
        If WCPrice(pos2) > WCPrice(pos2 + 1) Then
            TempPrice = WCPrice(pos2)
            WCPrice(pos2) = WCPrice(pos2 + 1)
            WCPrice(pos2 + 1) = TempPrice
            TempName = WCName(pos2)
            WCName(pos2) = WCName(pos2 + 1)
            WCName(pos2 + 1) = TempName
            Temp = InventoryW(pos2)
            InventoryW(pos2) = InventoryW(pos2 + 1)
            InventoryW(pos2 + 1) = Temp
        End If
    Next pos2
Next Pass

picWomen.Cls
picWomen.Print "#"; Tab(6); "Name"; Tab(31); "Price"
picWomen.Print "*********************************************"

For pos2 = 1 To CTR2
    picWomen.Print InventoryW(pos2); Tab(6); WCName(pos2); Tab(31); FormatCurrency(WCPrice(pos2))
Next pos2

'Next Column

For Pass = 1 To CTR3 - 1
    For pos3 = 1 To CTR3 - Pass
        If KCPrice(pos3) > KCPrice(pos3 + 1) Then
            TempPrice = KCPrice(pos3)
            KCPrice(pos3) = KCPrice(pos3 + 1)
            KCPrice(pos3 + 1) = TempPrice
            TempName = KCName(pos3)
            KCName(pos3) = KCName(pos3 + 1)
            KCName(pos3 + 1) = TempName
            Temp = InventoryK(pos3)
            InventoryK(pos3) = InventoryK(pos3 + 1)
            InventoryK(pos3 + 1) = Temp
        End If
    Next pos3
Next Pass

picKids.Cls
picKids.Print "#"; Tab(6); "Name"; Tab(31); "Price"
picKids.Print "*********************************************"

For pos3 = 1 To CTR3
    picKids.Print InventoryK(pos3); Tab(6); KCName(pos3); Tab(31); FormatCurrency(KCPrice(pos3))
Next pos3

End Sub
