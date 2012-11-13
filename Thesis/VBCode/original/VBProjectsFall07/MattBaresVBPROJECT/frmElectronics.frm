VERSION 5.00
Begin VB.Form frmElectronics 
   Caption         =   "Target"
   ClientHeight    =   6855
   ClientLeft      =   4170
   ClientTop       =   4545
   ClientWidth     =   9525
   LinkTopic       =   "Form1"
   ScaleHeight     =   6855
   ScaleWidth      =   9525
   Begin VB.Frame fraToys 
      Caption         =   "Welcome to the Toy Department"
      Height          =   6615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9255
      Begin VB.Frame fraDepts 
         Caption         =   "When you are done shopping,"
         Height          =   1935
         Left            =   6360
         TabIndex        =   6
         Top             =   4560
         Width           =   2775
         Begin VB.CommandButton cmdReturn 
            Caption         =   "Leave the Electronics Department"
            Height          =   615
            Left            =   120
            TabIndex        =   7
            Top             =   1080
            Width           =   2535
         End
         Begin VB.Label lblWhenDone 
            Caption         =   "you can return to the department selection menu by pressing the button below."
            Height          =   1095
            Left            =   120
            TabIndex        =   8
            Top             =   240
            Width           =   2535
         End
      End
      Begin VB.CommandButton cmdSortbyName 
         Caption         =   "Sort by Name"
         Height          =   735
         Left            =   120
         TabIndex        =   5
         Top             =   4560
         Width           =   1815
      End
      Begin VB.CommandButton cmdSortbyPrice 
         Caption         =   "Sort by Price"
         Height          =   735
         Left            =   120
         TabIndex        =   4
         Top             =   3120
         Width           =   1815
      End
      Begin VB.PictureBox picElectronics 
         Height          =   4575
         Left            =   2040
         ScaleHeight     =   4515
         ScaleWidth      =   4155
         TabIndex        =   3
         Top             =   1920
         Width           =   4215
      End
      Begin VB.CommandButton cmdBegin 
         Caption         =   "Display Electronics"
         Height          =   495
         Left            =   120
         TabIndex        =   2
         Top             =   1200
         Width           =   9015
      End
      Begin VB.CommandButton cmdAddtoCart 
         Caption         =   "Add Item to Cart"
         Height          =   735
         Left            =   4320
         TabIndex        =   1
         Top             =   240
         Width           =   4695
      End
      Begin VB.PictureBox picBkgd 
         Height          =   5535
         Left            =   0
         Picture         =   "frmElectronics.frx":0000
         ScaleHeight     =   5475
         ScaleWidth      =   9195
         TabIndex        =   9
         Top             =   1080
         Width           =   9255
      End
      Begin VB.Label lblSelectItem 
         Caption         =   $"frmElectronics.frx":128E9
         Height          =   735
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   4215
      End
   End
End
Attribute VB_Name = "frmElectronics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CTR As Integer, pos As Integer, EName(0 To 50) As String, EPrice(0 To 50) As Single
Dim InventoryE(0 To 50) As Integer
Private Sub cmdAddtoCart_Click()
Dim counter As Integer, Cart(0 To 100) As Integer, AdditiontoCart As Integer, X As Integer
'This subroutine loads users item selections from an inputbox into a 'electronics cart' array and then to a file
'for display during checkout. If the user puts anything in their cart, the program will allow
'the user to now use the "proceed to checkout" button in the department selection form.
Open App.Path & "\ElectronicsCart.txt" For Output As #1

AdditiontoCart = InputBox("Would you like to make a purchase? Input the item's number, type '-1' to quit.", "Add to cart")

Do While AdditiontoCart <> -1
    counter = counter + 1
    'If the user inputs a number that doesn't correspond to an item from this department, they will be warned
    'via a msgbox and asked to try again.
    If AdditiontoCart < 88 Or AdditiontoCart > 103 Then
        MsgBox "That item number does not correspond to an item from this department.  Please try again.", , "Wrong Item Number"
    End If
    'If the user puts anything in their cart, the program will allow
    'the user to now use the "proceed to checkout" button in the department selection form.
    If counter > 0 Then
        frmDepartments.cmdCheckout.Enabled = True
    End If
    
    Cart(counter) = AdditiontoCart
    AdditiontoCart = InputBox("Would you like to make another purchase? Input the item's number, type '-1' to quit.", "Add to cart")
Loop

For X = 1 To counter
    Print #1, Cart(X)
Next X

Close #1
End Sub

Private Sub cmdBegin_Click()
'This subroutine opens the file for electronics into 3 arrays and
'displays them in a picture box.

Open App.Path & "\Electronics.txt" For Input As #1

CTR = 0

Do While Not EOF(1)
    CTR = CTR + 1
    Input #1, InventoryE(CTR), EName(CTR), EPrice(CTR)
Loop
Close #1

picElectronics.Cls
picElectronics.Print "#"; Tab(6); "Name"; Tab(45); "Price"
picElectronics.Print "********************************************************************"

For pos = 1 To CTR
    picElectronics.Print InventoryE(pos); Tab(6); EName(pos); Tab(45); FormatCurrency(EPrice(pos))
Next pos
End Sub

Private Sub cmdReturn_Click()
'This subroutine shows the department selection form and hides the electronics department form.
frmElectronics.Hide
frmDepartments.Show
End Sub

Private Sub cmdSortbyName_Click()
Dim Pass As Integer, Temp As Integer, TempName As String, TempPrice As Single
'This subroutine sorts the electronics list by using bubble sort. It sorts by name.
'It then clears the picture boxes and redisplays the sorted lists.

For Pass = 1 To CTR - 1
    For pos = 1 To CTR - Pass
        If EName(pos) > EName(pos + 1) Then
            TempPrice = EPrice(pos)
            EPrice(pos) = EPrice(pos + 1)
            EPrice(pos + 1) = TempPrice
            TempName = EName(pos)
            EName(pos) = EName(pos + 1)
            EName(pos + 1) = TempName
            Temp = InventoryE(pos)
            InventoryE(pos) = InventoryE(pos + 1)
            InventoryE(pos + 1) = Temp
        End If
    Next pos
Next Pass

picElectronics.Cls
picElectronics.Print "#"; Tab(6); "Name"; Tab(45); "Price"
picElectronics.Print "********************************************************************"

For pos = 1 To CTR
    picElectronics.Print InventoryE(pos); Tab(6); EName(pos); Tab(45); FormatCurrency(EPrice(pos))
Next pos
End Sub

Private Sub cmdSortbyPrice_Click()
Dim Pass As Integer, Temp As Integer, TempName As String, TempPrice As Single
'This subroutine sorts the electronics list by using bubble sort. It sorts by price.
'It then clears the picture boxes and redisplays the sorted lists.

For Pass = 1 To CTR - 1
    For pos = 1 To CTR - Pass
        If EPrice(pos) > EPrice(pos + 1) Then
            TempPrice = EPrice(pos)
            EPrice(pos) = EPrice(pos + 1)
            EPrice(pos + 1) = TempPrice
            TempName = EName(pos)
            EName(pos) = EName(pos + 1)
            EName(pos + 1) = TempName
            Temp = InventoryE(pos)
            InventoryE(pos) = InventoryE(pos + 1)
            InventoryE(pos + 1) = Temp
        End If
    Next pos
Next Pass

picElectronics.Cls
picElectronics.Print "#"; Tab(6); "Name"; Tab(45); "Price"
picElectronics.Print "********************************************************************"

For pos = 1 To CTR
    picElectronics.Print InventoryE(pos); Tab(6); EName(pos); Tab(45); FormatCurrency(EPrice(pos))
Next pos
End Sub

