VERSION 5.00
Begin VB.Form frmHelmets 
   Caption         =   "Ben's Hockey Goods"
   ClientHeight    =   6675
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9270
   LinkTopic       =   "Form1"
   ScaleHeight     =   6675
   ScaleWidth      =   9270
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraKitchen 
      Caption         =   "Welcome to the Helmets Department"
      Height          =   6615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9255
      Begin VB.CommandButton cmdAddtoCart 
         Caption         =   "Add Item to Cart"
         Height          =   495
         Left            =   240
         TabIndex        =   8
         Top             =   1200
         Width           =   3975
      End
      Begin VB.CommandButton cmdBegin 
         Caption         =   "Display Helmets"
         Height          =   495
         Left            =   4800
         TabIndex        =   7
         Top             =   360
         Width           =   3375
      End
      Begin VB.PictureBox picHelmets 
         Height          =   4575
         Left            =   240
         ScaleHeight     =   4515
         ScaleWidth      =   4155
         TabIndex        =   6
         Top             =   1920
         Width           =   4215
      End
      Begin VB.CommandButton cmdSortbyPrice 
         Caption         =   "Sort by Price"
         Height          =   735
         Left            =   4800
         TabIndex        =   5
         Top             =   960
         Width           =   1815
      End
      Begin VB.CommandButton cmdSortbyName 
         Caption         =   "Sort by Name"
         Height          =   735
         Left            =   6720
         TabIndex        =   4
         Top             =   960
         Width           =   1455
      End
      Begin VB.Frame fraDepts 
         Caption         =   "When you are done shopping,"
         Height          =   1935
         Left            =   5400
         TabIndex        =   1
         Top             =   4560
         Width           =   2775
         Begin VB.CommandButton cmdReturn 
            Caption         =   "Leave the Helmets Department"
            Height          =   615
            Left            =   120
            TabIndex        =   2
            Top             =   1080
            Width           =   2535
         End
         Begin VB.Label lblWhenDone 
            Caption         =   "you can return to the department selection menu by pressing the button below."
            Height          =   1095
            Left            =   120
            TabIndex        =   3
            Top             =   240
            Width           =   2535
         End
      End
      Begin VB.Label lblSelectItem 
         Caption         =   $"frmHelments.frx":0000
         Height          =   735
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   4215
      End
   End
End
Attribute VB_Name = "frmHelmets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CTR As Integer, pos As Integer, HName(0 To 50) As String, HPrice(0 To 50) As Single
Dim InventoryH(0 To 50) As Integer
Private Sub cmdAddtoCart_Click()
Dim counter As Integer, Cart(0 To 100) As Integer, AdditiontoCart As Integer, X As Integer
'This subroutine loads users item selections from an inputbox into a 'Helmets' array and then to a file
'for display during checkout. If the user puts anything in their cart, the program will allow
'the user to now use the "proceed to checkout" button in the department selection form.
Open App.Path & "\Helmets.txt" For Output As #1

AdditiontoCart = InputBox("Would you like to make a purchase? Input the item's number, type '-1' to quit.", "Add to cart")

Do While AdditiontoCart <> -1
    counter = counter + 1
    'If the user inputs a number that doesn't correspond to an item from this department, they will be warned
    'via a msgbox and asked to try again.
    If AdditiontoCart < 41 Or AdditiontoCart > 60 Then
        MsgBox "That item number does not correspond to an item from this department.  Please try again.", , "Wrong Item Number"
    End If
    'If the user puts anything in their cart, the program will allow
    'the user to now use the "proceed to checkout" button in the department selection form.
    If counter > 0 Then
        frmDepartments.cmdCheckOut.Enabled = True
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
Open App.Path & "\Helmet.txt" For Input As #1
'This subroutine opens the file for Helmet items into 3 arrays and
'displays them in a picture box.

CTR = 0

Do While Not EOF(1)
    CTR = CTR + 1
    Input #1, InventoryH(CTR), HName(CTR), HPrice(CTR)
Loop
Close #1

picHelmet.Cls
picHelmet.Print "#"; Tab(6); "Name"; Tab(45); "Price"
picHelmet.Print "********************************************************************"

For pos = 1 To CTR
    picHelmet.Print InventoryH(pos); Tab(6); HName(pos); Tab(45); FormatCurrency(HPrice(pos))
Next pos
End Sub

Private Sub cmdReturn_Click()
'This subroutine shows the department selection form and hides the Helmet department form.
frmHelmet.Hide
frmDepartments.Show
End Sub

Private Sub cmdSortbyName_Click()
Dim Pass As Integer, Temp As Integer, TempName As String, TempPrice As Single
'This subroutine sorts the Helmet items list by using bubble sort. It sorts by name.
'It then clears the picture boxes and redisplays the sorted lists.

For Pass = 1 To CTR - 1
    For pos = 1 To CTR - Pass
        If HName(pos) > HName(pos + 1) Then
            TempPrice = HPrice(pos)
            HPrice(pos) = HPrice(pos + 1)
            HPrice(pos + 1) = TempPrice
            TempName = HName(pos)
            HName(pos) = HName(pos + 1)
            HName(pos + 1) = TempName
            Temp = InventoryH(pos)
            InventoryH(pos) = InventoryH(pos + 1)
            InventoryH(pos + 1) = Temp
        End If
    Next pos
Next Pass

picHelmet.Cls
picHelmet.Print "#"; Tab(6); "Name"; Tab(45); "Price"
picHelmet.Print "********************************************************************"

For pos = 1 To CTR
    picHelmet.Print InventoryH(pos); Tab(6); HName(pos); Tab(45); FormatCurrency(HPrice(pos))
Next pos
End Sub

Private Sub cmdSortbyPrice_Click()
Dim Pass As Integer, Temp As Integer, TempName As String, TempPrice As Single
'This subroutine sorts the Helmets items list by using bubble sort. It sorts by price.
'It then clears the picture boxes and redisplays the sorted lists.

For Pass = 1 To CTR - 1
    For pos = 1 To CTR - Pass
        If HPrice(pos) > HPrice(pos + 1) Then
            TempPrice = HPrice(pos)
            HPrice(pos) = HPrice(pos + 1)
            HPrice(pos + 1) = TempPrice
            TempName = HName(pos)
            HName(pos) = HName(pos + 1)
            HName(pos + 1) = TempName
            Temp = InventoryH(pos)
            InventoryH(pos) = InventoryH(pos + 1)
            InventoryH(pos + 1) = Temp
        End If
    Next pos
Next Pass

picHelmet.Cls
picHelmet.Print "#"; Tab(6); "Name"; Tab(45); "Price"
picHelmet.Print "********************************************************************"

For pos = 1 To CTR
    picHelmet.Print InventoryH(pos); Tab(6); HName(pos); Tab(45); FormatCurrency(HPrice(pos))
Next pos

End Sub

