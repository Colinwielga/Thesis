VERSION 5.00
Begin VB.Form frmSkates 
   Caption         =   "Ben's Hockey Goods"
   ClientHeight    =   6825
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9615
   LinkTopic       =   "Form1"
   ScaleHeight     =   6825
   ScaleWidth      =   9615
   Begin VB.Frame fraSkates 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Welcome to the Skates Department"
      Height          =   6615
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   9615
      Begin VB.Frame fraDepts 
         BackColor       =   &H00FFFFFF&
         Caption         =   "When you are done shopping,"
         Height          =   1935
         Left            =   5400
         TabIndex        =   6
         Top             =   4200
         Width           =   2775
         Begin VB.CommandButton cmdReturn 
            BackColor       =   &H000000FF&
            Caption         =   "Leave the Skates Department"
            BeginProperty Font 
               Name            =   "Goudy Old Style"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   1080
            Width           =   2535
         End
         Begin VB.Label lblWhenDone 
            BackColor       =   &H00FFFFFF&
            Caption         =   "you can return to the department selection menu by pressing the button below."
            BeginProperty Font 
               Name            =   "Goudy Old Style"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1095
            Left            =   120
            TabIndex        =   8
            Top             =   240
            Width           =   2535
         End
      End
      Begin VB.CommandButton cmdSortbyName 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Sort by Name"
         BeginProperty Font 
            Name            =   "Goudy Old Style"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6720
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   2520
         Width           =   2295
      End
      Begin VB.CommandButton cmdSortbyPrice 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Sort by Price"
         BeginProperty Font 
            Name            =   "Goudy Old Style"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4320
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   2520
         Width           =   2295
      End
      Begin VB.PictureBox picSkates 
         BackColor       =   &H00FFFFFF&
         Height          =   6375
         Left            =   0
         ScaleHeight     =   6315
         ScaleWidth      =   4155
         TabIndex        =   3
         Top             =   240
         Width           =   4215
      End
      Begin VB.CommandButton cmdBegin 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Display Skates"
         BeginProperty Font 
            Name            =   "Goudy Old Style"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4680
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1080
         Width           =   3975
      End
      Begin VB.CommandButton cmdAddtoCart 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Make a purchase"
         BeginProperty Font 
            Name            =   "Goudy Old Style"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   4320
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   1680
         Width           =   4695
      End
      Begin VB.Image imageskate 
         Height          =   6030
         Left            =   3600
         Picture         =   "frmSkates.frx":0000
         Top             =   960
         Width           =   6165
      End
      Begin VB.Label lblSelectItem 
         BackColor       =   &H00FFFFFF&
         Caption         =   $"frmSkates.frx":14CC3
         BeginProperty Font 
            Name            =   "Goudy Old Style"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   4680
         TabIndex        =   9
         Top             =   240
         Width           =   4215
      End
   End
End
Attribute VB_Name = "frmSkates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Ben's Hockey Store
'frmSkates
'Ben Bartelt
'3/26/08
'This form allows the user to purchases Skates. It sorts all of the skates by price and name of the skate.
'It then allows the user to select the skates they want to purchase. When they have selected one item they
'may return to the departments form so they can checkout.
'I have other comments under most of the subroutines.
Option Explicit
Dim CTR As Integer, pos As Integer, SName(0 To 50) As String, SPrice(0 To 50) As Single
Dim InventoryS(0 To 50) As Integer
Private Sub cmdAddtoCart_Click()
Dim counter As Integer, Cart(0 To 100) As Integer, AdditiontoCart As Integer, X As Integer
'This subroutine loads users item selections from an inputbox into a skates array and then to a file
'for display during checkout. If the user puts anything in their cart, the program will allow
'the user to now use the "proceed to checkout" button in the department selection form.
Open App.Path & "\Skatescart.txt" For Output As #1

AdditiontoCart = InputBox("Would you like to make a purchase? Input the item's number, type '-1' to quit.", "Add to cart")

Do While AdditiontoCart <> -1
    counter = counter + 1
    'If the user inputs a number that doesn't correspond to an item from this department, they will be warned
    'via a msgbox and asked to try again.
    If AdditiontoCart < 60 Or AdditiontoCart > 87 Then
        MsgBox "That item number does not correspond to an item from this department.  Please try again.", , "Wrong Item Number"
    End If
    'If the user puts anything in their cart, the program will allow
    'the user to now use the "proceed to checkout" button in the departmet selection form.
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
Open App.Path & "\Skates.txt" For Input As #1
'This subroutine opens the file for Skates into 3 arrays and
'displays them in a picture box.

CTR = 0

Do While Not EOF(1)
    CTR = CTR + 1
    Input #1, InventoryS(CTR), SName(CTR), SPrice(CTR)
Loop
Close #1

picSkates.Cls
picSkates.Print "#"; Tab(6); "Name"; Tab(45); "Price"
picSkates.Print "********************************************************************"

For pos = 1 To CTR
    picSkates.Print InventoryS(pos); Tab(6); SName(pos); Tab(45); FormatCurrency(SPrice(pos))
Next pos
End Sub

Private Sub cmdReturn_Click()
'This subroutine shows the department selection form and hides the toys department form.
frmSkates.Hide
frmDepartments.Show
End Sub

Private Sub cmdSortbyName_Click()
Dim Pass As Integer, Temp As Integer, TempName As String, TempPrice As Single
'This subroutine sorts the skates by using bubble sort. It sorts by name.
'It then clears the picture boxes and redisplays the sorted lists.

For Pass = 1 To CTR - 1
    For pos = 1 To CTR - Pass
        If SName(pos) > SName(pos + 1) Then
            TempPrice = SPrice(pos)
            SPrice(pos) = SPrice(pos + 1)
            SPrice(pos + 1) = TempPrice
            TempName = SName(pos)
            SName(pos) = SName(pos + 1)
            SName(pos + 1) = TempName
            Temp = InventoryS(pos)
            InventoryS(pos) = InventoryS(pos + 1)
            InventoryS(pos + 1) = Temp
        End If
    Next pos
Next Pass

picSkates.Cls
picSkates.Print "#"; Tab(6); "Name"; Tab(45); "Price"
picSkates.Print "********************************************************************"

For pos = 1 To CTR
    picSkates.Print InventoryS(pos); Tab(6); SName(pos); Tab(45); FormatCurrency(SPrice(pos))
Next pos
End Sub

Private Sub cmdSortbyPrice_Click()
Dim Pass As Integer, Temp As Integer, TempName As String, TempPrice As Single
'This subroutine sorts the skates by price then clears the picture boxes and redisplays the sorted lists.

For Pass = 1 To CTR - 1
    For pos = 1 To CTR - Pass
        If SPrice(pos) > SPrice(pos + 1) Then
            TempPrice = SPrice(pos)
            SPrice(pos) = SPrice(pos + 1)
            SPrice(pos + 1) = TempPrice
            TempName = SName(pos)
            SName(pos) = SName(pos + 1)
            SName(pos + 1) = TempName
            Temp = InventoryS(pos)
            InventoryS(pos) = InventoryS(pos + 1)
            InventoryS(pos + 1) = Temp
        End If
    Next pos
Next Pass

picSkates.Cls
picSkates.Print "#"; Tab(6); "Name"; Tab(45); "Price"
picSkates.Print "********************************************************************"

For pos = 1 To CTR
    picSkates.Print InventoryS(pos); Tab(6); SName(pos); Tab(45); FormatCurrency(SPrice(pos))
Next pos
End Sub

Private Sub Image1_Click()

End Sub
