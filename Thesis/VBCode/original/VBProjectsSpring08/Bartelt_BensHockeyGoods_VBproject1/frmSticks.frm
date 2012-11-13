VERSION 5.00
Begin VB.Form frmSticks 
   Caption         =   "Ben's Hockey Goods"
   ClientHeight    =   7005
   ClientLeft      =   285
   ClientTop       =   8550
   ClientWidth     =   9510
   LinkTopic       =   "Form1"
   ScaleHeight     =   7005
   ScaleWidth      =   9510
   Begin VB.Frame fraSticks 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Welcome to the Sticks Department"
      Height          =   6975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9495
      Begin VB.CommandButton cmdAddtoCart 
         BackColor       =   &H000080FF&
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
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1200
         Width           =   3735
      End
      Begin VB.CommandButton cmdBegin 
         BackColor       =   &H000080FF&
         Caption         =   "Display Sticks"
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
         TabIndex        =   5
         Top             =   360
         Width           =   4575
      End
      Begin VB.PictureBox picSticks 
         BackColor       =   &H00FFFFFF&
         Height          =   4575
         Left            =   4440
         ScaleHeight     =   4515
         ScaleWidth      =   4395
         TabIndex        =   4
         Top             =   1320
         Width           =   4455
      End
      Begin VB.CommandButton cmdReturn 
         BackColor       =   &H000000FF&
         Caption         =   "Leave the Sticks Department"
         BeginProperty Font 
            Name            =   "Goudy Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   6000
         Width           =   2535
      End
      Begin VB.CommandButton cmdSortbyPrice 
         BackColor       =   &H000080FF&
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
         Height          =   735
         Left            =   7080
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   6000
         Width           =   1815
      End
      Begin VB.CommandButton cmdSortbyName 
         BackColor       =   &H000080FF&
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
         Height          =   735
         Left            =   4440
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   6000
         Width           =   1815
      End
      Begin VB.Image imagecrossedsticks 
         Height          =   2700
         Left            =   720
         Picture         =   "frmSticks.frx":0000
         Top             =   2400
         Width           =   2700
      End
      Begin VB.Label lblSelectItem 
         BackColor       =   &H00FFFFFF&
         Caption         =   $"frmSticks.frx":FD6E
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
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   4215
      End
      Begin VB.Label lblSticks 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Sticks"
         BeginProperty Font 
            Name            =   "Goudy Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6240
         TabIndex        =   7
         Top             =   960
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmSticks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Ben's Hockey Store
'frmsticks
'Ben Bartelt
'3/26/08
'This form allows the user to purchases sticks. The sticks can be sorted by smallest price to largest or
'alphabetical order. They than can purchase their stick of choice. They then can proceed back to the departments
'form. Where they can then checkout.
'I have other comments under most of the subroutines.
Option Explicit
Dim CTR As Integer, pos As Integer, SName(0 To 50) As String, SPrice(0 To 50) As Single
Dim InventoryS(0 To 50) As Integer
Private Sub cmdAddtoCart_Click()
Dim counter As Integer, Cart(0 To 200) As Integer, AdditiontoCart As Integer, X As Integer
'This subroutine loads users item selections from an inputbox into a sticks array and then to a file
'for display during checkout. If the user puts anything in their cart, the program will allow
'the user to now use the "proceed to checkout" button in the department selection form.
Open App.Path & "\Stickscart.txt" For Output As #1

AdditiontoCart = InputBox("Would you like to make a purchase? Input the item's number, type '-1' to quit.", "Add to cart")

Do While AdditiontoCart <> -1
    counter = counter + 1
    'If the user inputs a number that doesn't correspond to an item from this department, they will be warned
    'via a msgbox and asked to try again.
    If AdditiontoCart < 27 Or AdditiontoCart > 40 Then
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
'This subroutine opens the file sticks into 3 arrays and displays them in a picture box.

Open App.Path & "\Sticks.txt" For Input As #1

CTR = 0

Do While Not EOF(1)
    CTR = CTR + 1
    Input #1, InventoryS(CTR), SName(CTR), SPrice(CTR)
Loop
Close #1

picSticks.Cls
picSticks.Print "#"; Tab(6); "Name"; Tab(45); "Price"
picSticks.Print "********************************************************************"

For pos = 1 To CTR
    picSticks.Print InventoryS(pos); Tab(6); SName(pos); Tab(45); FormatCurrency(SPrice(pos))
Next pos
End Sub

Private Sub cmdReturn_Click()
'This subroutine shows the department selection form and hides the furniture department form.
frmSticks.Hide
frmDepartments.Show
End Sub


Private Sub cmdSortbyName_Click()
Dim Pass As Integer, Temp As Integer, TempName As String, TempPrice As Single
'This subroutine sorts the sticks list by name then clears the picture boxes and redisplays the sorted lists.

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

picSticks.Cls
picSticks.Print "#"; Tab(6); "Name"; Tab(45); "Price"
picSticks.Print "********************************************************************"

For pos = 1 To CTR
    picSticks.Print InventoryS(pos); Tab(6); SName(pos); Tab(45); FormatCurrency(SPrice(pos))
Next pos

End Sub

Private Sub cmdSortbyPrice_Click()
Dim Pass As Integer, Temp As Integer, TempName As String, TempPrice As Single
'This subroutine sorts the sticks list by price it then clears the picture boxes and redisplays the sorted lists.

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

picSticks.Cls
picSticks.Print "#"; Tab(6); "Name"; Tab(45); "Price"
picSticks.Print "********************************************************************"

For pos = 1 To CTR
    picSticks.Print InventoryS(pos); Tab(6); SName(pos); Tab(45); FormatCurrency(SPrice(pos))
Next pos

End Sub

