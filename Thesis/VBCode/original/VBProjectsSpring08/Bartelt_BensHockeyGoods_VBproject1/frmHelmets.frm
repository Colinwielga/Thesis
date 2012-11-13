VERSION 5.00
Begin VB.Form frmHelmets 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Ben's Hockey Goods"
   ClientHeight    =   7275
   ClientLeft      =   15585
   ClientTop       =   8325
   ClientWidth     =   9645
   LinkTopic       =   "Form1"
   ScaleHeight     =   7275
   ScaleWidth      =   9645
   Begin VB.Frame fraHelmet 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Welcome to the Helmets Department"
      Height          =   6975
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   9255
      Begin VB.CommandButton cmdAddtoCart 
         BackColor       =   &H00FFFF00&
         Caption         =   "Make a Purchase"
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
         Left            =   240
         MaskColor       =   &H00FFFF80&
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1200
         Width           =   3975
      End
      Begin VB.CommandButton cmdBegin 
         BackColor       =   &H00FFFF80&
         Caption         =   "Display Helmets"
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
         Left            =   4800
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   360
         Width           =   3375
      End
      Begin VB.PictureBox picHelmets 
         BackColor       =   &H00FFFFFF&
         Height          =   4935
         Left            =   240
         ScaleHeight     =   4875
         ScaleWidth      =   4155
         TabIndex        =   6
         Top             =   1920
         Width           =   4215
      End
      Begin VB.CommandButton cmdSortbyPrice 
         BackColor       =   &H00FFFF80&
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
         Left            =   4800
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   960
         Width           =   1815
      End
      Begin VB.CommandButton cmdSortbyName 
         BackColor       =   &H00FFFF80&
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
         Left            =   6720
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   960
         Width           =   1455
      End
      Begin VB.Frame fraDepts 
         BackColor       =   &H00FFFFFF&
         Caption         =   "When you are done shopping,"
         Height          =   1935
         Left            =   5400
         TabIndex        =   1
         Top             =   4800
         Width           =   2775
         Begin VB.CommandButton cmdReturn 
            BackColor       =   &H000000FF&
            Caption         =   "Leave the Helmets Department"
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
            TabIndex        =   2
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
            TabIndex        =   3
            Top             =   240
            Width           =   2535
         End
      End
      Begin VB.Image ImageHelmet 
         Height          =   5340
         Left            =   4200
         Picture         =   "frmHelmets.frx":0000
         Top             =   240
         Width           =   5400
      End
      Begin VB.Label lblSelectItem 
         BackColor       =   &H00FFFFFF&
         Caption         =   $"frmHelmets.frx":CD90
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
'Ben's Hockey Store
'frmHelmets
'Ben Bartelt
'3/26/08
'This form allows the user to purchases Helmets. It all sorts the Helmets by price and name
'The form also allows the user to go back to the departments form so they can proceed to checkout.
'I have other comments under most of the subroutines.
Option Explicit
Dim CTR As Integer, pos As Integer, HName(0 To 50) As String, HPrice(0 To 50) As Single
Dim InventoryH(0 To 50) As Integer
Private Sub cmdAddtoCart_Click()
Dim counter As Integer, Cart(0 To 100) As Integer, AdditiontoCart As Integer, X As Integer
'This subroutine loads users item selections from an inputbox into a 'Helmets' array.
' If an item is added it allows the user to proceed to checkout on the departments form.
Open App.Path & "\Helmetscart.txt" For Output As #1

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
Open App.Path & "\Helmets.txt" For Input As #1
'This subroutine opens the helmets file and displays the items in arrays.

CTR = 0

Do While Not EOF(1)
    CTR = CTR + 1
    Input #1, InventoryH(CTR), HName(CTR), HPrice(CTR)
Loop
Close #1

picHelmets.Cls
picHelmets.Print "#"; Tab(6); "Name"; Tab(45); "Price"
picHelmets.Print "********************************************************************"

For pos = 1 To CTR
    picHelmets.Print InventoryH(pos); Tab(6); HName(pos); Tab(45); FormatCurrency(HPrice(pos))
Next pos
End Sub

Private Sub cmdReturn_Click()
'This subroutine shows the department selection form and hides the Helmet department form.
frmHelmets.Hide
frmDepartments.Show
End Sub

Private Sub cmdSortbyName_Click()
Dim Pass As Integer, Temp As Integer, TempName As String, TempPrice As Single
'This subroutine sorts the Helmets by name then clears and redisplays
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

picHelmets.Cls
picHelmets.Print "#"; Tab(6); "Name"; Tab(45); "Price"
picHelmets.Print "********************************************************************"

For pos = 1 To CTR
    picHelmets.Print InventoryH(pos); Tab(6); HName(pos); Tab(45); FormatCurrency(HPrice(pos))
Next pos
End Sub

Private Sub cmdSortbyPrice_Click()
Dim Pass As Integer, Temp As Integer, TempName As String, TempPrice As Single
'This subroutine sorts the Helmets items list by price it then clears so can be redisplayed

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

picHelmets.Cls
picHelmets.Print "#"; Tab(6); "Name"; Tab(45); "Price"
picHelmets.Print "********************************************************************"

For pos = 1 To CTR
    picHelmets.Print InventoryH(pos); Tab(6); HName(pos); Tab(45); FormatCurrency(HPrice(pos))
Next pos

End Sub

