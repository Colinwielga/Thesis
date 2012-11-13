VERSION 5.00
Begin VB.Form frmAccessories 
   Caption         =   "Ben's Hockey Goods"
   ClientHeight    =   7275
   ClientLeft      =   17385
   ClientTop       =   450
   ClientWidth     =   7620
   LinkTopic       =   "Form1"
   Picture         =   "frmAccessories.frx":0000
   ScaleHeight     =   7275
   ScaleWidth      =   7620
   Begin VB.Frame fraAccessories 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Welcome to the Accessories Department"
      Height          =   7215
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7575
      Begin VB.CommandButton cmdAddtoCart 
         BackColor       =   &H0000FFFF&
         Caption         =   "Make a Purchase"
         BeginProperty Font 
            Name            =   "Goudy Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   4680
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1080
         Width           =   2415
      End
      Begin VB.CommandButton cmdBegin 
         BackColor       =   &H0000FFFF&
         Caption         =   "Display Accessories"
         BeginProperty Font 
            Name            =   "Goudy Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   4680
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   240
         Width           =   2415
      End
      Begin VB.PictureBox picAccessories 
         BackColor       =   &H00FFFFFF&
         Height          =   4575
         Left            =   120
         ScaleHeight     =   4515
         ScaleWidth      =   4155
         TabIndex        =   6
         Top             =   1080
         Width           =   4215
      End
      Begin VB.CommandButton cmdSortbyPrice 
         BackColor       =   &H0000FFFF&
         Caption         =   "Sort by Price"
         BeginProperty Font 
            Name            =   "Goudy Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   6000
         Width           =   1815
      End
      Begin VB.CommandButton cmdSortbyName 
         BackColor       =   &H0000FFFF&
         Caption         =   "Sort by Name"
         BeginProperty Font 
            Name            =   "Goudy Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   2520
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   6000
         Width           =   1815
      End
      Begin VB.Frame fraDepts 
         BackColor       =   &H00FFFFFF&
         Caption         =   "When you are done shopping,"
         Height          =   1935
         Left            =   4680
         TabIndex        =   1
         Top             =   1920
         Width           =   2775
         Begin VB.CommandButton cmdReturn 
            BackColor       =   &H000000FF&
            Caption         =   "Leave the Accessories Department"
            BeginProperty Font 
               Name            =   "Goudy Old Style"
               Size            =   9.75
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
            Height          =   1095
            Left            =   120
            TabIndex        =   3
            Top             =   360
            Width           =   2535
         End
      End
      Begin VB.Image imageaccessories 
         Height          =   1155
         Left            =   5280
         Picture         =   "frmAccessories.frx":17B05
         Top             =   4680
         Width           =   1500
      End
      Begin VB.Label lblSelectItem 
         BackColor       =   &H00FFFFFF&
         Caption         =   $"frmAccessories.frx":198E1
         Height          =   735
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   4215
      End
   End
End
Attribute VB_Name = "frmAccessories"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Ben's Hockey Store
'frmAccessories
'Ben Bartelt
'3/26/08
'The purupose of this form is to display hockey accessories the are not 100% necessary but some things like a mouth guard
'pucks, and a jock-strap are highly required to play the game of hockey.
'The users is also allowed to arrange the accessories in alphabetical order and from lowest to highest price
'I have comments under most subroutines to describe what each button is doing.
Option Explicit
Dim CTR As Integer, pos As Integer, AName(0 To 50) As String, APrice(0 To 50) As Single
Dim InventoryA(0 To 50) As Integer
Private Sub cmdAddtoCart_Click()
Dim counter As Integer, Cart(0 To 200) As Integer, AdditiontoCart As Integer, X As Integer
'This subroutine loads users item selections from an inputbox into a 'accessories' array and then to a file
'for display during checkout. If the user puts anything in their cart, the program will allow
'the user to now use the "proceed to checkout" button in the department selection form.
Open App.Path & "\AccessoriesCart.txt" For Output As #1

AdditiontoCart = InputBox("Would you like to make a purchase? Input the item's number, type '-1' to quit.", "Add to cart")

Do While AdditiontoCart <> -1
    counter = counter + 1
    'If number doeesn't correspond to the restrictions they will be warned by message box to try again.
    If AdditiontoCart < 88 Or AdditiontoCart > 105 Then
        MsgBox "That item number does not correspond to an item from this department.  Please try again.", , "Wrong Item Number"
    End If
    'if the users adds anything to cart the user now has access to proceed to checkout on the departments form.
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
'This subroutine opens the file for Accessories into arrays and displays them in a picture box.

Open App.Path & "\accessories.txt" For Input As #1

CTR = 0

Do While Not EOF(1)
    CTR = CTR + 1
    Input #1, InventoryA(CTR), AName(CTR), APrice(CTR)
Loop
Close #1

picAccessories.Cls
picAccessories.Print "#"; Tab(6); "Name"; Tab(45); "Price"
picAccessories.Print "********************************************************************"

For pos = 1 To CTR
    picAccessories.Print InventoryA(pos); Tab(6); AName(pos); Tab(45); FormatCurrency(APrice(pos))
Next pos
End Sub

Private Sub cmdReturn_Click()
'This subroutine shows the department selection form and hides the accessories department form.
frmAccessories.Hide
frmDepartments.Show
End Sub

Private Sub cmdSortbyName_Click()
Dim Pass As Integer, Temp As Integer, TempName As String, TempPrice As Single
'This subroutine sorts the accessories list. It sorts by name, clears the picture boxes and redisplays the sorted lists.

For Pass = 1 To CTR - 1
    For pos = 1 To CTR - Pass
        If AName(pos) > AName(pos + 1) Then
            TempPrice = APrice(pos)
            APrice(pos) = APrice(pos + 1)
            APrice(pos + 1) = TempPrice
            TempName = AName(pos)
            AName(pos) = AName(pos + 1)
            AName(pos + 1) = TempName
            Temp = InventoryA(pos)
            InventoryA(pos) = InventoryA(pos + 1)
            InventoryA(pos + 1) = Temp
        End If
    Next pos
Next Pass

picAccessories.Cls
picAccessories.Print "#"; Tab(6); "Name"; Tab(45); "Price"
picAccessories.Print "********************************************************************"

For pos = 1 To CTR
    picAccessories.Print InventoryA(pos); Tab(6); AName(pos); Tab(45); FormatCurrency(APrice(pos))
Next pos
End Sub

Private Sub cmdSortbyPrice_Click()
Dim Pass As Integer, Temp As Integer, TempName As String, TempPrice As Single
'This subroutine sorts the accessories list by price, clears the pic box then redisplays by name.

For Pass = 1 To CTR - 1
    For pos = 1 To CTR - Pass
        If APrice(pos) > APrice(pos + 1) Then
            TempPrice = APrice(pos)
            APrice(pos) = APrice(pos + 1)
            APrice(pos + 1) = TempPrice
            TempName = AName(pos)
            AName(pos) = AName(pos + 1)
            AName(pos + 1) = TempName
            Temp = InventoryA(pos)
            InventoryA(pos) = InventoryA(pos + 1)
            InventoryA(pos + 1) = Temp
        End If
    Next pos
Next Pass

picAccessories.Cls
picAccessories.Print "#"; Tab(6); "Name"; Tab(45); "Price"
picAccessories.Print "********************************************************************"

For pos = 1 To CTR
    picAccessories.Print InventoryA(pos); Tab(6); AName(pos); Tab(45); FormatCurrency(APrice(pos))
Next pos
End Sub

Private Sub imageaccessories_Click()
'just a image of hockey accessories
End Sub
