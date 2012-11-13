VERSION 5.00
Begin VB.Form frmPadding 
   Caption         =   "Ben's Hockey Goods"
   ClientHeight    =   11280
   ClientLeft      =   6360
   ClientTop       =   4500
   ClientWidth     =   14760
   LinkTopic       =   "Form1"
   ScaleHeight     =   11280
   ScaleWidth      =   14760
   Begin VB.Frame fraPadding 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Welcome to the Padding Department"
      Height          =   10095
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   12855
      Begin VB.PictureBox picShoulder 
         BackColor       =   &H00FFFFFF&
         Height          =   5895
         Left            =   8520
         ScaleHeight     =   5835
         ScaleWidth      =   3315
         TabIndex        =   15
         Top             =   2400
         Width           =   3375
      End
      Begin VB.PictureBox picSE 
         BackColor       =   &H00FFFFFF&
         Height          =   5895
         Left            =   3840
         ScaleHeight     =   5835
         ScaleWidth      =   3915
         TabIndex        =   14
         Top             =   2400
         Width           =   3975
      End
      Begin VB.PictureBox picBreezers 
         BackColor       =   &H00FFFFFF&
         Height          =   5895
         Left            =   240
         ScaleHeight     =   5835
         ScaleWidth      =   3195
         TabIndex        =   13
         Top             =   2400
         Width           =   3255
      End
      Begin VB.CommandButton cmdReturn 
         BackColor       =   &H000000FF&
         Caption         =   "Leave the Padding Department"
         BeginProperty Font 
            Name            =   "Goudy Old Style"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   8400
         Width           =   2535
      End
      Begin VB.CommandButton cmdBegin 
         BackColor       =   &H00FF00FF&
         Caption         =   "Display Padding"
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
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   480
         Width           =   3255
      End
      Begin VB.CommandButton cmdAddtoCart 
         BackColor       =   &H00FF00FF&
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
         Height          =   495
         Index           =   1
         Left            =   4800
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1200
         Width           =   2295
      End
      Begin VB.CommandButton cmdSortbyPrice 
         BackColor       =   &H00FF00FF&
         Caption         =   "Sort Padding by Price"
         BeginProperty Font 
            Name            =   "Goudy Old Style"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   10440
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton cmdSortbyName 
         BackColor       =   &H00FF00FF&
         Caption         =   "Sort Padding by Name"
         BeginProperty Font 
            Name            =   "Goudy Old Style"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   8640
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label lblSelectItem 
         BackColor       =   &H00FFFFFF&
         Caption         =   $"frmPadding.frx":0000
         Height          =   735
         Index           =   1
         Left            =   4080
         TabIndex        =   11
         Top             =   360
         Width           =   4335
      End
      Begin VB.Label lblBreezers 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Breezers"
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
         Left            =   1200
         TabIndex        =   10
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label lblSE 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Shin pads and Elbow Pads"
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
         Left            =   4560
         TabIndex        =   9
         Top             =   1920
         Width           =   2775
      End
      Begin VB.Label lblShoulder 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Shoulder Pads"
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
         Left            =   9480
         TabIndex        =   8
         Top             =   1920
         Width           =   1455
      End
      Begin VB.Label lblWhenDone 
         BackColor       =   &H00FFFFFF&
         Caption         =   "When you are done shopping in this department, you can return to the department selection menu by pressing the button to the left."
         BeginProperty Font 
            Name            =   "Goudy Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   3600
         TabIndex        =   7
         Top             =   8400
         Width           =   6015
      End
   End
   Begin VB.CommandButton cmdAddtoCart 
      Caption         =   "Add Item to Cart"
      Height          =   735
      Index           =   0
      Left            =   480
      Picture         =   "frmPadding.frx":00AC
      TabIndex        =   0
      Top             =   1680
      Width           =   2535
   End
   Begin VB.Label lblSelectItem 
      Caption         =   $"frmPadding.frx":1460B
      Height          =   735
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   4335
   End
End
Attribute VB_Name = "frmPadding"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Ben's Hockey Store
'frmPadding
'Ben Bartelt
'3/26/08
'This form allows the user to display the padding the is required by hockey players. There is three lists
'The first list are breezers or the pants of a hockey player. The second list is shin pads and elbow pads
'The final list is shoulder pads. It allows the user to purchase all of these types of equipment. Then when they are
'done they can proceed back to the departments form where they may checkout.
'I have other comments under most of the subroutines.
Option Explicit
Dim InventoryB(0 To 50) As Integer, InventorySE(0 To 50) As Integer, InventoryS(0 To 50) As Integer
Dim BName(0 To 50) As String, BPrice(0 To 50) As Single, SEName(0 To 50) As String, SEPrice(0 To 50) As Single, SName(0 To 50) As String, SPrice(0 To 50) As Single
Dim CTR3 As Integer, pos3 As Integer, CTR2 As Integer, pos2 As Integer, CTR As Integer, pos As Integer

Private Sub cmdAddtoCart_Click(Index As Integer)
'This subroutine loads users item selections from an inputbox into a padding array and then to a file
'for display during checkout. If the user puts anything in their cart, the program will allow
'the user to now use the "proceed to checkout" button in the department selection form.
Dim Cart(0 To 100) As Integer, AdditiontoCart As Integer, X As Integer, counter As Integer

Open App.Path & "\Paddingcart.txt" For Output As #1

AdditiontoCart = InputBox("Would you like to make a purchase? Input the item's number, type '-1' to quit.", "Add to cart")

Do While AdditiontoCart <> -1
    counter = counter + 1
    'If the user inputs a number that doesn't correspond to an item from this department, they will be warned
    'via a msgbox and asked to try again.
    If AdditiontoCart < -1 Or AdditiontoCart > 26 And AdditiontoCart < 106 Or AdditiontoCart > 125 Then
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
'This subroutine opens the file for Breezers, Shin, Elbow and Shoulders pads into 9 arrays and
'displays them in their relative picture box.

Open App.Path & "\Breezers.txt" For Input As #1

CTR = 0

Do While Not EOF(1)
    CTR = CTR + 1
    Input #1, InventoryB(CTR), BName(CTR), BPrice(CTR)
Loop
Close #1

picBreezers.Cls
picBreezers.Print "#"; Tab(6); "Name"; Tab(31); "Price"
picBreezers.Print "*********************************************"

For pos = 1 To CTR
    picBreezers.Print InventoryB(pos); Tab(6); BName(pos); Tab(31); FormatCurrency(BPrice(pos))
Next pos

'Next Column

Open App.Path & "\SPE.txt" For Input As #2

CTR2 = 0

Do While Not EOF(2)
    CTR2 = CTR2 + 1
    Input #2, InventorySE(CTR2), SEName(CTR2), SEPrice(CTR2)
Loop
Close #2

picSE.Cls
picSE.Print "#"; Tab(6); "Name"; Tab(31); "Price"
picSE.Print "***********************************************************"

For pos2 = 1 To CTR2
    picSE.Print InventorySE(pos2); Tab(6); SEName(pos2); Tab(31); FormatCurrency(SEPrice(pos2))
Next pos2

'Next Column

Open App.Path & "\Shoulder.txt" For Input As #3

CTR3 = 0

Do While Not EOF(3)
    CTR3 = CTR3 + 1
    Input #3, InventoryS(CTR3), SName(CTR3), SPrice(CTR3)
Loop
Close #3

picShoulder.Cls
picShoulder.Print "#"; Tab(6); "Name"; Tab(31); "Price"
picShoulder.Print "*********************************************"

For pos3 = 1 To CTR3
    picShoulder.Print InventoryS(pos3); Tab(6); SName(pos3); Tab(31); FormatCurrency(SPrice(pos3))
Next pos3

End Sub





Private Sub cmdReturn_Click()
'This subroutine shows the department selection form and hides the clothing department form.
frmPadding.Hide
frmDepartments.Show
End Sub

Private Sub cmdSortbyName_Click()
Dim Pass As Integer, Temp As Integer, TempName As String, TempPrice As Single
'This subroutine sorts the padding by name then clears to be resorted.
For Pass = 1 To CTR - 1
    For pos = 1 To CTR - Pass
        If BName(pos) > BName(pos + 1) Then
            TempPrice = BPrice(pos)
            BPrice(pos) = BPrice(pos + 1)
            BPrice(pos + 1) = TempPrice
            TempName = BName(pos)
            BName(pos) = BName(pos + 1)
            BName(pos + 1) = TempName
            Temp = InventoryB(pos)
            InventoryB(pos) = InventoryB(pos + 1)
            InventoryB(pos + 1) = Temp
        End If
    Next pos
Next Pass

picBreezers.Cls
picBreezers.Print "#"; Tab(6); "Name"; Tab(31); "Price"
picBreezers.Print "*********************************************"

For pos = 1 To CTR
    picBreezers.Print InventoryB(pos); Tab(6); BName(pos); Tab(31); FormatCurrency(BPrice(pos))
Next pos

'Next Column

For Pass = 1 To CTR2 - 1
    For pos2 = 1 To CTR2 - Pass
        If SEName(pos2) > SEName(pos2 + 1) Then
            TempPrice = SEPrice(pos2)
            SEPrice(pos2) = SEPrice(pos2 + 1)
            SEPrice(pos2 + 1) = TempPrice
            TempName = SEName(pos2)
            SEName(pos2) = SEName(pos2 + 1)
            SEName(pos2 + 1) = TempName
            Temp = InventorySE(pos2)
            InventorySE(pos2) = InventorySE(pos2 + 1)
            InventorySE(pos2 + 1) = Temp
        End If
    Next pos2
Next Pass

picSE.Cls
picSE.Print "#"; Tab(6); "Name"; Tab(31); "Price"
picSE.Print "*********************************************"

For pos2 = 1 To CTR2
    picSE.Print InventorySE(pos2); Tab(6); SEName(pos2); Tab(31); FormatCurrency(SEPrice(pos2))
Next pos2

'Next Column

For Pass = 1 To CTR3 - 1
    For pos3 = 1 To CTR3 - Pass
        If SName(pos3) > SName(pos3 + 1) Then
            TempPrice = SPrice(pos3)
            SPrice(pos3) = SPrice(pos3 + 1)
            SPrice(pos3 + 1) = TempPrice
            TempName = SName(pos3)
            SName(pos3) = SName(pos3 + 1)
            SName(pos3 + 1) = TempName
            Temp = InventoryS(pos3)
            InventoryS(pos3) = InventoryS(pos3 + 1)
            InventoryS(pos3 + 1) = Temp
        End If
    Next pos3
Next Pass

picShoulder.Cls
picShoulder.Print "#"; Tab(6); "Name"; Tab(31); "Price"
picShoulder.Print "*********************************************"

For pos3 = 1 To CTR3
    picShoulder.Print InventoryS(pos3); Tab(6); SName(pos3); Tab(31); FormatCurrency(SPrice(pos3))
Next pos3

End Sub

Private Sub cmdSortbyPrice_Click()
Dim Pass As Integer, Temp As Integer, TempName As String, TempPrice As Single
'This subroutine sorts all of the breezers by price then clears the picture boxes and redisplays the sorted lists.

For Pass = 1 To CTR - 1
    For pos = 1 To CTR - Pass
        If BPrice(pos) > BPrice(pos + 1) Then
            TempPrice = BPrice(pos)
            BPrice(pos) = BPrice(pos + 1)
            BPrice(pos + 1) = TempPrice
            TempName = BName(pos)
            BName(pos) = BName(pos + 1)
            BName(pos + 1) = TempName
            Temp = InventoryB(pos)
            InventoryB(pos) = InventoryB(pos + 1)
            InventoryB(pos + 1) = Temp
        End If
    Next pos
Next Pass

picBreezers.Cls
picBreezers.Print "#"; Tab(6); "Name"; Tab(31); "Price"
picBreezers.Print "*********************************************"

For pos = 1 To CTR
    picBreezers.Print InventoryB(pos); Tab(6); BName(pos); Tab(31); FormatCurrency(BPrice(pos))
Next pos

'Next Column

For Pass = 1 To CTR2 - 1
    For pos2 = 1 To CTR2 - Pass
        If SEPrice(pos2) > SEPrice(pos2 + 1) Then
            TempPrice = SEPrice(pos2)
            SEPrice(pos2) = SEPrice(pos2 + 1)
            SEPrice(pos2 + 1) = TempPrice
            TempName = SEName(pos2)
            SEName(pos2) = SEName(pos2 + 1)
            SEName(pos2 + 1) = TempName
            Temp = InventorySE(pos2)
            InventorySE(pos2) = InventorySE(pos2 + 1)
            InventorySE(pos2 + 1) = Temp
        End If
    Next pos2
Next Pass

picSE.Cls
picSE.Print "#"; Tab(6); "Name"; Tab(31); "Price"
picSE.Print "*********************************************"

For pos2 = 1 To CTR2
    picSE.Print InventorySE(pos2); Tab(6); SEName(pos2); Tab(31); FormatCurrency(SEPrice(pos2))
Next pos2

'Next Column

For Pass = 1 To CTR3 - 1
    For pos3 = 1 To CTR3 - Pass
        If SPrice(pos3) > SPrice(pos3 + 1) Then
            TempPrice = SPrice(pos3)
            SPrice(pos3) = SPrice(pos3 + 1)
            SPrice(pos3 + 1) = TempPrice
            TempName = SName(pos3)
            SName(pos3) = SName(pos3 + 1)
            SName(pos3 + 1) = TempName
            Temp = InventoryS(pos3)
            InventoryS(pos3) = InventoryS(pos3 + 1)
            InventoryS(pos3 + 1) = Temp
        End If
    Next pos3
Next Pass

picShoulder.Cls
picShoulder.Print "#"; Tab(6); "Name"; Tab(31); "Price"
picShoulder.Print "*********************************************"

For pos3 = 1 To CTR3
    picShoulder.Print InventoryS(pos3); Tab(6); SName(pos3); Tab(31); FormatCurrency(SPrice(pos3))
Next pos3

End Sub

