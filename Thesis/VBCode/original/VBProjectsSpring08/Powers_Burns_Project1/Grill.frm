VERSION 5.00
Begin VB.Form Grill 
   Caption         =   "Form3"
   ClientHeight    =   5265
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10635
   LinkTopic       =   "Form3"
   Picture         =   "Grill.frx":0000
   ScaleHeight     =   5265
   ScaleWidth      =   10635
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command20 
      BackColor       =   &H0000FFFF&
      Caption         =   "Quit"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton Command19 
      BackColor       =   &H000000C0&
      Caption         =   "Clear Items and Return to Sotre"
      Enabled         =   0   'False
      Height          =   855
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   3720
      Width           =   2055
   End
   Begin VB.CommandButton Command18 
      BackColor       =   &H000000C0&
      Caption         =   "Keep Items and Retun to Store"
      Enabled         =   0   'False
      Height          =   855
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   2760
      Width           =   2055
   End
   Begin VB.CommandButton Command17 
      BackColor       =   &H000000C0&
      Caption         =   "Clear"
      Enabled         =   0   'False
      Height          =   855
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   1800
      Width           =   2055
   End
   Begin VB.CommandButton Command16 
      BackColor       =   &H000000C0&
      Caption         =   "Start"
      Height          =   855
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   840
      Width           =   2055
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   5055
      Left            =   7200
      ScaleHeight     =   4995
      ScaleWidth      =   3315
      TabIndex        =   16
      Top             =   120
      Width           =   3375
   End
   Begin VB.CommandButton Command15 
      BackColor       =   &H0000C000&
      Caption         =   "Pretzel / Cheesebread"
      Enabled         =   0   'False
      Height          =   735
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   4080
      Width           =   1455
   End
   Begin VB.CommandButton Command14 
      BackColor       =   &H0000C000&
      Caption         =   "Whole Pizza"
      Enabled         =   0   'False
      Height          =   735
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   4080
      Width           =   1455
   End
   Begin VB.CommandButton Command13 
      BackColor       =   &H0000C000&
      Caption         =   "Pizza Slice"
      Enabled         =   0   'False
      Height          =   735
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4080
      Width           =   1455
   End
   Begin VB.CommandButton Command12 
      BackColor       =   &H0000C000&
      Caption         =   "Shrimp Basket"
      Enabled         =   0   'False
      Height          =   735
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3240
      Width           =   1455
   End
   Begin VB.CommandButton Command11 
      BackColor       =   &H0000C000&
      Caption         =   "Chicken Basket"
      Enabled         =   0   'False
      Height          =   735
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3240
      Width           =   1455
   End
   Begin VB.CommandButton Command10 
      BackColor       =   &H0000C000&
      Caption         =   "Tator Tots"
      Enabled         =   0   'False
      Height          =   735
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3240
      Width           =   1455
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H0000C000&
      Caption         =   "Onion Rings"
      Enabled         =   0   'False
      Height          =   735
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2400
      Width           =   1455
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H0000C000&
      Caption         =   "Fries"
      Enabled         =   0   'False
      Height          =   735
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2400
      Width           =   1455
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H0000C000&
      Caption         =   "Grilled Chicken"
      Enabled         =   0   'False
      Height          =   735
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2400
      Width           =   1455
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H0000C000&
      Caption         =   "Hot Dog"
      Enabled         =   0   'False
      Height          =   735
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1560
      Width           =   1455
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H0000C000&
      Caption         =   "Chicken Burger w/Cheese"
      Enabled         =   0   'False
      Height          =   735
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1560
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H0000C000&
      Caption         =   "Chicken Burger"
      Enabled         =   0   'False
      Height          =   735
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1560
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H0000C000&
      Caption         =   "Bacon Burger"
      Enabled         =   0   'False
      Height          =   735
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   720
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0000C000&
      Caption         =   "Cheeseburger"
      Enabled         =   0   'False
      Height          =   735
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   720
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000C000&
      Caption         =   "Hamburger"
      Enabled         =   0   'False
      Height          =   735
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   720
      Width           =   1455
   End
   Begin VB.Line Line2 
      X1              =   7080
      X2              =   7080
      Y1              =   120
      Y2              =   5160
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Grill Items"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   15
      Top             =   240
      Width           =   1335
   End
   Begin VB.Line Line1 
      X1              =   4800
      X2              =   4800
      Y1              =   600
      Y2              =   4800
   End
End
Attribute VB_Name = "Grill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim Subtotal As Single

Private Sub Command1_Click()
Dim Ham As Single
Ham = 2.55
Subtotal = Subtotal + Ham
Picture1.Print "Hamburger", FormatCurrency(Subtotal)
Command17.Enabled = True
Command18.Enabled = True
Command19.Enabled = True
End Sub

Private Sub Command10_Click()
Dim tots As Single
tots = 1.25
Subtotal = Subtotal + tots
Picture1.Print "Tator Tots "; FormatCurrency(Subtotal)
Command17.Enabled = True
Command18.Enabled = True
Command19.Enabled = True
End Sub

Private Sub Command11_Click()
Dim Cbas As Single
Cbas = 4.25
Subtotal = Subtotal + Cbas
Picture1.Print "Chicken Basket "; FormatCurrency(Subtotal)
Command17.Enabled = True
Command18.Enabled = True
Command19.Enabled = True
End Sub

Private Sub Command12_Click()
Dim Shrimp As Single
Shrimp = 5.25
Subtotal = Subtotal + Shrimp
Picture1.Print "Shrimp Basket "; FormatCurrency(Subtotal)
Command17.Enabled = True
Command18.Enabled = True
Command19.Enabled = True
End Sub

Private Sub Command13_Click()
Dim PS As Single
PS = 2#
Subtotal = Subtotal + PS
Picture1.Print "Pizza Slice "; FormatCurrency(Subtotal)
Command17.Enabled = True
Command18.Enabled = True
Command19.Enabled = True
End Sub

Private Sub Command14_Click()
Dim WP As Single
WP = 13.5
Subtotal = Subtotal + WP
Picture1.Print "Whole Pizza "; FormatCurrency(Subtotal)
Command17.Enabled = True
Command18.Enabled = True
Command19.Enabled = True
End Sub

Private Sub Command15_Click()
Dim Pret As Single
Pret = 1.75
Subtotal = Subtotal + Pret
Picture1.Print "Pretzal/Cheesey Bread "; FormatCurrency(Subtotal)
Command17.Enabled = True
Command18.Enabled = True
Command19.Enabled = True
End Sub

Private Sub Command16_Click()
Picture1.Print "Item", "Subtotal"
Picture1.Print "****************************"
Command1.Enabled = True
Command2.Enabled = True
Command3.Enabled = True
Command4.Enabled = True
Command5.Enabled = True
Command6.Enabled = True
Command7.Enabled = True
Command8.Enabled = True
Command9.Enabled = True
Command10.Enabled = True
Command11.Enabled = True
Command12.Enabled = True
Command13.Enabled = True
Command14.Enabled = True
Command15.Enabled = True
End Sub

Private Sub Command17_Click()
Picture1.Cls
Subtotal = 0
Command1.Enabled = False
Command2.Enabled = False
Command3.Enabled = False
Command4.Enabled = False
Command5.Enabled = False
Command6.Enabled = False
Command7.Enabled = False
Command8.Enabled = False
Command9.Enabled = False
Command10.Enabled = False
Command11.Enabled = False
Command12.Enabled = False
Command13.Enabled = False
Command14.Enabled = False
Command15.Enabled = False
Command17.Enabled = False
Command18.Enabled = False
Command19.Enabled = False
End Sub

Private Sub Command18_Click()
Dim answer As Single
answer = InputBox("Do you want to checkout? If yes you can't get anymore grill items. If yes enter 1, if no enter 0?")
If answer = 1 Then
    GrillTotal = Subtotal
    Home.Show
    Grill.Hide
    Home.Command7.Enabled = False
    Home.Command13.Enabled = True
    ElseIf answer = 0 Then
Else
End If
End Sub

Private Sub Command19_Click()
Subtotal = 0
Home.Show
Grill.Hide
End Sub

Private Sub Command2_Click()
Dim CZ As Single
CZ = 2.65
Subtotal = Subtotal + CZ
Picture1.Print "Cheeseburger "; FormatCurrency(Subtotal)
Command17.Enabled = True
Command18.Enabled = True
Command19.Enabled = True
End Sub

Private Sub Command20_Click()
End
End Sub

Private Sub Command3_Click()
Dim BB As Single
BB = 2.9
Subtotal = Subtotal + BB
Picture1.Print "Bacon Burger "; FormatCurrency(Subtotal)
Command17.Enabled = True
Command18.Enabled = True
Command19.Enabled = True
End Sub

Private Sub Command4_Click()
Dim CB As Single
CB = 2.95
Subtotal = Subtotal + CB
Picture1.Print "Chicken Burger "; FormatCurrency(Subtotal)
Command17.Enabled = True
Command18.Enabled = True
Command19.Enabled = True
End Sub

Private Sub Command5_Click()
Dim CBc As Single
CBc = 3.45
Subtotal = Subtotal + CBc
Picture1.Print "Chicken Burger w/c"; FormatCurrency(Subtotal)
Command17.Enabled = True
Command18.Enabled = True
Command19.Enabled = True
End Sub

Private Sub Command6_Click()
Dim Dog As Single
Dog = 1.75
Subtotal = Subtotal + Dog
Picture1.Print "Hot Dog "; FormatCurrency(Subtotal)
Command17.Enabled = True
Command18.Enabled = True
Command19.Enabled = True
End Sub

Private Sub Command7_Click()
Dim GC As Single
GC = 3.05
Subtotal = Subtotal + GC
Picture1.Print "Grilled Chicken "; FormatCurrency(Subtotal)
Command17.Enabled = True
Command18.Enabled = True
Command19.Enabled = True
End Sub

Private Sub Command8_Click()
Dim Fry As Single
Fry = 1#
Subtotal = Subtotal + Fry
Picture1.Print "Fries ", FormatCurrency(Subtotal)
Command17.Enabled = True
Command18.Enabled = True
Command19.Enabled = True
End Sub

Private Sub Command9_Click()
Dim Os As Single
Os = 1.75
Subtotal = Subtotal + Os
Picture1.Print "Onion Rings "; FormatCurrency(Subtotal)
Command17.Enabled = True
Command18.Enabled = True
Command19.Enabled = True
End Sub
