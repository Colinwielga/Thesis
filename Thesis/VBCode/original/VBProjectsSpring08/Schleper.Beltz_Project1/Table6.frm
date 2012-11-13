VERSION 5.00
Begin VB.Form Table6 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8835
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10710
   LinkTopic       =   "Form1"
   Picture         =   "Table6.frx":0000
   ScaleHeight     =   8835
   ScaleWidth      =   10710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdKeyboard 
      Height          =   735
      Left            =   2640
      Picture         =   "Table6.frx":1619F
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   7920
      Width           =   1215
   End
   Begin VB.CommandButton cmdUpload 
      Caption         =   "Upload Text"
      Height          =   735
      Left            =   3840
      TabIndex        =   30
      Top             =   7920
      Width           =   1095
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "<==Tables"
      Height          =   615
      Left            =   240
      TabIndex        =   29
      Top             =   8040
      Width           =   1455
   End
   Begin VB.CommandButton cmdCheeseFries 
      Caption         =   "Cheese Fries"
      Height          =   615
      Index           =   0
      Left            =   3240
      TabIndex        =   23
      Top             =   4320
      Width           =   1455
   End
   Begin VB.CommandButton cmdHamburger 
      Caption         =   "Hamburger"
      Height          =   615
      Index           =   1
      Left            =   1800
      TabIndex        =   22
      Top             =   5400
      Width           =   1455
   End
   Begin VB.CommandButton cmdCheeseBurger 
      Caption         =   "CheeseBurger"
      Height          =   615
      Index           =   1
      Left            =   360
      TabIndex        =   21
      Top             =   5400
      Width           =   1455
   End
   Begin VB.CommandButton cmdSteak 
      Caption         =   "Steak"
      Height          =   615
      Index           =   1
      Left            =   1800
      TabIndex        =   20
      Top             =   6000
      Width           =   1455
   End
   Begin VB.CommandButton cmdSalmon 
      Caption         =   "Salmon"
      Height          =   615
      Index           =   1
      Left            =   360
      TabIndex        =   19
      Top             =   6000
      Width           =   1455
   End
   Begin VB.CommandButton cmdIceCream 
      Caption         =   "Ice Cream"
      Height          =   615
      Index           =   0
      Left            =   360
      TabIndex        =   18
      Top             =   7200
      Width           =   1455
   End
   Begin VB.CommandButton cmdCheeseSticks 
      Caption         =   "Cheese Sticks"
      Height          =   615
      Index           =   0
      Left            =   360
      TabIndex        =   17
      Top             =   4320
      Width           =   1455
   End
   Begin VB.CommandButton cmdPotatoSkins 
      Caption         =   "Potato Skins"
      Height          =   615
      Index           =   0
      Left            =   360
      TabIndex        =   16
      Top             =   3720
      Width           =   1455
   End
   Begin VB.CommandButton cmdBread 
      Caption         =   "Bread"
      Height          =   615
      Index           =   0
      Left            =   1800
      TabIndex        =   15
      Top             =   3720
      Width           =   1455
   End
   Begin VB.CommandButton cmdPepsi 
      Caption         =   "Pepsi"
      Height          =   615
      Index           =   0
      Left            =   360
      TabIndex        =   14
      Top             =   2640
      Width           =   1455
   End
   Begin VB.CommandButton cmdSprite 
      Caption         =   "Sprite"
      Height          =   615
      Index           =   0
      Left            =   1800
      TabIndex        =   13
      Top             =   2640
      Width           =   1455
   End
   Begin VB.CommandButton cmdMountainDew 
      Caption         =   "Mountain Dew"
      Height          =   615
      Index           =   0
      Left            =   1800
      TabIndex        =   12
      Top             =   2040
      Width           =   1455
   End
   Begin VB.CommandButton cmdRootBeer 
      Caption         =   "Root Beer"
      Height          =   615
      Index           =   0
      Left            =   360
      TabIndex        =   11
      Top             =   2040
      Width           =   1455
   End
   Begin VB.CommandButton cmdBuffaloWings 
      Caption         =   "Buffalo Wings"
      Height          =   615
      Index           =   0
      Left            =   1800
      TabIndex        =   10
      Top             =   4320
      Width           =   1455
   End
   Begin VB.CommandButton cmdNachos 
      Caption         =   "Nachos"
      Height          =   615
      Index           =   0
      Left            =   3240
      TabIndex        =   9
      Top             =   3720
      Width           =   1455
   End
   Begin VB.CommandButton cmdFruitPunch 
      Caption         =   "Fruit Punch"
      Height          =   615
      Index           =   0
      Left            =   3240
      TabIndex        =   8
      Top             =   2040
      Width           =   1455
   End
   Begin VB.CommandButton cmdWater 
      Caption         =   "Water"
      Height          =   615
      Index           =   0
      Left            =   3240
      TabIndex        =   7
      Top             =   2640
      Width           =   1455
   End
   Begin VB.CommandButton cmdRibs 
      Caption         =   "Ribs"
      Height          =   615
      Index           =   0
      Left            =   3240
      TabIndex        =   6
      Top             =   5400
      Width           =   1455
   End
   Begin VB.CommandButton cmdAlfredo 
      Caption         =   "Alfredo"
      Height          =   615
      Index           =   0
      Left            =   3240
      TabIndex        =   5
      Top             =   6000
      Width           =   1455
   End
   Begin VB.CommandButton cmdCheeseCake 
      Caption         =   "Cheesecake"
      Height          =   615
      Index           =   0
      Left            =   1800
      TabIndex        =   4
      Top             =   7200
      Width           =   1455
   End
   Begin VB.CommandButton cmdApplePie 
      Caption         =   "Apple Pie"
      Height          =   615
      Index           =   0
      Left            =   3240
      TabIndex        =   3
      Top             =   7200
      Width           =   1455
   End
   Begin VB.PictureBox picResults 
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6015
      Left            =   5640
      ScaleHeight     =   5955
      ScaleWidth      =   4515
      TabIndex        =   2
      Top             =   1800
      Width           =   4575
   End
   Begin VB.CommandButton cmdTotal 
      BackColor       =   &H0000FF00&
      Caption         =   "Total"
      Height          =   735
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7920
      Width           =   1695
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H000000FF&
      Caption         =   "Clear"
      Height          =   735
      Left            =   8760
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7920
      Width           =   1695
   End
   Begin VB.Label Label5 
      Caption         =   "Table #6"
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7080
      TabIndex        =   28
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Appetizers"
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1800
      TabIndex        =   27
      Top             =   3240
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Entrees"
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2040
      TabIndex        =   26
      Top             =   4920
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Drinks"
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2040
      TabIndex        =   25
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "Deserts"
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2040
      TabIndex        =   24
      Top             =   6600
      Width           =   1095
   End
End
Attribute VB_Name = "Table6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Vinnie Joe's Pub
'Table6
'Vinnie Schleper, Joey Beltz
'3/25/08
' this form shows the tables and all of the items that a customer may order.
'   When items are chosen the total may be found. Typing in messages is also
'   alowed.
Option Explicit
Dim runningTotal As Single
Private OldX As Integer
  Private OldY As Integer
  Private DragMode As Boolean
  Dim MoveMe As Boolean

  Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

     MoveMe = True
     OldX = X
     OldY = Y

 End Sub

 Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)


     If MoveMe = True Then
         Me.Left = Me.Left + (X - OldX)
         Me.Top = Me.Top + (Y - OldY)
     End If

 End Sub

 Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)


     Me.Left = Me.Left + (X - OldX)
     Me.Top = Me.Top + (Y - OldY)
     MoveMe = False

 End Sub


Private Sub cmdAlfredo_Click(Index As Integer)
picResults.Print "Alfredo", FormatCurrency(10.95, 2)
runningTotal = 10.95 + runningTotal
End Sub

Private Sub cmdApplePie_Click(Index As Integer)
picResults.Print "A.Pie", FormatCurrency(3, 2)
runningTotal = 3 + runningTotal
End Sub

Private Sub cmdBack_Click()
Table6.Hide
Tables.Show
End Sub

Private Sub cmdBread_Click(Index As Integer)
picResults.Print "Bread", FormatCurrency(3.95, 2)
runningTotal = 3.95 + runningTotal
End Sub

Private Sub cmdBuffaloWings_Click(Index As Integer)
picResults.Print "Wings", FormatCurrency(4.95, 2)
runningTotal = 4.95 + runningTotal
End Sub

Private Sub cmdCheeseBurger_Click(Index As Integer)
picResults.Print "Chzburger", FormatCurrency(6.95, 2)
runningTotal = 6.95 + runningTotal
End Sub

Private Sub cmdCheeseCake_Click(Index As Integer)
picResults.Print "Cheesecake", FormatCurrency(3, 2)
runningTotal = 3 + runningTotal
End Sub

Private Sub cmdCheeseFries_Click(Index As Integer)
picResults.Print "ChzFries", FormatCurrency(2.95, 2)
runningTotal = 2.95 + runningTotal
End Sub

Private Sub cmdCheeseSticks_Click(Index As Integer)
picResults.Print "ChzSticks", FormatCurrency(2.95, 2)
runningTotal = 2.95 + runningTotal
End Sub

Private Sub cmdClear_Click()
picResults.Cls
runningTotal = 0
End Sub

Private Sub cmdFruitPunch_Click(Index As Integer)
picResults.Print "FPunch", FormatCurrency(0.99, 2)
runningTotal = 0.99 + runningTotal
End Sub

Private Sub cmdHamburger_Click(Index As Integer)
picResults.Print "Hamburger", FormatCurrency(6.5, 2)
runningTotal = 6.5 + runningTotal
End Sub

Private Sub cmdIceCream_Click(Index As Integer)
picResults.Print "IceCream", FormatCurrency(3, 2)
runningTotal = 3# + runningTotal
End Sub

Private Sub cmdKeyboard_Click()
Table6.Hide
Keyboard6.Show
End Sub

Private Sub cmdMountainDew_Click(Index As Integer)
picResults.Print "MDew", FormatCurrency(0.99, 2)
runningTotal = 0.99 + runningTotal
End Sub

Private Sub cmdNachos_Click(Index As Integer)
picResults.Print "Nachos", FormatCurrency(5.95, 2)
runningTotal = 5.95 + runningTotal
End Sub

Private Sub cmdPepsi_Click(Index As Integer)
picResults.Print "Pepsi", FormatCurrency(0.99, 2)
runningTotal = 0.99 + runningTotal
End Sub

Private Sub cmdPotatoSkins_Click(Index As Integer)
picResults.Print "PSkins", FormatCurrency(5.95, 2)
runningTotal = 5.95 + runningTotal
End Sub

Private Sub cmdRibs_Click(Index As Integer)
picResults.Print "Ribs", FormatCurrency(10.95, 2)
runningTotal = 10.95 + runningTotal
End Sub

Private Sub cmdRootBeer_Click(Index As Integer)
picResults.Print "RBeer", FormatCurrency(0.99, 2)
runningTotal = 0.99 + runningTotal
End Sub

Private Sub cmdSalmon_Click(Index As Integer)
picResults.Print "Salmon", FormatCurrency(11.95, 2)
runningTotal = 11.95 + runningTotal
End Sub

Private Sub cmdSprite_Click(Index As Integer)
picResults.Print "Sprite", FormatCurrency(0.99, 2)
runningTotal = 0.99 + runningTotal
End Sub

Private Sub cmdSteak_Click(Index As Integer)
picResults.Print "Steak", FormatCurrency(15.95, 2)
runningTotal = 15.95 + runningTotal
End Sub

Private Sub cmdTotal_Click()
Dim Tax As Single, Total As Single
picResults.Print "---------------------"
picResults.Print "Subtotal", FormatCurrency(runningTotal)
Tax = runningTotal * 0.07
picResults.Print "Tax", FormatCurrency(Tax)
Total = runningTotal * 1.07
picResults.Print "Total", FormatCurrency(Total)
End Sub

Private Sub cmdUpload_Click()
picResults.Print message6
End Sub

Private Sub cmdWater_Click(Index As Integer)
picResults.Print "Water", FormatCurrency(0, 2)
runningTotal = 0 + runningTotal
End Sub

