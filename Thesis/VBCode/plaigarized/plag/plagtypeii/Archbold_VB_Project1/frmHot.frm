VERSION 5.00
Begin VB.Form frmHot
   BackColor       =   &H000000FF&
   Caption         =   "Hot Food"
   ClientHeight    =   7965
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6510
   LinkTopic       =   "Form1"
   ScaleHeight     =   7965
   ScaleWidth      =   6510
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack
      Caption         =   "Back to the Store"
      Height          =   975
      Left            =   4080
      TabIndex        =   18
      Top             =   960
      Width           =   2175
   End
   Begin VB.PictureBox PicResults
      Height          =   735
      Left            =   3960
      ScaleHeight     =   675
      ScaleWidth      =   2355
      TabIndex        =   17
      Top             =   5520
      Width           =   2415
   End
   Begin VB.PictureBox picFood
      AutoSize        =   -1  'True
      Height          =   1335
      Left            =   4320
      ScaleHeight     =   1275
      ScaleWidth      =   1635
      TabIndex        =   15
      Top             =   2280
      Width           =   1695
   End
   Begin VB.CommandButton cmdAdd
      Caption         =   "Add to Total"
      Height          =   495
      Left            =   3960
      TabIndex        =   14
      Top             =   7320
      Width           =   2175
   End
   Begin VB.TextBox txtown
      Height          =   615
      Left            =   3960
      TabIndex        =   13
      Top             =   6600
      Width           =   2055
   End
   Begin VB.CommandButton cmdSupreme
      Caption         =   "Nachos Supreme"
      Height          =   975
      Left            =   1800
      TabIndex        =   11
      Top             =   5520
      Width           =   1455
   End
   Begin VB.CommandButton cmdNachos
      Caption         =   "Nachos"
      Height          =   975
      Left            =   1800
      TabIndex        =   10
      Top             =   4440
      Width           =   1455
   End
   Begin VB.CommandButton cmdPizza
      Caption         =   "Pizza Slice"
      Height          =   975
      Left            =   240
      TabIndex        =   9
      Top             =   5520
      Width           =   1455
   End
   Begin VB.CommandButton cmdBowl
      Caption         =   "Bowl of Soup"
      Height          =   975
      Left            =   1800
      TabIndex        =   8
      Top             =   3360
      Width           =   1455
   End
   Begin VB.CommandButton cmdCup
      Caption         =   "Cup of Soup"
      Height          =   975
      Left            =   1800
      TabIndex        =   7
      Top             =   2280
      Width           =   1455
   End
   Begin VB.CommandButton cmdTaco
      Caption         =   "Taco"
      Height          =   975
      Left            =   1800
      TabIndex        =   6
      Top             =   1200
      Width           =   1455
   End
   Begin VB.CommandButton cmdBasket
      Caption         =   "Chicken Basket"
      Height          =   975
      Left            =   1800
      TabIndex        =   5
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton cmdGrilled
      Caption         =   "Grilled Chicken"
      Height          =   975
      Left            =   240
      TabIndex        =   4
      Top             =   4440
      Width           =   1455
   End
   Begin VB.CommandButton cmdPatty
      Caption         =   "Chicken Patty"
      Height          =   975
      Left            =   240
      TabIndex        =   3
      Top             =   3360
      Width           =   1455
   End
   Begin VB.CommandButton cmdFries
      Caption         =   "French Fries"
      Height          =   975
      Left            =   240
      TabIndex        =   2
      Top             =   2280
      Width           =   1455
   End
   Begin VB.CommandButton cmdCheese
      Caption         =   "Cheese Burger"
      Height          =   975
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton cmdHam
      Caption         =   "Hamburger"
      Height          =   975
      Left            =   240
      TabIndex        =   0
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Label2
      Caption         =   "Cost of Item:"
      BeginProperty Font
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4200
      TabIndex        =   16
      Top             =   4680
      Width           =   1935
   End
   Begin VB.Label Label1
      Caption         =   "If you know the price of an item not listed here check the cold food section or enter the price here ==>"
      BeginProperty Font
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   12
      Top             =   6600
      Width           =   3615
   End
End
Attribute VB_Name = "frmHot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Project Name Sexton dining
'Form Name Hot
'Author Nick Archbold
'Date written 2/24/10
'Objective To plan out what to buy at sexton before punching out at the end of a week
Dim someAmount As Single

'all food items follow the same pattern reference the first one to see detailed comments
Private Sub cmdAdd_Click()
'Subtracts the price form the total value of the user input punches and send a message box telling the remaning amount or overage, and displays the cost and a picture of the item
PicResults.Cls
Cashola = (Punches * 4.85)
someAmount = txtown.Text
'adds to running total
RunningTotal = RunningTotal + someAmount
'displays ammount left or overage
If Cashola > RunningTotal Then
    MsgBox ("You have " & FormatCurrency((Cashola - RunningTotal)) & " Left")
Else
    MsgBox ("Whoa, with that purchace you went over by " & FormatCurrency(RunningTotal - Cashola))
End If
'prints ammount and picture
PicResults.Print FormatCurrency(someAmount)
picFood.Picture = LoadPicture(App.Path & "\question.jpg")
End Sub

Private Sub cmdBack_Click()
'goes to the store front
frmHot.Visible = False
frmStore.Visible = True

End Sub

Private Sub cmdBasket_Click()
'Subtracts the price form the total value of the user input punches and send a message box telling the remaning amount or overage, and displays the cost and a picture of the item
PicResults.Cls
Cashola = (Punches * 4.85)
RunningTotal = RunningTotal + 4.5
If Cashola > RunningTotal Then
    MsgBox ("You have " & FormatCurrency((Cashola - RunningTotal)) & " Left")
ElseIf Cashola = RunningTotal Then
    MsgBox ("Nice you made it to exactly " & Punches & " punches")
Else
    MsgBox ("Whoa, with that purchace you went over by " & FormatCurrency(RunningTotal - Cashola))
End If
PicResults.Print FormatCurrency(4.5)
picFood.Picture = LoadPicture(App.Path & "\basket.jpg")
End Sub

Private Sub cmdBowl_Click()
'Subtracts the price form the total value of the user input punches and send a message box telling the remaning amount or overage, and displays the cost and a picture of the item
PicResults.Cls
Cashola = (Punches * 4.85)
RunningTotal = RunningTotal + 2.99
If Cashola > RunningTotal Then
    MsgBox ("You have " & FormatCurrency((Cashola - RunningTotal)) & " Left")
Else
    MsgBox ("Whoa, with that purchace you went over by " & FormatCurrency(RunningTotal - Cashola))
End If
PicResults.Print FormatCurrency(2.99)
picFood.Picture = LoadPicture(App.Path & "\bowl.jpg")
End Sub

Private Sub cmdCheese_Click()
'Subtracts the price form the total value of the user input punches and send a message box telling the remaning amount or overage, and displayes a picture of the item.
PicResults.Cls
Cashola = (Punches * 4.85)
'adds to running total
RunningTotal = RunningTotal + 2.99
'displays amt left or overage
If Cashola > RunningTotal Then
    MsgBox ("You have " & FormatCurrency((Cashola - RunningTotal)) & " Left")
Else
    MsgBox ("Whoa, with that purchace you went over by " & FormatCurrency(RunningTotal - Cashola))

End If
'prints the ammount and a picture
PicResults.Print FormatCurrency(2.99)
picFood.Picture = LoadPicture(App.Path & "\cheese.jpg")
End Sub

Private Sub cmdCup_Click()
'Subtracts the price form the total value of the user input punches and send a message box telling the remaning amount or overage, and displays the cost and a picture of the item
PicResults.Cls
Cashola = (Punches * 4.85)
RunningTotal = RunningTotal + 1.99
If Cashola > RunningTotal Then
    MsgBox ("You have " & FormatCurrency((Cashola - RunningTotal)) & " Left")
Else
    MsgBox ("Whoa, with that purchace you went over by " & FormatCurrency(RunningTotal - Cashola))
End If
PicResults.Print FormatCurrency(1.99)
picFood.Picture = LoadPicture(App.Path & "\Cup.jpg")
End Sub

Private Sub cmdFries_Click()
'Subtracts the price form the total value of the user input punches and send a message box telling the remaning amount or overage, and displays the cost and a picture of the item
PicResults.Cls
Cashola = (Punches * 4.85)
RunningTotal = RunningTotal + 1.2
If Cashola > RunningTotal Then
    MsgBox ("You have " & FormatCurrency((Cashola - RunningTotal)) & " Left")
Else
    MsgBox ("Whoa, with that purchace you went over by " & FormatCurrency(RunningTotal - Cashola))
End If
PicResults.Print FormatCurrency(1.2)
picFood.Picture = LoadPicture(App.Path & "\fries.jpg")
End Sub

Private Sub cmdGrilled_Click()
'Subtracts the price form the total value of the user input punches and send a message box telling the remaning amount or overage, and displays the cost and a picture of the item
PicResults.Cls
Cashola = (Punches * 4.85)
RunningTotal = RunningTotal + 3.75
If Cashola > RunningTotal Then
    MsgBox ("You have " & FormatCurrency((Cashola - RunningTotal)) & " Left")
Else
    MsgBox ("Whoa, with that purchace you went over by " & FormatCurrency(RunningTotal - Cashola))
End If
PicResults.Print FormatCurrency(3.75)
picFood.Picture = LoadPicture(App.Path & "\grilled.jpg")
End Sub

Private Sub cmdHam_Click()
'Subtracts the price form the total value of the user input punches and send a message box telling the remaning amount or overage, and displays the cost and a picture of the item
PicResults.Cls
Cashola = (Punches * 4.85)
RunningTotal = RunningTotal + 2.79
If Cashola > RunningTotal Then
    MsgBox ("You have " & FormatCurrency((Cashola - RunningTotal)) & " Left")
Else
    MsgBox ("Whoa, with that purchace you went over by " & FormatCurrency(RunningTotal - Cashola))
End If
PicResults.Print FormatCurrency(2.79)
picFood.Picture = LoadPicture(App.Path & "\ham.jpg")
End Sub

Private Sub cmdNachos_Click()
'Subtracts the price form the total value of the user input punches and send a message box telling the remaning amount or overage, and displays the cost and a picture of the item
PicResults.Cls
Cashola = (Punches * 4.85)
RunningTotal = RunningTotal + 2.25
If Cashola > RunningTotal Then
    MsgBox ("You have " & FormatCurrency((Cashola - RunningTotal)) & " Left")
Else
    MsgBox ("Whoa, with that purchace you went over by " & FormatCurrency(RunningTotal - Cashola))
End If
PicResults.Print FormatCurrency(2.25)
picFood.Picture = LoadPicture(App.Path & "\nacho.jpg")
End Sub

Private Sub cmdPatty_Click()
'Subtracts the price form the total value of the user input punches and send a message box telling the remaning amount or overage, and displays the cost and a picture of the item
PicResults.Cls
Cashola = (Punches * 4.85)
RunningTotal = RunningTotal + 3.25
If Cashola > RunningTotal Then
    MsgBox ("You have " & FormatCurrency((Cashola - RunningTotal)) & " Left")
Else
    MsgBox ("Whoa, with that purchace you went over by " & FormatCurrency(RunningTotal - Cashola))
End If
PicResults.Print FormatCurrency(3.25)
picFood.Picture = LoadPicture(App.Path & "\patty.jpg")
End Sub

Private Sub cmdPizza_Click()
'Subtracts the price form the total value of the user input punches and send a message box telling the remaning amount or overage, and displays the cost and a picture of the item
PicResults.Cls
Cashola = (Punches * 4.85)
RunningTotal = RunningTotal + 2.15
If Cashola > RunningTotal Then
    MsgBox ("You have " & FormatCurrency((Cashola - RunningTotal)) & " Left")
Else
    MsgBox ("Whoa, with that purchace you went over by " & FormatCurrency(RunningTotal - Cashola))
End If
PicResults.Print FormatCurrency(2.15)
picFood.Picture = LoadPicture(App.Path & "\pizza.jpg")
End Sub

Private Sub cmdSupreme_Click()
'Subtracts the price form the total value of the user input punches and send a message box telling the remaning amount or overage, and displays the cost and a picture of the item
PicResults.Cls
Cashola = (Punches * 4.85)
RunningTotal = RunningTotal + 4.5
If Cashola > RunningTotal Then
    MsgBox ("You have " & FormatCurrency((Cashola - RunningTotal)) & " Left")
Else
    MsgBox ("Whoa, with that purchace you went over by " & FormatCurrency(RunningTotal - Cashola))
End If
PicResults.Print FormatCurrency(4.5)
picFood.Picture = LoadPicture(App.Path & "\supreme.jpg")
End Sub

Private Sub cmdTaco_Click()
'Subtracts the price form the total value of the user input punches and send a message box telling the remaning amount or overage, and displays the cost and a picture of the item
PicResults.Cls
Cashola = (Punches * 4.85)
RunningTotal = RunningTotal + 1.45
If Cashola > RunningTotal Then
    MsgBox ("You have " & FormatCurrency((Cashola - RunningTotal)) & " Left")
Else
    MsgBox ("Whoa, with that purchace you went over by " & FormatCurrency(RunningTotal - Cashola))
End If
PicResults.Print FormatCurrency(1.45)
picFood.Picture = LoadPicture(App.Path & "\Taco.jpg")
End Sub

Private Sub Text1_Change()

End Sub
