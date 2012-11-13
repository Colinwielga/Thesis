VERSION 5.00
Begin VB.Form frmCold 
   BackColor       =   &H00FF0000&
   Caption         =   "Clod Food"
   ClientHeight    =   8250
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6405
   LinkTopic       =   "Form1"
   ScaleHeight     =   8250
   ScaleWidth      =   6405
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picFood 
      AutoSize        =   -1  'True
      Height          =   1335
      Left            =   3960
      ScaleHeight     =   1275
      ScaleWidth      =   1635
      TabIndex        =   20
      Top             =   2040
      Width           =   1695
   End
   Begin VB.PictureBox PicResults 
      Height          =   735
      Left            =   3720
      ScaleHeight     =   675
      ScaleWidth      =   1995
      TabIndex        =   19
      Top             =   5640
      Width           =   2055
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back to Store"
      Height          =   975
      Left            =   3600
      TabIndex        =   17
      Top             =   480
      Width           =   2415
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add to total "
      Height          =   495
      Left            =   3960
      TabIndex        =   16
      Top             =   7560
      Width           =   1815
   End
   Begin VB.TextBox txtCold 
      Height          =   495
      Left            =   3960
      TabIndex        =   15
      Top             =   6960
      Width           =   1815
   End
   Begin VB.CommandButton cmdVeggies 
      Caption         =   "Veggies"
      Height          =   855
      Index           =   1
      Left            =   1680
      TabIndex        =   13
      Top             =   6000
      Width           =   1335
   End
   Begin VB.CommandButton cmdEnergy 
      Caption         =   "Energy Drink"
      Height          =   855
      Index           =   1
      Left            =   1680
      TabIndex        =   12
      Top             =   5040
      Width           =   1335
   End
   Begin VB.CommandButton cmdYogurt 
      Caption         =   "Yogurt"
      Height          =   855
      Index           =   1
      Left            =   1680
      TabIndex        =   11
      Top             =   4080
      Width           =   1335
   End
   Begin VB.CommandButton cmdJuice 
      Caption         =   "Juice"
      Height          =   855
      Index           =   1
      Left            =   1680
      TabIndex        =   10
      Top             =   3120
      Width           =   1335
   End
   Begin VB.CommandButton cmdPop 
      Caption         =   "Bottle Pop"
      Height          =   855
      Index           =   1
      Left            =   1680
      TabIndex        =   9
      Top             =   2160
      Width           =   1335
   End
   Begin VB.CommandButton cmdCookie 
      Caption         =   "Cookie"
      Height          =   855
      Index           =   1
      Left            =   1680
      TabIndex        =   8
      Top             =   1200
      Width           =   1335
   End
   Begin VB.CommandButton cmdBagel 
      Caption         =   "Bagel"
      Height          =   855
      Index           =   1
      Left            =   1680
      TabIndex        =   7
      Top             =   240
      Width           =   1335
   End
   Begin VB.CommandButton cmdThird 
      Caption         =   "Third Sub"
      Height          =   855
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   6000
      Width           =   1335
   End
   Begin VB.CommandButton cmdHalf 
      Caption         =   "Half Sub"
      Height          =   855
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   5040
      Width           =   1335
   End
   Begin VB.CommandButton cmdFruit 
      Caption         =   "Fruit"
      Height          =   855
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   4080
      Width           =   1335
   End
   Begin VB.CommandButton cmdCandy 
      Caption         =   "Candy"
      Height          =   855
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   3120
      Width           =   1335
   End
   Begin VB.CommandButton cmdGummies 
      Caption         =   "Gummies"
      Height          =   855
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   2160
      Width           =   1335
   End
   Begin VB.CommandButton cmdlittle 
      Caption         =   "Little Bag Chips"
      Height          =   855
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   1335
   End
   Begin VB.CommandButton cmdBig 
      Caption         =   "Big Bag Chips"
      Height          =   855
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1335
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
      Left            =   3720
      TabIndex        =   18
      Top             =   4800
      Width           =   2055
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
      TabIndex        =   14
      Top             =   6960
      Width           =   3615
   End
End
Attribute VB_Name = "frmCold"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name Sexton dining
'Form Name Cold
'Author Nick Archbold
'Date written 2/24/10
'Objective To plan out what to buy at sexton before punching out at the end of a week

'all food items follow the same patten and the first one should be referenced as an example
Private Sub cmdAdd_Click()
'Subtracts the price form the total value of the user input punches and send a message box telling the remaning amount or overage, and displays the cost and a picture of the item
PicResults.Cls
Money = (Punches * 4.85)
cold = txtCold.Text
'adds to the running toatl
RunningTotal = RunningTotal + cold
'displays the ammount left or overage
If Money > RunningTotal Then
    MsgBox ("You have " & FormatCurrency((Money - RunningTotal)) & " Left")
Else
    MsgBox ("Whoa, with that purchace you went over by " & FormatCurrency(RunningTotal - Money))
End If
'prints the ammount spent and a picture of the item
PicResults.Print FormatCurrency(cold)
picFood.Picture = LoadPicture(App.Path & "\question.jpg")
End Sub

Private Sub cmdBack_Click()
'goes to the store front
frmCold.Visible = False
frmStore.Visible = True
End Sub

Private Sub cmdBagel_Click(Index As Integer)
'Subtracts the price form the total value of the user input punches and send a message box telling the remaning amount or overage, and displays the cost and a picture of the item
PicResults.Cls
Money = (Punches * 4.85)
RunningTotal = RunningTotal + 1.05
If Money > RunningTotal Then
    MsgBox ("You have " & FormatCurrency((Money - RunningTotal)) & " Left")
Else
    MsgBox ("Whoa, with that purchace you went over by " & FormatCurrency(RunningTotal - Money))
    
End If
PicResults.Print FormatCurrency(1.05)
picFood.Picture = LoadPicture(App.Path & "\bagel.jpg")
End Sub

Private Sub cmdBig_Click(Index As Integer)
'Subtracts the price form the total value of the user input punches and send a message box telling the remaning amount or overage, and displays the cost and a picture of the item
PicResults.Cls
Money = (Punches * 4.85)
RunningTotal = RunningTotal + 3.99
If Money > RunningTotal Then
    MsgBox ("You have " & FormatCurrency((Money - RunningTotal)) & " Left")
Else
    MsgBox ("Whoa, with that purchace you went over by " & FormatCurrency(RunningTotal - Money))
    
End If
PicResults.Print FormatCurrency(3.99)
picFood.Picture = LoadPicture(App.Path & "\big.jpg")
End Sub

Private Sub cmdCandy_Click(Index As Integer)
'Subtracts the price form the total value of the user input punches and send a message box telling the remaning amount or overage, and displays the cost and a picture of the item
PicResults.Cls
Money = (Punches * 4.85)
RunningTotal = RunningTotal + 0.99
If Money > RunningTotal Then
    MsgBox ("You have " & FormatCurrency((Money - RunningTotal)) & " Left")
Else
    MsgBox ("Whoa, with that purchace you went over by " & FormatCurrency(RunningTotal - Money))
    
End If
PicResults.Print FormatCurrency(0.99)
picFood.Picture = LoadPicture(App.Path & "\candy.jpg")
End Sub

Private Sub cmdCookie_Click(Index As Integer)
'Subtracts the price form the total value of the user input punches and send a message box telling the remaning amount or overage, and displays the cost and a picture of the item
PicResults.Cls
Money = (Punches * 4.85)
RunningTotal = RunningTotal + 0.75
If Money > RunningTotal Then
    MsgBox ("You have " & FormatCurrency((Money - RunningTotal)) & " Left")
Else
    MsgBox ("Whoa, with that purchace you went over by " & FormatCurrency(RunningTotal - Money))
    
End If
PicResults.Print FormatCurrency(0.75)
picFood.Picture = LoadPicture(App.Path & "\cookie.jpg")
End Sub

Private Sub cmdEnergy_Click(Index As Integer)
'Subtracts the price form the total value of the user input punches and send a message box telling the remaning amount or overage, and displays the cost and a picture of the item
PicResults.Cls
Money = (Punches * 4.85)
RunningTotal = RunningTotal + 2.25
If Money > RunningTotal Then
    MsgBox ("You have " & FormatCurrency((Money - RunningTotal)) & " Left")
Else
    MsgBox ("Whoa, with that purchace you went over by " & FormatCurrency(RunningTotal - Money))
    
End If
PicResults.Print FormatCurrency(2.25)
picFood.Picture = LoadPicture(App.Path & "\energy.jpg")
End Sub

Private Sub cmdFruit_Click(Index As Integer)
'Subtracts the price form the total value of the user input punches and send a message box telling the remaning amount or overage, and displays the cost and a picture of the item
PicResults.Cls
Money = (Punches * 4.85)
RunningTotal = RunningTotal + 0.7
If Money > RunningTotal Then
    MsgBox ("You have " & FormatCurrency((Money - RunningTotal)) & " Left")
Else
    MsgBox ("Whoa, with that purchace you went over by " & FormatCurrency(RunningTotal - Money))
    
End If
PicResults.Print FormatCurrency(0.7)
picFood.Picture = LoadPicture(App.Path & "\apple.jpg")
End Sub

Private Sub cmdGummies_Click(Index As Integer)
'Subtracts the price form the total value of the user input punches and send a message box telling the remaning amount or overage, and displays the cost and a picture of the item
PicResults.Cls
Money = (Punches * 4.85)
RunningTotal = RunningTotal + 2.99
If Money > RunningTotal Then
    MsgBox ("You have " & FormatCurrency((Money - RunningTotal)) & " Left")
Else
    MsgBox ("Whoa, with that purchace you went over by " & FormatCurrency(RunningTotal - Money))
    
End If
PicResults.Print FormatCurrency(2.99)
picFood.Picture = LoadPicture(App.Path & "\gummies.jpg")
End Sub

Private Sub cmdHalf_Click(Index As Integer)
'Subtracts the price form the total value of the user input punches and send a message box telling the remaning amount or overage, and displays the cost and a picture of the item
PicResults.Cls
Money = (Punches * 4.85)
RunningTotal = RunningTotal + 3.99
If Money > RunningTotal Then
    MsgBox ("You have " & FormatCurrency((Money - RunningTotal)) & " Left")
Else
    MsgBox ("Whoa, with that purchace you went over by " & FormatCurrency(RunningTotal - Money))
    
End If
PicResults.Print FormatCurrency(3.99)
picFood.Picture = LoadPicture(App.Path & "\half.jpg")
End Sub

Private Sub cmdJuice_Click(Index As Integer)
'Subtracts the price form the total value of the user input punches and send a message box telling the remaning amount or overage, and displays the cost and a picture of the item
PicResults.Cls
Money = (Punches * 4.85)
RunningTotal = RunningTotal + 1.5
If Money > RunningTotal Then
    MsgBox ("You have " & FormatCurrency((Money - RunningTotal)) & " Left")
Else
    MsgBox ("Whoa, with that purchace you went over by " & FormatCurrency(RunningTotal - Money))
    
End If
PicResults.Print FormatCurrency(1.5)
picFood.Picture = LoadPicture(App.Path & "\juice.jpg")
End Sub

Private Sub cmdlittle_Click(Index As Integer)
'Subtracts the price form the total value of the user input punches and send a message box telling the remaning amount or overage, and displays the cost and a picture of the item
PicResults.Cls
Money = (Punches * 4.85)
RunningTotal = RunningTotal + 0.99
If Money > RunningTotal Then
    MsgBox ("You have " & FormatCurrency((Money - RunningTotal)) & " Left")
Else
    MsgBox ("Whoa, with that purchace you went over by " & FormatCurrency(RunningTotal - Money))
    
End If
PicResults.Print FormatCurrency(0.99)
picFood.Picture = LoadPicture(App.Path & "\little.jpg")
End Sub

Private Sub cmdPop_Click(Index As Integer)
'Subtracts the price form the total value of the user input punches and send a message box telling the remaning amount or overage, and displays the cost and a picture of the item
PicResults.Cls
Money = (Punches * 4.85)
RunningTotal = RunningTotal + 1.5
If Money > RunningTotal Then
    MsgBox ("You have " & FormatCurrency((Money - RunningTotal)) & " Left")
Else
    MsgBox ("Whoa, with that purchace you went over by " & FormatCurrency(RunningTotal - Money))
    
End If
PicResults.Print FormatCurrency(1.5)
picFood.Picture = LoadPicture(App.Path & "\pop.jpg")
End Sub

Private Sub cmdThird_Click(Index As Integer)
'Subtracts the price form the total value of the user input punches and send a message box telling the remaning amount or overage, and displays the cost and a picture of the item
PicResults.Cls
Money = (Punches * 4.85)
RunningTotal = RunningTotal + 2.99
If Money > RunningTotal Then
    MsgBox ("You have " & FormatCurrency((Money - RunningTotal)) & " Left")
Else
    MsgBox ("Whoa, with that purchace you went over by " & FormatCurrency(RunningTotal - Money))
    
End If
PicResults.Print FormatCurrency(2.99)
picFood.Picture = LoadPicture(App.Path & "\third.jpg")
End Sub

Private Sub cmdVeggies_Click(Index As Integer)
'Subtracts the price form the total value of the user input punches and send a message box telling the remaning amount or overage, and displays the cost and a picture of the item
PicResults.Cls
Money = (Punches * 4.85)
RunningTotal = RunningTotal + 1.75
If Money > RunningTotal Then
    MsgBox ("You have " & FormatCurrency((Money - RunningTotal)) & " Left")
Else
    MsgBox ("Whoa, with that purchace you went over by " & FormatCurrency(RunningTotal - Money))
    
End If
PicResults.Print FormatCurrency(1.75)
picFood.Picture = LoadPicture(App.Path & "\veggies.jpg")
End Sub

Private Sub cmdYogurt_Click(Index As Integer)
'Subtracts the price form the total value of the user input punches and send a message box telling the remaning amount or overage, and displays the cost and a picture of the item
PicResults.Cls
Money = (Punches * 4.85)
RunningTotal = RunningTotal + 1.15
If Money > RunningTotal Then
    MsgBox ("You have " & FormatCurrency((Money - RunningTotal)) & " Left")
Else
    MsgBox ("Whoa, with that purchace you went over by " & FormatCurrency(RunningTotal - Money))
    
End If
PicResults.Print FormatCurrency(1.15)
picFood.Picture = LoadPicture(App.Path & "\Yogurt.jpg")
End Sub

