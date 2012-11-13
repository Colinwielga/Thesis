VERSION 5.00
Begin VB.Form Drinks 
   Caption         =   "Form2"
   ClientHeight    =   5100
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10680
   LinkTopic       =   "Form2"
   Picture         =   "Drinks.frx":0000
   ScaleHeight     =   5100
   ScaleWidth      =   10680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command15 
      BackColor       =   &H000000C0&
      Caption         =   "Clear and Go Back to Store"
      Height          =   975
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   3960
      Width           =   1455
   End
   Begin VB.CommandButton Command14 
      BackColor       =   &H0000FFFF&
      Caption         =   "Quit"
      Height          =   375
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton Command13 
      BackColor       =   &H000000C0&
      Caption         =   "Keep and Return to Store"
      Height          =   975
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   2880
      Width           =   1455
   End
   Begin VB.CommandButton Command12 
      BackColor       =   &H000000C0&
      Caption         =   "Clear"
      Height          =   975
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   1800
      Width           =   1455
   End
   Begin VB.CommandButton Command11 
      BackColor       =   &H000000C0&
      Caption         =   "Start"
      Height          =   975
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   720
      Width           =   1455
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   4695
      Left            =   7440
      ScaleHeight     =   4635
      ScaleWidth      =   3075
      TabIndex        =   11
      Top             =   240
      Width           =   3135
   End
   Begin VB.CommandButton Command10 
      BackColor       =   &H00FF8080&
      Caption         =   "Gatorade/Powerade"
      Enabled         =   0   'False
      Height          =   975
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3960
      Width           =   1695
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H00FF8080&
      Caption         =   "Sobe/Fuse Drinks"
      Enabled         =   0   'False
      Height          =   975
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2880
      Width           =   1695
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00FF8080&
      Caption         =   "Bottled Water"
      Enabled         =   0   'False
      Height          =   975
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2880
      Width           =   1695
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00FF8080&
      Caption         =   "Energy Drinks"
      Enabled         =   0   'False
      Height          =   975
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2880
      Width           =   1695
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00FF8080&
      Caption         =   "Cold Coffee"
      Enabled         =   0   'False
      Height          =   975
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1800
      Width           =   1695
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FF8080&
      Caption         =   "Small Coffee"
      Enabled         =   0   'False
      Height          =   975
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1800
      Width           =   1695
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FF8080&
      Caption         =   "Large Coffee"
      Enabled         =   0   'False
      Height          =   975
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1800
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FF8080&
      Caption         =   "Juice"
      Enabled         =   0   'False
      Height          =   975
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   720
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FF8080&
      Caption         =   "Soda"
      Enabled         =   0   'False
      Height          =   975
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   720
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FF8080&
      Caption         =   "Fountain Drinks/Milk"
      Enabled         =   0   'False
      Height          =   975
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   720
      Width           =   1695
   End
   Begin VB.Line Line2 
      X1              =   7320
      X2              =   7320
      Y1              =   120
      Y2              =   5040
   End
   Begin VB.Line Line1 
      X1              =   5640
      X2              =   5640
      Y1              =   600
      Y2              =   5040
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Drinks"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "Drinks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Subtotal As Single

Private Sub Command1_Click()
Dim Drink As Single
Drink = 0.9
Subtotal = Subtotal + Drink
Picture1.Print "Fountain Drink "; FormatCurrency(Subtotal)

End Sub

Private Sub Command10_Click()
Dim Ade As Single
Ade = 1.39
Subtotal = Subtotal + Ade
Picture1.Print "Sports Drink", FormatCurrency(Subtotal)
End Sub

Private Sub Command11_Click()
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
End Sub

Private Sub Command12_Click()
Picture1.Cls
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
Subtotal = 0
End Sub

Private Sub Command13_Click()
Dim answer As Single
answer = InputBox("Do you want to checkout? If yes you can't get anymore drink items. If yes enter 1, if no enter 0?")
If answer = 1 Then
    DrinkTotal = Subtotal
    Home.Show
    Drinks.Hide
    Home.Command8.Enabled = False
    Home.Command13.Enabled = True
    ElseIf answer = 0 Then
Else
End If
End Sub

Private Sub Command14_Click()
End
End Sub

Private Sub Command15_Click()
Home.Show
Drinks.Hide
End Sub

Private Sub Command2_Click()
Dim Soda As Single
Soda = 1.39
Subtotal = Subtotal + Soda
Picture1.Print "Soda", FormatCurrency(Subtotal)
End Sub

Private Sub Command3_Click()
Dim Juice As Single
Juice = 1.75
Subtotal = Subtotal + Juice
Picture1.Print "Juice", FormatCurrency(Subtotal)
End Sub

Private Sub Command4_Click()
Dim LCoffee As Single
LCoffee = 1.25
Subtotal = Subtotal + LCoffee
Picture1.Print "Large Coffee "; FormatCurrency(Subtotal)
End Sub

Private Sub Command5_Click()
Dim Coffee As Single
Coffee = 1.05
Subtotal = Subtotal + Coffee
Picture1.Print "Small Coffee "; FormatCurrency(Subtotal)
End Sub

Private Sub Command6_Click()
Dim CC As Single
CC = 2.05
Subtotal = Subtotal + CC
Picture1.Print "Cold Coffee", FormatCurrency(Subtotal)
End Sub

Private Sub Command7_Click()
Dim En As Single
En = 2.5
Subtotal = Subtotal + En
Picture1.Print "Energy Drink", FormatCurrency(Subtotal)
End Sub

Private Sub Command8_Click()
Dim Water As Single
Water = 1.39
Subtotal = Subtotal + Water
Picture1.Print "Water", FormatCurrency(Subtotal)
End Sub

Private Sub Command9_Click()
Dim Sobe As Single
Sobe = 2.05
Subtotal = Subtotal + Sobe
Picture1.Print "Sobe/Fuse", FormatCurrency(Subtotal)
End Sub
