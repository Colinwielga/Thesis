VERSION 5.00
Begin VB.Form BSS 
   Caption         =   "Form1"
   ClientHeight    =   5385
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10470
   LinkTopic       =   "Form1"
   Picture         =   "BSS.frx":0000
   ScaleHeight     =   5385
   ScaleWidth      =   10470
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command14 
      BackColor       =   &H0000FFFF&
      Caption         =   "Quit"
      Height          =   495
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   4800
      Width           =   1335
   End
   Begin VB.CommandButton Command13 
      BackColor       =   &H00C0C000&
      Caption         =   "Bowl of Soup"
      Enabled         =   0   'False
      Height          =   735
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   4080
      Width           =   2055
   End
   Begin VB.CommandButton Command12 
      BackColor       =   &H00C0C000&
      Caption         =   "Cup of Soup"
      Enabled         =   0   'False
      Height          =   735
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   3240
      Width           =   2055
   End
   Begin VB.CommandButton Command11 
      BackColor       =   &H00C0C000&
      Caption         =   "Dinner Salad"
      Enabled         =   0   'False
      Height          =   735
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   2400
      Width           =   2055
   End
   Begin VB.CommandButton Command10 
      BackColor       =   &H00C0C000&
      Caption         =   "Chef Salad"
      Enabled         =   0   'False
      Height          =   735
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   1560
      Width           =   2055
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H00C0C000&
      Caption         =   "Garden Salad"
      Enabled         =   0   'False
      Height          =   735
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   720
      Width           =   2055
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00C0C000&
      Caption         =   "Breakfast Sandwich"
      Enabled         =   0   'False
      Height          =   855
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3720
      Width           =   2055
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00C0C000&
      Caption         =   "Scone"
      Enabled         =   0   'False
      Height          =   855
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2760
      Width           =   2055
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0C000&
      Caption         =   "Muffin"
      Enabled         =   0   'False
      Height          =   855
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1800
      Width           =   2055
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C000&
      Caption         =   "Bagel"
      Enabled         =   0   'False
      Height          =   855
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   840
      Width           =   2055
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H000000C0&
      Caption         =   "Clear Items and Return to Store"
      Enabled         =   0   'False
      Height          =   975
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3840
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H000000C0&
      Caption         =   "Keep Items and Return to Store"
      Enabled         =   0   'False
      Height          =   975
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2760
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000C0&
      Caption         =   "Clear"
      Enabled         =   0   'False
      Height          =   975
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1680
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000000C0&
      Caption         =   "Start"
      Height          =   975
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   600
      Width           =   1815
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000009&
      Height          =   5175
      Left            =   7080
      ScaleHeight     =   5115
      ScaleWidth      =   3195
      TabIndex        =   0
      Top             =   120
      Width           =   3255
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Salad / Soup"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   6
      Top             =   120
      Width           =   2055
   End
   Begin VB.Line Line2 
      X1              =   2520
      X2              =   2520
      Y1              =   240
      Y2              =   4920
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Bakery"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   5
      Top             =   120
      Width           =   1215
   End
   Begin VB.Line Line1 
      X1              =   5040
      X2              =   5040
      Y1              =   240
      Y2              =   5040
   End
End
Attribute VB_Name = "BSS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Subtotal As Single

Private Sub Command1_Click()
Picture1.Print "Item", "Subtotal"
Picture1.Print "****************************"
Command5.Enabled = True
Command6.Enabled = True
Command7.Enabled = True
Command8.Enabled = True
Command9.Enabled = True
Command10.Enabled = True
Command11.Enabled = True
Command12.Enabled = True
Command13.Enabled = True
End Sub

Private Sub Command10_Click()
Dim CS As Single
CS = 2.75
Subtotal = Subtotal + CS
Picture1.Print "Chef Salad ", FormatCurrency(Subtotal)
Command2.Enabled = True
Command3.Enabled = True
Command4.Enabled = True
End Sub

Private Sub Command11_Click()
Dim DS As Single
DS = 3.75
Subtotal = Subtotal + DS
Picture1.Print "Din. Salad  ", FormatCurrency(Subtotal)
Command2.Enabled = True
Command3.Enabled = True
Command4.Enabled = True
End Sub

Private Sub Command12_Click()
Dim Cup As Single
Cup = 1.75
Subtotal = Subtotal + Cup
Picture1.Print "Cup of Soup ", FormatCurrency(Subtotal)
Command2.Enabled = True
Command3.Enabled = True
Command4.Enabled = True
End Sub

Private Sub Command13_Click()
Dim Bowl As Single
Bowl = 2.65
Subtotal = Subtotal + Bowl
Picture1.Print "B. of Soup  ", FormatCurrency(Subtotal)
Command2.Enabled = True
Command3.Enabled = True
Command4.Enabled = True
End Sub

Private Sub Command14_Click()
End
End Sub

Private Sub Command2_Click()
Command5.Enabled = False
Command6.Enabled = False
Command7.Enabled = False
Command8.Enabled = False
Command9.Enabled = False
Command10.Enabled = False
Command11.Enabled = False
Command12.Enabled = False
Command13.Enabled = False
Picture1.Cls
Subtotal = 0
End Sub

Private Sub Command3_Click()
Dim answer As Single
answer = InputBox("Do you want to checkout? If yes you can't get anymore bakery, salad or soup items. If yes enter 1, if no enter 0?")
If answer = 1 Then
    BakeryTotal = Subtotal
    Home.Show
    BSS.Hide
    Home.Command10.Enabled = False
    Home.Command13.Enabled = True
    ElseIf answer = 0 Then
Else
End If
End Sub

Private Sub Command4_Click()
Subtotal = 0
Home.Show
BSS.Hide
End Sub

Private Sub Command5_Click()
Dim Bagel As Single
Bagel = 1.05
Subtotal = Subtotal + Bagel
Picture1.Print "Bagel ", FormatCurrency(Subtotal)
Command2.Enabled = True
Command3.Enabled = True
Command4.Enabled = True
End Sub

Private Sub Command6_Click()
Dim Muffin As Single
Muffin = 1.05
Subtotal = Subtotal + Muffin
Picture1.Print "Muffin ", FormatCurrency(Subtotal)
Command2.Enabled = True
Command3.Enabled = True
Command4.Enabled = True
End Sub

Private Sub Command7_Click()
Dim Scone As Single
Scone = 1.25
Subtotal = Subtotal + Scone
Picture1.Print "Scone ", FormatCurrency(Subtotal)
Command2.Enabled = True
Command3.Enabled = True
Command4.Enabled = True
End Sub

Private Sub Command8_Click()
Dim Break As Single
Break = 2.95
Subtotal = Subtotal + Break
Picture1.Print "B-fast Sand. ", FormatCurrency(Subtotal)
Command2.Enabled = True
Command3.Enabled = True
Command4.Enabled = True
End Sub

Private Sub Command9_Click()
Dim GS As Single
GS = 2.25
Subtotal = Subtotal + GS
Picture1.Print "Gard. Salad  ", FormatCurrency(Subtotal)
Command2.Enabled = True
Command3.Enabled = True
Command4.Enabled = True
End Sub
