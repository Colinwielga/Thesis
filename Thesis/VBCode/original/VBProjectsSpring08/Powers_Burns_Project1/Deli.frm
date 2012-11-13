VERSION 5.00
Begin VB.Form Deli 
   Caption         =   "Form1"
   ClientHeight    =   5415
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10725
   LinkTopic       =   "Form1"
   Picture         =   "Deli.frx":0000
   ScaleHeight     =   5415
   ScaleWidth      =   10725
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command15 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Bagel Sandwich"
      Enabled         =   0   'False
      Height          =   975
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   3840
      Width           =   1575
   End
   Begin VB.CommandButton Command14 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Whole Sub"
      Enabled         =   0   'False
      Height          =   975
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   2760
      Width           =   1575
   End
   Begin VB.CommandButton Command13 
      BackColor       =   &H00C0C0C0&
      Caption         =   "1/2 Sub"
      Enabled         =   0   'False
      Height          =   975
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   2760
      Width           =   1575
   End
   Begin VB.CommandButton Command12 
      BackColor       =   &H00C0C0C0&
      Caption         =   "1/3 Sub"
      Enabled         =   0   'False
      Height          =   975
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2760
      Width           =   1575
   End
   Begin VB.CommandButton Command11 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Nacho Supreme"
      Enabled         =   0   'False
      Height          =   975
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1680
      Width           =   1575
   End
   Begin VB.CommandButton Command10 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Croissant"
      Enabled         =   0   'False
      Height          =   975
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1680
      Width           =   1575
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Wrap"
      Enabled         =   0   'False
      Height          =   975
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1680
      Width           =   1575
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Nachos"
      Enabled         =   0   'False
      Height          =   975
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   600
      Width           =   1575
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Hard Taco"
      Enabled         =   0   'False
      Height          =   975
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   600
      Width           =   1575
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Soft Taco"
      Enabled         =   0   'False
      Height          =   975
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   600
      Width           =   1575
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H000000C0&
      Caption         =   "Clear Items and Return to Store"
      Enabled         =   0   'False
      Height          =   1095
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3960
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H000000C0&
      Caption         =   "Keep Items and Return to Store"
      Enabled         =   0   'False
      Height          =   1095
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2760
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H000000C0&
      Caption         =   "Clear"
      Enabled         =   0   'False
      Height          =   1095
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1560
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000C0&
      Caption         =   "Start"
      Height          =   1095
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   360
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H8000000E&
      Height          =   5175
      Left            =   7080
      ScaleHeight     =   5115
      ScaleWidth      =   3435
      TabIndex        =   1
      Top             =   120
      Width           =   3495
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FFFF&
      Caption         =   "Quit"
      Height          =   375
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
   Begin VB.Line Line1 
      X1              =   5280
      X2              =   5280
      Y1              =   120
      Y2              =   5160
   End
End
Attribute VB_Name = "Deli"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Subtotal As Single

Private Sub Command1_Click()
End
End Sub

Private Sub Command10_Click()
Dim Croy As Single
Croy = 3.1
Subtotal = Subtotal + Croy
Picture1.Print "Croissant", FormatCurrency(Subtotal)
Command3.Enabled = True
Command4.Enabled = True
Command5.Enabled = True
End Sub

Private Sub Command11_Click()
Dim Ns As Single
Ns = 4
Subtotal = Subtotal + Ns
Picture1.Print "Nachos Sup.", FormatCurrency(Subtotal)
Command3.Enabled = True
Command4.Enabled = True
Command5.Enabled = True
End Sub

Private Sub Command12_Click()
Dim Tres As Single
Tres = 2.55
Subtotal = Subtotal + Tres
Picture1.Print "1/3 Sub", FormatCurrency(Subtotal)
Command3.Enabled = True
Command4.Enabled = True
Command5.Enabled = True
End Sub

Private Sub Command13_Click()
Dim Dub As Single
Dub = 3.55
Subtotal = Subtotal + Dub
Picture1.Print "1/2 Sub", FormatCurrency(Subtotal)
Command3.Enabled = True
Command4.Enabled = True
Command5.Enabled = True
End Sub

Private Sub Command14_Click()
Dim Whole As Single
Whole = 7.1
Subtotal = Subtotal + Whole
Picture1.Print "Whole Sub", FormatCurrency(Subtotal)
Command3.Enabled = True
Command4.Enabled = True
Command5.Enabled = True
End Sub

Private Sub Command15_Click()
Dim Bagel As Single
Bagel = 3.25
Subtotal = Subtotal + Bagel
Picture1.Print "Bagel Sand.", FormatCurrency(Subtotal)
Command3.Enabled = True
Command4.Enabled = True
Command5.Enabled = True
End Sub

Private Sub Command2_Click()
Picture1.Print "Item", "Subtotal"
Picture1.Print "****************************"
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

Private Sub Command3_Click()
Picture1.Cls
Subtotal = 0
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
End Sub

Private Sub Command4_Click()
Dim answer As Single
answer = InputBox("Do you want to checkout? If yes you can't get anymore deli items. If yes enter 1, if no enter 0?")
If answer = 1 Then
    DeliTotal = Subtotal
    Home.Show
    Deli.Hide
    Home.Command6.Enabled = False
    Home.Command13.Enabled = True
    ElseIf answer = 0 Then
Else
End If
End Sub

Private Sub Command5_Click()
Deli.Hide
Home.Show
End Sub

Private Sub Command6_Click()
Dim ST As Single
ST = 1.45
Subtotal = Subtotal + ST
Picture1.Print "Soft Taco", FormatCurrency(Subtotal)
Command3.Enabled = True
Command4.Enabled = True
Command5.Enabled = True
End Sub

Private Sub Command7_Click()
Dim HT As Single
HT = 1.35
Subtotal = Subtotal + HT
Picture1.Print "Hard Taco", FormatCurrency(Subtotal)
Command3.Enabled = True
Command4.Enabled = True
Command5.Enabled = True
End Sub

Private Sub Command8_Click()
Dim Na As Single
Na = 2.25
Subtotal = Subtotal + Na
Picture1.Print "Nachos", FormatCurrency(Subtotal)
Command3.Enabled = True
Command4.Enabled = True
Command5.Enabled = True
End Sub

Private Sub Command9_Click()
Dim Wrap As Single
Wrap = 1.6
Subtotal = Subtotal + Wrap
Picture1.Print "Wrap", FormatCurrency(Subtotal)
Command3.Enabled = True
Command4.Enabled = True
Command5.Enabled = True
End Sub
