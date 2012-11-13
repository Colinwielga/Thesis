VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H000080FF&
   Caption         =   "Form2"
   ClientHeight    =   8370
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12165
   LinkTopic       =   "Form2"
   ScaleHeight     =   8370
   ScaleWidth      =   12165
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtSize2 
      Height          =   615
      Left            =   3000
      TabIndex        =   12
      Top             =   5160
      Width           =   1455
   End
   Begin VB.TextBox txtSize 
      Height          =   495
      Left            =   2760
      TabIndex        =   10
      Top             =   1200
      Width           =   1335
   End
   Begin VB.PictureBox Picture1 
      Height          =   5895
      Left            =   7800
      ScaleHeight     =   5835
      ScaleWidth      =   3555
      TabIndex        =   9
      Top             =   720
      Width           =   3615
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   975
      Left            =   9120
      TabIndex        =   8
      Top             =   7200
      Width           =   2055
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   855
      Left            =   5280
      TabIndex        =   7
      Top             =   7200
      Width           =   2535
   End
   Begin VB.CommandButton cmdTotal 
      Caption         =   "Total"
      Height          =   1095
      Left            =   2280
      TabIndex        =   6
      Top             =   6960
      Width           =   1815
   End
   Begin VB.CommandButton cmdJersey2 
      Height          =   1575
      Left            =   5160
      Picture         =   "Form2.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5160
      Width           =   1935
   End
   Begin VB.CommandButton cmdBanner 
      Height          =   2415
      Left            =   5160
      Picture         =   "Form2.frx":1175
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2520
      Width           =   1815
   End
   Begin VB.CommandButton cmdShirt 
      Height          =   1815
      Left            =   720
      Picture         =   "Form2.frx":206A
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4320
      Width           =   1935
   End
   Begin VB.CommandButton cmdPuck 
      Height          =   1575
      Left            =   960
      Picture         =   "Form2.frx":2B13
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2400
      Width           =   1335
   End
   Begin VB.CommandButton cmdHat 
      BackColor       =   &H000080FF&
      Height          =   2055
      Left            =   4800
      Picture         =   "Form2.frx":3827
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   2535
   End
   Begin VB.CommandButton cmdJersey 
      Height          =   1815
      Left            =   840
      Picture         =   "Form2.frx":76E0
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label lblSize2 
      Caption         =   "Choose Size"
      Height          =   375
      Left            =   3240
      TabIndex        =   13
      Top             =   4560
      Width           =   975
   End
   Begin VB.Label lblSize 
      Caption         =   "Choose Size"
      Height          =   375
      Left            =   2880
      TabIndex        =   11
      Top             =   600
      Width           =   1215
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim runningTotal As Single

Private Sub cmdBanner_Click()
runningTotal = runningTotal + 27.99
picResults.Print "Banner", FormatCurrency(27.99)
End Sub

Private Sub cmdClear_Click()
runningTotal = 0
picResults.Cls
End Sub

Private Sub cmdHat_Click()
runningTotal = runningTotal + 29.99
picResults.Print "Oilers Hat"; Tab(15); , FormatCurrency(29.99)
End Sub

Private Sub cmdJersey_Click()
Dim Jsize As String
Jsize = txtSize.Text
runningTotal = runningTotal + 132.99
picResults.Print "Oilers Jersey"; "("; Jsize; ") "; FormatCurrency(132.99); ""
End Sub

Private Sub cmdJersey2_Click()
runningTotal = runningTotal + 1391.99
picResults.Print "Gretzky Signed Jersey", FormatCurrency(1391.99)
End Sub


Private Sub cmdPuck_Click()
runningTotal = runningTotal + 38.99
picResults.Print "Signed Puck", FormatCurrency(38.99)
End Sub

Private Sub cmdQuit_Click()
End
End Sub

Private Sub cmdShirt_Click()
Dim Jsize As String
Jsize = txtSize2.Text
runningTotal = runningTotal + 27.99
picResults.Print "Oilers Shirt"; "("; Jsize; ") "; FormatCurrency(27.99)
End Sub

Private Sub cmdTotal_Click()
Dim subTotal As Single
Dim tax As Single
Dim total As Single
tax = runningTotal * 0.07
total = runningTotal + tax
picResults.Print "----------------------------"
picResults.Print "Subtotal", FormatCurrency(runningTotal)
picResults.Print "Tax", FormatCurrency(tax)
picResults.Print "Total", FormatCurrency(total)
End Sub

