VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H000000C0&
   Caption         =   "Form3"
   ClientHeight    =   8025
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12705
   LinkTopic       =   "Form3"
   ScaleHeight     =   8025
   ScaleWidth      =   12705
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtSize2 
      Height          =   615
      Left            =   3000
      TabIndex        =   12
      Top             =   5640
      Width           =   975
   End
   Begin VB.TextBox txtSize 
      Height          =   495
      Left            =   2880
      TabIndex        =   10
      Top             =   1440
      Width           =   1095
   End
   Begin VB.PictureBox picResults 
      Height          =   5655
      Left            =   7920
      ScaleHeight     =   5595
      ScaleWidth      =   3435
      TabIndex        =   9
      Top             =   600
      Width           =   3495
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   975
      Left            =   9000
      TabIndex        =   8
      Top             =   6840
      Width           =   2535
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   975
      Left            =   5160
      TabIndex        =   7
      Top             =   6840
      Width           =   2295
   End
   Begin VB.CommandButton cmdTotal 
      Caption         =   "Total"
      Height          =   1095
      Left            =   1800
      TabIndex        =   6
      Top             =   6720
      Width           =   2295
   End
   Begin VB.CommandButton cmdMug 
      Height          =   2295
      Left            =   4920
      Picture         =   "Form3.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4440
      Width           =   1815
   End
   Begin VB.CommandButton cmdShirt 
      Height          =   2055
      Left            =   600
      Picture         =   "Form3.frx":10DC
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4560
      Width           =   2175
   End
   Begin VB.CommandButton cmdPicture 
      Height          =   2055
      Left            =   4920
      Picture         =   "Form3.frx":2399
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2280
      Width           =   2055
   End
   Begin VB.CommandButton cmdPuck 
      Height          =   1935
      Left            =   840
      Picture         =   "Form3.frx":3E76
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2400
      Width           =   1815
   End
   Begin VB.CommandButton cmdHat 
      Height          =   2055
      Left            =   4800
      Picture         =   "Form3.frx":4E61
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   2175
   End
   Begin VB.CommandButton cmdJersey 
      Height          =   2055
      Left            =   720
      Picture         =   "Form3.frx":89C9
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   240
      Width           =   1935
   End
   Begin VB.Label lblSize2 
      Caption         =   "Choose Size"
      Height          =   375
      Left            =   2880
      TabIndex        =   13
      Top             =   5040
      Width           =   1215
   End
   Begin VB.Label lblSize 
      Caption         =   "Choose Size"
      Height          =   375
      Left            =   2880
      TabIndex        =   11
      Top             =   960
      Width           =   1215
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim runningTotal As Single

Private Sub cmdPicture_Click()
runningTotal = runningTotal + 116.99
picResults.Print "Picture", FormatCurrency(116.99)
End Sub

Private Sub cmdClear_Click()
runningTotal = 0
picResults.Cls
End Sub

Private Sub cmdHat_Click()
runningTotal = runningTotal + 29.99
picResults.Print "Flames Hat"; Tab(15); , FormatCurrency(29.99)
End Sub

Private Sub cmdJersey_Click()
Dim Jsize As String
Jsize = txtSize.Text
runningTotal = runningTotal + 132.99
picResults.Print "Flames Jersey"; "("; Jsize; ") "; FormatCurrency(132.99); ""
End Sub

Private Sub cmdMug_Click()
runningTotal = runningTotal + 33.99
picResults.Print "Flames Mug", FormatCurrency(33.99)
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
runningTotal = runningTotal + 60.99
picResults.Print "Flames Hoody"; "("; Jsize; ") "; FormatCurrency(60.99)
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


