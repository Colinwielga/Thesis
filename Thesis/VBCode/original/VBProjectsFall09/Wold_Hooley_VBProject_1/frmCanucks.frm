VERSION 5.00
Begin VB.Form Form5 
   BackColor       =   &H00FF0000&
   Caption         =   "Form5"
   ClientHeight    =   9030
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13440
   LinkTopic       =   "Form5"
   ScaleHeight     =   9030
   ScaleWidth      =   13440
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtSize2 
      Height          =   375
      Left            =   3480
      TabIndex        =   12
      Top             =   6240
      Width           =   1455
   End
   Begin VB.TextBox txtSize 
      Height          =   615
      Left            =   3360
      TabIndex        =   10
      Top             =   1680
      Width           =   1455
   End
   Begin VB.PictureBox Picture1 
      Height          =   5895
      Left            =   8400
      ScaleHeight     =   5835
      ScaleWidth      =   4395
      TabIndex        =   9
      Top             =   1080
      Width           =   4455
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   1095
      Left            =   9360
      TabIndex        =   8
      Top             =   7560
      Width           =   3015
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   1335
      Left            =   5040
      TabIndex        =   7
      Top             =   7440
      Width           =   2535
   End
   Begin VB.CommandButton cmdTotal 
      Caption         =   "Total"
      Height          =   1335
      Left            =   1440
      TabIndex        =   6
      Top             =   7320
      Width           =   2415
   End
   Begin VB.CommandButton cmdStick 
      Height          =   1935
      Left            =   6120
      Picture         =   "Form5.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5040
      Width           =   1935
   End
   Begin VB.CommandButton cmdShirt 
      Height          =   1935
      Left            =   480
      Picture         =   "Form5.frx":0883
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4800
      Width           =   2535
   End
   Begin VB.CommandButton cmdBanner 
      Height          =   2175
      Left            =   5760
      Picture         =   "Form5.frx":188B
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2760
      Width           =   2295
   End
   Begin VB.CommandButton cmdPuck 
      Height          =   1695
      Left            =   720
      Picture         =   "Form5.frx":2B6E
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2640
      Width           =   2175
   End
   Begin VB.CommandButton cmdHat 
      Height          =   2055
      Left            =   5520
      Picture         =   "Form5.frx":389A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   360
      Width           =   2655
   End
   Begin VB.CommandButton cmdJersey 
      Height          =   1935
      Left            =   840
      Picture         =   "Form5.frx":794A
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   480
      Width           =   2055
   End
   Begin VB.Label lblSize2 
      Caption         =   "Choose Size"
      Height          =   495
      Left            =   3600
      TabIndex        =   13
      Top             =   5640
      Width           =   1215
   End
   Begin VB.Label lblSize 
      Caption         =   "Choose Size"
      Height          =   495
      Left            =   3480
      TabIndex        =   11
      Top             =   960
      Width           =   1215
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim runningTotal As Single

Private Sub cmdBanner_Click()
runningTotal = runningTotal + 27.99
picResults.Print "Vancouver Picture", FormatCurrency(27.99)
End Sub

Private Sub cmdClear_Click()
runningTotal = 0
picResults.Cls
End Sub

Private Sub cmdHat_Click()
runningTotal = runningTotal + 29.99
picResults.Print "Vancouver Hat"; Tab(15); , FormatCurrency(29.99)
End Sub

Private Sub cmdJersey_Click()
Dim Jsize As String
Jsize = txtSize.Text
runningTotal = runningTotal + 132.99
picResults.Print "Vancouver Jersey"; "("; Jsize; ") "; FormatCurrency(132.99); ""
End Sub

Private Sub cmdStick_Click()
runningTotal = runningTotal + 11.99
picResults.Print "Hockey Stick", FormatCurrency(11.99)
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
picResults.Print "Vancouver Shirt"; "("; Jsize; ") "; FormatCurrency(27.99)
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


