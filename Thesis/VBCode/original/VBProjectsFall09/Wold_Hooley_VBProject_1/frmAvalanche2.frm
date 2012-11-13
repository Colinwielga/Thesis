VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000080&
   Caption         =   "Form1"
   ClientHeight    =   9675
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14280
   LinkTopic       =   "Form1"
   ScaleHeight     =   9675
   ScaleWidth      =   14280
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtSize 
      Height          =   615
      Left            =   2400
      TabIndex        =   11
      Top             =   1440
      Width           =   1455
   End
   Begin VB.TextBox txtSize2 
      Height          =   615
      Left            =   2160
      TabIndex        =   10
      Top             =   6360
      Width           =   1575
   End
   Begin VB.CommandButton cmdMug 
      Height          =   1935
      Left            =   4320
      Picture         =   "Form1.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5520
      Width           =   1935
   End
   Begin VB.CommandButton cmdShirt 
      Height          =   1935
      Left            =   480
      Picture         =   "Form1.frx":0E87
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5400
      Width           =   1575
   End
   Begin VB.CommandButton cmdBanner 
      Height          =   2175
      Left            =   4560
      Picture         =   "Form1.frx":1D37
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2880
      Width           =   1575
   End
   Begin VB.CommandButton cmdPuck 
      Height          =   1815
      Left            =   600
      Picture         =   "Form1.frx":34B5
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3000
      Width           =   1815
   End
   Begin VB.CommandButton cmdHat 
      Height          =   2055
      Left            =   4440
      Picture         =   "Form1.frx":43D1
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   600
      Width           =   2055
   End
   Begin VB.CommandButton cmdJersey 
      Height          =   2055
      Left            =   600
      Picture         =   "Form1.frx":9C64
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   480
      Width           =   1335
   End
   Begin VB.CommandButton cmdTotal 
      Caption         =   "Total"
      Height          =   975
      Left            =   1440
      TabIndex        =   3
      Top             =   7800
      Width           =   3015
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   1095
      Left            =   5880
      TabIndex        =   2
      Top             =   7680
      Width           =   2295
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   1095
      Left            =   9600
      TabIndex        =   1
      Top             =   7680
      Width           =   2175
   End
   Begin VB.PictureBox picResults 
      Height          =   6135
      Left            =   6840
      ScaleHeight     =   6075
      ScaleWidth      =   5835
      TabIndex        =   0
      Top             =   600
      Width           =   5895
   End
   Begin VB.Label lblSize2 
      Caption         =   "Choose your size"
      Height          =   375
      Left            =   2280
      TabIndex        =   13
      Top             =   5880
      Width           =   1335
   End
   Begin VB.Label lblSize 
      Caption         =   "Choose your size"
      Height          =   375
      Left            =   2520
      TabIndex        =   12
      Top             =   840
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim runningTotal As Single

Private Sub cmdBanner_Click()
runningTotal = runningTotal + 24.99
picResults.Print "Banner", FormatCurrency(24.99)
End Sub

Private Sub cmdClear_Click()
runningTotal = 0
picResults.Cls
End Sub

Private Sub cmdHat_Click()
runningTotal = runningTotal + 21.99
picResults.Print "Avs Hat"; Tab(15); , FormatCurrency(21.99)
End Sub

Private Sub cmdJersey_Click()
Dim Jsize As String
Jsize = txtSize.Text
runningTotal = runningTotal + 114.99
picResults.Print "Avs Jersey"; "("; Jsize; ") "; FormatCurrency(114.99); ""
End Sub

Private Sub cmdMug_Click()
runningTotal = runningTotal + 29.99
picResults.Print "Stanley Cup Mug", FormatCurrency(29.99)
End Sub

Private Sub cmdPuck_Click()
runningTotal = runningTotal + 29.99
picResults.Print "Signed Puck", FormatCurrency(29.99)
End Sub

Private Sub cmdQuit_Click()
End
End Sub

Private Sub cmdShirt_Click()
Dim Jsize As String
Jsize = txtSize2.Text
runningTotal = runningTotal + 19.99
picResults.Print "Avs Shirt"; "("; Jsize; ") "; FormatCurrency(19.99)
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

