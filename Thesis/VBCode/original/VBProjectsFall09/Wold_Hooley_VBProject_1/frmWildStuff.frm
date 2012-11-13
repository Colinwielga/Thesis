VERSION 5.00
Begin VB.Form frmWildStuff 
   BackColor       =   &H00004000&
   Caption         =   "Form4"
   ClientHeight    =   9720
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11910
   FillColor       =   &H00004000&
   LinkTopic       =   "Form4"
   ScaleHeight     =   9720
   ScaleWidth      =   11910
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdMainMenu 
      Height          =   1935
      Left            =   8880
      Picture         =   "frmWildStuff.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   7680
      Width           =   2175
   End
   Begin VB.TextBox txtSize2 
      Height          =   495
      Left            =   2760
      TabIndex        =   11
      Top             =   6120
      Width           =   1335
   End
   Begin VB.TextBox txtSize 
      Height          =   495
      Left            =   2880
      TabIndex        =   10
      Top             =   1680
      Width           =   975
   End
   Begin VB.PictureBox picResults 
      Height          =   6135
      Left            =   7080
      ScaleHeight     =   6075
      ScaleWidth      =   3675
      TabIndex        =   9
      Top             =   720
      Width           =   3735
   End
   Begin VB.CommandButton cmdQuit 
      Height          =   1335
      Left            =   5760
      Picture         =   "frmWildStuff.frx":B472
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   8040
      Width           =   2295
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H80000002&
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "NancyBlue"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   8280
      Width           =   2055
   End
   Begin VB.CommandButton cmdTotal 
      BackColor       =   &H80000002&
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "NancyBlue"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   8280
      Width           =   2535
   End
   Begin VB.CommandButton cmdMug 
      Height          =   1695
      Left            =   4800
      Picture         =   "frmWildStuff.frx":14424
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5160
      Width           =   1815
   End
   Begin VB.CommandButton cmdShirt 
      Height          =   2175
      Left            =   600
      Picture         =   "frmWildStuff.frx":1540A
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4920
      Width           =   2055
   End
   Begin VB.CommandButton cmdBanner 
      Height          =   1815
      Left            =   4440
      Picture         =   "frmWildStuff.frx":162CE
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3000
      Width           =   2175
   End
   Begin VB.CommandButton cmdPuck 
      Height          =   1815
      Left            =   720
      Picture         =   "frmWildStuff.frx":17894
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2760
      Width           =   1935
   End
   Begin VB.CommandButton cmdHat 
      Height          =   1815
      Left            =   4320
      Picture         =   "frmWildStuff.frx":1889E
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   600
      Width           =   2295
   End
   Begin VB.CommandButton cmdJersey 
      Height          =   1935
      Left            =   600
      Picture         =   "frmWildStuff.frx":19384
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   360
      Width           =   2175
   End
   Begin VB.Label lblSize2 
      Caption         =   "Choose Size"
      Height          =   375
      Left            =   2880
      TabIndex        =   13
      Top             =   5640
      Width           =   1215
   End
   Begin VB.Label lblSize 
      Caption         =   "Choose Size"
      Height          =   255
      Left            =   2880
      TabIndex        =   12
      Top             =   1320
      Width           =   1095
   End
End
Attribute VB_Name = "frmWildStuff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim runningTotal As Single

Private Sub cmdBanner_Click()
runningTotal = runningTotal + 14.99
picResults.Print "Wild Picture"; Tab(30); FormatCurrency(14.99)
End Sub

Private Sub cmdClear_Click()
runningTotal = 0
picResults.Cls
End Sub

Private Sub cmdHat_Click()
runningTotal = runningTotal + 29.99
picResults.Print "Wild Hat"; Tab(30); , FormatCurrency(29.99)
End Sub

Private Sub cmdJersey_Click()
Dim Jsize As String
Jsize = txtSize.Text
runningTotal = runningTotal + 132.99
picResults.Print "Wild Jersey"; "("; Jsize; ")"; Tab(30); FormatCurrency(132.99); ""
End Sub

Private Sub cmdMainMenu_Click()
frmWildStuff.Hide
frmMainMenu.Show
End Sub

Private Sub cmdMug_Click()
runningTotal = runningTotal + 9.99
picResults.Print "Wild Zamboni"; Tab(30); FormatCurrency(9.99)
End Sub

Private Sub cmdPuck_Click()
runningTotal = runningTotal + 38.99
picResults.Print "Signed Puck"; Tab(30); FormatCurrency(38.99)
End Sub

Private Sub cmdQuit_Click()
End
End Sub

Private Sub cmdShirt_Click()
Dim Jsize As String
Jsize = txtSize2.Text
runningTotal = runningTotal + 27.99
picResults.Print "Wild Shirt"; "("; Jsize; ")"; Tab(30); FormatCurrency(27.99)
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


