VERSION 5.00
Begin VB.Form frmAvalancheStuff 
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
   Begin VB.CommandButton cmdMainMenu 
      Height          =   2175
      Left            =   11880
      Picture         =   "frmAvalancheStuff.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   7200
      Width           =   2055
   End
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
      Picture         =   "frmAvalancheStuff.frx":B472
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5520
      Width           =   1935
   End
   Begin VB.CommandButton cmdShirt 
      Height          =   1935
      Left            =   480
      Picture         =   "frmAvalancheStuff.frx":C2F9
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5400
      Width           =   1575
   End
   Begin VB.CommandButton cmdBanner 
      Height          =   2175
      Left            =   4680
      Picture         =   "frmAvalancheStuff.frx":D1A9
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2880
      Width           =   1575
   End
   Begin VB.CommandButton cmdPuck 
      Height          =   1815
      Left            =   600
      Picture         =   "frmAvalancheStuff.frx":E927
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3000
      Width           =   1815
   End
   Begin VB.CommandButton cmdHat 
      Height          =   2055
      Left            =   4200
      Picture         =   "frmAvalancheStuff.frx":F843
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   600
      Width           =   2055
   End
   Begin VB.CommandButton cmdJersey 
      Height          =   2055
      Left            =   600
      Picture         =   "frmAvalancheStuff.frx":150D6
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   480
      Width           =   1335
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
      Height          =   975
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7800
      Width           =   3015
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
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7800
      Width           =   2295
   End
   Begin VB.CommandButton cmdQuit 
      Height          =   1335
      Left            =   8640
      Picture         =   "frmAvalancheStuff.frx":1B365
      Style           =   1  'Graphical
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
Attribute VB_Name = "frmAvalancheStuff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim runningTotal As Single

Private Sub cmdBanner_Click()
runningTotal = runningTotal + 24.99
picResults.Print "Banner"; Tab(30); FormatCurrency(24.99)
End Sub

Private Sub cmdClear_Click()
runningTotal = 0
picResults.Cls
End Sub

Private Sub cmdHat_Click()
runningTotal = runningTotal + 21.99
picResults.Print "Avs Hat"; Tab(30); FormatCurrency(21.99)
End Sub

Private Sub cmdJersey_Click()
Dim Jsize As String
Jsize = txtSize.Text
runningTotal = runningTotal + 114.99
picResults.Print "Avs Jersey"; "("; Jsize; ")"; Tab(30); FormatCurrency(114.99); ""
End Sub

Private Sub cmdMainMenu_Click()
frmAvalancheStuff.Hide
frmMainMenu.Show
End Sub

Private Sub cmdMug_Click()
runningTotal = runningTotal + 29.99
picResults.Print "Stanley Cup Mug"; Tab(30); FormatCurrency(29.99)
End Sub

Private Sub cmdPuck_Click()
runningTotal = runningTotal + 29.99
picResults.Print "Signed Puck"; Tab(30); FormatCurrency(29.99)
End Sub

Private Sub cmdQuit_Click()
End
End Sub

Private Sub cmdShirt_Click()
Dim Jsize As String
Jsize = txtSize2.Text
runningTotal = runningTotal + 19.99
picResults.Print "Avs Shirt"; "("; Jsize; ")"; Tab(30); FormatCurrency(19.99)
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

