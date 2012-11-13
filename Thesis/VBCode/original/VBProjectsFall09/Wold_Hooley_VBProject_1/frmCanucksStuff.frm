VERSION 5.00
Begin VB.Form frmCanucksstuff 
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
   Begin VB.CommandButton Command1 
      Height          =   1935
      Left            =   11280
      Picture         =   "frmCanucksStuff.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   7080
      Width           =   2055
   End
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
   Begin VB.PictureBox picResults 
      Height          =   5895
      Left            =   8400
      ScaleHeight     =   5835
      ScaleWidth      =   4395
      TabIndex        =   9
      Top             =   1080
      Width           =   4455
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H80000002&
      Height          =   1335
      Left            =   7560
      Picture         =   "frmCanucksStuff.frx":B472
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   7200
      Width           =   2415
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
      Height          =   1335
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   7320
      Width           =   2535
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
      Height          =   1335
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7320
      Width           =   2415
   End
   Begin VB.CommandButton cmdStick 
      Height          =   1935
      Left            =   6120
      Picture         =   "frmCanucksStuff.frx":14424
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5040
      Width           =   1935
   End
   Begin VB.CommandButton cmdShirt 
      Height          =   1935
      Left            =   360
      Picture         =   "frmCanucksStuff.frx":14CA7
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4800
      Width           =   2535
   End
   Begin VB.CommandButton cmdBanner 
      Height          =   2175
      Left            =   5760
      Picture         =   "frmCanucksStuff.frx":15CAF
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2760
      Width           =   2295
   End
   Begin VB.CommandButton cmdPuck 
      Height          =   1695
      Left            =   1200
      Picture         =   "frmCanucksStuff.frx":16F92
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2640
      Width           =   1695
   End
   Begin VB.CommandButton cmdHat 
      Height          =   2055
      Left            =   6000
      Picture         =   "frmCanucksStuff.frx":17CBE
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   360
      Width           =   2055
   End
   Begin VB.CommandButton cmdJersey 
      Height          =   1935
      Left            =   840
      Picture         =   "frmCanucksStuff.frx":1BD6E
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
Attribute VB_Name = "frmCanucksstuff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim runningTotal As Single

Private Sub cmdBanner_Click()
runningTotal = runningTotal + 27.99
picResults.Print "Vancouver Picture"; Tab(35); FormatCurrency(27.99)
End Sub

Private Sub cmdClear_Click()
runningTotal = 0
picResults.Cls
End Sub

Private Sub cmdHat_Click()
runningTotal = runningTotal + 29.99
picResults.Print "Vancouver Hat"; Tab(30); , FormatCurrency(29.99)
End Sub

Private Sub cmdJersey_Click()
Dim Jsize As String
Jsize = txtSize.Text
runningTotal = runningTotal + 132.99
picResults.Print "Vancouver Jersey"; "("; Jsize; ")"; Tab(35); FormatCurrency(132.99); ""
End Sub

Private Sub cmdStick_Click()
runningTotal = runningTotal + 11.99
picResults.Print "Hockey Stick"; Tab(35); FormatCurrency(11.99)
End Sub

Private Sub cmdPuck_Click()
runningTotal = runningTotal + 38.99
picResults.Print "Signed Puck"; Tab(35); FormatCurrency(38.99)
End Sub

Private Sub cmdQuit_Click()
End
End Sub

Private Sub cmdShirt_Click()
Dim Jsize As String
Jsize = txtSize2.Text
runningTotal = runningTotal + 27.99
picResults.Print "Vancouver Shirt"; "("; Jsize; ")"; Tab(35); FormatCurrency(27.99)
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


Private Sub Command1_Click()
frmCanucksstuff.Hide
frmMainMenu.Show
End Sub
