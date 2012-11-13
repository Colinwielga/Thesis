VERSION 5.00
Begin VB.Form frmFlamesStuff 
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
   Begin VB.CommandButton Command1 
      Height          =   1935
      Left            =   10200
      Picture         =   "frmFlamesStuff.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   6000
      Width           =   2175
   End
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
      Top             =   240
      Width           =   3495
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H80000009&
      Height          =   1095
      Left            =   6840
      Picture         =   "frmFlamesStuff.frx":B472
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6720
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
      Height          =   975
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6840
      Width           =   2295
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
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6720
      Width           =   2295
   End
   Begin VB.CommandButton cmdMug 
      Height          =   2295
      Left            =   4920
      Picture         =   "frmFlamesStuff.frx":14424
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4440
      Width           =   1815
   End
   Begin VB.CommandButton cmdShirt 
      Height          =   2055
      Left            =   600
      Picture         =   "frmFlamesStuff.frx":15500
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4560
      Width           =   2175
   End
   Begin VB.CommandButton cmdPicture 
      Height          =   2055
      Left            =   4920
      Picture         =   "frmFlamesStuff.frx":167BD
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2280
      Width           =   2055
   End
   Begin VB.CommandButton cmdPuck 
      Height          =   1935
      Left            =   840
      Picture         =   "frmFlamesStuff.frx":1829A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2400
      Width           =   1815
   End
   Begin VB.CommandButton cmdHat 
      Height          =   2055
      Left            =   4800
      Picture         =   "frmFlamesStuff.frx":19285
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   2175
   End
   Begin VB.CommandButton cmdJersey 
      Height          =   2055
      Left            =   720
      Picture         =   "frmFlamesStuff.frx":1CDED
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
Attribute VB_Name = "frmFlamesStuff"
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
picResults.Print "Flames Hat"; Tab(30); FormatCurrency(29.99)
End Sub

Private Sub cmdJersey_Click()
Dim Jsize As String
Jsize = txtSize.Text
runningTotal = runningTotal + 132.99
picResults.Print "Flames Jersey"; "("; Jsize; ")"; Tab(30); FormatCurrency(132.99); ""
End Sub

Private Sub cmdMug_Click()
runningTotal = runningTotal + 33.99
picResults.Print "Flames Mug"; Tab(30); FormatCurrency(33.99)
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
runningTotal = runningTotal + 60.99
picResults.Print "Flames Hoody"; "("; Jsize; ")"; Tab(30); FormatCurrency(60.99)
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
frmFlamesStuff.Hide
frmMainMenu.Show
End Sub
