VERSION 5.00
Begin VB.Form frmBookStore 
   BackColor       =   &H003D30AD&
   Caption         =   "Buy a Book "
   ClientHeight    =   10500
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   ScaleHeight     =   10500
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture6 
      BorderStyle     =   0  'None
      Height          =   2175
      Left            =   10440
      Picture         =   "frmBookStore.frx":0000
      ScaleHeight     =   2175
      ScaleWidth      =   2655
      TabIndex        =   15
      Top             =   480
      Width           =   2655
   End
   Begin VB.PictureBox Picture5 
      BackColor       =   &H003D30AD&
      BorderStyle     =   0  'None
      Height          =   2895
      Left            =   7920
      Picture         =   "frmBookStore.frx":41C9
      ScaleHeight     =   2895
      ScaleWidth      =   2655
      TabIndex        =   14
      Top             =   120
      Width           =   2655
   End
   Begin VB.CommandButton cmdClr 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   9840
      TabIndex        =   12
      Top             =   9120
      Width           =   1695
   End
   Begin VB.CommandButton cmdGoBack 
      Caption         =   "Return to Homepage"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   11880
      TabIndex        =   11
      Top             =   9120
      Width           =   2415
   End
   Begin VB.CommandButton cmdTotal 
      Caption         =   "Calculate Total+Tax"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   7800
      TabIndex        =   10
      Top             =   9120
      Width           =   1695
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H0084C11E&
      Height          =   5535
      Left            =   8640
      ScaleHeight     =   5475
      ScaleWidth      =   4515
      TabIndex        =   9
      Top             =   3240
      Width           =   4575
   End
   Begin VB.PictureBox Picture4 
      Height          =   3015
      Left            =   4440
      Picture         =   "frmBookStore.frx":74E5
      ScaleHeight     =   2955
      ScaleWidth      =   1995
      TabIndex        =   7
      Top             =   5760
      Width           =   2055
   End
   Begin VB.PictureBox Picture2 
      Height          =   3135
      Left            =   3600
      Picture         =   "frmBookStore.frx":9694
      ScaleHeight     =   3075
      ScaleWidth      =   3675
      TabIndex        =   5
      Top             =   1080
      Width           =   3735
   End
   Begin VB.PictureBox Picture1 
      Height          =   3135
      Left            =   120
      Picture         =   "frmBookStore.frx":1494C
      ScaleHeight     =   3075
      ScaleWidth      =   2955
      TabIndex        =   4
      Top             =   1080
      Width           =   3015
   End
   Begin VB.CommandButton cmdQuixote 
      Caption         =   "Click to Purchase:             Don Quixote     $19.95"
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   4200
      TabIndex        =   3
      Top             =   9000
      Width           =   2535
   End
   Begin VB.CommandButton cmdJane 
      Caption         =   "Click to Purchase:                  Jane Eyre               $12.95"
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   360
      TabIndex        =   2
      Top             =   9000
      Width           =   2535
   End
   Begin VB.CommandButton cmdGrapes 
      Caption         =   "Click to Purchase:       The Grapes Of Wrath $14.95"
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   4080
      TabIndex        =   1
      Top             =   4440
      Width           =   2775
   End
   Begin VB.CommandButton cmdPnP 
      Caption         =   "Click to Purchase:  Pride and Prejudice $9.95"
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   240
      TabIndex        =   0
      Top             =   4440
      Width           =   2655
   End
   Begin VB.PictureBox Picture3 
      Height          =   3015
      Left            =   120
      Picture         =   "frmBookStore.frx":1AD3D
      ScaleHeight     =   2955
      ScaleWidth      =   3075
      TabIndex        =   6
      Top             =   5760
      Width           =   3135
   End
   Begin VB.Label lblHelp 
      BackColor       =   &H003D30AD&
      Caption         =   "Will you help me buy some books?"
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0084C11E&
      Height          =   735
      Left            =   480
      TabIndex        =   13
      Top             =   240
      Width           =   6615
   End
   Begin VB.Label lblClickbook 
      BackColor       =   &H003D30AD&
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0084C11E&
      Height          =   375
      Left            =   480
      TabIndex        =   8
      Top             =   1320
      Width           =   6735
   End
End
Attribute VB_Name = "frmBookStore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'allow the user to choose different products, add the cost
'and calculate price with tax and subtotal

'declare a Running Total
Dim RunningTotal As Single

Private Sub cmdClr_Click()
'gives an option to clear total so it does not add to previous totals
picResults.Cls
RunningTotal = 0
End Sub

Private Sub cmdGoBack_Click()
'returns user to main page
frmBookStore.Hide
frmPersonality.Show

End Sub

Private Sub cmdGrapes_Click()
'adds price to the RunningTotal with every click
RunningTotal = RunningTotal + 14.95
'displays name of book and price in picturebox
picResults.Print "The Grapes of Wrath  ", FormatCurrency(14.95)
MsgBox "Brilliant!"
End Sub

Private Sub cmdJane_Click()
'adds price to the RunningTotal with every click
RunningTotal = RunningTotal + 12.95
'displays name of book and price in picturebox
picResults.Print "Jane Eyre ", , FormatCurrency(12.95)
MsgBox "Jane Eyre is one of my favorites!"
End Sub

Private Sub cmdPnP_Click()
'adds price to the RunningTotal with every click
RunningTotal = RunningTotal + 9.95
'displays name of book and price in picturebox
picResults.Print "Pride and Prejudice  ", FormatCurrency(9.95)
MsgBox "Good Choice! I love Jane Austin!"
End Sub

Private Sub cmdQuixote_Click()
'adds price to the RunningTotal with every click
RunningTotal = RunningTotal + 19.95
picResults.Print "Don Quixote  ", , FormatCurrency(19.95)
'displays name of book and price in picturebox
MsgBox "Good Choice! A Classic!"
End Sub

Private Sub cmdTotal_Click()
'Declare all variables
Dim Tax As Single
Dim PnP As Single
Dim Grapes As Single
Dim DonQ As Single
Dim Jane As Single
Dim Subtotal As Single
Dim Total As Single

'set the value of necessary variables

Tax = RunningTotal * 0.065
Subtotal = RunningTotal
Total = RunningTotal + Tax

'display Total results with tax, formatting for currency
picResults.Print "---------------------------"
picResults.Print "%6.5 Sales Tax   ", FormatCurrency(Tax)
picResults.Print "Subtotal  ", , FormatCurrency(Subtotal)
picResults.Print "Total   ", , FormatCurrency(Total)

End Sub


