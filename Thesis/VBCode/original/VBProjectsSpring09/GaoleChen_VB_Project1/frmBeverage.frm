VERSION 5.00
Begin VB.Form frmBeverage 
   BackColor       =   &H00FF8080&
   Caption         =   "Form1"
   ClientHeight    =   8025
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10965
   LinkTopic       =   "Form1"
   Picture         =   "frmBeverage.frx":0000
   ScaleHeight     =   8025
   ScaleWidth      =   10965
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdTotal 
      Caption         =   "Total Price"
      Height          =   615
      Left            =   9360
      TabIndex        =   12
      Top             =   6120
      Width           =   1215
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   615
      Left            =   1680
      TabIndex        =   11
      Top             =   6120
      Width           =   855
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   615
      Left            =   480
      TabIndex        =   10
      Top             =   6120
      Width           =   855
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next"
      Height          =   615
      Left            =   1680
      TabIndex        =   9
      Top             =   5280
      Width           =   855
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      Height          =   615
      Left            =   480
      TabIndex        =   8
      Top             =   5280
      Width           =   855
   End
   Begin VB.TextBox txtWine 
      BackColor       =   &H8000000D&
      Height          =   615
      Left            =   1800
      TabIndex        =   7
      Top             =   3960
      Width           =   1095
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H8000000D&
      Height          =   5055
      Left            =   7800
      Picture         =   "frmBeverage.frx":160C7
      ScaleHeight     =   4995
      ScaleWidth      =   2715
      TabIndex        =   5
      Top             =   840
      Width           =   2775
   End
   Begin VB.TextBox txtPops 
      BackColor       =   &H8000000D&
      Height          =   1455
      Left            =   1800
      TabIndex        =   4
      Top             =   2400
      Width           =   1095
   End
   Begin VB.TextBox txtBeer 
      BackColor       =   &H8000000D&
      Height          =   1215
      Left            =   1800
      TabIndex        =   2
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Label4 
      BackColor       =   &H8000000D&
      Caption         =   "Wine is only $9 per bottle! Enter number here==>"
      Height          =   615
      Left            =   360
      TabIndex        =   6
      Top             =   3960
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000D&
      Caption         =   "Pops are just the price of its squareroot's price without coins! $2=>$1.41=>$1   How many do you want?==>"
      Height          =   1455
      Left            =   360
      TabIndex        =   3
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000D&
      Caption         =   "Enter number of beers you wish to buy--You'll get  10% off for 10 or above purchases! ==>"
      Height          =   1215
      Left            =   360
      TabIndex        =   1
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      Caption         =   "All drinks are on sale!"
      BeginProperty Font 
         Name            =   "Script MT Bold"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      TabIndex        =   0
      Top             =   120
      Width           =   9975
   End
End
Attribute VB_Name = "frmBeverage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Digital Menu
'Form Name: frmBeverage
'Authors: Gaole Chen
'Date Written: 3/8/09
'Objective: The user can order the  by clicking the vivid pictures.
'The form also calculate the total price automatically for the user.

Option Explicit

Private Sub cmdBack_Click()
frmBeverage.Hide
frmDessert.Show
End Sub

Private Sub cmdClear_Click()
picResults.Cls
End Sub

Private Sub cmdNext_Click()
frmBeverage.Hide
frmCheck.Show
End Sub

Private Sub cmdQuit_Click()
End
End Sub

Private Sub cmdTotal_Click()
'this button calculates the total price of the drinks after the discounts, and shows the result in the picturebox.

'declare the variables
Dim SubTotal As Single, Totalpop As Single, Pops As Integer, Wine As Integer, Beer As Integer, Tax As Single, Total As Single
Dim Totalwine As Single, Totalbeer As Single
'initialize with the value of SubTotal to zero
SubTotal = 0
'get input from the textbox of beer
Beer = txtBeer.Text
'computing costs
If Beer >= 10 Then
    Totalbeer = 1.8 * Beer
    picResults.Print Beer; " Beer: ", FormatCurrency(Totalbeer)
Else
    Totalbeer = 2 * Beer
    picResults.Print Beer; " Beer: ", FormatCurrency(Totalbeer)
End If

'get input from the textbox of pops
Pops = txtPops.Text
'computing costs
Totalpop = 1 * Pops
picResults.Print Pops; " Pops: ", FormatCurrency(Round(Sqr(Totalpop)))

'get input from the textbox of wine
Wine = txtWine.Text
Totalwine = 9 * Wine
picResults.Print Wine; " Wines: ", FormatCurrency(Totalwine)

SubTotal = Totalbeer + Totalpop + Totalwine
Tax = 0.08 * SubTotal
Total = SubTotal + Tax
picResults.Print
picResults.Print "All drinks cost:"; FormatCurrency(FormatNumber(SubTotal, 2))
picResults.Print "Taxes:", FormatCurrency(FormatNumber(Tax, 2))
picResults.Print "In total:", FormatCurrency(FormatNumber(Total, 2))

Totaldrinkcost = Totaldrinkcost + Total
Totalcost = Totalcost + Totaldrinkcost
End Sub

