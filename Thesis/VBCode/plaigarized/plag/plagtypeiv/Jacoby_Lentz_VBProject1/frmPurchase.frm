VERSION 5.00
Begin VB.Form frmPurchase
   Caption         =   "Form1"
   ClientHeight    =   8235
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11370
   LinkTopic       =   "Form1"
   Picture         =   "frmPurchase.frx":0000
   ScaleHeight     =   8235
   ScaleWidth      =   11370
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGoBack
      Caption         =   "Return to Main"
      Height          =   375
      Left            =   9600
      TabIndex        =   3
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton cmdSet
      Caption         =   "Buy the Complete Set"
      Height          =   855
      Left            =   360
      TabIndex        =   2
      Top             =   3960
      Width           =   2295
   End
   Begin VB.CommandButton cmdSeason
      BackColor       =   &H000000FF&
      Caption         =   "Buy 1 Season"
      Height          =   855
      Left            =   360
      MaskColor       =   &H000000FF&
      TabIndex        =   1
      Top             =   2040
      UseMaskColor    =   -1  'True
      Width           =   2295
   End
   Begin VB.PictureBox picResults
      BackColor       =   &H0000FFFF&
      Height          =   3375
      Left            =   3240
      ScaleHeight     =   3315
      ScaleWidth      =   5115
      TabIndex        =   0
      Top             =   2400
      Width           =   5175
   End
End
Attribute VB_Name = "frmPurchase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdGoBack_Click()
frmMain.Show
frmPurchase.Hide
End Sub

Private Sub cmdSeason_Click()
    Dim X As Single
    Dim Total, SubTotal As Single
    Dim Tax As Single
    Tax = 0.07 * Total
    X = 28.99
    Total = X
    SubTotal = Total + Tax
    picResults.Print "1 Season  "; Total
    picResults.Print "Tax       "; FormatCurrency(Tax, 2)
    picResults.Print "------------------"
    picResults.Print "Total for 1 season "; FormatCurrency(SubTotal)
End Sub

Private Sub cmdSet_Click()
Dim X As Single
    Dim Tax As Single
    Dim Total, SubTotal As Single
    picResults.Print
    Tax = 0.07 * Total
    X = 145.99
    Total = X
    SubTotal = Total + Tax
    picResults.Print "Complete Set  "; Total
    picResults.Print "Tax           "; FormatCurrency(Tax, 2)
    picResults.Print "------------------"
    picResults.Print "Total for the complete set "; FormatCurrency(SubTotal)
    picResults.Print
    picResults.Print "***You save 30 bucks by buying the complete set***"
End Sub
