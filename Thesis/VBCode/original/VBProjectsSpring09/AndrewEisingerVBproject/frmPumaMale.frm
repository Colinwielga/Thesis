VERSION 5.00
Begin VB.Form frmPumaMale 
   BackColor       =   &H0000C0C0&
   Caption         =   "PumaMale"
   ClientHeight    =   9015
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13935
   LinkTopic       =   "Form1"
   Picture         =   "frmPumaMale.frx":0000
   ScaleHeight     =   9015
   ScaleWidth      =   13935
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H0000FF00&
      Caption         =   "Clear"
      Height          =   1215
      Left            =   8640
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   7680
      Width           =   1815
   End
   Begin VB.CommandButton cmdTotal 
      BackColor       =   &H000000FF&
      Caption         =   "Total"
      Height          =   1215
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   6360
      Width           =   1815
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H0000FF00&
      Caption         =   "Quit"
      Height          =   1215
      Left            =   10920
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   7680
      Width           =   1815
   End
   Begin VB.CommandButton cmdGoBackHome 
      BackColor       =   &H000000FF&
      Caption         =   "Go Back To Store Home"
      Height          =   1215
      Left            =   10080
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   6360
      Width           =   1815
   End
   Begin VB.CommandButton cmdPumaGo 
      BackColor       =   &H000000FF&
      Caption         =   "Go Back To Puma Store"
      Height          =   1215
      Left            =   12000
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6360
      Width           =   1815
   End
   Begin VB.PictureBox picResults 
      Height          =   6135
      Left            =   8160
      ScaleHeight     =   6075
      ScaleWidth      =   5475
      TabIndex        =   8
      Top             =   120
      Width           =   5535
   End
   Begin VB.CommandButton cmdItalia 
      BackColor       =   &H000080FF&
      Caption         =   "Italia Graphic Tee"
      Height          =   1215
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   7080
      Width           =   1815
   End
   Begin VB.CommandButton cmdKing 
      BackColor       =   &H00FF0000&
      Caption         =   "King Legend's Shirt"
      Height          =   1215
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5760
      Width           =   1815
   End
   Begin VB.CommandButton cmdSF 
      BackColor       =   &H000080FF&
      Caption         =   "SF Graphic Tee"
      Height          =   1215
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7080
      Width           =   1815
   End
   Begin VB.CommandButton cmdGP 
      BackColor       =   &H000080FF&
      Caption         =   "GP Cat Shoes"
      Height          =   1215
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7080
      Width           =   1815
   End
   Begin VB.CommandButton cmdFurioV 
      BackColor       =   &H000080FF&
      Caption         =   "Furio V Mesh Shoes"
      Height          =   1215
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7080
      Width           =   1815
   End
   Begin VB.CommandButton cmdSpeedCat 
      BackColor       =   &H00FF0000&
      Caption         =   "Speed Cat Shoes"
      Height          =   1215
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5760
      Width           =   1815
   End
   Begin VB.CommandButton cmdSFPOLO 
      BackColor       =   &H00FF0000&
      Caption         =   "SF Polo Shirt"
      Height          =   1215
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5760
      Width           =   1815
   End
   Begin VB.CommandButton cmdJagoII 
      BackColor       =   &H00FF0000&
      Caption         =   "Jago II Shoes"
      Height          =   1215
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5760
      Width           =   1815
   End
End
Attribute VB_Name = "frmPumaMale"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'AthleticStore
' PumaMale
' Andrew Eisinger
' 3/19/09
'This program lets the female select what she would like to buy and then
'this program also adds up the subtotal, multiplys a tax to it and adds the tax and subtotal into a total price
Dim Total As Single, Jago As Single, Furio As Single, King As Single, SFPOLO As Single, SFGraphic As Single, SubTotal As Single, Tax As Single
Dim Italia As Single, Speed As Single, GP As Single

Private Sub cmdClear_Click()
'   Clear the picture box
    picResults.Cls
    Total = 0
    Tax = 0
    SubTotal = 0
End Sub


Private Sub cmdFurioV_Click()
 Furio = 101.79
    SubTotal = Furio + SubTotal
    picResults.Print "Furio V Mesh Shoes: "; FormatCurrency(Furio)
End Sub

Private Sub cmdGoBackHome_Click()
frmStoreHome.Show
frmPumaMale.Hide
End Sub

Private Sub cmdGP_Click()
 GP = 60.45
    SubTotal = GP + SubTotal
    picResults.Print "GP Cat Shoes: "; FormatCurrency(GP)
End Sub

Private Sub cmdItalia_Click()
 Italia = 23.21
    SubTotal = Italia + SubTotal
    picResults.Print "Italia Graphic Tee: "; FormatCurrency(Italia)
End Sub

Private Sub cmdJagoII_Click()
  Jago = 53.45
    SubTotal = Jago + SubTotal
    picResults.Print "Jago II Shoes: "; FormatCurrency(Jago)
End Sub

Private Sub cmdKing_Click()
 King = 29.87
    SubTotal = King + SubTotal
    picResults.Print "King Legend's Shirt: "; FormatCurrency(King)
End Sub

Private Sub cmdPumaGo_Click()
frmPuma1.Show
frmPumaMale.Hide
End Sub

Private Sub cmdQuit_Click()
End
End Sub

Private Sub cmdSF_Click()
 SFGraphic = 31.29
    SubTotal = SFGraphic + SubTotal
    picResults.Print "SF Graphic Tee: "; FormatCurrency(SFGraphic)
End Sub

Private Sub cmdSFPOLO_Click()
 SFPOLO = 45
    SubTotal = SFPOLO + SubTotal
    picResults.Print "SF Polo Shirt: "; FormatCurrency(SFPOLO)
End Sub

Private Sub cmdSpeedCat_Click()
 Speed = 79.89
    SubTotal = Speed + SubTotal
    picResults.Print "Speed Cat Shoes: "; FormatCurrency(Speed)
End Sub


Private Sub cmdTotal_Click()
   picResults.Print "*************"
' Add the number here
    Tax = SubTotal * 0.09
    Total = SubTotal + Tax
    picResults.Print "SubTotal: "; FormatCurrency(SubTotal)
    picResults.Print "Tax: "; FormatCurrency(Tax)
    picResults.Print "Total: "; FormatCurrency(Total)
End Sub

