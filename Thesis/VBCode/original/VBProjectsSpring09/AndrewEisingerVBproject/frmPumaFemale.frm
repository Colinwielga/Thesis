VERSION 5.00
Begin VB.Form frmPumaFemale 
   BackColor       =   &H00FF00FF&
   Caption         =   "PumaFemale"
   ClientHeight    =   9495
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14250
   LinkTopic       =   "Form1"
   Picture         =   "frmPumaFemale.frx":0000
   ScaleHeight     =   9495
   ScaleWidth      =   14250
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGoPuma 
      BackColor       =   &H0000FFFF&
      Caption         =   "Go Back To Puma Home"
      Height          =   1215
      Left            =   11400
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   7800
      Width           =   1815
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00C000C0&
      Caption         =   "Quit"
      Height          =   1215
      Left            =   12120
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   6480
      Width           =   1815
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00C000C0&
      Caption         =   "Clear"
      Height          =   1215
      Left            =   10200
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   6480
      Width           =   1815
   End
   Begin VB.PictureBox picResults 
      Height          =   6495
      Left            =   8880
      ScaleHeight     =   6435
      ScaleWidth      =   5115
      TabIndex        =   10
      Top             =   0
      Width           =   5175
   End
   Begin VB.CommandButton cmdStoreHome 
      BackColor       =   &H0000FFFF&
      Caption         =   "Go Back To Store Home"
      Height          =   1215
      Left            =   9120
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   7800
      Width           =   1815
   End
   Begin VB.CommandButton cmdTotal 
      BackColor       =   &H00C000C0&
      Caption         =   "Total"
      Height          =   1215
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6480
      Width           =   1815
   End
   Begin VB.CommandButton cmdRun 
      BackColor       =   &H0000FFFF&
      Caption         =   "Run Skirt"
      Height          =   1215
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   7800
      Width           =   1815
   End
   Begin VB.CommandButton cmdWheelspin 
      BackColor       =   &H0000FFFF&
      Caption         =   "Wheelspin Patent Shoes"
      Height          =   1215
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7800
      Width           =   1815
   End
   Begin VB.CommandButton cmdTP 
      BackColor       =   &H0000FFFF&
      Caption         =   "TP 3/4 Tight"
      Height          =   1215
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7800
      Width           =   1815
   End
   Begin VB.CommandButton cmdEtoile 
      BackColor       =   &H0000FFFF&
      Caption         =   "Etoile Shoes"
      Height          =   1215
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7800
      Width           =   1815
   End
   Begin VB.CommandButton cmdFT 
      BackColor       =   &H00C000C0&
      Caption         =   "FT Graphic Tee"
      Height          =   1215
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6480
      Width           =   1815
   End
   Begin VB.CommandButton cmdDrift 
      BackColor       =   &H00C000C0&
      Caption         =   "Puma Drift Shoes"
      Height          =   1215
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6480
      Width           =   1815
   End
   Begin VB.CommandButton cmdVoltaic 
      BackColor       =   &H00C000C0&
      Caption         =   "Voltaic Shoes"
      Height          =   1215
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6480
      Width           =   1815
   End
   Begin VB.CommandButton cmdSpeedCat 
      BackColor       =   &H00C000C0&
      Caption         =   "Speed Cat Shoes"
      Height          =   1215
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6480
      Width           =   1815
   End
End
Attribute VB_Name = "frmPumaFemale"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'AthleticStore
' PumaFemale
' Andrew Eisinger
' 3/17/09
'This program lets the female select what she would like to buy and then
'this program also adds up the subtotal, multiplys a tax to it and adds the tax and subtotal into a total price
Dim Total As Single, Drift As Single, FT As Single, Etoile As Single, Run As Single, SubTotal As Single, Tax As Single
Dim SpeedCat As Single, Voltaic As Single, Wheelspin As Single, TP As Single
Private Sub cmdClear_Click()
'   Clear the picture box
    picResults.Cls
    Total = 0
    Tax = 0
    SubTotal = 0
End Sub

Private Sub cmdDrift_Click()
   Drift = 79.95
    SubTotal = Drift + SubTotal
' Add the number here
    picResults.Print "Puma Drift Shoes: "; FormatCurrency(Drift)
End Sub

Private Sub cmdEtoile_Click()
   Etoile = 88.39
    SubTotal = Etoile + SubTotal
' Add the number here
    picResults.Print "Etoile Shoes: "; FormatCurrency(Etoile)
End Sub

Private Sub cmdFT_Click()
FT = 32.21
SubTotal = FT + SubTotal
' Add the number here
    picResults.Print "FT Graphic Tee: "; FormatCurrency(FT)
End Sub

Private Sub cmdGoPuma_Click()
frmPuma1.Show
frmPumaFemale.Hide
End Sub

Private Sub cmdQuit_Click()
End
End Sub

Private Sub cmdRun_Click()
   Run = 47
    SubTotal = Run + SubTotal
' Add the number here
    picResults.Print "Run Skirt: "; FormatCurrency(Run)
End Sub

Private Sub cmdSpeedCat_Click()
    SpeedCat = 79.95
    SubTotal = SpeedCat + SubTotal
' Add the number here
    picResults.Print "Speed Cat Shoes: "; FormatCurrency(SpeedCat)
End Sub

Private Sub cmdStoreHome_Click()
frmStoreHome.Show
frmPumaFemale.Hide
End Sub

Private Sub cmdTotal_Click()
   picResults.Print "*************"
' Add the number here
    Tax = SubTotal * 0.1
    Total = SubTotal + Tax
    picResults.Print "SubTotal: "; FormatCurrency(SubTotal)
    picResults.Print "Tax: "; FormatCurrency(Tax)
    picResults.Print "Total: "; FormatCurrency(Total)
End Sub

Private Sub cmdTP_Click()
    TP = 35.34
    SubTotal = TP + SubTotal
' Add the number here
    picResults.Print "TP 3/4 Tight: "; FormatCurrency(TP)
End Sub

Private Sub cmdVoltaic_Click()
    Voltaic = 55.54
    SubTotal = Voltaic + SubTotal
' Add the number here
    picResults.Print "Voltaic Shoes: "; FormatCurrency(Voltaic)
End Sub

Private Sub cmdWheelspin_Click()
    Wheelspin = 101.98
    SubTotal = Wheelspin + SubTotal
' Add the number here
    picResults.Print "Wheelspin Patent Shoes: "; FormatCurrency(Wheelspin)
End Sub


