VERSION 5.00
Begin VB.Form frmDining2 
   BackColor       =   &H00FF8080&
   Caption         =   "Alaskan Dining"
   ClientHeight    =   7035
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11070
   LinkTopic       =   "Form1"
   ScaleHeight     =   7035
   ScaleWidth      =   11070
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picResults 
      Height          =   6135
      Left            =   3600
      ScaleHeight     =   6075
      ScaleWidth      =   6435
      TabIndex        =   5
      Top             =   720
      Width           =   6495
   End
   Begin VB.CommandButton cmdGoBacktoHome 
      BackColor       =   &H00FF80FF&
      Caption         =   "Return to Alaskan Home Page"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5160
      Width           =   2775
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H80000013&
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3960
      Width           =   2775
   End
   Begin VB.CommandButton cmdElegant2 
      BackColor       =   &H80000013&
      Caption         =   "Click to view elegant dining options"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2760
      Width           =   2775
   End
   Begin VB.CommandButton cmdCasual2 
      BackColor       =   &H80000013&
      Caption         =   "Click to view casual dining options"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1560
      Width           =   2775
   End
   Begin VB.Label lblAlaskanDining 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Alaskan Dining"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2775
   End
End
Attribute VB_Name = "frmDining2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Sunshine & Snow Cruise Lines
'Form Name: frmDining2
'Authors: Brittany Nosal & Kelly Sunder
'Date Written: 3/14/2009
'Objective: This form includes command buttons that shows the user either a casual dinner menu with prices
'included or an elegant dinner menu with prices included.

Option Explicit
Private Sub cmdCasual2_Click()
Dim CasualFoods(1 To 30) As String, CasualPrices(1 To 100) As Single
Dim CTR As Integer
CTR = 0

picResults.Cls
Open App.Path & "\CasualDiningOptions2.txt" For Input As #1

picResults.Print "Food Item"; Tab(45); "Price"
picResults.Print "***************************************************************"

Do While Not EOF(1)
    CTR = CTR + 1
    Input #1, CasualFoods(CTR), CasualPrices(CTR)
    picResults.Print CasualFoods(CTR); Tab(45); FormatCurrency(CasualPrices(CTR), 2)
Loop
Close #1
End Sub

Private Sub cmdElegant2_Click()
Dim ElegantFoods(1 To 30) As String, ElegantPrices(1 To 100) As Single
Dim CTR As Integer
CTR = 0

picResults.Cls
Open App.Path & "\ElegantDiningOptions2.txt" For Input As #1

picResults.Print "Food Item"; Tab(45); "Price"
picResults.Print "***************************************************************"

Do While Not EOF(1)
    CTR = CTR + 1
    Input #1, ElegantFoods(CTR), ElegantPrices(CTR)
    picResults.Print ElegantFoods(CTR); Tab(45); FormatCurrency(ElegantPrices(CTR), 2)
Loop
Close #1

End Sub

Private Sub cmdGoBacktoHome_Click()
frmDining2.Hide
frmAlaskanHome.Show
End Sub

Private Sub cmdClear_Click()
picResults.Cls
End Sub

