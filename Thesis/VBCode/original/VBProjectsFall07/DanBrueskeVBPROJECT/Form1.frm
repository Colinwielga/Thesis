VERSION 5.00
Begin VB.Form FrmStore 
   BackColor       =   &H00000000&
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   15240
   ScaleWidth      =   25080
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   975
      Left            =   10080
      TabIndex        =   11
      Top             =   5400
      Width           =   2055
   End
   Begin VB.PictureBox picResults2 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H000000C0&
      Height          =   8175
      Left            =   6600
      ScaleHeight     =   8115
      ScaleWidth      =   3075
      TabIndex        =   10
      Top             =   1800
      Width           =   3135
   End
   Begin VB.CommandButton cmdCompute 
      Caption         =   "Compute Total"
      Height          =   975
      Left            =   7080
      TabIndex        =   9
      Top             =   600
      Width           =   2055
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   975
      Left            =   4100
      TabIndex        =   8
      Top             =   9000
      Width           =   2055
   End
   Begin VB.CommandButton cmdStudy 
      Caption         =   "Study Guide"
      Height          =   975
      Left            =   4100
      TabIndex        =   7
      Top             =   7800
      Width           =   2055
   End
   Begin VB.CommandButton cmdSunglasses 
      Caption         =   "Sunglasses"
      Height          =   975
      Left            =   4080
      TabIndex        =   6
      Top             =   6600
      Width           =   2055
   End
   Begin VB.CommandButton cmdSweatpants 
      Caption         =   "Sweat Pants"
      Height          =   975
      Left            =   4100
      TabIndex        =   5
      Top             =   4200
      Width           =   2055
   End
   Begin VB.CommandButton cmdSweatshirt 
      Caption         =   "Sweatshirt"
      Height          =   975
      Left            =   4100
      TabIndex        =   4
      Top             =   3000
      Width           =   2055
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load Prices"
      Height          =   975
      Left            =   1080
      TabIndex        =   3
      Top             =   600
      Width           =   2055
   End
   Begin VB.CommandButton cmdTee 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Tee-Shirt"
      Height          =   975
      Left            =   4100
      TabIndex        =   2
      Top             =   1800
      Width           =   2055
   End
   Begin VB.CommandButton cmdHat 
      Caption         =   "Hat"
      Height          =   975
      Left            =   4100
      TabIndex        =   1
      Top             =   5400
      Width           =   2055
   End
   Begin VB.PictureBox picResults1 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H000000C0&
      Height          =   8175
      Left            =   600
      ScaleHeight     =   8115
      ScaleWidth      =   3075
      TabIndex        =   0
      Top             =   1800
      Width           =   3135
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "The Store"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   855
      Left            =   3600
      TabIndex        =   12
      Top             =   600
      Width           =   3015
   End
End
Attribute VB_Name = "FrmStore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Ctr As Integer
Dim Pos As Integer
Dim Merchandise(1 To 20) As String
Dim Price(1 To 20) As Single
Dim RunningTotal As Single

Private Sub cmdClear_Click()

picResults2.Cls
RunningTotal = 0

End Sub

Private Sub cmdCompute_Click()

picResults2.Print ""
picResults2.Print "**********************************************************"
picResults2.Print "Subtotal", FormatCurrency(RunningTotal)
picResults2.Print "Tax", FormatCurrency(0.07 * RunningTotal)
picResults2.Print "Total", FormatCurrency(RunningTotal + (0.07 * RunningTotal))

End Sub

Private Sub cmdHat_Click()

RunningTotal = RunningTotal + 17
picResults2.Print "Hat", FormatCurrency(17)

End Sub

Private Sub cmdLoad_Click()

Open App.Path & "\merchandise.txt" For Input As #1
    Ctr = 0
    Do Until EOF(1)
        Ctr = Ctr + 1
        Input #1, Merchandise(Ctr), Price(Ctr)
    Loop
Close #1

picResults1.Cls
picResults1.Print "Item", "Price"
picResults1.Print "**********************************************************"

For Pos = 1 To Ctr
    picResults1.Print Merchandise(Pos), FormatCurrency(Price(Pos))
Next Pos

End Sub

Private Sub cmdQuit_Click()
End
End Sub

Private Sub cmdStudy_Click()
RunningTotal = RunningTotal + 28
picResults2.Print "Study Guide", FormatCurrency(28)
End Sub

Private Sub cmdSunglasses_Click()
RunningTotal = RunningTotal + 12
picResults2.Print "Sunglasses", FormatCurrency(12)
End Sub

Private Sub cmdSweatpants_Click()
RunningTotal = RunningTotal + 30.5
picResults2.Print "Sweatpants", FormatCurrency(30.5)
End Sub

Private Sub cmdSweatshirt_Click()
RunningTotal = RunningTotal + 35.75
picResults2.Print "Sweatshirt", FormatCurrency(35.75)
End Sub

Private Sub cmdTee_Click()
RunningTotal = RunningTotal + 21.5
picResults2.Print "Tee-Shirt", FormatCurrency(21.5)
End Sub
