VERSION 5.00
Begin VB.Form frmGrocery_Store 
   BackColor       =   &H00FF8080&
   Caption         =   "Form1"
   ClientHeight    =   8385
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10260
   LinkTopic       =   "Form1"
   ScaleHeight     =   8385
   ScaleWidth      =   10260
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CommandButton Command15 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5040
      TabIndex        =   18
      Top             =   6960
      Width           =   1695
   End
   Begin VB.CommandButton Command14 
      Caption         =   "Back to Main Page"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      TabIndex        =   17
      Top             =   6960
      Width           =   1695
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Proceed to Check Out!"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   7320
      TabIndex        =   16
      Top             =   6960
      Width           =   1695
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Tomato Juice"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3360
      TabIndex        =   15
      Top             =   5760
      Width           =   1095
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Yogurt"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3360
      TabIndex        =   14
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Flax Seed"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1800
      TabIndex        =   13
      Top             =   4560
      Width           =   1095
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Whole Grain Pasta"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1800
      TabIndex        =   12
      Top             =   3360
      Width           =   1095
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Green Beans"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      TabIndex        =   11
      Top             =   5760
      Width           =   1095
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Berries"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1800
      TabIndex        =   10
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Salmon"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3360
      TabIndex        =   9
      Top             =   3360
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Eggs"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3360
      TabIndex        =   8
      Top             =   4560
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Organic Rice"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      TabIndex        =   7
      Top             =   3360
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "100 % Fruit Juice"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1800
      TabIndex        =   6
      Top             =   5760
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Granola"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      TabIndex        =   5
      Top             =   4560
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Apple"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      TabIndex        =   4
      Top             =   2160
      Width           =   1095
   End
   Begin VB.PictureBox picRunningTotal 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5055
      Left            =   4920
      ScaleHeight     =   4995
      ScaleWidth      =   4155
      TabIndex        =   0
      Top             =   1680
      Width           =   4215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF8080&
      Caption         =   $"frmGrocery_Store.frx":0000
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1455
      Left            =   4920
      TabIndex        =   3
      Top             =   240
      Width           =   4695
   End
   Begin VB.Label lblBodyMass 
      BackColor       =   &H00FF8080&
      Caption         =   "Wholesome"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   615
      Left            =   600
      TabIndex        =   2
      Top             =   120
      Width           =   4815
   End
   Begin VB.Label lblCalculator 
      BackColor       =   &H00FF8080&
      Caption         =   "Foods"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   39.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   975
      Left            =   1800
      TabIndex        =   1
      Top             =   600
      Width           =   3975
   End
End
Attribute VB_Name = "frmGrocery_Store"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim runningTotal As Single

Private Sub Command1_Click()
Dim Apple As Single
    Apple = 1#
    runningTotal = runningTotal + Apple
    picRunningTotal.Print " Apple"; , , FormatCurrency(Apple)
End Sub

Private Sub Command10_Click()
Dim FlaxSeed As Single
    FlaxSeed = 2.5
    runningTotal = runningTotal + FlaxSeed
    picRunningTotal.Print " Flax Seed"; , , FormatCurrency(FlaxSeed)
End Sub

Private Sub Command11_Click()
Dim Yogurt As Single
    Yogurt = 0.99
    runningTotal = runningTotal + Yogurt
    picRunningTotal.Print " Yogurt"; , , FormatCurrency(Yogurt)
End Sub

Private Sub Command12_Click()
Dim TomatoJuice As Single
    TomatoJuice = 1.5
    runningTotal = runningTotal + TomatoJuice
    picRunningTotal.Print " Tomato Juice "; , FormatCurrency(TomatoJuice)
End Sub

Private Sub Command13_Click()

Dim Response As Integer
      ' Displays a message box with the yes and no options.
      Response = MsgBox(prompt:="Are you sure you would like to purchase these things? ", Buttons:=vbYesNo)
      ' If statement to check if the yes button was selected.
      If Response = vbYes Then
         MsgBox "Your total comes to " & FormatCurrency(runningTotal)
      Else
      ' The no button was selected.
      frmTargetHeartRate.Show
      End If

End Sub

Private Sub Command14_Click()
frmMainpage.Show
frmGrocery_Store.Hide

End Sub

Private Sub Command15_Click()
    picRunningTotal.Cls
    
End Sub

Private Sub Command2_Click()
Dim Granola As Single
    Granola = 4#
    runningTotal = runningTotal + Granola
    picRunningTotal.Print " Granola "; , , FormatCurrency(Granola)
End Sub

Private Sub Command3_Click()
Dim FruitJuice As Single
    FruitJuice = 3.75
    runningTotal = runningTotal + FruitJuice
    picRunningTotal.Print " 100 % Fruit Juice "; , FormatCurrency(FruitJuice)
End Sub

Private Sub Command4_Click()
Dim OrganicRice As Single
    OrganicRice = 2.5
    runningTotal = runningTotal + OrganicRice
    picRunningTotal.Print " Organic Rice "; , FormatCurrency(OrganicRice)
End Sub

Private Sub Command5_Click()
Dim Eggs As Single
    Eggs = 6#
    runningTotal = runningTotal + Eggs
    picRunningTotal.Print " Eggs (a dozen) "; , FormatCurrency(Eggs)
End Sub

Private Sub Command6_Click()
Dim Salmon As Single
    Salmon = 10#
    runningTotal = runningTotal + Salmon
    picRunningTotal.Print " Salmon (per pound) "; , FormatCurrency(Salmon)
End Sub

Private Sub Command7_Click()
Dim Berries As Single
    Berries = 4#
    runningTotal = runningTotal + Berries
    picRunningTotal.Print " Berries "; , , FormatCurrency(Berries)
End Sub

Private Sub Command8_Click()
Dim GreenBeans As Single
    GreenBeans = 3#
    runningTotal = runningTotal + GreenBeans
    picRunningTotal.Print " Green Beans "; , FormatCurrency(GreenBeans)
End Sub

Private Sub Command9_Click()
Dim WholeGrainPasta As Single
    WholeGrainPasta = 5#
    runningTotal = runningTotal + WholeGrainPasta
    picRunningTotal.Print " Whole Grain Pasta "; , FormatCurrency(WholeGrainPasta)
End Sub
