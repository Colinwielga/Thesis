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
   Begin VB.CommandButton cmdOrder 
      Caption         =   "List By Price"
      Height          =   975
      Left            =   1680
      TabIndex        =   13
      Top             =   10200
      Width           =   2055
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   975
      Left            =   9240
      TabIndex        =   11
      Top             =   10200
      Width           =   2055
   End
   Begin VB.PictureBox picResults2 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   8175
      Left            =   7920
      ScaleHeight     =   8115
      ScaleWidth      =   4275
      TabIndex        =   10
      Top             =   1800
      Width           =   4335
   End
   Begin VB.CommandButton cmdCompute 
      Caption         =   "Compute Total"
      Height          =   975
      Left            =   9240
      TabIndex        =   9
      Top             =   600
      Width           =   2055
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   975
      Left            =   5400
      TabIndex        =   8
      Top             =   9000
      Width           =   2055
   End
   Begin VB.CommandButton cmdStudy 
      Caption         =   "Study Guide"
      Height          =   975
      Left            =   5400
      TabIndex        =   7
      Top             =   7800
      Width           =   2055
   End
   Begin VB.CommandButton cmdSunglasses 
      Caption         =   "Sunglasses"
      Height          =   975
      Left            =   5400
      TabIndex        =   6
      Top             =   6600
      Width           =   2055
   End
   Begin VB.CommandButton cmdSweatpants 
      Caption         =   "Sweat Pants"
      Height          =   975
      Left            =   5400
      TabIndex        =   5
      Top             =   4200
      Width           =   2055
   End
   Begin VB.CommandButton cmdSweatshirt 
      Caption         =   "Sweatshirt"
      Height          =   975
      Left            =   5400
      TabIndex        =   4
      Top             =   3000
      Width           =   2055
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load Prices"
      Height          =   975
      Left            =   1680
      TabIndex        =   3
      Top             =   600
      Width           =   2055
   End
   Begin VB.CommandButton cmdTee 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Tee-Shirt"
      Height          =   975
      Left            =   5400
      TabIndex        =   2
      Top             =   1800
      Width           =   2055
   End
   Begin VB.CommandButton cmdHat 
      Caption         =   "Hat"
      Height          =   975
      Left            =   5400
      TabIndex        =   1
      Top             =   5400
      Width           =   2055
   End
   Begin VB.PictureBox picResults1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   8175
      Left            =   600
      ScaleHeight     =   8115
      ScaleWidth      =   4275
      TabIndex        =   0
      Top             =   1800
      Width           =   4335
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
      Left            =   4920
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
Dim PosOrder As Integer
Dim Pass As Integer
Dim Temp As String
Dim I As Integer
Dim Merchandise(1 To 20) As String
Dim Price(1 To 20) As Single
Dim RunningTotal As Single

    'This form is for the store. It loads a list of items and prices to an array.  The array is printed and can be sorted.
    'Also, there are command buttons that have prices stored in them so the consumer can make purchases of one or more items.
    'The program then calculates the total cost of all the items selected including tax.

Private Sub cmdClear_Click()
    'This clears all values in the picResults2 column and sets the total to "0" to start from the beginning.
picResults2.Cls
RunningTotal = 0

End Sub

Private Sub cmdCompute_Click()
    'This prints the total, tax, and them added together to show how much the list of items costs for the purchaser.
picResults2.Print ""
picResults2.Print "**********************************************************"

picResults2.Print "Subtotal"; Tab(20); FormatCurrency(RunningTotal)
picResults2.Print "Tax"; Tab(20); FormatCurrency(0.07 * RunningTotal)
picResults2.Print "Total"; Tab(20); FormatCurrency(RunningTotal + (0.07 * RunningTotal))

End Sub

Private Sub cmdHat_Click()
    'Stores the price for Hats, and then adds the amount to the total and prints
RunningTotal = RunningTotal + 17
picResults2.Print "Hat"; Tab(20); FormatCurrency(17)

End Sub

Private Sub cmdLoad_Click()
    'Loads the program with the list of merchandise and prices to use for the form.
Open App.Path & "\merchandise.txt" For Input As #1
    Ctr = 0
    Do Until EOF(1)
        Ctr = Ctr + 1
        Input #1, Merchandise(Ctr), Price(Ctr)
    Loop
Close #1

    'Clears every time clicked on and prints the list of items loaded from the program.
picResults1.Cls
picResults1.Print "Item"; Tab(20); "Price"
picResults1.Print "**********************************************************"

For Pos = 1 To Ctr
    picResults1.Print Merchandise(Pos); Tab(20); FormatCurrency(Price(Pos))
Next Pos

End Sub

Private Sub cmdOrder_Click()
    'Clears and prints the list of items in order of highest price to lowest price.
picResults1.Cls
picResults1.Print "Item"; Tab(20); "Price"
picResults1.Print "**********************************************************"

For Pass = 1 To Ctr - 1
    For PosOrder = 1 To Ctr - Pass
        If Price(PosOrder) < Price(PosOrder + 1) Then
            Temp = Price(PosOrder)
            Price(PosOrder) = Price(PosOrder + 1)
            Price(PosOrder + 1) = Temp
            Temp = Merchandise(PosOrder)
            Merchandise(PosOrder) = Merchandise(PosOrder + 1)
            Merchandise(PosOrder + 1) = Temp
        End If
    Next PosOrder
Next Pass

For I = 1 To Ctr
    picResults1.Print Merchandise(I); Tab(20); FormatCurrency(Price(I))
Next I

End Sub

Private Sub cmdQuit_Click()
    'Quits the program.
End
End Sub

Private Sub cmdStudy_Click()
    'Stores the price for the study guide, and then adds the amount to the total and prints
RunningTotal = RunningTotal + 28
picResults2.Print "Study Guide"; Tab(20); FormatCurrency(28)

End Sub

Private Sub cmdSunglasses_Click()
    'Stores the price for the sunglasses, and then adds the amount to the total and prints
RunningTotal = RunningTotal + 12
picResults2.Print "Sunglasses"; Tab(20); FormatCurrency(12)

End Sub

Private Sub cmdSweatpants_Click()
    'Stores the price for the sweatpants, and then adds the amount to the total and prints
RunningTotal = RunningTotal + 30.5
picResults2.Print "Sweatpants"; Tab(20); FormatCurrency(30.5)

End Sub

Private Sub cmdSweatshirt_Click()
    'Stores the price for the sweatshirt, and then adds the amount to the total and prints
RunningTotal = RunningTotal + 35.75
picResults2.Print "Sweatshirt"; Tab(20); FormatCurrency(35.75)

End Sub

Private Sub cmdTee_Click()
    'Stores the price for the Tee-shirt, and then adds the amount to the total and prints
RunningTotal = RunningTotal + 21.5
picResults2.Print "Tee-Shirt"; Tab(20); FormatCurrency(21.5)

End Sub
