VERSION 5.00
Begin VB.Form frmFrozen 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form1"
   ClientHeight    =   11010
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton Option4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Salisbury Steak"
      Height          =   495
      Left            =   8280
      TabIndex        =   21
      Top             =   8040
      Width           =   2655
   End
   Begin VB.OptionButton Option3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Meatloaf"
      Height          =   495
      Left            =   8280
      TabIndex        =   20
      Top             =   7560
      Width           =   2535
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Lasanga"
      Height          =   375
      Left            =   8280
      TabIndex        =   19
      Top             =   7080
      Width           =   2415
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Turkey and Mashed Potatoes"
      Height          =   375
      Left            =   8280
      TabIndex        =   17
      Top             =   6600
      Width           =   2535
   End
   Begin VB.CommandButton cmdSpecial 
      BackColor       =   &H000000FF&
      Caption         =   "Cub Foods has a pizza special. Buy 5 pizzas, get 1 free! Click here to add pizzas to your cart."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   3600
      Width           =   2775
   End
   Begin VB.TextBox txtSupreme 
      Height          =   375
      Left            =   4080
      TabIndex        =   13
      Top             =   5280
      Width           =   975
   End
   Begin VB.TextBox txtSausage 
      Height          =   375
      Left            =   4080
      TabIndex        =   12
      Top             =   4680
      Width           =   975
   End
   Begin VB.TextBox txtCheese 
      Height          =   375
      Left            =   4080
      TabIndex        =   11
      Top             =   4080
      Width           =   975
   End
   Begin VB.TextBox txtPepperoni 
      Height          =   375
      Left            =   4080
      TabIndex        =   10
      Top             =   3480
      Width           =   975
   End
   Begin VB.CommandButton cmdContinue 
      BackColor       =   &H00FF00FF&
      Caption         =   "Continue Shopping or Check Out"
      Height          =   1215
      Left            =   11400
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   9120
      Width           =   3015
   End
   Begin VB.CommandButton cmdCalculateFrozen 
      BackColor       =   &H00FF8080&
      Caption         =   "Calculate Frozen Foods Subtotal"
      Height          =   1215
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   9120
      Width           =   3015
   End
   Begin VB.PictureBox picResults 
      Height          =   4095
      Left            =   8760
      ScaleHeight     =   4035
      ScaleWidth      =   4395
      TabIndex        =   2
      Top             =   1080
      Width           =   4455
   End
   Begin VB.CommandButton cmdSearchIceCream 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Click Here to Shop Ice Cream"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   960
      Width           =   2895
   End
   Begin VB.Label lbldirections 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Click on flavors below to add to cart."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8040
      TabIndex        =   18
      Top             =   6120
      Width           =   3255
   End
   Begin VB.Image Image4 
      Height          =   1335
      Left            =   11400
      Picture         =   "frmFrozen.frx":0000
      Top             =   5400
      Width           =   1905
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Frozen Dinners"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   8760
      TabIndex        =   16
      Top             =   5640
      Width           =   3735
   End
   Begin VB.Label lblWelcomeFrozen 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Welcome to the Frozen Food Section"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3120
      TabIndex        =   0
      Top             =   120
      Width           =   8055
   End
   Begin VB.Label lblPizza 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Frozen Pizzas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   3360
      TabIndex        =   5
      Top             =   2400
      Width           =   3135
   End
   Begin VB.Label lblpizzaprice 
      BackColor       =   &H00FFFFFF&
      Caption         =   "All Pizzas Cost $5.50.  Please enter amount in adjacent box. Enter 0 if you would not like that flavor."
      Height          =   375
      Left            =   3120
      TabIndex        =   15
      Top             =   2880
      Width           =   4575
   End
   Begin VB.Image Image2 
      Height          =   2400
      Left            =   240
      Picture         =   "frmFrozen.frx":1094
      Top             =   3120
      Width           =   2400
   End
   Begin VB.Label lblSupreme 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Supreme"
      Height          =   375
      Left            =   3240
      TabIndex        =   9
      Top             =   5280
      Width           =   855
   End
   Begin VB.Label lblSausage 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Sausage"
      Height          =   375
      Left            =   3240
      TabIndex        =   8
      Top             =   4680
      Width           =   855
   End
   Begin VB.Label lblCheese 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Cheese"
      Height          =   375
      Left            =   3240
      TabIndex        =   7
      Top             =   4080
      Width           =   855
   End
   Begin VB.Label lblPepperoni 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Pepperoni"
      Height          =   375
      Left            =   3240
      TabIndex        =   6
      Top             =   3480
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   5025
      Left            =   120
      Picture         =   "frmFrozen.frx":134E0
      Top             =   5880
      Width           =   7500
   End
   Begin VB.Image Image3 
      Height          =   4500
      Left            =   -720
      Picture         =   "frmFrozen.frx":23A2D
      Top             =   240
      Width           =   4395
   End
End
Attribute VB_Name = "frmFrozen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim IceCreamFlavors(1 To 50) As String, Prices(1 To 50) As Single, CTR As Integer, CostOfPizza As Single, CostOfIceCream As Single

Private Sub cmdCalculateFrozen_Click()
'displays frozen food total spent
picResults.Print "**********************************************************"
picResults.Print "Frozen Foods Subtotal: "; FormatCurrency(FrozenRunningTotal)
picResults.Print "**********************************************************"

End Sub

Private Sub cmdContinue_Click()
'adds frozen food total to the runnign total and displays
RunningTotal = BakeryRunningTotal + ProduceRunningTotal + FrozenRunningTotal
MsgBox "Total spent so far is: " & FormatCurrency(RunningTotal)
'takes user back to enter form
frmProduce.Hide
frmBakery.Hide
frmFrozen.Hide
frmCheckOut.Hide
frmEnter.Show

End Sub

Private Sub cmdSearchIceCream_Click()
Dim Pos As Integer
Dim Found As Boolean
Dim Sname As String

CTR = 0

Open App.Path & "\IceCreamFlavors.txt" For Input As #1
'this opens a data file and reads it into an array
Do Until EOF(1)
    CTR = CTR + 1
    Input #1, IceCreamFlavors(CTR), Prices(CTR)
Loop
Close #1
'the user is asked to input a search flavor
Sname = InputBox("Due to an extensive ice cream selection, please enter flavor to search for")
'a match and stop search is done
Do While Found = False And Pos < CTR
    Pos = Pos + 1
    If LCase(Sname) = LCase(IceCreamFlavors(Pos)) Then
        Found = True
    End If
Loop
'the results of the search are displayed, and if the item was found than it is added to the frozen foods running total
    If Found = True Then
        MsgBox "Yes! We do do have that flavor and it has been added to your cart."
        CostOfIceCream = CostOfIceCream + Prices(Pos)
        picResults.Print IceCreamFlavors(Pos); " ice cream"; Tab(35); FormatCurrency(Prices(Pos))
        FrozenRunningTotal = FrozenRunningTotal + CostOfIceCream
    Else
        MsgBox "I'm sorry, we don't carry that flavor"
    End If
    
End Sub

Private Sub cmdSpecial_Click()
Dim Pepperoni As Integer, Sausage As Integer, Cheese As Integer, Supreme As Integer
Dim SumPizzas As Integer, Price As Single

Price = 5.5
Pepperoni = txtPepperoni.Text
Sausage = txtSausage.Text
Cheese = txtCheese.Text
Supreme = txtSupreme.Text
'the user inputs a number by textbox and this calculates how many pizzas are desired
SumPizzas = Pepperoni + Sausage + Cheese + Supreme
'the amount of pizzas is greater than 5 then one pizza is subtracted and the cost is calculated
    If SumPizzas > 5 Then
        CostOfPizza = (SumPizzas - 1) * Price
        FrozenRunningTotal = FrozenRunningTotal + CostOfPizza
    Else 'amount of pizza is less than 5 and the cost is calculated
        CostOfPizza = SumPizzas * Price
        FrozenRunningTotal = FrozenRunningTotal + CostOfPizza
    End If

'results are displayed
picResults.Print SumPizzas; " Pizza(s) "; Tab(35); FormatCurrency(CostOfPizza)



End Sub


Private Sub Option1_Click()
Dim Turkey As Single

Turkey = 3.75
'if the user clicks the option button then the price of the option is added to the frozen running total
FrozenRunningTotal = FrozenRunningTotal + Turkey

picResults.Print "Turkey Frozen Dinner"; Tab(35); FormatCurrency(Turkey)


End Sub

Private Sub Option2_Click()
Dim Lasanga As Single

Lasanga = 4
'if the user clicks the option button then the price of the option is added to the frozen running total
FrozenRunningTotal = FrozenRunningTotal + Lasanga

picResults.Print "Lasanga Frozen Dinner"; Tab(35); FormatCurrency(Lasanga)


End Sub

Private Sub Option3_Click()
Dim Meatloaf As Single

Meatloaf = 3.25
'if the user clicks the option button then the price of the option is added to the frozen running total
FrozenRunningTotal = FrozenRunningTotal + Meatloaf

picResults.Print "Meatloaf Frozen Dinner"; Tab(35); FormatCurrency(Meatloaf)
End Sub

Private Sub Option4_Click()
Dim Steak As Single

Steak = 4.25
'if the user clicks the option button then the price of the option is added to the frozen running total
FrozenRunningTotal = FrozenRunningTotal + Steak

picResults.Print "Steak Frozen Dinner", Tab(35); FormatCurrency(Steak)

End Sub
