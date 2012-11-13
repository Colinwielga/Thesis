VERSION 5.00
Begin VB.Form frmBigBowlOrder 
   BackColor       =   &H00000040&
   Caption         =   "Form1"
   ClientHeight    =   8940
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12360
   LinkTopic       =   "Form1"
   ScaleHeight     =   8940
   ScaleWidth      =   12360
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return to Big Bowl "
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   14160
      TabIndex        =   22
      Top             =   10560
      Width           =   3375
   End
   Begin VB.CommandButton cmdSubmit 
      Caption         =   "Click here to Submit Order"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   18240
      TabIndex        =   21
      Top             =   9840
      Width           =   1095
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   15000
      TabIndex        =   19
      Top             =   11400
      Width           =   1695
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear order form"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   15000
      TabIndex        =   18
      Top             =   9600
      Width           =   1695
   End
   Begin VB.CommandButton cmdTotal 
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   15000
      TabIndex        =   17
      Top             =   8880
      Width           =   1695
   End
   Begin VB.PictureBox picRunningTotal 
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7335
      Left            =   13440
      ScaleHeight     =   7275
      ScaleWidth      =   5955
      TabIndex        =   16
      Top             =   1440
      Width           =   6015
   End
   Begin VB.CommandButton cmdVeggie2 
      Caption         =   "Yellow Curry Vegetable with Tofu"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1920
      TabIndex        =   15
      Top             =   10680
      Width           =   2055
   End
   Begin VB.CommandButton cmdChicken2 
      Caption         =   "Lemon Chicken"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1800
      TabIndex        =   14
      Top             =   4200
      Width           =   2055
   End
   Begin VB.CommandButton cmdVeggie1 
      Caption         =   "Thai Green Vegetable Curry"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1920
      TabIndex        =   7
      Top             =   9720
      Width           =   2055
   End
   Begin VB.CommandButton cmdSeafood2 
      Caption         =   "Teriyaki Glazed Fresh Salmon"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1920
      TabIndex        =   6
      Top             =   8640
      Width           =   2055
   End
   Begin VB.CommandButton Seafood1 
      Caption         =   "Sweet Ginger Sea Scallops and Shrimp"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1920
      TabIndex        =   5
      Top             =   7680
      Width           =   2055
   End
   Begin VB.CommandButton Beef2 
      Caption         =   "Mongolian Beef"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1800
      TabIndex        =   4
      Top             =   4920
      Width           =   2055
   End
   Begin VB.CommandButton Beef1 
      Caption         =   "Beef and Broccoli"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1920
      TabIndex        =   3
      Top             =   6720
      Width           =   2055
   End
   Begin VB.CommandButton cmdChicken1 
      Caption         =   "Kung Pao Chicken"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1920
      TabIndex        =   2
      Top             =   5880
      Width           =   2055
   End
   Begin VB.CommandButton cmdApp2 
      Caption         =   "Chinese Chicken Lettuce Wraps"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      TabIndex        =   1
      Top             =   2400
      Width           =   1695
   End
   Begin VB.CommandButton cmdApp1 
      Caption         =   "Crispy Chicken Egg Rolls"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   2400
      Width           =   1695
   End
   Begin VB.Image imgfood5 
      Height          =   2415
      Left            =   10680
      Picture         =   "frmBigBowlOrder.frx":0000
      Stretch         =   -1  'True
      Top             =   6840
      Width           =   2535
   End
   Begin VB.Image imgfood6 
      Height          =   2610
      Left            =   10680
      Picture         =   "frmBigBowlOrder.frx":CA4A
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   2595
   End
   Begin VB.Image imgfood4 
      Height          =   2655
      Left            =   8160
      Picture         =   "frmBigBowlOrder.frx":16819
      Stretch         =   -1  'True
      Top             =   9240
      Width           =   2895
   End
   Begin VB.Image imgfood3 
      Height          =   2535
      Left            =   5520
      Picture         =   "frmBigBowlOrder.frx":36091
      Stretch         =   -1  'True
      Top             =   6840
      Width           =   2655
   End
   Begin VB.Image imgfood2 
      Height          =   2535
      Left            =   7920
      Picture         =   "frmBigBowlOrder.frx":ECEE4
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   2775
   End
   Begin VB.Image imgfood1 
      Height          =   2535
      Left            =   5160
      Picture         =   "frmBigBowlOrder.frx":10C28C
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   2775
   End
   Begin VB.Line Line4 
      X1              =   4920
      X2              =   4920
      Y1              =   1200
      Y2              =   12000
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   4920
      Y1              =   9480
      Y2              =   9480
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   3840
      Y1              =   7440
      Y2              =   7440
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   4920
      Y1              =   5640
      Y2              =   5640
   End
   Begin VB.Label lblOrder 
      Alignment       =   2  'Center
      BackColor       =   &H00400000&
      Caption         =   "Order the chef's weekly specials "
      BeginProperty Font 
         Name            =   "Myriad Pro Cond"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000A&
      Height          =   1095
      Left            =   0
      TabIndex        =   20
      Top             =   120
      Width           =   19455
   End
   Begin VB.Label lblVeg 
      Alignment       =   2  'Center
      Caption         =   "Vegetable-->>>"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   13
      Top             =   9840
      Width           =   1440
   End
   Begin VB.Label lblSeafood 
      Alignment       =   2  'Center
      Caption         =   "Seafood-->>>"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   -120
      TabIndex        =   12
      Top             =   8160
      Width           =   1575
   End
   Begin VB.Label lblMeat 
      Alignment       =   2  'Center
      Caption         =   "Beef-->>>"
      BeginProperty Font 
         Name            =   "Myriad Web Pro"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   11
      Top             =   6600
      Width           =   1575
   End
   Begin VB.Label lblChicken 
      Alignment       =   2  'Center
      Caption         =   "Chicken-->>>"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   10
      Top             =   4560
      Width           =   1455
   End
   Begin VB.Label lblMain 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "Main Entrees"
      BeginProperty Font 
         Name            =   "Myriad Pro Cond"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   0
      TabIndex        =   9
      Top             =   3480
      Width           =   3855
   End
   Begin VB.Label lblApps 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "Appetizers"
      BeginProperty Font 
         Name            =   "Myriad Pro Cond"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   0
      TabIndex        =   8
      Top             =   1440
      Width           =   3735
   End
End
Attribute VB_Name = "frmBigBowlOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'CSCI VB Project: Big Bowl
'BigBowlOrder
'Elizabeth K. Sturlaugson
'Due Date: Friday, March 28th, 2008
'The purpose of this form is to allow the user to enter items of the weekly specials and "order" them for pickup
'Within this form, the user also sees a total cost for the items and a description of the items purchased before it submits its order to the restaurant


Option Explicit
Dim RunningTotal As Single

Private Sub Beef1_Click()
'declare the variables
'cost of Beef and Broccoli
Dim Beef1 As Single
Beef1 = 12.95

RunningTotal = Beef1 + RunningTotal
picRunningTotal.Print "Beef and Broccoli", FormatCurrency(Beef1)

'description of Beef and Broccoli
MsgBox "Tender beef, broccoli florets, shitake mushrooms, light garlic wine sauce", , "Description"


End Sub

Private Sub Beef2_Click()
'declare the variables
'cost of Mongolian Beef
Dim Beef2 As Single
Beef2 = 13.95

RunningTotal = Beef2 + RunningTotal
picRunningTotal.Print "Mongolian Beef", FormatCurrency(Beef2)

'description of Mongolian Beef
MsgBox "Tender beef, mushrooms and green onions", , "Description"

End Sub

Private Sub cmdApp1_Click()
'delcare the varaibles
'cost of Egg Rolls
Dim App1 As Single
App1 = 5.95

RunningTotal = App1 + RunningTotal
picRunningTotal.Print "Crispy Chicken Egg Rolls", FormatCurrency(App1)

'description of egg rolls
MsgBox "Chicken, napa cabbage, bamboo shoots, shitake mushrooms, plum & sesame mustard sauces", , "Description"



End Sub

Private Sub cmdApp2_Click()
'declare the variables
'cost of Lettuce Wraps
Dim App2 As Single
App2 = 6.25

RunningTotal = App2 + RunningTotal
picRunningTotal.Print "Chinese Chicken Lettuce Wraps", FormatCurrency(App2)

'description of Lettuce Wraps
MsgBox "Chicken, bibb lettuce, hoisin sauce", , "Description"


End Sub

Private Sub cmdChicken1_Click()
'declare the variables
'cost of Kung Pao Chicken
Dim Chicken1 As Single
Chicken1 = 11.95

RunningTotal = Chicken1 + RunningTotal
picRunningTotal.Print "Kung Pao Chicken", FormatCurrency(Chicken1)

'description of Kung Pao Chicken
MsgBox "Chicken, blackened chillies, roasted peanuts, sweet spicy sauce", , "Description"


End Sub

Private Sub cmdChicken2_Click()
'declare the variables
'cost of Lemon Chicken
Dim Chicken2 As Single
Chicken2 = 12.95

RunningTotal = Chicken2 + RunningTotal
picRunningTotal.Print "Lemon Chicken", FormatCurrency(Chicken2)

'description of Lemon Chicken
MsgBox "Crispy golden chicken, fresh lemon, ginger, red pepper", , "Description"


End Sub

Private Sub cmdClear_Click()
'clears form and running total
picRunningTotal.Cls
RunningTotal = 0

End Sub

Private Sub cmdQuit_Click()
'ends/quits program
End
End Sub

Private Sub cmdReturn_Click()
frmBigBowl.Show
frmBigBowlOrder.Hide

End Sub

Private Sub cmdSeafood2_Click()
'declare the variables
'cost of Salmon
Dim Seafood2 As Single
Seafood2 = 15.9

RunningTotal = Seafood2 + RunningTotal
picRunningTotal.Print "Teriyaki Glazed Fresh Salmon", FormatCurrency(Seafood2)

'description of Salmon
MsgBox "Naturally raised pacific northwest salmon, our own teriyaki sauce, fried rice", , "Description"


End Sub

Private Sub cmdSubmit_Click()
Dim Order As String
Dim Number As Single

'gathering contact information
Order = InputBox("Please enter your first name", "Order")
Number = InputBox("Please enter your telephone number (no hyphens)", "Contact Information")
picRunningTotal.Print "Thank you "; Order; " for your order."
picRunningTotal.Print " We will contact you when it is ready for pick-up."

'based on the total of the order, the waitng time varies

If RunningTotal < 50 Then
picRunningTotal.Print "Your estimated waiting time is 20 to 30 minutes."
ElseIf RunningTotal > 50 Then
picRunningTotal.Print "Your estimated waiting time is 30 to 45 minutes."

End If




End Sub
Private Sub cmdTotal_Click()
picRunningTotal.Print "***************************"
picRunningTotal.Print "Sub Total", FormatCurrency(RunningTotal)

'declare the variables
'calculations
Dim Tax As Single
Tax = RunningTotal * 0.065
picRunningTotal.Print "Tax", FormatCurrency(Tax)

'declare the variables
'calculations
Dim Total As Single
Total = RunningTotal + Tax
picRunningTotal.Print "Total", FormatCurrency(Total)

MsgBox "Please click on the Submit Order form.", , "Submit Order"

End Sub

Private Sub cmdVeggie1_Click()
'declare the variables
'cost of Thai Green Vegetable Curry
Dim Veggie1 As Single
Veggie1 = 9.95

RunningTotal = Veggie1 + RunningTotal
picRunningTotal.Print "Thai Green Vegetable Curry", FormatCurrency(Veggie1)

'description of Thai Green Vegetable Curry
MsgBox "Seasonal vegetables, baby bok choy, sweet green beans, green curry", , "Description"


End Sub

Private Sub cmdVeggie2_Click()
'declare the variables
'cost of Vegetables with Tofu
Dim Veggie2 As Single
Veggie2 = 10.95

picRunningTotal.Print "Yellow Curry Vegetable with Tofu", FormatCurrency(Veggie2)

'description of Vegetables with Currry
MsgBox "Bock choy, green beans, seasonal vegetables, yellow coconut curry sauce", , "Description"


End Sub

Private Sub Seafood1_Click()
'declare the variables
'cost of Sea Scallops and Shrimp

Dim Seafood1 As Single
Seafood1 = 14.9

RunningTotal = Seafood1 + RunningTotal
picRunningTotal.Print "Sweet Ginger Sea Scallops and Shrimp", FormatCurrency(Seafood1)

'description of Sea Scallops and Shrimp
MsgBox "Fresh sea scallops, gulf shrimp, mushrooms, sweet vinegar-soy glaze", , "Description"



End Sub
