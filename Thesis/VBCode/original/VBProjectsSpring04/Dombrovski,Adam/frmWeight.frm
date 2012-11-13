VERSION 5.00
Begin VB.Form frmWeight 
   BackColor       =   &H80000009&
   Caption         =   "Weight Management Page"
   ClientHeight    =   7710
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9060
   LinkTopic       =   "Form1"
   ScaleHeight     =   8750.06
   ScaleMode       =   0  'User
   ScaleWidth      =   9060
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCart2 
      Caption         =   "Click to Checkout"
      Height          =   975
      Left            =   6240
      TabIndex        =   24
      Top             =   6360
      Width           =   2535
   End
   Begin VB.PictureBox picWCart 
      Height          =   495
      Left            =   6480
      ScaleHeight     =   435
      ScaleWidth      =   2235
      TabIndex        =   23
      Top             =   4680
      Width           =   2295
   End
   Begin VB.PictureBox picWpResults 
      Height          =   375
      Left            =   4680
      ScaleHeight     =   315
      ScaleWidth      =   795
      TabIndex        =   21
      Top             =   6360
      Width           =   855
   End
   Begin VB.CommandButton cmdAW6 
      Caption         =   "Add"
      Height          =   375
      Left            =   7800
      TabIndex        =   20
      Top             =   4080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdAW5 
      Caption         =   "Add"
      Height          =   375
      Left            =   7800
      TabIndex        =   19
      Top             =   3600
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdAW4 
      Caption         =   "Add"
      Height          =   375
      Left            =   7800
      TabIndex        =   18
      Top             =   3120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdAW3 
      Caption         =   "Add"
      Height          =   375
      Left            =   7800
      TabIndex        =   17
      Top             =   2640
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdAW2 
      Caption         =   "Add"
      Height          =   375
      Left            =   7800
      TabIndex        =   16
      Top             =   2160
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdAW1 
      Caption         =   "Add"
      Height          =   375
      Left            =   7800
      TabIndex        =   15
      Top             =   1680
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtW6 
      Height          =   375
      Left            =   6720
      TabIndex        =   14
      Text            =   "0"
      Top             =   4080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtW5 
      Height          =   375
      Left            =   6720
      TabIndex        =   13
      Text            =   "0"
      Top             =   3600
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtW4 
      Height          =   375
      Left            =   6720
      TabIndex        =   12
      Text            =   "0"
      Top             =   3120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtW3 
      Height          =   375
      Left            =   6720
      TabIndex        =   11
      Text            =   "0"
      Top             =   2640
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtW2 
      Height          =   375
      Left            =   6720
      TabIndex        =   10
      Text            =   "0"
      Top             =   2160
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtW1 
      Height          =   375
      Left            =   6720
      TabIndex        =   9
      Text            =   "0"
      Top             =   1680
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox picWtResults 
      BackColor       =   &H0080C0FF&
      Height          =   1935
      Left            =   3120
      ScaleHeight     =   1875
      ScaleWidth      =   3075
      TabIndex        =   8
      Top             =   4080
      Width           =   3135
   End
   Begin VB.CommandButton cmdW6 
      Caption         =   "CorePlex® Multiple Vitamin and Mineral Dietary Supplement"
      Height          =   735
      Left            =   120
      TabIndex        =   7
      Top             =   5520
      Width           =   1935
   End
   Begin VB.CommandButton cmdW5 
      Caption         =   "Cherry Spark Nutritional Supplement Drink Mix (Can)"
      Height          =   735
      Left            =   120
      TabIndex        =   6
      Top             =   4680
      Width           =   1935
   End
   Begin VB.CommandButton cmdW4 
      Caption         =   "MNS Yellow Multinutrient Dietary Supplement"
      Height          =   735
      Left            =   120
      TabIndex        =   5
      Top             =   3840
      Width           =   1935
   End
   Begin VB.CommandButton cmdW3 
      Caption         =   "MNS Gold Multinutrient Dietary Supplement "
      Height          =   735
      Left            =   120
      TabIndex        =   4
      Top             =   3000
      Width           =   1935
   End
   Begin VB.CommandButton cmdW2 
      Caption         =   "MNS Platinum Multinutrient Dietary Supplement"
      Height          =   735
      Left            =   120
      TabIndex        =   3
      Top             =   2160
      Width           =   1935
   End
   Begin VB.CommandButton cmdW1 
      Caption         =   "NEW! A Perfect You® Platinum Multinutrient Dietary Supplements"
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   1935
   End
   Begin VB.CommandButton cmdBackWeight 
      Caption         =   "Click to go Back"
      Height          =   975
      Left            =   7200
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000009&
      Caption         =   "Price:"
      Height          =   375
      Left            =   4080
      TabIndex        =   22
      Top             =   6360
      Width           =   495
   End
   Begin VB.Image imgW6 
      Height          =   2475
      Left            =   3360
      Picture         =   "frmWeight.frx":0000
      Top             =   1440
      Visible         =   0   'False
      Width           =   2700
   End
   Begin VB.Image imgW5 
      Height          =   2475
      Left            =   3360
      Picture         =   "frmWeight.frx":0F65
      Top             =   1440
      Visible         =   0   'False
      Width           =   2700
   End
   Begin VB.Image imgW4 
      Height          =   2475
      Left            =   3360
      Picture         =   "frmWeight.frx":2BDE
      Top             =   1440
      Visible         =   0   'False
      Width           =   2700
   End
   Begin VB.Image imgW3 
      Height          =   2475
      Left            =   3360
      Picture         =   "frmWeight.frx":48D3
      Top             =   1440
      Visible         =   0   'False
      Width           =   2700
   End
   Begin VB.Image imgW2 
      Height          =   2475
      Left            =   3360
      Picture         =   "frmWeight.frx":751A
      Top             =   1440
      Visible         =   0   'False
      Width           =   2700
   End
   Begin VB.Image imgW1 
      Height          =   2475
      Left            =   3360
      Picture         =   "frmWeight.frx":8F5C
      Top             =   1440
      Visible         =   0   'False
      Width           =   2700
   End
   Begin VB.Label Label1 
      BackColor       =   &H000000FF&
      Caption         =   "Click buttons to view products, descriptions and to add to your shopping cart. Enter only numbers into Add field."
      Height          =   975
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   2295
   End
   Begin VB.Image ImageB 
      Height          =   975
      Left            =   2520
      Picture         =   "frmWeight.frx":B34F
      Top             =   120
      Width           =   4050
   End
End
Attribute VB_Name = "frmWeight"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : AdvoCare Store (Dombrovski,Adam.vbp)
'Form Name : frmWeight (Weigh Management Page)
'Author: Adam Dombrovski
'Date Written: March 15, 2004
'Purpose: The purpose of this form is to allow the user to learn about the
    'Weight Management products that AdvoCare puts out. They can add
    'products to their shopping cart as well.
    
Private Sub cmdAW1_Click()
picWCart.Cls
If txtW1 >= 1 Then
ctr = ctr + 1
W1 = txtW1 * Wprice(1)
runningTotal = runningTotal + W1
Sctr = Sctr + txtW1
Cart(ctr) = "A Perfect You Platinum"
CartPrice(ctr) = W1
picWCart.Print "Your cart has "; Sctr; "items."
txtW1.Text = "0"
End If
End Sub

Private Sub cmdAW2_Click()
picWCart.Cls
If txtW2 >= 1 Then
ctr = ctr + 1
W2 = txtW2 * Wprice(2)
runningTotal = runningTotal + W2
Sctr = Sctr + txtW2
Cart(ctr) = "MNS Platinum"
CartPrice(ctr) = W2
picWCart.Print "Your cart has "; Sctr; "items."
txtW2.Text = "0"
End If
End Sub

Private Sub cmdAW3_Click()
picWCart.Cls
If txtW3 >= 1 Then
ctr = ctr + 1
W3 = txtW3 * Wprice(3)
runningTotal = runningTotal + W3
Sctr = Sctr + txtW3
Cart(ctr) = "MNS Gold"
CartPrice(ctr) = W3
picWCart.Print "Your cart has "; Sctr; "items."
txtW3.Text = "0"
End If
End Sub

Private Sub cmdAW4_Click()
picWCart.Cls
If txtW4 >= 1 Then
ctr = ctr + 1
W4 = txtW4 * Wprice(4)
runningTotal = runningTotal + W4
Sctr = Sctr + txtW4
Cart(ctr) = "MNS Yellow"
CartPrice(ctr) = W4
picWCart.Print "Your cart has "; Sctr; "items."
txtW4.Text = "0"
End If
End Sub

Private Sub cmdAW5_Click()
picWCart.Cls
If txtW5 >= 1 Then
ctr = ctr + 1
W5 = txtW5 * Wprice(5)
runningTotal = runningTotal + W5
Sctr = Sctr + txtW5
Cart(ctr) = "Cherry Spark"
CartPrice(ctr) = W5
picWCart.Print "Your cart has "; Sctr; "items."
txtW5.Text = "0"
End If
End Sub

Private Sub cmdAW6_Click()
picWCart.Cls
If txtW6 >= 1 Then
ctr = ctr + 1
W6 = txtW6 * Wprice(6)
runningTotal = runningTotal + W6
Sctr = Sctr + txtW6
Cart(ctr) = "CorePlex Multivitamin"
CartPrice(ctr) = W6
picWCart.Print "Your cart has "; Sctr; "items."
txtW6.Text = "0"
End If
End Sub

Private Sub cmdBackWeight_Click()
frmWeight.Hide
frmChoose.Show
End Sub



Private Sub cmdCart2_Click()
frmWeight.Hide
frmCart.Show
End Sub

Private Sub cmdW1_Click()
picWpResults.Cls
picWtResults.Cls
txtW1.Visible = True
cmdAW1.Visible = True
imgW1.Visible = True
txtW2.Visible = False
cmdAW2.Visible = False
imgW2.Visible = False
txtW3.Visible = False
cmdAW3.Visible = False
imgW3.Visible = False
txtW4.Visible = False
cmdAW4.Visible = False
imgW4.Visible = False
txtW5.Visible = False
cmdAW5.Visible = False
imgW5.Visible = False
txtW6.Visible = False
cmdAW6.Visible = False
imgW6.Visible = False
picWpResults.Print FormatCurrency(Wprice(1))
Open PATH & "Texts\APY.txt" For Input As #10
Input #10, APY
picWtResults.Print APY
Close #10
End Sub

Private Sub cmdW2_Click()
picWpResults.Cls
picWtResults.Cls
txtW1.Visible = False
cmdAW1.Visible = False
imgW1.Visible = False
txtW2.Visible = True
cmdAW2.Visible = True
imgW2.Visible = True
txtW3.Visible = False
cmdAW3.Visible = False
imgW3.Visible = False
txtW4.Visible = False
cmdAW4.Visible = False
imgW4.Visible = False
txtW5.Visible = False
cmdAW5.Visible = False
imgW5.Visible = False
txtW6.Visible = False
cmdAW6.Visible = False
imgW6.Visible = False
picWpResults.Print FormatCurrency(Wprice(2))
Open PATH & "Texts\MNSPlat.txt" For Input As #11
Input #11, MNSPlat
picWtResults.Print MNSPlat
Close #11
End Sub

Private Sub cmdW3_Click()
picWpResults.Cls
picWtResults.Cls
txtW1.Visible = False
cmdAW1.Visible = False
imgW1.Visible = False
txtW2.Visible = False
cmdAW2.Visible = False
imgW2.Visible = False
txtW3.Visible = True
cmdAW3.Visible = True
imgW3.Visible = True
txtW4.Visible = False
cmdAW4.Visible = False
imgW4.Visible = False
txtW5.Visible = False
cmdAW5.Visible = False
imgW5.Visible = False
txtW6.Visible = False
cmdAW6.Visible = False
imgW6.Visible = False
picWpResults.Print FormatCurrency(Wprice(3))
Open PATH & "Texts\MNSGold.txt" For Input As #12
Input #12, MNSGold
picWtResults.Print MNSGold
Close #12
End Sub

Private Sub cmdW4_Click()
picWpResults.Cls
picWtResults.Cls
txtW1.Visible = False
cmdAW1.Visible = False
imgW1.Visible = False
txtW2.Visible = False
cmdAW2.Visible = False
imgW2.Visible = False
txtW3.Visible = False
cmdAW3.Visible = False
imgW3.Visible = False
txtW4.Visible = True
cmdAW4.Visible = True
imgW4.Visible = True
txtW5.Visible = False
cmdAW5.Visible = False
imgW5.Visible = False
txtW6.Visible = False
cmdAW6.Visible = False
imgW6.Visible = False
picWpResults.Print FormatCurrency(Wprice(4))
Open PATH & "Texts\MNSYell.txt" For Input As #13
Input #13, MNSYell
picWtResults.Print MNSYell
Close #13
End Sub

Private Sub cmdW5_Click()
picWpResults.Cls
picWtResults.Cls
txtW1.Visible = False
cmdAW1.Visible = False
imgW1.Visible = False
txtW2.Visible = False
cmdAW2.Visible = False
imgW2.Visible = False
txtW3.Visible = False
cmdAW3.Visible = False
imgW3.Visible = False
txtW4.Visible = False
cmdAW4.Visible = False
imgW4.Visible = False
txtW5.Visible = True
cmdAW5.Visible = True
imgW5.Visible = True
txtW6.Visible = False
cmdAW6.Visible = False
imgW6.Visible = False
picWpResults.Print FormatCurrency(Wprice(5))
Open PATH & "Texts\Spark.txt" For Input As #14
Input #14, Spark
picWtResults.Print Spark
Close #14
End Sub

Private Sub cmdW6_Click()
picWpResults.Cls
picWtResults.Cls
txtW1.Visible = False
cmdAW1.Visible = False
imgW1.Visible = False
txtW2.Visible = False
cmdAW2.Visible = False
imgW2.Visible = False
txtW3.Visible = False
cmdAW3.Visible = False
imgW3.Visible = False
txtW4.Visible = False
cmdAW4.Visible = False
imgW4.Visible = False
txtW5.Visible = False
cmdAW5.Visible = False
imgW5.Visible = False
txtW6.Visible = True
cmdAW6.Visible = True
imgW6.Visible = True
picWpResults.Print FormatCurrency(Wprice(6))
Open PATH & "Texts\CorePlex1.txt" For Input As #15
Input #15, CorePlex
picWtResults.Print CorePlex
Close #15
End Sub

Private Sub Form_Load()
picWpResults.Cls
MsgBox "Welcome to the Weight Management page! You will find all the products you need to lose or maintain weight effectively and safely."
Open PATH & "Texts\weight.txt" For Input As #9
For j = 1 To 6
Input #9, WID(j), Wprice(j)
Next j
End Sub


Private Sub txtW1_Click()
txtW1.Text = ""
End Sub
Private Sub txtW2_Click()
txtW2.Text = ""
End Sub
Private Sub txtW3_Click()
txtW3.Text = ""
End Sub
Private Sub txtW4_Click()
txtW4.Text = ""
End Sub
Private Sub txtW5_Click()
txtW5.Text = ""
End Sub
Private Sub txtW6_Click()
txtW6.Text = ""
End Sub




