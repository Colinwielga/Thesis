VERSION 5.00
Begin VB.Form frmSports 
   BackColor       =   &H80000009&
   Caption         =   "Athletic Performance Page"
   ClientHeight    =   7710
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9060
   LinkTopic       =   "Form1"
   ScaleHeight     =   7710
   ScaleWidth      =   9060
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCart1 
      Caption         =   "Click to Checkout"
      Height          =   975
      Left            =   6240
      TabIndex        =   27
      Top             =   6480
      Width           =   2535
   End
   Begin VB.PictureBox picSCart 
      Height          =   495
      Left            =   6600
      ScaleHeight     =   435
      ScaleWidth      =   2235
      TabIndex        =   26
      Top             =   5760
      Width           =   2295
   End
   Begin VB.PictureBox picStResults 
      BackColor       =   &H008080FF&
      Height          =   1935
      Left            =   3240
      ScaleHeight     =   1875
      ScaleWidth      =   3075
      TabIndex        =   25
      Top             =   4080
      Width           =   3135
   End
   Begin VB.PictureBox picSpResults 
      Height          =   375
      Left            =   4680
      ScaleHeight     =   315
      ScaleWidth      =   795
      TabIndex        =   23
      Top             =   6600
      Width           =   855
   End
   Begin VB.CommandButton cmdAProtein 
      Caption         =   "Add"
      Height          =   375
      Left            =   8160
      TabIndex        =   22
      Top             =   5160
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtAProtein 
      Height          =   375
      Left            =   7200
      TabIndex        =   21
      Text            =   "0"
      Top             =   5160
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdAPerGold 
      Caption         =   "Add"
      Height          =   375
      Left            =   8160
      TabIndex        =   20
      Top             =   4560
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtAPerGold 
      Height          =   375
      Left            =   7200
      TabIndex        =   19
      Text            =   "0"
      Top             =   4560
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdA5 
      Caption         =   "Add"
      Height          =   375
      Left            =   8160
      TabIndex        =   18
      Top             =   3960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtA5 
      Height          =   375
      Left            =   7200
      TabIndex        =   17
      Text            =   "0"
      Top             =   3960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdA4 
      Caption         =   "Add"
      Height          =   375
      Left            =   8160
      TabIndex        =   16
      Top             =   3360
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtA4 
      Height          =   375
      Left            =   7200
      TabIndex        =   15
      Text            =   "0"
      Top             =   3360
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdA3 
      Caption         =   "Add"
      Height          =   375
      Left            =   8160
      TabIndex        =   14
      Top             =   2760
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtA3 
      Height          =   375
      Left            =   7200
      TabIndex        =   13
      Text            =   "0"
      Top             =   2760
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdA2 
      Caption         =   "Add"
      Height          =   375
      Left            =   8160
      TabIndex        =   12
      Top             =   2160
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtA2 
      Height          =   375
      Left            =   7200
      TabIndex        =   11
      Text            =   "0"
      Top             =   2160
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdA1 
      Caption         =   "Add"
      Height          =   375
      Left            =   8160
      TabIndex        =   10
      Top             =   1560
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtA1 
      Height          =   375
      Left            =   7200
      TabIndex        =   9
      Text            =   "0"
      Top             =   1560
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdPerGold 
      Caption         =   "Performance Gold  Herbal Dietary Supplement "
      Height          =   735
      Left            =   240
      TabIndex        =   8
      Top             =   5520
      Width           =   1935
   End
   Begin VB.CommandButton cmdProtein 
      Caption         =   "BodyLean Vanilla Shake (Can)"
      Height          =   735
      Left            =   240
      TabIndex        =   7
      Top             =   6360
      Width           =   1935
   End
   Begin VB.CommandButton cmd5 
      Caption         =   "POS 5 Dietary Supplement with HMB and Suma"
      Height          =   735
      Left            =   240
      TabIndex        =   6
      Top             =   4680
      Width           =   1935
   End
   Begin VB.CommandButton cmd4 
      Caption         =   "POS 4 Arginine Dietary Supplement Drink Mix "
      Height          =   735
      Left            =   240
      TabIndex        =   5
      Top             =   3840
      Width           =   1935
   End
   Begin VB.CommandButton cmd3 
      Caption         =   "NEW! POS 3 Fruit Punch Flavor Sports Drink Mix (Can)"
      Height          =   735
      Left            =   240
      TabIndex        =   4
      Top             =   3000
      Width           =   1935
   End
   Begin VB.CommandButton cmd2 
      Caption         =   "POS 2 Chocolate Recovery Drink Mix (Pouches)"
      Height          =   735
      Left            =   240
      TabIndex        =   3
      Top             =   2160
      Width           =   1935
   End
   Begin VB.CommandButton cmd1 
      Caption         =   "POS 1 Amino Acid and Herbal Dietary Supplement"
      Height          =   735
      Left            =   240
      TabIndex        =   1
      Top             =   1320
      Width           =   1935
   End
   Begin VB.CommandButton cmdBackSport 
      Caption         =   "Click to go Back"
      Height          =   975
      Left            =   7560
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000009&
      Caption         =   "Price:"
      Height          =   375
      Left            =   4080
      TabIndex        =   24
      Top             =   6600
      Width           =   495
   End
   Begin VB.Image imgProtein 
      Height          =   2475
      Left            =   3480
      Picture         =   "frmSports.frx":0000
      Top             =   1440
      Visible         =   0   'False
      Width           =   2700
   End
   Begin VB.Image imgPerGold 
      Height          =   2475
      Left            =   3480
      Picture         =   "frmSports.frx":20D7
      Top             =   1440
      Visible         =   0   'False
      Width           =   2700
   End
   Begin VB.Image img5 
      Height          =   2475
      Left            =   3480
      Picture         =   "frmSports.frx":3366
      Top             =   1440
      Visible         =   0   'False
      Width           =   2700
   End
   Begin VB.Image img4 
      Height          =   2475
      Left            =   3480
      Picture         =   "frmSports.frx":4572
      Top             =   1440
      Visible         =   0   'False
      Width           =   2700
   End
   Begin VB.Image img3 
      Height          =   2475
      Left            =   3480
      Picture         =   "frmSports.frx":579E
      Top             =   1440
      Visible         =   0   'False
      Width           =   2700
   End
   Begin VB.Image img2 
      Height          =   2475
      Left            =   3480
      Picture         =   "frmSports.frx":7014
      Top             =   1440
      Visible         =   0   'False
      Width           =   2700
   End
   Begin VB.Image img1 
      Height          =   2475
      Left            =   3480
      Picture         =   "frmSports.frx":9703
      Top             =   1440
      Visible         =   0   'False
      Width           =   2700
   End
   Begin VB.Label Label1 
      BackColor       =   &H000000FF&
      Caption         =   "Click buttons to view products, descriptions and to add to your shopping cart.  Enter only numbers to Add field."
      Height          =   975
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2055
   End
   Begin VB.Image ImageA 
      Height          =   975
      Left            =   2400
      Picture         =   "frmSports.frx":AD6D
      Top             =   120
      Width           =   4050
   End
End
Attribute VB_Name = "frmSports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : AdvoCare Store (Dombrovski,Adam.vbp)
'Form Name : frmSports (Athletic Performance Page)
'Author: Adam Dombrovski
'Date Written: March 15, 2004
'Purpose: The purpose of this form is to allow the user to learn about the
    'Athletic Performance products that AdvoCare puts out. They can add
    'products to their shopping cart as well.
Private Sub cmd1_Click()
picSpResults.Cls
txtA1.Visible = True
cmdA1.Visible = True
img1.Visible = True
txtA2.Visible = False
cmdA2.Visible = False
img2.Visible = False
txtA3.Visible = False
cmdA3.Visible = False
img3.Visible = False
txtA4.Visible = False
cmdA4.Visible = False
img4.Visible = False
txtA5.Visible = False
cmdA5.Visible = False
img5.Visible = False
txtAPerGold.Visible = False
cmdAPerGold.Visible = False
imgPerGold.Visible = False
txtAProtein.Visible = False
cmdAProtein.Visible = False
imgProtein.Visible = False
picStResults.Cls
picSpResults.Print FormatCurrency(Sprice(1))
Open PATH & "Texts\POS1.txt" For Input As #2
Input #2, POS1
picStResults.Print POS1
Close #2
End Sub

Private Sub cmd2_Click()
picSpResults.Cls
txtA1.Visible = False
cmdA1.Visible = False
img1.Visible = False
txtA2.Visible = True
cmdA2.Visible = True
img2.Visible = True
txtA3.Visible = False
cmdA3.Visible = False
img3.Visible = False
txtA4.Visible = False
cmdA4.Visible = False
img4.Visible = False
txtA5.Visible = False
cmdA5.Visible = False
img5.Visible = False
txtAPerGold.Visible = False
cmdAPerGold.Visible = False
imgPerGold.Visible = False
txtAProtein.Visible = False
cmdAProtein.Visible = False
imgProtein.Visible = False
picSpResults.Print FormatCurrency(Sprice(2))
picStResults.Cls
Open PATH & "Texts\POS2.txt" For Input As #3
Input #3, POS2
picStResults.Print POS2
Close #3
End Sub

Private Sub cmd3_Click()
picSpResults.Cls
txtA1.Visible = False
cmdA1.Visible = False
img1.Visible = False
txtA2.Visible = False
cmdA2.Visible = False
img2.Visible = False
txtA3.Visible = True
cmdA3.Visible = True
img3.Visible = True
txtA4.Visible = False
cmdA4.Visible = False
img4.Visible = False
txtA5.Visible = False
cmdA5.Visible = False
img5.Visible = False
txtAPerGold.Visible = False
cmdAPerGold.Visible = False
imgPerGold.Visible = False
txtAProtein.Visible = False
cmdAProtein.Visible = False
imgProtein.Visible = False
picSpResults.Print FormatCurrency(Sprice(3))
picStResults.Cls
Open PATH & "Texts\POS3.txt" For Input As #4
Input #4, POS3
picStResults.Print POS3
Close #4
End Sub

Private Sub cmd4_Click()
picSpResults.Cls
txtA1.Visible = False
cmdA1.Visible = False
img1.Visible = False
txtA2.Visible = False
cmdA2.Visible = False
img2.Visible = False
txtA3.Visible = False
cmdA3.Visible = False
img3.Visible = False
txtA4.Visible = True
cmdA4.Visible = True
img4.Visible = True
txtA5.Visible = False
cmdA5.Visible = False
img5.Visible = False
txtAPerGold.Visible = False
cmdAPerGold.Visible = False
imgPerGold.Visible = False
txtAProtein.Visible = False
cmdAProtein.Visible = False
imgProtein.Visible = False
picSpResults.Print FormatCurrency(Sprice(4))
picStResults.Cls
Open PATH & "Texts\POS4.txt" For Input As #5
Input #5, POS4
picStResults.Print POS4
Close #5
End Sub

Private Sub cmd5_Click()
picSpResults.Cls
txtA1.Visible = False
cmdA1.Visible = False
img1.Visible = False
txtA2.Visible = False
cmdA2.Visible = False
img2.Visible = False
txtA3.Visible = False
cmdA3.Visible = False
img3.Visible = False
txtA4.Visible = False
cmdA4.Visible = False
img4.Visible = False
txtA5.Visible = True
cmdA5.Visible = True
img5.Visible = True
txtAPerGold.Visible = False
cmdAPerGold.Visible = False
imgPerGold.Visible = False
txtAProtein.Visible = False
cmdAProtein.Visible = False
imgProtein.Visible = False
picSpResults.Print FormatCurrency(Sprice(5))
picStResults.Cls
Open PATH & "Texts\POS5.txt" For Input As #6
Input #6, POS5
picStResults.Print POS5
Close #6
End Sub

Private Sub cmdA1_Click()
picSCart.Cls
If txtA1 >= 1 Then
ctr = ctr + 1
A1 = txtA1 * Sprice(1)
runningTotal = runningTotal + A1
Sctr = Sctr + txtA1
Cart(ctr) = "POS 1"
CartPrice(ctr) = A1
picSCart.Print "Your cart has "; Sctr; "items."
txtA1.Text = "0"
End If
End Sub

Private Sub cmdA2_Click()
picSCart.Cls
If txtA2 >= 1 Then
ctr = ctr + 1
A2 = txtA2 * Sprice(2)
runningTotal = runningTotal + A2
Sctr = Sctr + txtA2
Cart(ctr) = "POS 2"
CartPrice(ctr) = A2
picSCart.Print "Your cart has "; Sctr; "items."
txtA2.Text = "0"
End If
End Sub

Private Sub cmdA3_Click()
picSCart.Cls
If txtA3 >= 1 Then
ctr = ctr + 1
Sctr = Sctr + txtA3
A3 = txtA3 * Sprice(3)
runningTotal = runningTotal + A3
Cart(ctr) = "POS 3"
CartPrice(ctr) = A3
picSCart.Print "Your cart has "; Sctr; "items."
txtA3.Text = "0"
End If
End Sub

Private Sub cmdA4_Click()

picSCart.Cls
If txtA4 >= 1 Then
ctr = ctr + 1
Sctr = Sctr + txtA4
A4 = txtA4 * Sprice(4)
runningTotal = runningTotal + A4
Cart(ctr) = "POS 4"
CartPrice(ctr) = A4
picSCart.Print "Your cart has "; Sctr; "items."
txtA4.Text = "0"
End If
End Sub

Private Sub cmdA5_Click()
picSCart.Cls
If txtA5 >= 1 Then
ctr = ctr + 1
Sctr = Sctr + txtA5
A5 = txtA5 * Sprice(5)
runningTotal = runningTotal + A5
Cart(ctr) = "POS 5"
CartPrice(ctr) = A5
picSCart.Print "Your cart has "; Sctr; "items."
txtA5.Text = "0"
End If
End Sub

Private Sub cmdAPerGold_Click()

picSCart.Cls
If txtAPerGold >= 1 Then
ctr = ctr + 1
Sctr = Sctr + txtAPerGold
A6 = txtAPerGold * Sprice(6)
runningTotal = runningTotal + A6
Cart(ctr) = "Performance Gold"
CartPrice(ctr) = A6
picSCart.Print "Your cart has "; Sctr; "items."
txtAPerGold.Text = "0"
End If
End Sub

Private Sub cmdAProtein_Click()
picSCart.Cls
If txtAProtein >= 1 Then
ctr = ctr + 1
Sctr = Sctr + txtAProtein
A7 = txtAProtein * Sprice(7)
runningTotal = runningTotal + A7
Cart(ctr) = "BodyLean Protein"
CartPrice(ctr) = A7
picSCart.Print "Your cart has "; Sctr; "items."
txtAProtein.Text = "0"
End If
End Sub

Private Sub cmdBackSport_Click()
frmSports.Hide
frmChoose.Show
End Sub


Private Sub cmdCart1_Click()
frmSports.Hide
frmCart.Show
End Sub

Private Sub cmdPerGold_Click()
picSpResults.Cls
txtA1.Visible = False
cmdA1.Visible = False
img1.Visible = False
txtA2.Visible = False
cmdA2.Visible = False
img2.Visible = False
txtA3.Visible = False
cmdA3.Visible = False
img3.Visible = False
txtA4.Visible = False
cmdA4.Visible = False
img4.Visible = False
txtA5.Visible = False
cmdA5.Visible = False
img5.Visible = False
txtAPerGold.Visible = True
cmdAPerGold.Visible = True
imgPerGold.Visible = True
txtAProtein.Visible = False
cmdAProtein.Visible = False
imgProtein.Visible = False
picSpResults.Print FormatCurrency(Sprice(6))
picStResults.Cls
Open PATH & "Texts\PerGold.txt" For Input As #7
Input #7, PerGold
picStResults.Print PerGold
Close #7
End Sub

Private Sub cmdProtein_Click()
picSpResults.Cls
txtA1.Visible = False
cmdA1.Visible = False
img1.Visible = False
txtA2.Visible = False
cmdA2.Visible = False
img2.Visible = False
txtA3.Visible = False
cmdA3.Visible = False
img3.Visible = False
txtA4.Visible = False
cmdA4.Visible = False
img4.Visible = False
txtA5.Visible = False
cmdA5.Visible = False
img5.Visible = False
txtAPerGold.Visible = False
cmdAPerGold.Visible = False
imgPerGold.Visible = False
txtAProtein.Visible = True
cmdAProtein.Visible = True
imgProtein.Visible = True
picSpResults.Print FormatCurrency(Sprice(7))
picStResults.Cls
Open PATH & "Texts\Protein.txt" For Input As #8
Input #8, Protein
picStResults.Print Protein
Close #8
End Sub



Private Sub Form_Load()
picSpResults.Cls
MsgBox "Welcome to the Athletic Performance page! You will find all the products you need to add strength and lean body mass quickly and safely."
PATH = "N:\CS130\handin\Dombrovski, Adam\"
Open PATH & "Texts\sports.txt" For Input As #1
For j = 1 To 7
Input #1, SID(j), Sprice(j)
Next j

End Sub


Private Sub txtA1_Click()
txtA1.Text = ""
End Sub

Private Sub txtA2_Click()
txtA2.Text = ""
End Sub

Private Sub txtA3_Click()
txtA3.Text = ""
End Sub
Private Sub txtA4_Click()
txtA4.Text = ""
End Sub
Private Sub txtA5_Click()
txtA5.Text = ""
End Sub
Private Sub txtAPerGold_Click()
txtAPerGold.Text = ""
End Sub
Private Sub txtAProtein_Click()
txtAProtein.Text = ""
End Sub
