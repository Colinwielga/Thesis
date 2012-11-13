VERSION 5.00
Begin VB.Form frmFlowers 
   BackColor       =   &H8000000D&
   Caption         =   "FLOWERS"
   ClientHeight    =   7020
   ClientLeft      =   6270
   ClientTop       =   4575
   ClientWidth     =   9330
   LinkTopic       =   "Form1"
   ScaleHeight     =   7020
   ScaleWidth      =   9330
   Begin VB.CommandButton cmdboughtown 
      Caption         =   "If you have bought your wedding flowers somewhere else. Click Here to input the cost to the budget."
      Height          =   1335
      Left            =   7200
      TabIndex        =   14
      Top             =   2880
      Width           =   1575
   End
   Begin VB.CommandButton cmdotheroption 
      Caption         =   "If You Do Not Like The Selection Available and Would Like To See If The Kind Of Flower You Want Is Avaliable Please Click Here"
      Height          =   1935
      Left            =   7200
      TabIndex        =   13
      Top             =   840
      Width           =   1575
   End
   Begin VB.CommandButton cmdUhOh 
      Caption         =   "Click here if you made a selection and would like to change it."
      Height          =   735
      Left            =   3960
      TabIndex        =   12
      Top             =   6120
      Width           =   2775
   End
   Begin VB.PictureBox picDaffodils 
      Height          =   1575
      Left            =   5160
      Picture         =   "Katie Wasness Flowers.frx":0000
      ScaleHeight     =   1515
      ScaleWidth      =   1755
      TabIndex        =   11
      Top             =   1560
      Width           =   1815
   End
   Begin VB.PictureBox picLilies 
      Height          =   1575
      Left            =   3600
      Picture         =   "Katie Wasness Flowers.frx":635C
      ScaleHeight     =   1515
      ScaleWidth      =   1395
      TabIndex        =   10
      Top             =   1560
      Width           =   1455
   End
   Begin VB.PictureBox picTulips 
      Height          =   1575
      Left            =   2040
      Picture         =   "Katie Wasness Flowers.frx":782A
      ScaleHeight     =   1515
      ScaleWidth      =   1395
      TabIndex        =   9
      Top             =   1560
      Width           =   1455
   End
   Begin VB.PictureBox picRoses 
      Height          =   1575
      Left            =   120
      Picture         =   "Katie Wasness Flowers.frx":A902
      ScaleHeight     =   1515
      ScaleWidth      =   1755
      TabIndex        =   8
      Top             =   1560
      Width           =   1815
   End
   Begin VB.CommandButton cmdDaffodils 
      Caption         =   "Daffodils"
      Height          =   975
      Left            =   5160
      TabIndex        =   7
      Top             =   3240
      Width           =   1815
   End
   Begin VB.CommandButton cmdLilies 
      Caption         =   "Lilies "
      Height          =   975
      Left            =   3600
      TabIndex        =   6
      Top             =   3240
      Width           =   1455
   End
   Begin VB.CommandButton cmdTulips 
      Caption         =   "Tulips"
      Height          =   975
      Left            =   2040
      TabIndex        =   5
      Top             =   3240
      Width           =   1455
   End
   Begin VB.PictureBox picFlowers 
      BackColor       =   &H8000000D&
      ForeColor       =   &H8000000E&
      Height          =   1695
      Left            =   600
      ScaleHeight     =   1635
      ScaleWidth      =   7755
      TabIndex        =   3
      Top             =   4320
      Width           =   7815
   End
   Begin VB.CommandButton cmdRoses 
      Caption         =   "Roses"
      Height          =   975
      Left            =   120
      TabIndex        =   2
      Top             =   3240
      Width           =   1815
   End
   Begin VB.CommandButton cmdBackTo 
      Caption         =   "Click to Go Back to Main Menu"
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   6240
      Width           =   2655
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   495
      Left            =   8040
      TabIndex        =   0
      Top             =   6360
      Width           =   1095
   End
   Begin VB.Label labtitle 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      Caption         =   "CHOOSE THE KIND OF FLOWERS YOU WOULD LIKE."
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   600
      TabIndex        =   4
      Top             =   120
      Width           =   6855
   End
End
Attribute VB_Name = "frmFlowers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'pjtWeddingBudget(Katie Wasness Wedding)
'frmFlowers(Katie Wasness Flowers)
'Author = Katie Wasness
'Date Written: March 10-15
'Purpose--The purpose of this form is choose the Flowers and put the price of that into the Running Total.
Dim FlowerCTR As Integer





Private Sub cmdboughtown_Click()
'this button is in case an individual has choosen a ceremony site somewhere else
CostOfFlowers = InputBox("What is the cost of your Flowers?", "Other Flowers")
picFlowers.Print "The Cost Of Your Flowers Is "; FormatCurrency(CostOfFlowers); "."
RunningTotal = RunningTotal + CostOfFlowers
picFlowers.Print "The total amount of money you have budgeted so far is "; FormatCurrency(RunningTotal); "."
cmdTulips.Enabled = False
cmdRoses.Enabled = False
cmdLilies.Enabled = False
cmdDaffodils.Enabled = False
cmdboughtown.Enabled = False
cmdotheroption.Enabled = False
cmdUhOh.Enabled = True
End Sub

Private Sub cmdBackTo_Click()
'this button is used to go back to the main menu
frmWedding.Show
frmFlowers.Hide
End Sub

Private Sub cmdDaffodils_Click()
'this button is used if the individual wants daffodils at their wedding
WhichFlower = "Daffodils"
TotalFlowers = InputBox("How many Daffodils are you planning to use in your wedding?", "Daffodil for $3.99 per daffodil")
CostOfFlowers = TotalFlowers * 3.99
RunningTotal = RunningTotal + CostOfFlowers
picFlowers.Print "You have decided to have "; TotalFlowers; "daffodils in your wedding.  The total cost for those flowers is "; FormatCurrency(CostOfFlowers); "."
picFlowers.Print "The total amount of money you have budgeted so far is "; FormatCurrency(RunningTotal); "."
cmdTulips.Enabled = False
cmdRoses.Enabled = False
cmdLilies.Enabled = False
cmdDaffodils.Enabled = False
cmdboughtown.Enabled = False
cmdotheroption.Enabled = False
cmdUhOh.Enabled = True
End Sub

Private Sub cmdLilies_Click()
'this button is used if the individual wants lilies at their wedding
WhichFlower = "Lilies"
TotalFlowers = InputBox("How many Lilies are you planning to use in your wedding?", "Lilies for $4.25 per Lilie")
CostOfFlowers = TotalFlowers * 4.25
RunningTotal = RunningTotal + CostOfFlowers
picFlowers.Print "You have decided to have "; TotalFlowers; "lilies in your wedding.  The total cost for those flowers is "; FormatCurrency(CostOfFlowers); "."
picFlowers.Print "The total amount of money you have budgeted so far is "; FormatCurrency(RunningTotal); "."
cmdTulips.Enabled = False
cmdRoses.Enabled = False
cmdLilies.Enabled = False
cmdDaffodils.Enabled = False
cmdboughtown.Enabled = False
cmdotheroption.Enabled = False
cmdUhOh.Enabled = True
End Sub

Private Sub cmdotheroption_Click()
'this button is used if the individual wants to use flowers on a different list but not pictured
Dim Found As Boolean
Dim CTR As Integer
PATH = "N:\CS130\handin\Wasness, Katie\"

Open PATH + "flowers.txt" For Input As #1
'Open "M:\CS130\Wasness, Katie\flowers.txt" For Input As #1
For CTR = 1 To 6
    Input #1, Flowers(CTR), FlowerCost(CTR)
Next CTR
Close
WhichFlower = InputBox("The Other Flowers we have are Orchids, Calla Lilies, Baby's Breath, Iris, Lilac, and Stephanotis. Please Pick One of These.", "Other Kind Of Flower")
CTR = 0
Found = False
Do While Found = False And CTR < 6
    CTR = CTR + 1
    If WhichFlower = Flowers(CTR) Then
        Found = True
    End If
Loop
If Found = True Then
    picFlowers.Print WhichFlower; " will be beautiful at your wedding."
    TotalFlowers = InputBox("How many flowers will you be having at your wedding?", "Other Flower Options")
    CostOfFlowers = TotalFlowers * FlowerCost(CTR)
    RunningTotal = RunningTotal + CostOfFlowers
    picFlowers.Print "You have decided to have "; TotalFlowers; WhichFlower; "s in your wedding.  The total cost for those flowers is "; FormatCurrency(CostOfFlowers); "."
    picFlowers.Print "The total amount of money you have spent on your wedding so far is "; FormatCurrency(RunningTotal); "."
    cmdTulips.Enabled = False
    cmdRoses.Enabled = False
    cmdLilies.Enabled = False
    cmdDaffodils.Enabled = False
    cmdboughtown.Enabled = False
    cmdotheroption.Enabled = False
    cmdUhOh.Enabled = True
Else
    picFlowers.Print WhichFlower; "is not a flower in this list. You will have to choose another flower we do not have that flower."
End If
End Sub

Private Sub cmdQuit_Click()
'this button is to end the program
End
End Sub

Private Sub cmdRoses_Click()
'this button is used if the individual wants roses at their wedding
WhichFlower = "Roses"
TotalFlowers = InputBox("How many Roses are you planning to use in your wedding?", "Roses for $6.50 per rose")
CostOfFlowers = TotalFlowers * 6.5
RunningTotal = RunningTotal + CostOfFlowers
picFlowers.Print "You have decided to have "; TotalFlowers; "roses in your wedding.  The total cost for those flowers is "; FormatCurrency(CostOfFlowers); "."
picFlowers.Print "The total amount of money you have budgeted so far is "; FormatCurrency(RunningTotal); "."
cmdTulips.Enabled = False
cmdRoses.Enabled = False
cmdLilies.Enabled = False
cmdDaffodils.Enabled = False
cmdboughtown.Enabled = False
cmdotheroption.Enabled = False
cmdUhOh.Enabled = True

End Sub

Private Sub cmdTulips_Click()
'this button is used if the individual wants tulips at their wedding
WhichFlower = "Tulips"
TotalFlowers = InputBox("How many Tulips are you planning to use in your wedding?", "Tulips for $2.00 per tulip")
CostOfFlowers = TotalFlowers * 2#
RunningTotal = RunningTotal + CostOfFlowers
picFlowers.Print "You have decided to have "; TotalFlowers; "tulips in your wedding.  The total cost for those flowers is "; FormatCurrency(CostOfFlowers); "."
picFlowers.Print "The total amount of money you have budgeted so far is "; FormatCurrency(RunningTotal); "."
cmdTulips.Enabled = False
cmdRoses.Enabled = False
cmdLilies.Enabled = False
cmdDaffodils.Enabled = False
cmdboughtown.Enabled = False
cmdotheroption.Enabled = False
cmdUhOh.Enabled = True
End Sub

Private Sub cmdUhOh_Click()
'this button is used if the user has made a selection and wants to change it. it clears the picture box and minuses the cost of the selection from the total cost.
picFlowers.Cls
RunningTotal = RunningTotal - CostOfFlowers
cmdTulips.Enabled = True
cmdRoses.Enabled = True
cmdLilies.Enabled = True
cmdDaffodils.Enabled = True
cmdboughtown.Enabled = True
cmdotheroption.Enabled = True
cmdUhOh.Enabled = False

End Sub

