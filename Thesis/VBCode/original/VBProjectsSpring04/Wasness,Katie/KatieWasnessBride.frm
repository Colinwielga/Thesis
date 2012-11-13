VERSION 5.00
Begin VB.Form frmBride 
   BackColor       =   &H8000000D&
   Caption         =   "BRIDAL ATTIRE"
   ClientHeight    =   8700
   ClientLeft      =   4020
   ClientTop       =   2700
   ClientWidth     =   10935
   LinkTopic       =   "Form1"
   ScaleHeight     =   8700
   ScaleWidth      =   10935
   Begin VB.CommandButton cmdUhOh 
      Caption         =   "Click here if you made a selection and have decided that you do not want that ensamble."
      Enabled         =   0   'False
      Height          =   1455
      Left            =   9240
      TabIndex        =   11
      Top             =   4080
      Width           =   1455
   End
   Begin VB.PictureBox picDressCost 
      BackColor       =   &H8000000D&
      ForeColor       =   &H8000000E&
      Height          =   1095
      Left            =   240
      ScaleHeight     =   1035
      ScaleWidth      =   10035
      TabIndex        =   9
      Top             =   6120
      Width           =   10095
   End
   Begin VB.CommandButton cmdDress3 
      Caption         =   "Please click here to select this dress ensamble. The total cost of this dress ensamble is $399.00"
      Height          =   1095
      Left            =   6960
      TabIndex        =   8
      Top             =   4920
      Width           =   1815
   End
   Begin VB.CommandButton cmdDress2 
      Caption         =   "Please click here to select this dress ensamble. The total cost of this dress ensamble is $799.00"
      Height          =   1095
      Left            =   3960
      TabIndex        =   7
      Top             =   4920
      Width           =   1815
   End
   Begin VB.CommandButton cmdDress1 
      Caption         =   "Please click here to select this dress ensamble. The total cost of this dress ensamble is $699.00."
      Height          =   1095
      Left            =   480
      TabIndex        =   6
      Top             =   4920
      Width           =   1935
   End
   Begin VB.CommandButton cmdboughtown 
      Caption         =   "I have bought the dress somewhere else, but I would like to input the cost of the Dress into the Budget."
      Height          =   1575
      Left            =   9120
      TabIndex        =   5
      Top             =   1920
      Width           =   1695
   End
   Begin VB.PictureBox picDress3 
      Height          =   3975
      Left            =   6720
      Picture         =   "Katie Wasness Bride.frx":0000
      ScaleHeight     =   3915
      ScaleWidth      =   2235
      TabIndex        =   4
      Top             =   720
      Width           =   2295
   End
   Begin VB.PictureBox picDress2 
      Height          =   3855
      Left            =   3600
      Picture         =   "Katie Wasness Bride.frx":5240
      ScaleHeight     =   3795
      ScaleWidth      =   2835
      TabIndex        =   3
      Top             =   840
      Width           =   2895
   End
   Begin VB.PictureBox picDress1 
      Height          =   3975
      Left            =   240
      Picture         =   "Katie Wasness Bride.frx":8909
      ScaleHeight     =   3915
      ScaleWidth      =   2835
      TabIndex        =   2
      Top             =   720
      Width           =   2895
   End
   Begin VB.CommandButton cmdBackTo 
      Caption         =   "Go Back to Attire Form"
      Height          =   615
      Left            =   240
      TabIndex        =   1
      Top             =   7320
      Width           =   1935
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   735
      Left            =   9000
      TabIndex        =   0
      Top             =   7440
      Width           =   1335
   End
   Begin VB.Label labtitle 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      Caption         =   "WHAT DRESS ENSAMBLE WOULD YOU LIKE THE BRIDE TO  HAVE?"
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   1320
      TabIndex        =   10
      Top             =   120
      Width           =   8175
   End
End
Attribute VB_Name = "frmBride"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'pjtWeddingBudget(Katie Wasness Wedding)
'frmBride(Katie Wasness Bride)
'Author = Katie Wasness
'Date Written: March 10-15
'Purpose--The purpose of this form is choose the bride's attire and to put the price of that into the Running Total.
Private Sub cmdBackTo_Click()
'This button brings you back to the Attire Form
frmAttire.Show
frmBride.Hide
End Sub

Private Sub cmdboughtown_Click()
'this button is in case an individual has boughten the bridal dress somewhere else
CostofDress = InputBox("Enter the amount of the dress that you bought elsewhere, using a decimal point to show dollars and cents.", "Amount of Dress")
picDressCost.Print "I am sure the dress that you have picked out will look beautiful on "; Bride; " The total cost of your ensamble will be "; FormatCurrency(CostofDress); "."
RunningTotal = RunningTotal + CostofDress
picDressCost.Print "Your total running total is now "; FormatCurrency(RunningTotal); "."
cmdDress1.Enabled = False
cmdDress2.Enabled = False
cmdDress3.Enabled = False
cmdboughtown.Enabled = False
cmdUhOh.Enabled = True

End Sub

Private Sub cmdDress1_Click()
'this button is used if the user wants to pic dress number 1
CostofDress = 699#
picDressCost.Print Bride; " will look beautiful in this dress. The total cost of your ensamble will be "; FormatCurrency(CostofDress); "."
RunningTotal = RunningTotal + CostofDress
picDressCost.Print "Your total running total is now "; FormatCurrency(RunningTotal); "."
cmdDress1.Enabled = False
cmdDress2.Enabled = False
cmdDress3.Enabled = False
cmdboughtown.Enabled = False
cmdUhOh.Enabled = True

End Sub

Private Sub cmdDress2_Click()
'this button is used if the user wants to pic dress number 2
CostofDress = 799#
picDressCost.Print Bride; " will look beautiful in this dress. The total cost of your ensamble will be "; FormatCurrency(CostofDress); "."
RunningTotal = RunningTotal + CostofDress
picDressCost.Print "Your total running total is now "; FormatCurrency(RunningTotal); "."
cmdDress1.Enabled = False
cmdDress2.Enabled = False
cmdDress3.Enabled = False
cmdboughtown.Enabled = False
cmdUhOh.Enabled = True


End Sub

Private Sub cmdDress3_Click()
'this button is used if the user wants to pic dress number 3
CostofDress = 399#
picDressCost.Print Bride; " will look beautiful in this dress. The total cost of your ensamble will be "; FormatCurrency(CostofDress); "."
RunningTotal = RunningTotal + CostofDress
picDressCost.Print "Your total running total is now "; FormatCurrency(RunningTotal); "."
cmdDress1.Enabled = False
cmdDress2.Enabled = False
cmdDress3.Enabled = False
cmdboughtown.Enabled = False
cmdUhOh.Enabled = True

End Sub

Private Sub cmdQuit_Click()
'this button is used to quit the program
End
End Sub

Private Sub cmdUhOh_Click()
'this button is used if the user made a selection and wants to change it. it clears the picture box and minuses the cost of the first selection from the total cost.
picDressCost.Cls
RunningTotal = RunningTotal - CostofDress
cmdDress1.Enabled = True
cmdDress2.Enabled = True
cmdDress3.Enabled = True
cmdboughtown.Enabled = True
End Sub

