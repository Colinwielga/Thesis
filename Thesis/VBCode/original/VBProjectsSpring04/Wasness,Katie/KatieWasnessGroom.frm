VERSION 5.00
Begin VB.Form frmGroom 
   BackColor       =   &H8000000D&
   Caption         =   "GROOM ATTIRE"
   ClientHeight    =   8460
   ClientLeft      =   4950
   ClientTop       =   3075
   ClientWidth     =   10620
   LinkTopic       =   "Form1"
   ScaleHeight     =   8460
   ScaleWidth      =   10620
   Begin VB.CommandButton CmdTux3 
      Caption         =   "Please click here to select this Tux ensamble. The total cost of this Tux ensamble is $109.00."
      Height          =   855
      Left            =   6600
      TabIndex        =   10
      Top             =   4920
      Width           =   2535
   End
   Begin VB.CommandButton cmdTux2 
      Caption         =   "Please click here to select this Tux ensamble. The total cost of this Tux ensamble is $99.00."
      Height          =   855
      Left            =   3360
      TabIndex        =   9
      Top             =   4920
      Width           =   3015
   End
   Begin VB.CommandButton cmdTux1 
      Caption         =   "Please click here to select this Tux ensamble. The total cost of this Tux ensamble is $89.00."
      Height          =   855
      Left            =   240
      TabIndex        =   8
      Top             =   4920
      Width           =   2895
   End
   Begin VB.CommandButton cmdUhOh 
      Caption         =   "Click here if you made a selection and have decided that you do not want that ensamble."
      Height          =   2415
      Left            =   9360
      TabIndex        =   7
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CommandButton cmdboughtown 
      Caption         =   "I have bought the Tux somewhere else, but I would like to input the cost of the Tux into the Budget."
      Height          =   2895
      Left            =   9360
      TabIndex        =   6
      Top             =   120
      Width           =   1095
   End
   Begin VB.PictureBox Picture4 
      Height          =   2895
      Left            =   6960
      Picture         =   "Katie Wasness Groom.frx":0000
      ScaleHeight     =   2835
      ScaleWidth      =   1515
      TabIndex        =   5
      Top             =   1320
      Width           =   1575
   End
   Begin VB.PictureBox Picture3 
      Height          =   2895
      Left            =   4080
      Picture         =   "Katie Wasness Groom.frx":2D8F
      ScaleHeight     =   2835
      ScaleWidth      =   1515
      TabIndex        =   4
      Top             =   1320
      Width           =   1575
   End
   Begin VB.PictureBox Picture2 
      Height          =   2895
      Left            =   960
      Picture         =   "Katie Wasness Groom.frx":6073
      ScaleHeight     =   2835
      ScaleWidth      =   1515
      TabIndex        =   3
      Top             =   1320
      Width           =   1575
   End
   Begin VB.PictureBox picTuxCost 
      BackColor       =   &H8000000D&
      ForeColor       =   &H8000000E&
      Height          =   1215
      Left            =   120
      ScaleHeight     =   1155
      ScaleWidth      =   10155
      TabIndex        =   2
      Top             =   6120
      Width           =   10215
   End
   Begin VB.CommandButton cmdBackTo 
      Caption         =   "Go Back to Attire Form"
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   7680
      Width           =   1935
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   735
      Left            =   9120
      TabIndex        =   0
      Top             =   7560
      Width           =   1335
   End
   Begin VB.Label labTitle 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      Caption         =   "WHAT TUX WILL YOUR GROOM BE WEARING?"
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   240
      TabIndex        =   11
      Top             =   120
      Width           =   8895
   End
End
Attribute VB_Name = "frmGroom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'pjtWeddingBudget(Katie Wasness Wedding)
'frmGroom(Katie Wasness Groom)
'Author = Katie Wasness
'Date Written: March 10-15
'Purpose--The purpose of this form is choose groom's attire and put the price of that into the Running Total.
Private Sub cmdBackTo_Click()
'this button is used to go back to the attire menu
frmAttire.Show
frmGroom.Hide
End Sub

Private Sub cmdboughtown_Click()
'this button is in case an individual has boughten the tux somewhere else
CostofTux = InputBox("Enter the amount of the Tux that you bought elsewhere, using a decimal point to show dollars and cents.", "Amount of Tux")
picTuxCost.Print Groom; " will look handsome in this tux. The total cost of your ensamble will be "; FormatCurrency(CostofTux); "."
RunningTotal = RunningTotal + CostofTux
picTuxCost.Print "Your total running total is now "; FormatCurrency(RunningTotal); "."
cmdTux1.Enabled = False
cmdTux2.Enabled = False
CmdTux3.Enabled = False
cmdboughtown.Enabled = False
cmdUhOh.Enabled = True

End Sub

Private Sub cmdQuit_Click()
'this button is to end the program
End
End Sub

Private Sub cmdTux1_Click()
'this button is used if the user wants to pic tux number 1
CostofTux = 89#
picTuxCost.Print Groom; " will look handsome in this tux. The total cost of your ensamble will be "; FormatCurrency(CostofTux); "."
RunningTotal = RunningTotal + CostofTux
picTuxCost.Print "Your total running total is now "; FormatCurrency(RunningTotal); "."
cmdTux1.Enabled = False
cmdTux2.Enabled = False
CmdTux3.Enabled = False
cmdboughtown.Enabled = False
cmdUhOh.Enabled = True

End Sub

Private Sub cmdTux2_Click()
'this button is used if the user wants to pic tux number 2
CostofTux = 99#
picTuxCost.Print Groom; " will look handsome in this tux. The total cost of your ensamble will be "; FormatCurrency(CostofTux); "."
RunningTotal = RunningTotal + CostofTux
picTuxCost.Print "Your total running total is now "; FormatCurrency(RunningTotal); "."
cmdTux1.Enabled = False
cmdTux2.Enabled = False
CmdTux3.Enabled = False
cmdboughtown.Enabled = False
cmdUhOh.Enabled = True

End Sub

Private Sub CmdTux3_Click()
'this button is used if the user wants to pic tux number 3
CostofTux = 109#
picTuxCost.Print Groom; " will look handsome in this tux. The total cost of your ensamble will be "; FormatCurrency(CostofTux); "."
RunningTotal = RunningTotal + CostofTux
picTuxCost.Print "Your total running total is now "; FormatCurrency(RunningTotal); "."
cmdTux1.Enabled = False
cmdTux2.Enabled = False
CmdTux3.Enabled = False
cmdboughtown.Enabled = False
cmdUhOh.Enabled = True

End Sub

Private Sub cmdUhOh_Click()
'this button is used if the user has made a selection and wants to change it. it clears the picture box and minuses the cost of the selection from the total cost.
picTuxCost.Cls
RunningTotal = RunningTotal - CostofTux
cmdTux1.Enabled = True
cmdTux2.Enabled = True
CmdTux3.Enabled = True
cmdboughtown.Enabled = True
End Sub

