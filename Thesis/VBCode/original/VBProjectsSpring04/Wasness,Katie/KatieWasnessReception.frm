VERSION 5.00
Begin VB.Form frmReception 
   BackColor       =   &H8000000D&
   Caption         =   "RECEPTION"
   ClientHeight    =   7095
   ClientLeft      =   5520
   ClientTop       =   3825
   ClientWidth     =   8850
   LinkTopic       =   "Form1"
   ScaleHeight     =   7095
   ScaleWidth      =   8850
   Begin VB.CommandButton cmdboughtown 
      Caption         =   "Click Here if you have choosen a different reception site and would like to input the price."
      Height          =   1095
      Left            =   4320
      TabIndex        =   11
      Top             =   3720
      Width           =   1935
   End
   Begin VB.PictureBox picReception 
      BackColor       =   &H8000000D&
      ForeColor       =   &H8000000E&
      Height          =   1215
      Left            =   120
      ScaleHeight     =   1155
      ScaleWidth      =   8115
      TabIndex        =   10
      Top             =   4920
      Width           =   8175
   End
   Begin VB.CommandButton CmdUhOh 
      Caption         =   "Click Here If You Have Made A Selection And Would Like To Change It."
      Enabled         =   0   'False
      Height          =   1095
      Left            =   2640
      TabIndex        =   9
      Top             =   3720
      Width           =   1575
   End
   Begin VB.CommandButton cmd3Rec 
      Caption         =   "Click Here To Have Your Reception At The Science Museum of Minnesota"
      Height          =   735
      Left            =   6360
      TabIndex        =   8
      Top             =   3960
      Width           =   2055
   End
   Begin VB.CommandButton cmd2Rec 
      Caption         =   "Click Here To Have Your Reception At The McNamara Alumni Center-----U of M"
      Height          =   735
      Left            =   3120
      TabIndex        =   7
      Top             =   2880
      Width           =   2775
   End
   Begin VB.CommandButton cmd1Rec 
      Caption         =   "Click Here To Have Your Reception At The Christos Union Depot Space"
      Height          =   735
      Left            =   240
      TabIndex        =   6
      Top             =   3960
      Width           =   2295
   End
   Begin VB.PictureBox pic3Rec 
      Height          =   3255
      Left            =   6360
      Picture         =   "Katie Wasness Reception.frx":0000
      ScaleHeight     =   3195
      ScaleWidth      =   1995
      TabIndex        =   4
      Top             =   600
      Width           =   2055
   End
   Begin VB.PictureBox pic2Rec 
      Height          =   2175
      Left            =   3120
      Picture         =   "Katie Wasness Reception.frx":64E3
      ScaleHeight     =   2115
      ScaleWidth      =   2715
      TabIndex        =   3
      Top             =   600
      Width           =   2775
   End
   Begin VB.PictureBox pic1Rec 
      Height          =   3255
      Left            =   240
      Picture         =   "Katie Wasness Reception.frx":D836
      ScaleHeight     =   3195
      ScaleWidth      =   2235
      TabIndex        =   2
      Top             =   600
      Width           =   2295
   End
   Begin VB.CommandButton cmdBackTo 
      Caption         =   "Go Back to Main Menu"
      Height          =   615
      Left            =   240
      TabIndex        =   1
      Top             =   6360
      Width           =   1935
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   735
      Left            =   6840
      TabIndex        =   0
      Top             =   6240
      Width           =   1455
   End
   Begin VB.Label labtitle 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      Caption         =   "WHERE WOULD YOU LIKE YOUR RECEPTION?"
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   8415
   End
End
Attribute VB_Name = "frmReception"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'pjtWeddingBudget(Katie Wasness Wedding)
'frmReception(Katie Wasness Reception)
'Author = Katie Wasness
'Date Written: March 10-15
'Purpose--The purpose of this form is choose the Reception site and put the price of that into the Running Total.

Private Sub cmd1Rec_Click()
'this button is used to select the Christos Union Depot Space for the reception site and input the cost into the Running Total
Choice2 = "Christos Union Depot Space"
CostOfReception = 8000#
picReception.Print "The Cost Of Your Reception At The "; Choice2; " Is "; FormatCurrency(CostOfReception); "."
picReception.Print "This includes the cost of all food, rentals, flowers for reception site, reception site reservation, and decorations."
RunningTotal = RunningTotal + CostOfReception
picReception.Print "The Total Cost of Your Wedding Thus Far Is "; FormatCurrency(RunningTotal); "."
cmd1Rec.Enabled = False
cmd2Rec.Enabled = False
cmd3Rec.Enabled = False
cmdboughtown.Enabled = False
cmdUhOh.Enabled = True
End Sub

Private Sub cmd2Rec_Click()
'this button is used to select the McNamara Alumni Center--U of M for the reception site and input the cost into the Running Total
Choice2 = "McNamara Alumni Center--U of M"
CostOfReception = 9000#
picReception.Print "The Cost Of Your Reception At The "; Choice2; " Is "; FormatCurrency(CostOfReception); "."
picReception.Print "This includes the cost of all food, rentals, flowers for reception site, reception site reservation, and decorations."
RunningTotal = RunningTotal + CostOfReception
picReception.Print "The Total Cost of Your Wedding Thus Far Is "; FormatCurrency(RunningTotal); "."
cmd1Rec.Enabled = False
cmd2Rec.Enabled = False
cmd3Rec.Enabled = False
cmdboughtown.Enabled = False
cmdUhOh.Enabled = True
End Sub

Private Sub cmd3Rec_Click()
'this button is used to select the Science Museum of Minnesota for the reception site and input the cost into the Running Total
Choice2 = "Science Museum of Minnesota"
CostOfReception = 5000#
picReception.Print "The Cost Of Your Reception At The "; Choice2; " Is "; FormatCurrency(CostOfReception); "."
picReception.Print "This includes the cost of all food, rentals, flowers for reception site, reception site reservation, and decorations."
RunningTotal = RunningTotal + CostOfReception
picReception.Print "The Total Cost of Your Wedding Thus Far Is "; FormatCurrency(RunningTotal); "."
cmd1Rec.Enabled = False
cmd2Rec.Enabled = False
cmd3Rec.Enabled = False
cmdboughtown.Enabled = False
cmdUhOh.Enabled = True
End Sub

Private Sub cmdBackTo_Click()
'this button is used to go back to the main menu
frmWedding.Show
frmReception.Hide
End Sub

Private Sub cmdboughtown_Click()
'this button is in case an individual has choosen a reception site somewhere else
CostOfReception = InputBox("What is the cost of your reception site?", "Other Reception Site")
Choice2 = "Another Reception Site"
picReception.Print "The Cost Of Your Wedding In Your Location Is "; FormatCurrency(CostOfReception); "."
RunningTotal = RunningTotal + CostOfReception
picReception.Print "The Total Cost Of Your Wedding Thus Far Is "; FormatCurrency(RunningTotal); "."
cmd1Rec.Enabled = False
cmd2Rec.Enabled = False
cmd3Rec.Enabled = False
cmdboughtown.Enabled = False
cmdUhOh.Enabled = True

End Sub

Private Sub cmdQuit_Click()
'this button is to end the program
End
End Sub

Private Sub cmdUhOh_Click()
'this button is used if the user has made a selection and wants to change it. it clears the picture box and minuses the cost of the selection from the total cost.
picReception.Cls
RunningTotal = RunningTotal - CostOfReception
cmd1Rec.Enabled = True
cmd2Rec.Enabled = True
cmd3Rec.Enabled = True
cmdboughtown.Enabled = True
cmdUhOh.Enabled = False
End Sub

