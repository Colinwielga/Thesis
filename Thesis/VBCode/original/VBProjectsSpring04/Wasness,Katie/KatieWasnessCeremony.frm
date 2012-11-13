VERSION 5.00
Begin VB.Form frmCeremony 
   BackColor       =   &H8000000D&
   Caption         =   "CEREMONY"
   ClientHeight    =   5850
   ClientLeft      =   6465
   ClientTop       =   4575
   ClientWidth     =   8580
   LinkTopic       =   "Form1"
   ScaleHeight     =   5850
   ScaleWidth      =   8580
   Begin VB.CommandButton cmdboughtown 
      Caption         =   "I have selected a different ceremony place, but would like to input teh cost into the budget"
      Height          =   1695
      Left            =   6960
      TabIndex        =   11
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton cmdChurch 
      Caption         =   "Click Here if You Would Like Your Wedding in this Church"
      Height          =   615
      Left            =   4800
      TabIndex        =   10
      Top             =   3240
      Width           =   2055
   End
   Begin VB.CommandButton cmdMoutian 
      Caption         =   "Click Here if You Would Like Your Wedding Outside Viewing this Mountian"
      Height          =   615
      Left            =   2520
      TabIndex        =   9
      Top             =   3240
      Width           =   2175
   End
   Begin VB.CommandButton cmdGarden 
      Caption         =   "Click Here if You Would Like Your Wedding in this Garden"
      Height          =   615
      Left            =   240
      TabIndex        =   8
      Top             =   3240
      Width           =   2175
   End
   Begin VB.PictureBox picCeremony 
      BackColor       =   &H8000000D&
      ForeColor       =   &H8000000E&
      Height          =   855
      Left            =   240
      ScaleHeight     =   795
      ScaleWidth      =   7875
      TabIndex        =   7
      Top             =   3960
      Width           =   7935
   End
   Begin VB.CommandButton cmdUhOh 
      Caption         =   "Click Here If You Are Not Satisfied With Your Choice And You Would Like To Change It"
      Enabled         =   0   'False
      Height          =   1815
      Left            =   6960
      TabIndex        =   6
      Top             =   240
      Width           =   1215
   End
   Begin VB.PictureBox picChurch 
      Height          =   2775
      Left            =   4800
      Picture         =   "Katie Wasness Ceremony.frx":0000
      ScaleHeight     =   2715
      ScaleWidth      =   1995
      TabIndex        =   5
      Top             =   360
      Width           =   2055
   End
   Begin VB.PictureBox picMoutian 
      Height          =   2775
      Left            =   2520
      Picture         =   "Katie Wasness Ceremony.frx":5E0A
      ScaleHeight     =   2715
      ScaleWidth      =   2115
      TabIndex        =   4
      Top             =   360
      Width           =   2175
   End
   Begin VB.PictureBox picGarden 
      Height          =   2775
      Left            =   240
      Picture         =   "Katie Wasness Ceremony.frx":B7A0
      ScaleHeight     =   2715
      ScaleWidth      =   2115
      TabIndex        =   3
      Top             =   360
      Width           =   2175
   End
   Begin VB.CommandButton cmdBackTo 
      Caption         =   "Go Back to Main Menu"
      Height          =   615
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   4920
      Width           =   1935
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   495
      Left            =   6840
      TabIndex        =   0
      Top             =   4920
      Width           =   1335
   End
   Begin VB.Label lbtitle 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      Caption         =   "WHERE WOULD YOU LIKE YOUR CEREMONY??"
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   7575
   End
End
Attribute VB_Name = "frmCeremony"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'pjtWeddingBudget(Katie Wasness Wedding)
'frmCeremony(Katie Wasness Ceremony)
'Author = Katie Wasness
'Date Written: March 10-15
'Purpose--The purpose of this form is choose the Ceremony site and put the price of that into the Running Total.
Private Sub cmdBackTo_Click(Index As Integer)
'this button brings the user back to the main menu
frmWedding.Show
frmCeremony.Hide
End Sub

Private Sub cmdboughtown_Click()
'this button is in case an individual has choosen a ceremony site somewhere else
CostOfCeremony = InputBox("What is the cost of your ceremony site?", "Other Ceremony Site")
Choice = "Other Ceremony Site"
picCeremony.Print "The Cost Of Your Wedding In Your Location Is "; FormatCurrency(CostOfCeremony); "."
RunningTotal = RunningTotal + CostOfCeremony
picCeremony.Print "The Total Cost Of Your Wedding Thus Far Is "; FormatCurrency(RunningTotal); "."
cmdMoutian.Enabled = False
cmdGarden.Enabled = False
cmdChurch.Enabled = False
cmdboughtown.Enabled = False
cmdUhOh.Enabled = True


End Sub

Private Sub cmdChurch_Click()
'this button is used to select the Church as the ceremony site and to input that as the price of the ceremony
CostOfCeremony = 400#
Choice = "Church"
picCeremony.Print "The Cost Of Your Wedding In This Church Is "; FormatCurrency(CostOfCeremony); "."
picCeremony.Print "This includes the cost of the ceremony site reservation, officiants, and decorations."
RunningTotal = RunningTotal + CostOfCeremony
picCeremony.Print "The Total Cost Of Your Wedding Thus Far Is "; FormatCurrency(RunningTotal); "."
cmdMoutian.Enabled = False
cmdGarden.Enabled = False
cmdChurch.Enabled = False
cmdboughtown.Enabled = False
cmdUhOh.Enabled = True
End Sub

Private Sub cmdGarden_Click()
'this button is used to select the Garden as the ceremony site and to input that as the price of the ceremony
Choice = "Garden"
CostOfCeremony = 1500#
picCeremony.Print "The Cost Of Your Wedding In This Garden Is "; FormatCurrency(CostOfCeremony); "."
picCeremony.Print "This includes the cost of all rentals, flowers for ceremony site, ceremony site reservation,"
picCeremony.Print "officiants, and decorations."
RunningTotal = RunningTotal + CostOfCeremony
picCeremony.Print "The Total Cost Of Your Wedding Thus Far Is "; FormatCurrency(RunningTotal); "."
cmdMoutian.Enabled = False
cmdGarden.Enabled = False
cmdChurch.Enabled = False
cmdboughtown.Enabled = False
cmdUhOh.Enabled = True
End Sub

Private Sub cmdMoutian_Click()
'this button is used to select the Mountian as the ceremony site and to input that as the price of the ceremony
Dim CostofMountian As Single
Choice = "Mountian"
CostOfCeremony = 750#
picCeremony.Print "The Cost Of Your Wedding Outside Viewing this Mountain Is "; FormatCurrency(CostOfCeremony); "."
picCeremony.Print "This includes the cost of all rentals, ceremony site reservation, officiants, and decorations."
RunningTotal = RunningTotal + CostOfCeremony
picCeremony.Print "The Total Cost Of Your Wedding Thus Far Is "; FormatCurrency(RunningTotal); "."
cmdMoutian.Enabled = False
cmdGarden.Enabled = False
cmdChurch.Enabled = False
cmdboughtown.Enabled = False
cmdUhOh.Enabled = True
End Sub

Private Sub cmdQuit_Click()
'this button is to quit the program
End
End Sub

Private Sub cmdUhOh_Click()
'this button is used if the user has made a selection and wants to change it. it clears the picture box and minuses the cost of the selection from the total cost.
picCeremony.Cls
RunningTotal = RunningTotal - CostOfCeremony
cmdMoutian.Enabled = True
cmdGarden.Enabled = True
cmdChurch.Enabled = True
cmdboughtown.Enabled = True
cmdUhOh.Enabled = False
End Sub

