VERSION 5.00
Begin VB.Form Dinner 
   BackColor       =   &H0080FF80&
   Caption         =   "Dinner"
   ClientHeight    =   6015
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8955
   LinkTopic       =   "Form1"
   ScaleHeight     =   6015
   ScaleWidth      =   8955
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDinner 
      Caption         =   "Dinner is served..."
      Enabled         =   0   'False
      Height          =   855
      Left            =   600
      TabIndex        =   16
      Top             =   4920
      Width           =   1455
   End
   Begin VB.CommandButton cmdNext4 
      Caption         =   "Now it is time to move on to the last part of our date..."
      Enabled         =   0   'False
      Height          =   855
      Left            =   3600
      TabIndex        =   15
      Top             =   4920
      Width           =   1575
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   615
      Left            =   6600
      TabIndex        =   14
      Top             =   5160
      Width           =   1095
   End
   Begin VB.OptionButton optSpecial 
      BackColor       =   &H0080FF80&
      Height          =   255
      Left            =   7320
      TabIndex        =   13
      Top             =   1800
      Width           =   255
   End
   Begin VB.PictureBox picSpecial 
      Height          =   1335
      Left            =   6480
      Picture         =   "StacyWurm4.frx":0000
      ScaleHeight     =   1275
      ScaleWidth      =   1875
      TabIndex        =   11
      Top             =   2160
      Width           =   1935
   End
   Begin VB.OptionButton optTacoBell 
      BackColor       =   &H0080FF80&
      Height          =   255
      Left            =   5400
      TabIndex        =   10
      Top             =   1800
      Width           =   255
   End
   Begin VB.PictureBox picTacoBell 
      Height          =   975
      Left            =   4800
      Picture         =   "StacyWurm4.frx":12BC
      ScaleHeight     =   915
      ScaleWidth      =   1275
      TabIndex        =   8
      Top             =   2280
      Width           =   1335
   End
   Begin VB.OptionButton optOliveGarden 
      BackColor       =   &H0080FF80&
      Height          =   255
      Left            =   3360
      TabIndex        =   7
      Top             =   1800
      Width           =   255
   End
   Begin VB.PictureBox picOliveGarden 
      Height          =   1215
      Left            =   2640
      Picture         =   "StacyWurm4.frx":1C5B
      ScaleHeight     =   1155
      ScaleWidth      =   1755
      TabIndex        =   6
      Top             =   2160
      Width           =   1815
   End
   Begin VB.PictureBox picResults 
      Height          =   855
      Left            =   720
      ScaleHeight     =   795
      ScaleWidth      =   7395
      TabIndex        =   4
      Top             =   3840
      Width           =   7455
   End
   Begin VB.OptionButton optApplebees 
      BackColor       =   &H0080FF80&
      Height          =   255
      Left            =   1320
      TabIndex        =   3
      Top             =   1800
      Width           =   255
   End
   Begin VB.PictureBox picApplebees 
      Height          =   1335
      Left            =   480
      Picture         =   "StacyWurm4.frx":2C1A
      ScaleHeight     =   1275
      ScaleWidth      =   1755
      TabIndex        =   2
      Top             =   2160
      Width           =   1815
   End
   Begin VB.Label Specialty 
      Alignment       =   2  'Center
      BackColor       =   &H0080FF80&
      Caption         =   "Local or Specialty Resturaunt"
      Height          =   255
      Left            =   6360
      TabIndex        =   12
      Top             =   1440
      Width           =   2295
   End
   Begin VB.Label TacoBell 
      Alignment       =   2  'Center
      BackColor       =   &H0080FF80&
      Caption         =   "Taco Bell"
      Height          =   255
      Left            =   4920
      TabIndex        =   9
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label OliveGarden 
      Alignment       =   2  'Center
      BackColor       =   &H0080FF80&
      Caption         =   "Olive Garden"
      Height          =   255
      Left            =   2880
      TabIndex        =   5
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label Applebees 
      Alignment       =   2  'Center
      BackColor       =   &H0080FF80&
      Caption         =   "Applebee's"
      Height          =   255
      Left            =   960
      TabIndex        =   1
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label DinnerO 
      Alignment       =   2  'Center
      BackColor       =   &H0080FF80&
      Caption         =   $"StacyWurm4.frx":3BC0
      Height          =   615
      Left            =   1080
      TabIndex        =   0
      Top             =   480
      Width           =   6855
   End
End
Attribute VB_Name = "Dinner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Project Name: Date Chooser (Wurm, Stacy - VB Project)
' Form Name: Dinner (StacyWurm4.frm)
' Author: Stacy Wurm
' Date Written: Monday, March 8th, 2004
' Purpose of this Form: ' Allow the user to choose where to go to dinner
                        ' Also print the choice and amount it will cost
                        ' and the overall amount spent

Private Sub cmdDinner_Click()
'  Dictates what will be printed for each option
    If optApplebees = True Then
        Cost = 20
        TotalCost = TotalCost + Cost
        Choice = "Applebee's!!  Always a great choice with lots of options!!"
        Decision3 = "Applebee's"
    ElseIf optOliveGarden = True Then
        Cost = 35
        Choice = "the Olive Garden!!  Italian is a great choice!!"
        TotalCost = TotalCost + Cost
        Decision3 = "the Olive Garden"
    ElseIf optTacoBell = True Then
        Cost = 10
        Choice = "Taco Bell!!  South of the border time huh??"
        TotalCost = TotalCost + Cost
        Decision3 = "Taco Bell"
    ElseIf optSpecial = True Then
        Cost = 45
        Choice = "a local or specialty resturaunt!  Wow, fancy, she must be really special!!"
        TotalCost = TotalCost + Cost
        Decision3 = "a local or specialty resturaunt"
    End If

'Prints results for the dinner page
picResults.Print "Dinner is served at "; Choice
picResults.Print "This is going to take "; FormatCurrency(Cost); " away from your total!!"
picResults.Print "You have spent "; FormatCurrency(TotalCost); " of your budget to this point"
cmdNext4.Enabled = True
End Sub

Private Sub cmdNext4_Click()
' This moves the user on to the next form
Dinner.Hide
AfterDinnerPlans.Show
End Sub

Private Sub cmdQuit_Click()
End
End Sub

Private Sub optApplebees_Click()
' Clears results and chooses this option
picResults.Cls
cmdDinner.Enabled = True
End Sub

Private Sub optOliveGarden_Click()
' Clears results and chooses this option
picResults.Cls
cmdDinner.Enabled = True
End Sub

Private Sub optSpecial_Click()
' Clears results and chooses this option
picResults.Cls
cmdDinner.Enabled = True
End Sub

Private Sub optTacoBell_Click()
' Clears results and chooses this option
picResults.Cls
cmdDinner.Enabled = True
End Sub
