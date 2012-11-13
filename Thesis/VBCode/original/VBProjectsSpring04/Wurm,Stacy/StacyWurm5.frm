VERSION 5.00
Begin VB.Form AfterDinnerPlans 
   BackColor       =   &H0080FFFF&
   Caption         =   "Choice4"
   ClientHeight    =   5820
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9390
   LinkTopic       =   "Form1"
   ScaleHeight     =   5820
   ScaleWidth      =   9390
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picHouse 
      BackColor       =   &H0080FFFF&
      Height          =   1455
      Left            =   6960
      Picture         =   "StacyWurm5.frx":0000
      ScaleHeight     =   1395
      ScaleWidth      =   1035
      TabIndex        =   16
      Top             =   1920
      Width           =   1095
   End
   Begin VB.PictureBox picDance 
      BackColor       =   &H0080FFFF&
      Height          =   1215
      Left            =   4560
      Picture         =   "StacyWurm5.frx":0BF7
      ScaleHeight     =   1155
      ScaleWidth      =   1635
      TabIndex        =   15
      Top             =   2040
      Width           =   1695
   End
   Begin VB.PictureBox picBar 
      BackColor       =   &H0080FFFF&
      Height          =   1575
      Left            =   2880
      Picture         =   "StacyWurm5.frx":1907
      ScaleHeight     =   1515
      ScaleWidth      =   1155
      TabIndex        =   14
      Top             =   1920
      Width           =   1215
   End
   Begin VB.PictureBox picCoffee 
      BackColor       =   &H0080FFFF&
      Height          =   1695
      Left            =   720
      Picture         =   "StacyWurm5.frx":2751
      ScaleHeight     =   1635
      ScaleWidth      =   1515
      TabIndex        =   13
      Top             =   1920
      Width           =   1575
   End
   Begin VB.CommandButton cmdAfter 
      Caption         =   "This is what we are doing after dinner..."
      Enabled         =   0   'False
      Height          =   975
      Left            =   480
      TabIndex        =   12
      Top             =   4680
      Width           =   1575
   End
   Begin VB.CommandButton cmdWrapUp 
      Caption         =   "Well we have reached the end of the date...time to see what happened..."
      Enabled         =   0   'False
      Height          =   975
      Left            =   3960
      TabIndex        =   11
      Top             =   4680
      Width           =   1815
   End
   Begin VB.OptionButton optHome 
      BackColor       =   &H0080FFFF&
      Height          =   255
      Left            =   7440
      TabIndex        =   10
      Top             =   1560
      Width           =   255
   End
   Begin VB.OptionButton optClub 
      BackColor       =   &H0080FFFF&
      Height          =   255
      Left            =   5280
      TabIndex        =   9
      Top             =   1560
      Width           =   255
   End
   Begin VB.OptionButton optBar 
      BackColor       =   &H0080FFFF&
      Height          =   255
      Left            =   3360
      TabIndex        =   8
      Top             =   1560
      Width           =   255
   End
   Begin VB.OptionButton optCoffee 
      BackColor       =   &H0080FFFF&
      Height          =   255
      Left            =   1320
      TabIndex        =   7
      Top             =   1560
      Width           =   255
   End
   Begin VB.PictureBox picResults4 
      Height          =   855
      Left            =   600
      ScaleHeight     =   795
      ScaleWidth      =   8235
      TabIndex        =   2
      Top             =   3600
      Width           =   8295
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   615
      Left            =   7800
      TabIndex        =   0
      Top             =   4920
      Width           =   975
   End
   Begin VB.Label Home 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "Hang out at my Place"
      Height          =   255
      Left            =   6720
      TabIndex        =   6
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label Club 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "Go Dancin' at a Club"
      Height          =   255
      Left            =   4560
      TabIndex        =   5
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label Bar 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "To a Bar"
      Height          =   255
      Left            =   2880
      TabIndex        =   4
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Coffee 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "Go out for Coffee"
      Height          =   255
      Left            =   840
      TabIndex        =   3
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label AfterDinner 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   $"StacyWurm5.frx":313C
      Height          =   495
      Left            =   960
      TabIndex        =   1
      Top             =   360
      Width           =   7455
   End
End
Attribute VB_Name = "AfterDinnerPlans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Project Name: Date Chooser (Wurm, Stacy - VB Project)
' Form Name: Dinner (StacyWurm5.frm)
' Author: Stacy Wurm
' Date Written: Wednesday, March 10th, 2004
' Purpose of this Form: ' Allows user to decide where to go after dinner
                        ' Displays the cost and option that the user decided
                        ' then displays the overall cost to this point

Private Sub cmdAfter_Click()
' Dictates what will be printed for each me option
    If optCoffee = True Then
        Cost = 7
        Choice = "go out for coffee.  What a nice way to wind down a great evening."
        TotalCost = TotalCost + Cost
        Decision4 = "going out for coffee"
    ElseIf optBar = True Then
        Cost = 12
        Choice = "go to a bar.  Isn't that where all this date nonsence started to begin with?"
        TotalCost = TotalCost + Cost
        Decision4 = "going out to a bar"
    ElseIf optClub = True Then
        Cost = 10
        Choice = "go dancing at a club!!  Wow lots of energy here!!"
        TotalCost = TotalCost + Cost
        Decision4 = "going dancing at a club"
    ElseIf optHome = True Then
        Cost = 0
        Choice = "go back to your place!  Would your mother approve?  Well at least this one is free too!"
        TotalCost = TotalCost + Cost
        Decision4 = "going back to your place"
    End If

' Print all the results
picResults4.Print "After dinner you will "; Choice
picResults4.Print "This is going to take "; FormatCurrency(Cost); " away from your total!!"
picResults4.Print "The amount of your budget that has been spent is "; FormatCurrency(TotalCost); "."
cmdWrapUp.Enabled = True
End Sub

Private Sub cmdQuit_Click()
End
End Sub

Private Sub cmdWrapUp_Click()
' this command brings us to the final page for a wrap up
' of the entire date
AfterDinnerPlans.Hide
InTheEnd.Show
End Sub

Private Sub optBar_Click()
' user can choose this option
picResults4.Cls
cmdAfter.Enabled = True
End Sub

Private Sub optClub_Click()
' user can choose this option
picResults4.Cls
cmdAfter.Enabled = True
End Sub

Private Sub optCoffee_Click()
' user can choose this option
picResults4.Cls
cmdAfter.Enabled = True
End Sub

Private Sub optHome_Click()
' user can choose this option
picResults4.Cls
cmdAfter.Enabled = True
End Sub
