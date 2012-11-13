VERSION 5.00
Begin VB.Form DateEvent 
   BackColor       =   &H00FF80FF&
   Caption         =   "Choice2"
   ClientHeight    =   6030
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6030
   ScaleWidth      =   9000
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picActivity 
      BackColor       =   &H00FF80FF&
      Height          =   1455
      Left            =   6600
      Picture         =   "StacyWurm3.frx":0000
      ScaleHeight     =   1395
      ScaleWidth      =   1275
      TabIndex        =   16
      Top             =   2160
      Width           =   1335
   End
   Begin VB.PictureBox picSports 
      BackColor       =   &H00FF80FF&
      Height          =   1575
      Left            =   4680
      Picture         =   "StacyWurm3.frx":0C67
      ScaleHeight     =   1515
      ScaleWidth      =   1275
      TabIndex        =   15
      Top             =   2040
      Width           =   1335
   End
   Begin VB.PictureBox picComedyClub 
      BackColor       =   &H00FF80FF&
      Height          =   1215
      Left            =   2640
      Picture         =   "StacyWurm3.frx":1DFC
      ScaleHeight     =   1155
      ScaleWidth      =   1155
      TabIndex        =   14
      Top             =   2040
      Width           =   1215
   End
   Begin VB.PictureBox picTheatre 
      BackColor       =   &H00FF80FF&
      Height          =   1455
      Left            =   720
      Picture         =   "StacyWurm3.frx":278C
      ScaleHeight     =   1395
      ScaleWidth      =   1035
      TabIndex        =   13
      Top             =   2040
      Width           =   1095
   End
   Begin VB.OptionButton optActivity 
      BackColor       =   &H00FF80FF&
      Height          =   255
      Left            =   7080
      TabIndex        =   12
      Top             =   1800
      Width           =   255
   End
   Begin VB.OptionButton optSports 
      BackColor       =   &H00FF80FF&
      Height          =   255
      Left            =   5160
      TabIndex        =   11
      Top             =   1680
      Width           =   255
   End
   Begin VB.OptionButton optComedy 
      BackColor       =   &H00FF80FF&
      Height          =   255
      Left            =   3120
      TabIndex        =   10
      Top             =   1680
      Width           =   255
   End
   Begin VB.OptionButton optMovie 
      BackColor       =   &H00FF80FF&
      Height          =   255
      Left            =   1080
      TabIndex        =   9
      Top             =   1680
      Width           =   255
   End
   Begin VB.PictureBox picResults3 
      Height          =   975
      Left            =   600
      ScaleHeight     =   915
      ScaleWidth      =   8115
      TabIndex        =   4
      Top             =   3720
      Width           =   8175
   End
   Begin VB.CommandButton cmdEvent 
      Caption         =   "I have choosen our big event!!!"
      Enabled         =   0   'False
      Height          =   975
      Left            =   480
      TabIndex        =   3
      Top             =   4800
      Width           =   1815
   End
   Begin VB.CommandButton cmdNext3 
      Caption         =   "Now time for dinner!!  Hope you are hungary!!"
      Enabled         =   0   'False
      Height          =   975
      Left            =   3720
      TabIndex        =   2
      Top             =   4800
      Width           =   2055
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   615
      Left            =   7200
      TabIndex        =   0
      Top             =   4920
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FF80FF&
      Caption         =   "Anything Active (go roller skating, play sports, etc.)"
      Height          =   495
      Left            =   6360
      TabIndex        =   8
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Label Sports 
      Alignment       =   2  'Center
      BackColor       =   &H00FF80FF&
      Caption         =   "Sporting Event"
      Height          =   255
      Left            =   4560
      TabIndex        =   7
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label ComedyClub 
      Alignment       =   2  'Center
      BackColor       =   &H00FF80FF&
      Caption         =   "Comedy Club"
      Height          =   255
      Left            =   2640
      TabIndex        =   6
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Movie 
      Alignment       =   2  'Center
      BackColor       =   &H00FF80FF&
      Caption         =   "Movie"
      Height          =   255
      Left            =   720
      TabIndex        =   5
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Event 
      Alignment       =   2  'Center
      BackColor       =   &H00FF80FF&
      Caption         =   "Now it is time to decide what to do before dinner.  There are many different options.  Hope you pick one that your date likes!!!"
      Height          =   495
      Left            =   1440
      TabIndex        =   1
      Top             =   480
      Width           =   6015
   End
End
Attribute VB_Name = "DateEvent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Project Name: Date Chooser (Wurm, Stacy - VB Project)
' Form Name: DateEvent (StacyWurm3.frm)
' Author: Stacy Wurm
' Date Written: Sunday, March 7th, 2004
' Purpose of this Form: ' Allows the user to choose what to do
                        ' For a movie it also gives concession options
                        ' it asks your choice and totals how much spent
                        ' Then also displays amount spent so far

Private Sub cmdEvent_Click()
' First option and display total and choice
    If optMovie = True Then
        Cost = 15
        TotalCost = TotalCost + Cost
        Choice = "a movie!!  Hope you picked a good one!!"
        Decision2 = "a movie"
    ElseIf optComedy = True Then
        Cost = 20
        Choice = "a comedy club how funny!!!"
        TotalCost = TotalCost + Cost
        Decision2 = "a comedy club"
    ElseIf optSports = True Then
        Cost = 30
        Choice = "a sporting event!  I hope she likes sports(especially the one you chose)!"
        TotalCost = TotalCost + Cost
        Decision2 = "a sporting event"
    ElseIf optActivity = True Then
        Cost = 0
        Choice = "do an activity is something different!!  Plus it is free!!"
        TotalCost = TotalCost + Cost
        Decision2 = "participate in an activity"
    End If
picResults3.Print "You have decided to go to "; Choice
    If optMovie = True Then
        picResults3.Print "If you had concessions at the movie you spent "; FormatCurrency(MovieCost); "."
    End If
picResults3.Print "This is going to take "; FormatCurrency(Cost); " away from your total!!"
picResults3.Print "You have spent "; FormatCurrency(TotalCost); " of your budget to this point"
cmdNext3.Enabled = True
End Sub

Private Sub cmdNext3_Click()
' moves the user on to the choice for dinner
DateEvent.Hide
Dinner.Show
End Sub

Private Sub cmdQuit_Click()
End
End Sub

Private Sub optActivity_Click()
' Allows this event to be choosen by user
picResults3.Cls
cmdEvent.Enabled = True
End Sub

Private Sub optComedy_Click()
' Allows this event to be choosen by user
picResults3.Cls
cmdEvent.Enabled = True
End Sub

Private Sub optMovie_Click()
' Allows this event to be choosen by user
DateEventMovie.Show
picResults3.Cls
cmdEvent.Enabled = True
End Sub

Private Sub optSports_Click()
' Allows this event to be choosen by user
picResults3.Cls
cmdEvent.Enabled = True
End Sub
