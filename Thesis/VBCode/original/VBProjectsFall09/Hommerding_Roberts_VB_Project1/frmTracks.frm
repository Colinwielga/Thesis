VERSION 5.00
Begin VB.Form frmTracks 
   Caption         =   "Race Tracks in the chase"
   ClientHeight    =   8805
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12375
   LinkTopic       =   "Form3"
   ScaleHeight     =   8805
   ScaleWidth      =   12375
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdFind 
      Caption         =   "Tracks that are shorter than the average distance"
      Enabled         =   0   'False
      Height          =   855
      Left            =   5040
      TabIndex        =   6
      Top             =   1320
      Width           =   1335
   End
   Begin VB.PictureBox picResults 
      Height          =   2895
      Left            =   1440
      ScaleHeight     =   2835
      ScaleWidth      =   3915
      TabIndex        =   5
      Top             =   2400
      Width           =   3975
   End
   Begin VB.CommandButton cmdAvg 
      Caption         =   "Average mileage of the tracks"
      Enabled         =   0   'False
      Height          =   855
      Left            =   3480
      TabIndex        =   4
      Top             =   1320
      Width           =   1455
   End
   Begin VB.CommandButton cmdSum 
      Caption         =   "Total mileage of these races"
      Enabled         =   0   'False
      Height          =   855
      Left            =   1920
      TabIndex        =   3
      Top             =   1320
      Width           =   1455
   End
   Begin VB.CommandButton cmdList 
      Caption         =   "List the 10 chase tracks"
      Height          =   855
      Left            =   360
      TabIndex        =   2
      Top             =   1320
      Width           =   1455
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return to Main Menu"
      Height          =   495
      Left            =   10440
      TabIndex        =   0
      Top             =   8040
      Width           =   1695
   End
   Begin VB.Label lblInfo 
      Caption         =   $"frmTracks.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   6255
   End
   Begin VB.Image Crash 
      Height          =   11520
      Left            =   -2880
      Picture         =   "frmTracks.frx":00B3
      Top             =   -1200
      Width           =   15360
   End
End
Attribute VB_Name = "frmTracks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Introduction to NASCAR
'Form Tracks
'Colin Roberts and Luke Hommerding
'Written 10/18/09
'Purpose is to load a file into parallel arrays and to average the tracks distances
'and sum the tracks mileages and display tracks that are shorter than the average
Option Explicit 'delcares variable
Dim Tracks(1 To 50) As String
Dim Miles(1 To 50) As Single
Dim Ctr As Integer
Dim I As Integer
Dim Sum As Integer
Dim Avg As Integer
'returns to main menu
Private Sub cmdReturn_Click()
    frmMain.Show
    frmTracks.Hide
End Sub
'load file into parallel arrays
Private Sub cmdList_Click()
    Open App.Path & "\ChaseTracks.txt" For Input As #1
    Ctr = 0
'searches through the file until all data has been read
    Do While Not EOF(1)
        Ctr = Ctr + 1
        Input #1, Tracks(Ctr), Miles(Ctr)
    Loop
    Close #1 'closes the file
    'prints the header
    picResults.Print "Tracks"; Tab(18); "Mileage"
    picResults.Print "*******************************"
    For I = 1 To Ctr
    'prints the trakcs and their distances
    picResults.Print Tracks(I); Tab(18); Miles(I)
    Next I
    'enables or disables buttons
    cmdList.Enabled = False
    cmdSum.Enabled = True
    cmdAvg.Enabled = False
    cmdFind.Enabled = False
End Sub
    'adds the data to find the sum
Private Sub cmdSum_Click()
picResults.Cls
Sum = 0
    For I = 1 To Ctr
        Sum = Sum + Miles(I)
    Next I
    'prints header
    picResults.Print "Total Miles Raced at these 10 Tracks is:"
    picResults.Print "----------------------------------------------------------------"
    'prints sum
    picResults.Print Sum; "Miles"
    'enables and disables buttons
    cmdList.Enabled = False
    cmdSum.Enabled = False
    cmdAvg.Enabled = True
    cmdFind.Enabled = False
End Sub
    'computes the average distances of the race tracks
Private Sub cmdAvg_Click()
picResults.Cls
Sum = 0
    For I = 1 To Ctr
        Sum = Sum + Miles(I)
    Next I
    'divides total miles by number of tracks
     Avg = Sum / Ctr
    'prints header
    picResults.Print "Average Miles Per Event Raced at these 10 Tracks is:"
    picResults.Print "-------------------------------------------------------------------------------------"
    'prints average of miles
    picResults.Print Avg; "Miles"
    'disables or enables buttons
    cmdList.Enabled = False
    cmdSum.Enabled = False
    cmdAvg.Enabled = False
    cmdFind.Enabled = True
End Sub
'searches for tracks shorter than the average
Private Sub cmdFind_Click()
picResults.Cls
'prints header
picResults.Print "Tracks that are shorter than the average race distance."
picResults.Print "--------------------------------------------------------------------------------------"

    For I = 1 To Ctr
    'compares track length to average length
        If Miles(I) < Avg Then
        'prints tracks shorter than the average
        picResults.Print Tracks(I), Miles(I); " Miles"
        End If
    Next I
    'enables or disables buttons
    cmdList.Enabled = False
    cmdSum.Enabled = False
    cmdAvg.Enabled = False
    cmdFind.Enabled = False

End Sub
