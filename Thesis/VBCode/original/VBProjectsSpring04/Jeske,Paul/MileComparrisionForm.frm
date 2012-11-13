VERSION 5.00
Begin VB.Form MileComparrisonForm 
   BackColor       =   &H000000C0&
   Caption         =   "Mile Time Comparison"
   ClientHeight    =   8295
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8295
   LinkTopic       =   "Form1"
   ScaleHeight     =   8295
   ScaleWidth      =   8295
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Quitbutton 
      Caption         =   "Quit"
      Height          =   5775
      Left            =   6480
      TabIndex        =   2
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton comparebutton 
      Caption         =   "Are you faster or slower than the Great Steve Prefontaine??"
      Height          =   1575
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   1815
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00FF8080&
      Height          =   1095
      Left            =   2280
      ScaleHeight     =   1035
      ScaleWidth      =   5595
      TabIndex        =   0
      Top             =   480
      Width           =   5655
   End
   Begin VB.Image Image1 
      Height          =   6045
      Left            =   240
      Picture         =   "MileComparrisionForm.frx":0000
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   5715
   End
End
Attribute VB_Name = "MileComparrisonForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: TrackandFieldProgram (TrackProgram)'
'Form Name: MileComparrisionForm (MileComparrisonForm.frm)'
'Written By: Paul Jeske'
'Date Written: March 15th, 2004'
'Purpose of this form:  This form takes the user's total time
                        'and allows them to see how far
                        'ahead or behind they are from
                        'the great running legend Steve Prefontaine's times.
                        'The user may then terminates usage of the program by
                        'selecting the "quit" button

Private Sub Comparebutton_Click()
'Dims needed variables'
Dim Pre As Single
Dim Diff As Single

Pre = 234.6 'sets "Pre" equal to the best time(in seconds) that Steve Prefontaine ran for the mile run'
Diff = 0 'Sets "Diff" equal to zero'

'Determines if the user's time is faster or slower than Prefontaine's then finds, and displays the difference'If Sum > Pre Then

    If Sum > Pre Then
    Diff = Sum - Pre
    picResults.Print "RUN FASTER!! You are"; Diff; "seconds slower than Steve Prefontaine"
    Else
    Diff = Pre - Sum
    picResults.Print "CONGRATULATIONS!! You are"; Diff; "seconds faster than Steve Prefontaine"
End If
End Sub


Private Sub Quitbutton_Click()
'Brings up a message box with an encouraging message to the user and then allows user to terminate the program'
MsgBox "KEEP RUNNING HARD!!", , "The Finish Line"
End
End Sub
