VERSION 5.00
Begin VB.Form FiveKComparisonForm 
   BackColor       =   &H00FF0000&
   Caption         =   "5K Time Comparison "
   ClientHeight    =   8625
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13560
   LinkTopic       =   "Form2"
   ScaleHeight     =   8625
   ScaleWidth      =   13560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Quitbutton 
      Caption         =   "Quit"
      Height          =   1695
      Left            =   6720
      TabIndex        =   2
      Top             =   2040
      Width           =   2655
   End
   Begin VB.CommandButton Comparebutton 
      Caption         =   "Are you faster or slower than the great Steve Prefontaine?"
      Height          =   1575
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   2655
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H0000C000&
      Height          =   1095
      Left            =   3120
      ScaleHeight     =   1035
      ScaleWidth      =   6435
      TabIndex        =   0
      Top             =   480
      Width           =   6495
   End
   Begin VB.Image Image1 
      Height          =   6285
      Left            =   600
      Picture         =   "FiveKComparrisonForm.frx":0000
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   5595
   End
End
Attribute VB_Name = "FiveKComparisonForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: TrackandFieldProgram (TrackProgram)'
'Form Name: FiveKComparisonForm (FiveKComparrisonForm.frm)'
'Written By: Paul Jeske'
'Date Written: March 15th, 2004'
'Purpose of this form:  This form takes the user's total time
                        'and allows them to see how far
                        'ahead or behind they are from
                        'the great running legend Steve Prefontaine's times.
                        'The user may then terminates usage of the program by
                        'selecting the "quit" button



Private Sub Comparebutton_Click()
'Dims the needed variables'
Dim Pre As Single
Dim Diff As Single

Pre = 803.8 'sets "Pre" equal to the best time(in seconds) that Steve Prefontaine ran for the 5k
Diff = 0 'Sets "Diff" equal to zero'

'Determines if the user's time is faster or slower than Prefontaine's then finds, and displays the difference'
If Sum > Pre Then
    Diff = Sum - Pre
    picResults.Print "RUN FASTER!! You are"; Diff; "seconds slower than Steve Prefontaine"
    Else
    Diff = Pre - Sum
    picResults.Print "CONGRATULATIONS!! You are"; Diff; "seconds faster than Steve Prefontaine"
End If
End Sub

Private Sub Quitbutton_Click()
'Brings up a message box with an encouraging message to the user and then allows the user to terminate the program'
MsgBox "Keep Running Hard!!", , "The Finish Line"
End
End Sub
