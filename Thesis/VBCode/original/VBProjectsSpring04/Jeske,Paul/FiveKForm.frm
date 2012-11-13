VERSION 5.00
Begin VB.Form FiveKForm 
   BackColor       =   &H000000FF&
   Caption         =   "5K Split Times"
   ClientHeight    =   7170
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10170
   LinkTopic       =   "Form1"
   ScaleHeight     =   7170
   ScaleWidth      =   10170
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Nextbutton 
      Caption         =   "Next"
      Enabled         =   0   'False
      Height          =   1215
      Left            =   7800
      TabIndex        =   7
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton Totalbutton 
      BackColor       =   &H0000FFFF&
      Caption         =   "Total Time"
      Enabled         =   0   'False
      Height          =   1215
      Left            =   240
      TabIndex        =   6
      Top             =   5280
      Width           =   1095
   End
   Begin VB.CommandButton Quitbutton 
      Caption         =   "Quit"
      Height          =   1695
      Left            =   7800
      TabIndex        =   5
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton Slowbutton 
      Caption         =   "Slowest Lap"
      Enabled         =   0   'False
      Height          =   1095
      Left            =   240
      TabIndex        =   4
      Top             =   3960
      Width           =   1095
   End
   Begin VB.CommandButton Bestbutton 
      Caption         =   "Best Lap"
      Enabled         =   0   'False
      Height          =   1095
      Left            =   240
      TabIndex        =   3
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton Avgbutton 
      Caption         =   "Average Lap"
      Enabled         =   0   'False
      Height          =   1095
      Left            =   240
      TabIndex        =   2
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton splitbutton 
      Caption         =   "Load Splits"
      Height          =   1095
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00FF8080&
      Height          =   6855
      Left            =   1560
      ScaleHeight     =   6795
      ScaleWidth      =   5235
      TabIndex        =   0
      Top             =   120
      Width           =   5295
   End
   Begin VB.Image Image1 
      Height          =   1125
      Left            =   6840
      Picture         =   "FiveKForm.frx":0000
      Top             =   720
      Width           =   3240
   End
End
Attribute VB_Name = "FiveKForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: TrackandFieldProgram (TrackProgram)'
'Form Name: FiveKForm (FiveKForm.frm)'
'Written By: Paul Jeske'
'Date Written: March 15th, 2004'
'Purpose of this form: Inputs a text file of the user's splits from a
                        'twelve lap, 5,000 meter race. The user can then use the
                        'loaded input to find their average lap, fastest lap,
                        'slowest lap, and total time.  The user then can
                        'move on to the next form or quit.


'Command that forces user to declare a variable as needed'
Option Explicit

'Dims these variables in the "General" portion so that they may be used on the within the
'entire form
Dim A As Integer
Dim Number(1 To 12) As Integer
Dim Laps(1 To 12) As Double

Private Sub Avgbutton_Click()
Sum = 0
A = 1

'Adds the 12 different laps together to equal the sum'
For A = 1 To 12
    Sum = Laps(A) + Sum
Next A
one = Sum / 12 'Takes sum of the 12 different laps and divides them by 12 (the number of laps)


picResults.Print "Your Average Lap was ==============>", Sum / 12; "seconds" 'Prints the user's average lap in
                                                                             'the output box
End Sub


Private Sub Bestbutton_Click()
'Dims variables'
Dim Temp As Single
Dim Pass As Single

A = 1 'sets A equal to 1"

'Sorts list of laps in a Descending order or from fastest laps to slowest laps'
For Pass = 1 To A - 1
    For A = 1 To A - Pass
        If Laps(A) > Laps(A + 1) Then
            Temp = Laps(A)
            Laps(A) = Laps(A + 1)
            Laps(A + 1) = Temp
        End If
    Next A
Next Pass

picResults.Print "Your Fastest Lap was ===============>", Laps(A) 'Prints the lowest value or the
                                                                  'fastest time in the output box'
End Sub
Private Sub nextbutton_Click()
'Hides the "FiveKForm" and shows the new "FiveKComparisonForm"'
FiveKComparisonForm.Show
FiveKForm.Hide

Close #1

End Sub

Private Sub Quitbutton_Click()
'Ends Program'
End
End Sub

Private Sub Slowbutton_Click()
'Dims needed variables'
Dim Temp As Single
Dim Pass As Single

A = 1 'sets A equal to one'

'Sorts the list of four numbers in ascending order or from fastest to slowest'
For Pass = 1 To 3
    For A = 1 To A - Pass
        If Laps(A) < Laps(A + 1) Then
            Temp = Laps(A)
            Laps(A) = Laps(A + 1)
            Laps(A + 1) = Temp
        End If
    Next A
Next Pass

picResults.Print "Your Slowest Lap was ===============>", Laps(4); "seconds"; 'Prints slot number four which is the biggest value or slowest lap'
End Sub
Private Sub splitbutton_Click()
'Prints a header for the data'
picResults.Print "LAP #", "SPLITS"
picResults.Print "-------------------------------------------------------------------------"

'Opens up data file'
Open Path & "5K.txt" For Input As #1

'Loads and fills the array'
For A = 1 To 12
    Input #1, Number(A), Laps(A)
    picResults.Print Number(A), Laps(A)
Next A

'Forces user to load arrays first, and then allows them to preform other actions'
Totalbutton.Enabled = True
Avgbutton.Enabled = True
Bestbutton.Enabled = True
Slowbutton.Enabled = True
Nextbutton.Enabled = True
splitbutton.Enabled = False

End Sub

Private Sub Totalbutton_Click()
'Dims needed variables'
Dim Total As Integer
Dim CTR As Single
Dim fivesum As Single

Sum = 0
Total = 0

'Adds all laps together, then formulates them into a suitable display format'    Sum = Laps(A) + Sum
For A = 1 To 12
    Sum = Laps(A) + Sum
    fivesum = Sum
Next A
fivesum = Sum
CTR = 0
Do While fivesum > 60
    fivesum = fivesum - 60
    CTR = CTR + 1
Loop
    
picResults.Print "Your Total Time was ================>"; CTR; ":"; fivesum 'outputs the total'
End Sub


