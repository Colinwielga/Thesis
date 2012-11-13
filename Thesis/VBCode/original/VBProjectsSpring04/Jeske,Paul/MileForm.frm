VERSION 5.00
Begin VB.Form MileForm 
   BackColor       =   &H00008000&
   Caption         =   "Mile Split Times"
   ClientHeight    =   7335
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10650
   FillColor       =   &H000000C0&
   LinkTopic       =   "Form1"
   ScaleHeight     =   7335
   ScaleWidth      =   10650
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton nextbutton 
      Caption         =   "Next"
      Enabled         =   0   'False
      Height          =   1215
      Left            =   7560
      TabIndex        =   7
      Top             =   4680
      Width           =   2175
   End
   Begin VB.CommandButton quitbutton 
      Caption         =   "Quit"
      Height          =   1935
      Left            =   7560
      TabIndex        =   6
      Top             =   2520
      Width           =   2055
   End
   Begin VB.CommandButton Totalbutton 
      Caption         =   "Total Time"
      Enabled         =   0   'False
      Height          =   855
      Left            =   360
      TabIndex        =   5
      Top             =   5160
      Width           =   1215
   End
   Begin VB.CommandButton Slowbutton 
      Caption         =   "Slowest Lap"
      Enabled         =   0   'False
      Height          =   975
      Left            =   360
      TabIndex        =   4
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton Bestbutton 
      Caption         =   "Best Lap"
      Enabled         =   0   'False
      Height          =   975
      Left            =   360
      TabIndex        =   3
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton Avgbutton 
      Caption         =   "Average Lap"
      Enabled         =   0   'False
      Height          =   1095
      Left            =   360
      TabIndex        =   2
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton Splitbutton 
      BackColor       =   &H008080FF&
      Caption         =   "Load Splits"
      Height          =   975
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H0000FFFF&
      Height          =   6255
      Left            =   1800
      ScaleHeight     =   6195
      ScaleWidth      =   4995
      TabIndex        =   0
      Top             =   240
      Width           =   5055
   End
   Begin VB.Image Image1 
      Height          =   1125
      Left            =   6960
      Picture         =   "MileForm.frx":0000
      Top             =   600
      Width           =   3240
   End
End
Attribute VB_Name = "MileForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: TrackandFieldProgram (TrackProgram)'
'Form Name: MileForm (MileForm.frm)'
'Written By: Paul Jeske'
'Date Written: March 15th, 2004'
'Purpose of this form:  Inputs a text file of the user's splits from a
                        'four lap, mile run race. The user can then use the
                        'loaded input to find their average lap, fastest lap,
                        'slowest lap, and total time.  The user then can
                        'move on to the next form or quit.

'Dims variables for use in the withing the entire form'
Dim A As Integer
Dim Number(1 To 4) As Single
Dim Laps(1 To 4) As Single

'Command that forces user to declare a variable as needed'
Option Explicit
Private Sub Avgbutton_Click()
'Dims needed variable'
Dim one As Single

A = 1 'Sets "A" equal to one"

'Adds the 4 different laps together to equal the sum'
For A = 1 To 4
    Sum = Laps(A) + Sum
Next A
one = Sum / 4 'Takes sum of the 12 different laps and divides them by 12 (the number of laps)'


picResults.Print "Your Average Lap was ============>", Sum / 4; "seconds" 'Prints the user's average lap in
                                                                          'the output box
End Sub

Private Sub Bestbutton_Click()
'Dims needed variables'
Dim Temp As Single
Dim Pass As Single

A = 1 'sets A equal to one'

'Sorts list of four laps in descending order or from fastest laps to slowest laps'
For Pass = 1 To 3
    For A = 1 To A - Pass
        If Laps(A) > Laps(A + 1) Then
            Temp = Laps(A)
            Laps(A) = Laps(A + 1)
            Laps(A + 1) = Temp
        End If
    Next A
Next Pass

picResults.Print "Your Fastest Lap was =============>", Laps(1); "seconds" 'Prints the lowest value or the
                                                                  'fastest time in the output box'
End Sub

Private Sub nextbutton_Click()
'Hides the "MileForm" and shows the new "MileComparrisionForm"'
MileComparrisonForm.Show
MileForm.Hide

End Sub

Private Sub Quitbutton_Click()
'Ends program'
End
End Sub

Private Sub Slowbutton_Click()
'Dims needed variables'
Dim Temp As Single
Dim Pass As Single


A = 1 'Sets "A" equal to one'

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

picResults.Print "Your Slowest Lap was =============>", Laps(4); "seconds" 'Prints the user's average lap in
                                                                           'the output box
End Sub

Private Sub splitbutton_Click()
'Prints a header for the data'
picResults.Print "LAP #", "SPLIT"
picResults.Print "-------------------------------------------------------------------------"

'Opens up data file'
Open Path & "1Mile.txt" For Input As #1

'Loads and fills the array'
For A = 1 To 4
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
Dim mysum As Single


Sum = 0
Total = 0

'Adds all laps together, then formulates them into a suitable display format'
For A = 1 To 4
    Sum = Laps(A) + Sum
Next A
mysum = Sum
CTR = 0
Do While mysum > 60
    mysum = mysum - 60
    CTR = CTR + 1
Loop
    
picResults.Print "Your Total Time was ==============>"; CTR; ":"; mysum 'outputs the total time'
End Sub

