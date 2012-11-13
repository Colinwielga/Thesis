VERSION 5.00
Begin VB.Form frmStats 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000040&
   Caption         =   "T-Tester"
   ClientHeight    =   8025
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8040
   LinkTopic       =   "Form1"
   ScaleHeight     =   8025
   ScaleWidth      =   8040
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDesc 
      Caption         =   "Descriptive Statistics"
      Height          =   615
      Left            =   120
      TabIndex        =   10
      Top             =   2640
      Width           =   1575
   End
   Begin VB.TextBox txtCrit 
      Height          =   285
      Left            =   120
      TabIndex        =   9
      Top             =   4680
      Width           =   1575
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   615
      Left            =   120
      TabIndex        =   7
      Top             =   7080
      Width           =   1335
   End
   Begin VB.CommandButton cmdDisplay 
      Caption         =   "Display Data"
      Height          =   615
      Left            =   2040
      TabIndex        =   6
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear display"
      Height          =   615
      Left            =   6360
      TabIndex        =   5
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton cmdMore 
      Caption         =   "Input secondary data set"
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   1575
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "T-Test"
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   3600
      Width           =   1575
   End
   Begin VB.CommandButton cmdSort 
      Caption         =   "Sort data"
      Height          =   615
      Left            =   2040
      TabIndex        =   2
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton cmdAcquire 
      Caption         =   "Acquire primary data set"
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1575
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00C0C0C0&
      Height          =   6135
      Left            =   1800
      ScaleHeight     =   6075
      ScaleWidth      =   5955
      TabIndex        =   0
      Top             =   1560
      Width           =   6015
   End
   Begin VB.Label lblCrit 
      BackColor       =   &H00004080&
      Caption         =   "Critical Region ( |t|> x)"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   4320
      Width           =   1575
   End
End
Attribute VB_Name = "frmStats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'This program is designed to calculate basic descriptive statistics about
'a set of data and to perform a basic t statistic hypothesis test.
'Visual Basic T-test
'Benjamin Casner
'March 10th, 2009
'frmStats
'This is the primary interface form for the program, it is also where
'all data entry and display will take place
Private Sub cmdAcquire_Click()
    'This button will load the first data set from a file that the user names
    Dim FileName As String, pos As Integer
    ctr1 = 0
    'user enters the name of the file to be read
    FileName = InputBox("Enter the filename", "Data Set 1")
    Open App.Path & "\" & FileName For Input As #1
    Do Until EOF(1)
        ctr1 = ctr1 + 1
        Input #1, Sample1(ctr1)
    Loop
    Close #1
    'xBar is the sample mean for the data set
    xBar = 0
    'Sx is the sample variance for the data set
    Sx = 0
    'This for loop calculates the sum of all the data entries
    For pos = 1 To ctr1
        xBar = xBar + Sample1(pos)
    Next pos
    'we then divide by the number of entries to find the sample mean
    xBar = xBar / ctr1
    'we add the square of the difference of the individual
    'entries from the mean
    For pos = 1 To ctr1
        Sx = Sx + ((Sample1(pos) - xBar) ^ 2)
    Next pos
    'and then divide by the number of entries -1 to find the variance
    Sx = Sx / (ctr1 - 1)
End Sub

Private Sub cmdClear_Click()
'clears the result box
    picResults.Cls
End Sub

Private Sub cmdDesc_Click()
    'if there is only one sample, then this subroutine will
    'display the descriptive statistics for that sample
    If Compare = False Then
        picResults.Cls
        picResults.Print Tab(5); "Mean"; Tab(25); "Variance"; Tab(45); "Standard Deviation"
        picResults.Print Tab(5); "*************************************************************************"
        picResults.Print Tab(5); Round(xBar, 2); Tab(25); Round(Sx, 2); Tab(45); Round(Sx ^ (1 / 2), 2)
    'otherwise the program will ask which data set the user is inquiring about
    Else
        frmDesc.Show
    End If
End Sub

Private Sub cmdDisplay_Click()
    'This subroutine displays the data sets
    picResults.Cls
    'if there is more than one data set, then it will ask
    'which data set the user is inquiring about
    If Compare = True Then
        frmDataDisplay.Show
    'otherwise it will just display the data set
    Else
        For pos = 1 To ctr1
            picResults.Print pos, Tab(15); Sample1(pos)
        Next pos
    End If
End Sub

Private Sub cmdMore_Click()
    'This button lets the user read a second data set from a file
    Dim FileName As String, pos As Integer
    ctr2 = 0
    'compare is a boolean variable that tells the program whether it has
    'one or two data sets loaded
    'otherwise this button is almost exactly the same as cmdAcquire
    Compare = True
    'User enters the name of the file to be read
    FileName = InputBox("Enter the filename", "Data Set 2")
    Open App.Path & "\" & FileName For Input As #1
    Do Until EOF(1)
        ctr2 = ctr2 + 1
        Input #1, Sample2(ctr2)
    Loop
    Close #1
    yBar = 0
    Sy = 0
    For pos = 1 To ctr2
        yBar = yBar + Sample2(pos)
    Next pos
    yBar = yBar / ctr2
    For pos = 1 To ctr2
        Sy = Sy + ((Sample2(pos) - yBar) ^ 2)
    Next pos
    Sy = Sy / (ctr2 - 1)
End Sub

Private Sub cmdQuit_Click()
'ends the program
End
End Sub

Private Sub cmdSort_Click()
    'Sorts the data sets
    Dim temp As Single, pos As Integer, pass As Integer
    'Automatically sorts if there is only one data set
    If Compare = False Then
        For pass = 1 To ctr1 - 1
            For pos = 1 To ctr1 - pass
                If Sample1(pos) > Sample1(pos + 1) Then
                    temp = Sample1(pos + 1)
                    Sample1(pos + 1) = Sample1(pos)
                    Sample1(pos) = temp
                End If
            Next pos
        Next pass
    'otherwise it asks the user to select the data set
    Else
        frmSort.Show
    End If
    
End Sub

Private Sub cmdTest_Click()
    Dim hyp As Single
    'crit is the region outside of which the null hypothesis is rejected
    'The user will use a t table to select a critical region with the
    'alpha level (probability of rejecting null when it's true) that
    'the user is most comfortable with
    crit = txtCrit.Text
    'if there is only one data set then the program will perform a basic
    'one sample t-test
    If Compare = False Then
        'hyp is the null hypothesis for the test
        hyp = InputBox("What is the null hypothesis value for the mean?", "Null Hypothesis")
        'calculates the t-statistic
        t = (xBar - hyp) / ((Sx / ctr1) ^ 1 / 2)
        'if the t statistic falls outside the critical region then
        'fail to reject the null
        If Abs(t) < crit Then
            MsgBox ("t = " & FormatNumber(t, 2) & " is less than " & crit & " so we fail to reject the null hypothesis")
        'otherwise reject
        Else
            MsgBox ("t = " & FormatNumber(t, 2) & " is greater than " & crit & " so we reject the null hypothesis")
        End If
    'if there is more than one data set then the user may wish to
    'perform a different sort of t-test.
    Else
        frmTest.Show
    End If
End Sub

