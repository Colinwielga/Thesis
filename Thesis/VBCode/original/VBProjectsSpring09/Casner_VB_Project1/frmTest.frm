VERSION 5.00
Begin VB.Form frmTest 
   BackColor       =   &H00000040&
   Caption         =   "T-Test Select"
   ClientHeight    =   1545
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   1545
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdPaired 
      Caption         =   "Paired-T"
      Height          =   615
      Left            =   3120
      TabIndex        =   2
      Top             =   720
      Width           =   1335
   End
   Begin VB.CommandButton cmdTwo 
      Caption         =   "Two Sample"
      Height          =   615
      Left            =   1680
      TabIndex        =   1
      Top             =   720
      Width           =   1335
   End
   Begin VB.CommandButton cmdOne 
      Caption         =   "One Sample"
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label lblTest 
      BackColor       =   &H00004080&
      Caption         =   "Which type of test do you wish to perform?"
      Height          =   255
      Left            =   720
      TabIndex        =   3
      Top             =   360
      Width           =   3015
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Visual Basic T-test
'Benjamin Casner
'March 13th, 2009
'frmStats
'This form allows the user to select what type of test to perform
Private Sub cmdOne_Click()
    'This button performs the one sample t-test on one of the samples
    'as selected by the user
    Dim selector As String
    'Hide the test selection form
    frmTest.Hide
    'selector is the variable the program uses to determine which
    'sample to test. Having a button would have required ANOTHER form
    selector = InputBox("Type x to test the primary sample and y for the secondary", "Sample selection")
    Select Case selector
        'perform test on first sample
        Case Is = x
            hyp = InputBox("What is the null hypothesis value for the mean?", "Null Hypothesis")
            t = (xBar - hyp) / ((Sx / ctr1) ^ 1 / 2)
            If Abs(t) < crit Then
                MsgBox ("t = " & FormatNumber(t, 2) & " is less than " & crit & " so we fail to reject the null hypothesis")
            Else
                MsgBox ("t = " & FormatNumber(t, 2) & " is greater than " & crit & " so we reject the null hypothesis")
            End If
        'perform test on secondary sample
        Case Is = y
            hyp = InputBox("What is the null hypothesis value for the mean?", "Null Hypothesis")
            t = (yBar - hyp) / ((Sy / ctr1) ^ 1 / 2)
            If Abs(t) < crit Then
                MsgBox ("t = " & FormatNumber(t, 2) & " is less than " & crit & " so we fail to reject the null hypothesis")
            Else
                MsgBox ("t = " & FormatNumber(t, 2) & " is greater than " & crit & " so we reject the null hypothesis")
            End If
        'if the user does not enter an x or a y, then the program
        'will display an error message
        Case Else
            MsgBox ("Error: Please enter either an x or a y")
    End Select
End Sub

Private Sub cmdPaired_Click()
    Dim Dbar As Single, Sd As Single, pos As Integer
    'Use this button if the data in the two samples are in some
    'way paired. Note that this test does not work if the data
    'sets are sorted by smallest to largest
    frmTest.Hide
    'Dbar is the mean of the differences between the paired elements
    'Sd is the sample variance
    'If the data has been sorted, then it needs to be loaded in
    'its original positions before the paired t will be valid
    
    Dbar = 0
    For pos = 1 To ctr1
        Dbar = Dbar + (Sample1(pos) - Sample2(pos))
    Next pos
    Dbar = Dbar / ctr1
    For pos = 1 To ctr1
        Sd = Sd + ((Sample1(pos) - Sample2(pos)) - Dbar) ^ 2
    Next pos
    Sd = Sd / (ctr1 - 1)
    t = Dbar / ((Sd / ctr1) ^ 1 / 2)
    If Abs(t) < crit Then
        MsgBox ("t = " & FormatNumber(t, 2) & " is less than " & crit & " so we fail to reject the null hypothesis")
    Else
        MsgBox ("t = " & FormatNumber(t, 2) & " is greater than " & crit & " so we reject the null hypothesis")
    End If

End Sub

Private Sub cmdTwo_Click()
    'This button performs a two sample t-test
    Dim Sp As Single
    frmTest.Hide
    'Sp is the pooled variance of the two samples
    Sp = ((((ctr1 - 1) * Sx) + ((ctr2 - 1) * Sy)) / (ctr1 + ctr2 - 2))
    t = (xBar - yBar) / ((Sp * ((1 / ctr1) + (1 / ctr2))) ^ (1 / 2))
    If Abs(t) < crit Then
        MsgBox ("t = " & FormatNumber(t, 2) & " is less than " & crit & " so we fail to reject the null hypothesis")
    Else
        MsgBox ("t = " & FormatNumber(t, 2) & " is greater than " & crit & " so we reject the null hypothesis")
    End If
End Sub

Private Sub Form_Load()
    frmDataDisplay.Hide
End Sub
