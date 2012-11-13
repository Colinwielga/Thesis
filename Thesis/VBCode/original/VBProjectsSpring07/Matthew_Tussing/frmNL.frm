VERSION 5.00
Begin VB.Form frmNL 
   BackColor       =   &H000000C0&
   Caption         =   "National League"
   ClientHeight    =   7605
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9870
   LinkTopic       =   "Form1"
   ScaleHeight     =   7605
   ScaleWidth      =   9870
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAvg 
      Caption         =   "Find the top 5 NL batters combined average for the 2006 season"
      Height          =   1095
      Left            =   1800
      TabIndex        =   8
      Top             =   4560
      Width           =   3135
   End
   Begin VB.CommandButton cmdGoback 
      Caption         =   "Go Back To The Main Screen"
      Height          =   1455
      Left            =   3600
      TabIndex        =   7
      Top             =   6000
      Width           =   2535
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   1335
      Left            =   240
      TabIndex        =   6
      Top             =   6000
      Width           =   2655
   End
   Begin VB.PictureBox picResults2 
      Height          =   3495
      Left            =   6240
      ScaleHeight     =   3435
      ScaleWidth      =   3315
      TabIndex        =   5
      Top             =   3960
      Width           =   3375
   End
   Begin VB.PictureBox picResults 
      Height          =   3495
      Left            =   6240
      ScaleHeight     =   3435
      ScaleWidth      =   3315
      TabIndex        =   4
      Top             =   240
      Width           =   3375
   End
   Begin VB.CommandButton cmdClear2 
      Caption         =   "Clear the second pictrue box     (The bottom one)"
      Height          =   1695
      Left            =   3240
      TabIndex        =   3
      Top             =   2520
      Width           =   2655
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "      Clear the first picture box           (The top one)"
      Height          =   1815
      Left            =   3240
      TabIndex        =   2
      Top             =   360
      Width           =   2655
   End
   Begin VB.CommandButton cmdSort 
      Caption         =   "Sort the top 5 NL batting averages from highest to lowest in 2006"
      Height          =   1695
      Left            =   240
      TabIndex        =   1
      Top             =   2520
      Width           =   2655
   End
   Begin VB.CommandButton cmdRead 
      Caption         =   $"frmNL.frx":0000
      Height          =   1815
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   2655
   End
End
Attribute VB_Name = "frmNL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim names(1 To 100) As String
Dim bat(1 To 100) As Single
Dim ctr As Integer
Dim A As Integer

Private Sub cmdAvg_Click()
 Dim avg As Single
    Dim Sum As Single
    Dim A As Integer

    Sum = 0

    For A = 1 To ctr
        Sum = Sum + bat(A) 'adds up the avgs.
    Next A
        avg = Sum / ctr 'computes their combined avg.
        MsgBox ("The Top 5 NL Batters Combined Average Is " & FormatNumber(avg, 3))
        
End Sub

Private Sub cmdClear_Click()
'clears the firs picture box
    picResults.Cls
 End Sub

Private Sub cmdClear2_Click()
'clears the second picture box
    picResults2.Cls
End Sub

Private Sub cmdGoback_Click()
'brings you back to the main page
    frmProject.Show
    frmNL.Hide
End Sub

Private Sub cmdQuit_Click()
End 'ends the program
End Sub

Private Sub cmdRead_Click()
    'Reads and Prints the top five NL batting leaders
    Open App.Path & "\nlba.txt" For Input As #1
    ctr = 0
    Do Until EOF(1)
        ctr = ctr + 1
        Input #1, names(ctr), bat(ctr)
    Loop
    Close #1
       picResults.Print "Top 5 NL Batting Averages for 2006" 'prints this message
       picResults.Print "---------------------------------------------------------"
    
    For A = 1 To ctr
        picResults.Print names(A), bat(A) 'prints the top 5 averages and names
    Next A
End Sub

Private Sub cmdSort_Click()
Dim d As Integer, e As Integer, M As Integer
Dim Tempb As Single, Tempn As String

'This sorts the AL batting ave. from high to low
    For d = 1 To ctr - 1
        For e = 1 To ctr - d
            If bat(e) < bat(e + 1) Then
                Tempb = bat(e)
                bat(e) = bat(e + 1)
                bat(e + 1) = Tempb
                
                Tempn = names(e) 'sorts the names to stay with their averages
                names(e) = names(e + 1)
                names(e + 1) = Tempn
            End If
        Next e
    Next d
    picResults2.Print "Top 2006 NL Averages From High To Low" 'prints this message
    picResults2.Print "--------------------------------------------------------------------"
    
    For M = 1 To ctr
        picResults2.Print names(M), FormatNumber(bat(M), 3) 'prints the names and averages in order
    Next M
    
End Sub

