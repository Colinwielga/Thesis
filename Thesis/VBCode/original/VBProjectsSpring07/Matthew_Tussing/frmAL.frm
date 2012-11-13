VERSION 5.00
Begin VB.Form frmAL 
   BackColor       =   &H00800000&
   Caption         =   "American League"
   ClientHeight    =   7320
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9735
   LinkTopic       =   "Form1"
   ScaleHeight     =   7320
   ScaleWidth      =   9735
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAvg 
      Caption         =   "Find the average of the the top 5 players batting averages for the AL in 2006"
      Height          =   1095
      Left            =   1440
      TabIndex        =   8
      Top             =   4200
      Width           =   2775
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   1575
      Left            =   3120
      TabIndex        =   7
      Top             =   5520
      Width           =   2415
   End
   Begin VB.CommandButton cmdclear2 
      Caption         =   "Clear The Bottom Picture Box (Bottom One)"
      Height          =   1575
      Left            =   3240
      TabIndex        =   6
      Top             =   2280
      Width           =   2415
   End
   Begin VB.PictureBox picResults2 
      Height          =   3255
      Left            =   6000
      ScaleHeight     =   3195
      ScaleWidth      =   3435
      TabIndex        =   5
      Top             =   3960
      Width           =   3495
   End
   Begin VB.CommandButton cmdSorta 
      Caption         =   "Sort The Top Five AL Batting Leaders From High to Low"
      Height          =   1575
      Left            =   240
      TabIndex        =   4
      Top             =   2280
      Width           =   2415
   End
   Begin VB.PictureBox picResults 
      Height          =   3255
      Left            =   6000
      ScaleHeight     =   3195
      ScaleWidth      =   3435
      TabIndex        =   3
      Top             =   360
      Width           =   3495
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear The Top Picture Box    (Top One)"
      Height          =   1575
      Left            =   3120
      TabIndex        =   2
      Top             =   240
      Width           =   2415
   End
   Begin VB.CommandButton cmdGoback 
      Caption         =   "Go Back To Main Page"
      Height          =   1575
      Left            =   240
      TabIndex        =   1
      Top             =   5520
      Width           =   2415
   End
   Begin VB.CommandButton cmdALbat 
      Caption         =   "CLICK HERE FIRST TO LOAD                                                   The Top 5 AL Batting Leaders For The 2006 Season"
      Height          =   1575
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2415
   End
End
Attribute VB_Name = "frmAL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim names(1 To 100) As String
Dim bat(1 To 100) As Single
Dim ctr2 As Integer

Private Sub cmdALbat_Click()
Dim C As Integer

'Reads the top 5 AL batting averages
    Open App.Path & "\alba.txt" For Input As #2
    ctr2 = 0
    Do Until EOF(2)
        ctr2 = ctr2 + 1
        Input #2, names(ctr2), bat(ctr2)
    Loop
    Close #2
        'prints the top 5 AL bat avg.
        picResults.Print "The Top 5 AL Batting Averages"
        picResults.Print "--------------------------------------------------"
    For C = 1 To ctr2
        picResults.Print names(C), FormatNumber(bat(C), 3)
    Next C
    
End Sub

Private Sub cmdAvg_Click()
    Dim avg As Single
    Dim Sum As Single
    Dim A As Integer

    Sum = 0

    For A = 1 To ctr2
        Sum = Sum + bat(A) 'adds up the avgs.
    Next A
        avg = Sum / ctr2 'computes the combined avg.
        MsgBox ("The Top 5 AL Batters Combined Average Is " & FormatNumber(avg, 3))
        
End Sub

Private Sub cmdClear_Click()
    'this clears the picture box
    picResults.Cls
End Sub

Private Sub cmdClear2_Click()
    'clears the second picture box
    picResults2.Cls
End Sub

Private Sub cmdGoback_Click()
    'brings you back to the main page
    frmProject.Show
    frmAL.Hide
End Sub

Private Sub cmdQuit_Click()
End 'ends the program
End Sub

Private Sub cmdSorta_Click()
Dim d As Integer, e As Integer, M As Integer
Dim Tempb As Single, Tempn As String

'This sorts the AL batting ave. from high to low
    For d = 1 To ctr2 - 1
        For e = 1 To ctr2 - d
            If bat(e) < bat(e + 1) Then 'sorts the batting averages
                Tempb = bat(e)
                bat(e) = bat(e + 1)
                bat(e + 1) = Tempb
                
                Tempn = names(e) 'sorts the names keeping them with their averages
                names(e) = names(e + 1)
                names(e + 1) = Tempn
            End If
        Next e
    Next d
    picResults2.Print "Top 2006 AL Averages From High To Low" 'prints this message
    picResults2.Print "--------------------------------------------------------------------"
    
    For M = 1 To ctr2
        picResults2.Print names(M), FormatNumber(bat(M), 3) 'prints the sorted averages
    Next M
    
End Sub

