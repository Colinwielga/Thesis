VERSION 5.00
Begin VB.Form frmRunningBack 
   Caption         =   "Frm Running Back"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   FillColor       =   &H00FF0000&
   ForeColor       =   &H8000000D&
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture5 
      Height          =   3015
      Left            =   12120
      Picture         =   "frmRunningBack.frx":0000
      ScaleHeight     =   2955
      ScaleWidth      =   2235
      TabIndex        =   23
      Top             =   840
      Width           =   2295
   End
   Begin VB.PictureBox Picture4 
      Height          =   4455
      Left            =   11280
      Picture         =   "frmRunningBack.frx":16162
      ScaleHeight     =   4395
      ScaleWidth      =   3555
      TabIndex        =   22
      Top             =   4200
      Width           =   3615
   End
   Begin VB.PictureBox Picture3 
      Height          =   3495
      Left            =   720
      Picture         =   "frmRunningBack.frx":58054
      ScaleHeight     =   3435
      ScaleWidth      =   3195
      TabIndex        =   21
      Top             =   6600
      Width           =   3255
   End
   Begin VB.PictureBox Picture2 
      Height          =   4695
      Left            =   720
      Picture         =   "frmRunningBack.frx":7D18E
      ScaleHeight     =   4635
      ScaleWidth      =   3315
      TabIndex        =   20
      Top             =   600
      Width           =   3375
   End
   Begin VB.PictureBox picOutputBox2 
      Height          =   1455
      Left            =   7200
      ScaleHeight     =   1395
      ScaleWidth      =   2835
      TabIndex        =   17
      Top             =   8640
      Width           =   2895
   End
   Begin VB.CommandButton cmdShowIndividual 
      Caption         =   "Show Rankings"
      Height          =   495
      Left            =   4920
      TabIndex        =   16
      Top             =   9480
      Width           =   1935
   End
   Begin VB.TextBox txtIndividual 
      Height          =   615
      Left            =   4800
      TabIndex        =   15
      Top             =   8640
      Width           =   2175
   End
   Begin VB.CommandButton cmdTouchdowns 
      Caption         =   "Compute"
      Height          =   495
      Left            =   8880
      TabIndex        =   14
      Top             =   2880
      Width           =   1935
   End
   Begin VB.CommandButton cmdPower 
      Caption         =   "Compute"
      Height          =   495
      Left            =   6960
      TabIndex        =   13
      Top             =   2880
      Width           =   1695
   End
   Begin VB.CommandButton cmdSpeed 
      Caption         =   "Compute"
      Height          =   495
      Left            =   4920
      TabIndex        =   12
      Top             =   2880
      Width           =   1815
   End
   Begin VB.PictureBox picOutputBoxRB 
      Height          =   1815
      Left            =   6000
      ScaleHeight     =   1755
      ScaleWidth      =   3315
      TabIndex        =   11
      Top             =   4080
      Width           =   3375
   End
   Begin VB.PictureBox Picture1 
      Height          =   1935
      Left            =   4800
      Picture         =   "frmRunningBack.frx":AE550
      ScaleHeight     =   1875
      ScaleWidth      =   5835
      TabIndex        =   1
      Top             =   360
      Width           =   5895
   End
   Begin VB.CommandButton cmdReturnRunningBack 
      Caption         =   "Return to Main Menu"
      Height          =   1215
      Left            =   11520
      TabIndex        =   0
      Top             =   8880
      Width           =   2295
   End
   Begin VB.Label lblName 
      Caption         =   "Designer: John Cloeter"
      Height          =   255
      Left            =   9840
      TabIndex        =   24
      Top             =   10200
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Input Players Full Name to See Average of His Rankings"
      Height          =   495
      Left            =   4920
      TabIndex        =   19
      Top             =   8040
      Width           =   2295
   End
   Begin VB.Label lblPlayer 
      Caption         =   "Player Rankings and Overall Average"
      Height          =   375
      Left            =   7920
      TabIndex        =   18
      Top             =   8160
      Width           =   1575
   End
   Begin VB.Label lblBest 
      Caption         =   "Runningbacks from Best to Worst for:"
      Height          =   255
      Left            =   6360
      TabIndex        =   10
      Top             =   3720
      Width           =   2775
   End
   Begin VB.Label lbltd 
      Caption         =   "Touchdowns (5-1)"
      Height          =   255
      Left            =   9120
      TabIndex        =   9
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label lblPower 
      Caption         =   "Power (5-1)"
      Height          =   255
      Left            =   7320
      TabIndex        =   8
      Top             =   2520
      Width           =   855
   End
   Begin VB.Label lblSpeed 
      Caption         =   "Speed (5-1)"
      Height          =   255
      Left            =   5400
      TabIndex        =   7
      Top             =   2520
      Width           =   855
   End
   Begin VB.Label lblrb5 
      Caption         =   "Jerome Bettis"
      Height          =   255
      Left            =   6840
      TabIndex        =   6
      Top             =   7680
      Width           =   1935
   End
   Begin VB.Label lblrb4 
      Caption         =   "Tiki Barber"
      Height          =   495
      Left            =   6840
      TabIndex        =   5
      Top             =   7320
      Width           =   1935
   End
   Begin VB.Label lblrb3 
      Caption         =   "Clinton Portis"
      Height          =   495
      Left            =   6840
      TabIndex        =   4
      Top             =   6960
      Width           =   1935
   End
   Begin VB.Label lblrb2 
      Caption         =   "Shaun Alexander"
      Height          =   495
      Left            =   6840
      TabIndex        =   3
      Top             =   6600
      Width           =   1935
   End
   Begin VB.Label lblrb1 
      Caption         =   "LaDanian Tomlinson"
      Height          =   495
      Left            =   6840
      TabIndex        =   2
      Top             =   6240
      Width           =   1935
   End
End
Attribute VB_Name = "frmRunningBack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : FantasyFootball (Project1.vhp)
'Form Name : Frm Runningback (frmRunningback.frm)
'Author : John Cloeter
'Date : October 23, 2005
'Purpose of the Form : To take 5 runningbacks' names along with their speed, power, and touchdowns based on a scale of 1 to 5.  In this form, you can sort them from best to worst for each category, or type in a players name to see their average of their trait rankings.
Option Explicit
    Dim Player(1 To 5) As String
    Dim Numbers(1 To 5) As Single
    Dim Temp As Single
    Dim Guy As String
    Dim I As Integer
    Dim Pass As Integer
Private Sub cmdPower_Click() 'ranks runningbacks from best to worst for power from a file list and prints the results.
    Open App.Path & "\RunningbackPower.txt" For Input As #1
    For I = 1 To 5
        Input #1, Player(I), Numbers(I)
    Next I
    Close #1
    picOutputBoxRB.Cls
    picOutputBoxRB.Print "Power"
    picOutputBoxRB.Print "--------------------------------------------------------"
    For Pass = 1 To 5
        For I = 1 To 5 - Pass
            If Numbers(I) > Numbers(I + 1) Then
                Temp = Numbers(I)
                Numbers(I) = Numbers(I + 1)
                Numbers(I + 1) = Temp
                Guy = Player(I)
                Player(I) = Player(I + 1)
                Player(I + 1) = Guy
            End If
        Next I
        
        picOutputBoxRB.Print Numbers(I), Player(I)
    Next Pass
End Sub

Private Sub cmdReturnRunningBack_Click() 'sends user back to frmMain.
    frmRunningBack.Hide
    frmMain.Show
End Sub

Private Sub cmdShowIndividual_Click() 'takes players name from text box, and prints out their average trait ranking.
    Open App.Path & "\RunningBack.txt" For Input As #1
    Dim NotFound As Boolean
    Dim Power(1 To 5) As Integer
    Dim Touchdowns(1 To 5) As Integer
    Dim Average As Single
    Dim N As String
    picOutputBox2.Cls
    For I = 1 To 5
        Input #1, Player(I), Numbers(I), Power(I), Touchdowns(I)
    Next I
    N = txtIndividual.Text
    I = 0
    NotFound = True
    Do While I < 5 And NotFound = True
        I = I + 1
        If Player(I) = N Then
            NotFound = False
            Average = (Numbers(I) + Power(I) + Touchdowns(I)) / 3
        End If
    Loop
    If NotFound Then
            MsgBox "Sorry, but you must enter the correct name of your desired player.", , "Error"
        Else
            
            picOutputBox2.Print "Average Ranking ="; Average
        
    End If
    Close #1
End Sub

Private Sub cmdSpeed_Click() 'ranks runningbacks from best to worst for speed from a file list and prints the results.
    Open App.Path & "\RunningbackSpeed.txt" For Input As #1
    For I = 1 To 5
        Input #1, Player(I), Numbers(I)
    Next I
    Close #1
    picOutputBoxRB.Cls
    picOutputBoxRB.Print "Speed"
    picOutputBoxRB.Print "--------------------------------------------------------"
    For Pass = 1 To 5
        For I = 1 To 5 - Pass
            If Numbers(I) > Numbers(I + 1) Then
                Temp = Numbers(I)
                Numbers(I) = Numbers(I + 1)
                Numbers(I + 1) = Temp
                Guy = Player(I)
                Player(I) = Player(I + 1)
                Player(I + 1) = Guy
            End If
        Next I
        
        picOutputBoxRB.Print Numbers(I), Player(I)
    Next Pass
End Sub

Private Sub cmdTouchdowns_Click() 'ranks runningbacks from best to worst for touchdowns from a file list and prints the results.
    Open App.Path & "\RunningbackTouchdowns.txt" For Input As #1
    For I = 1 To 5
        Input #1, Player(I), Numbers(I)
    Next I
    Close #1
    picOutputBoxRB.Cls
    picOutputBoxRB.Print "Touchdowns"
    picOutputBoxRB.Print "--------------------------------------------------------"
    For Pass = 1 To 5
        For I = 1 To 5 - Pass
            If Numbers(I) > Numbers(I + 1) Then
                Temp = Numbers(I)
                Numbers(I) = Numbers(I + 1)
                Numbers(I + 1) = Temp
                Guy = Player(I)
                Player(I) = Player(I + 1)
                Player(I + 1) = Guy
            End If
        Next I
        
        picOutputBoxRB.Print Numbers(I), Player(I)
    Next Pass
End Sub
