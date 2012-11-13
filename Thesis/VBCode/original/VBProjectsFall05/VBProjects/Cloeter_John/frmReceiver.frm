VERSION 5.00
Begin VB.Form frmReceiver 
   Caption         =   "Frm Receiver"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   FillColor       =   &H00FF0000&
   ForeColor       =   &H8000000D&
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   Begin VB.PictureBox Picture5 
      Height          =   3735
      Left            =   12120
      Picture         =   "frmReceiver.frx":0000
      ScaleHeight     =   3675
      ScaleWidth      =   2595
      TabIndex        =   23
      Top             =   5280
      Width           =   2655
   End
   Begin VB.PictureBox Picture4 
      Height          =   4215
      Left            =   11760
      Picture         =   "frmReceiver.frx":261A2
      ScaleHeight     =   4155
      ScaleWidth      =   3195
      TabIndex        =   22
      Top             =   840
      Width           =   3255
   End
   Begin VB.PictureBox Picture3 
      Height          =   3855
      Left            =   960
      Picture         =   "frmReceiver.frx":502C8
      ScaleHeight     =   3795
      ScaleWidth      =   3315
      TabIndex        =   21
      Top             =   6120
      Width           =   3375
   End
   Begin VB.PictureBox Picture2 
      Height          =   3615
      Left            =   1200
      Picture         =   "frmReceiver.frx":7C806
      ScaleHeight     =   3555
      ScaleWidth      =   2715
      TabIndex        =   20
      Top             =   1320
      Width           =   2775
   End
   Begin VB.PictureBox picOutputBox2 
      Height          =   1455
      Left            =   8280
      ScaleHeight     =   1395
      ScaleWidth      =   2835
      TabIndex        =   17
      Top             =   8160
      Width           =   2895
   End
   Begin VB.CommandButton cmdShowIndividual 
      Caption         =   "Show Rankings"
      Height          =   495
      Left            =   6000
      TabIndex        =   16
      Top             =   9000
      Width           =   1935
   End
   Begin VB.TextBox txtIndividual 
      Height          =   615
      Left            =   5880
      TabIndex        =   15
      Top             =   8160
      Width           =   2175
   End
   Begin VB.CommandButton cmdTouchdowns 
      Caption         =   "Compute"
      Height          =   495
      Left            =   9480
      TabIndex        =   14
      Top             =   2880
      Width           =   1695
   End
   Begin VB.CommandButton cmdHands 
      Caption         =   "Compute"
      Height          =   495
      Left            =   7440
      TabIndex        =   13
      Top             =   2880
      Width           =   1815
   End
   Begin VB.CommandButton cmdSpeed 
      Caption         =   "Compute"
      Height          =   495
      Left            =   5400
      TabIndex        =   12
      Top             =   2880
      Width           =   1815
   End
   Begin VB.PictureBox picOutputBoxR 
      Height          =   1455
      Left            =   6720
      ScaleHeight     =   1395
      ScaleWidth      =   3195
      TabIndex        =   11
      Top             =   4080
      Width           =   3255
   End
   Begin VB.PictureBox Picture1 
      Height          =   1695
      Left            =   5160
      Picture         =   "frmReceiver.frx":9C288
      ScaleHeight     =   1635
      ScaleWidth      =   5715
      TabIndex        =   1
      Top             =   360
      Width           =   5775
   End
   Begin VB.CommandButton cmdReturnReceiver 
      Caption         =   "Return to Main Menu"
      Height          =   1095
      Left            =   12120
      TabIndex        =   0
      Top             =   9120
      Width           =   1935
   End
   Begin VB.Label lblName 
      Caption         =   "Designer: John Cloeter"
      Height          =   255
      Left            =   9960
      TabIndex        =   24
      Top             =   10080
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Input Players Full Name to See Average of His Rankings"
      Height          =   495
      Left            =   6000
      TabIndex        =   19
      Top             =   7680
      Width           =   2295
   End
   Begin VB.Label lblPlayer 
      Caption         =   "Player Rankings and Overall Average"
      Height          =   375
      Left            =   9000
      TabIndex        =   18
      Top             =   7560
      Width           =   1575
   End
   Begin VB.Label lblbest 
      Caption         =   "Receivers from Best to Worst"
      Height          =   255
      Left            =   7320
      TabIndex        =   10
      Top             =   3720
      Width           =   2175
   End
   Begin VB.Label lbltd 
      Caption         =   "Touchdowns (5-1)"
      Height          =   255
      Left            =   9720
      TabIndex        =   9
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label lblhands 
      Caption         =   "Hands(5-1)"
      Height          =   255
      Left            =   7920
      TabIndex        =   8
      Top             =   2520
      Width           =   855
   End
   Begin VB.Label lblspeed 
      Caption         =   "Speed (5-1)"
      Height          =   255
      Left            =   5880
      TabIndex        =   7
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label lblr5 
      Caption         =   "Hines Ward"
      Height          =   255
      Left            =   7680
      TabIndex        =   6
      Top             =   7320
      Width           =   1335
   End
   Begin VB.Label lblr4 
      Caption         =   "Roy Williams"
      Height          =   615
      Left            =   7680
      TabIndex        =   5
      Top             =   6960
      Width           =   1335
   End
   Begin VB.Label lblr3 
      Caption         =   "Chad Johnson"
      Height          =   615
      Left            =   7680
      TabIndex        =   4
      Top             =   6600
      Width           =   1335
   End
   Begin VB.Label lblr2 
      Caption         =   "Terrell Owens"
      Height          =   615
      Left            =   7680
      TabIndex        =   3
      Top             =   6240
      Width           =   1335
   End
   Begin VB.Label lblr1 
      Caption         =   "Randy Moss"
      Height          =   495
      Left            =   7680
      TabIndex        =   2
      Top             =   5880
      Width           =   1335
   End
End
Attribute VB_Name = "frmReceiver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : FantasyFootball (Project1.vhp)
'Form Name : Frm Receiver (frmReceiver.frm)
'Author : John Cloeter
'Date : October 23, 2005
'Purpose of the Form : To take 5 receivers' names along with their speed, hands, and touchdowns based on a scale of 1 to 5.  In this form, you can sort them from best to worst for each category, or type in a players name to see their average of their trait rankings.
Option Explicit
    Dim Player(1 To 5) As String
    Dim Numbers(1 To 5) As Single
    Dim Temp As Single
    Dim Guy As String
    Dim I As Integer
    Dim Pass As Integer
Private Sub cmdHands_Click() 'ranks receivers from best to worst for hands from a file list and prints the results.
    Open App.Path & "\ReceiverHands.txt" For Input As #1
    For I = 1 To 5
        Input #1, Player(I), Numbers(I)
    Next I
    Close #1
    picOutputBoxR.Cls
    picOutputBoxR.Print "Hands"
    picOutputBoxR.Print "--------------------------------------------------------"
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
        
        picOutputBoxR.Print Numbers(I), Player(I)
    Next Pass
End Sub

Private Sub cmdReturnReceiver_Click() 'sends user back to frmMain
    frmReceiver.Hide
    frmMain.Show
End Sub

Private Sub cmdShowIndividual_Click() 'takes the players name from the text box, and prints out that players average trait ranking.
    Open App.Path & "\Receiver.txt" For Input As #1
    Dim NotFound As Boolean
    Dim Hands(1 To 5) As Integer
    Dim Touchdowns(1 To 5) As Integer
    Dim Average As Single
    Dim N As String
    picOutputBox2.Cls
    For I = 1 To 5
        Input #1, Player(I), Numbers(I), Hands(I), Touchdowns(I)
       Next I
    N = txtIndividual.Text
    I = 0
    NotFound = True
    Do While I < 5 And NotFound = True
        I = I + 1
        If Player(I) = N Then
            NotFound = False
            Average = (Numbers(I) + Hands(I) + Touchdowns(I)) / 3
        End If
    Loop
    If NotFound Then
            MsgBox "Sorry, but you must enter the correct name of your desired player.", , "Error"
        Else
            
            picOutputBox2.Print "Average Ranking ="; Average
        
    End If
    Close #1
End Sub

Private Sub cmdSpeed_Click() 'ranks receivers from best to worst for speed from a file list and prints the results.
    Open App.Path & "\ReceiverSpeed.txt" For Input As #1
    For I = 1 To 5
        Input #1, Player(I), Numbers(I)
    Next I
    Close #1
    picOutputBoxR.Cls
    picOutputBoxR.Print "Speed"
    picOutputBoxR.Print "--------------------------------------------------------"
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
        
        picOutputBoxR.Print Numbers(I), Player(I)
    Next Pass
End Sub

Private Sub cmdTouchdowns_Click() 'ranks receivers from best to worst for touchdowns from a file list and prints the results.
    Open App.Path & "\ReceiverTouchdown.txt" For Input As #1
    For I = 1 To 5
        Input #1, Player(I), Numbers(I)
    Next I
    Close #1
    picOutputBoxR.Cls
    picOutputBoxR.Print "Touchdowns"
    picOutputBoxR.Print "--------------------------------------------------------"
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
        
        picOutputBoxR.Print Numbers(I), Player(I)
    Next Pass
End Sub

