VERSION 5.00
Begin VB.Form frmQuarterBack 
   Caption         =   "Frm Quarterback"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   FillColor       =   &H00FF0000&
   ForeColor       =   &H8000000D&
   LinkTopic       =   "Form1"
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture5 
      Height          =   4575
      Left            =   12000
      Picture         =   "frmQuarterBack.frx":0000
      ScaleHeight     =   4515
      ScaleWidth      =   2835
      TabIndex        =   23
      Top             =   4560
      Width           =   2895
   End
   Begin VB.PictureBox Picture4 
      Height          =   5775
      Left            =   240
      Picture         =   "frmQuarterBack.frx":2BF5A
      ScaleHeight     =   5715
      ScaleWidth      =   3555
      TabIndex        =   22
      Top             =   4680
      Width           =   3615
   End
   Begin VB.PictureBox Picture2 
      Height          =   3855
      Left            =   240
      Picture         =   "frmQuarterBack.frx":E8124
      ScaleHeight     =   3795
      ScaleWidth      =   2715
      TabIndex        =   21
      Top             =   480
      Width           =   2775
   End
   Begin VB.PictureBox Picture3 
      Height          =   3255
      Left            =   12360
      Picture         =   "frmQuarterBack.frx":1090BE
      ScaleHeight     =   3195
      ScaleWidth      =   2115
      TabIndex        =   20
      Top             =   840
      Width           =   2175
   End
   Begin VB.PictureBox picOutputBox2 
      Height          =   1455
      Left            =   7560
      ScaleHeight     =   1395
      ScaleWidth      =   2835
      TabIndex        =   18
      Top             =   8400
      Width           =   2895
   End
   Begin VB.CommandButton cmdShowIndividual 
      Caption         =   "Show Rankings"
      Height          =   495
      Left            =   5280
      TabIndex        =   17
      Top             =   9240
      Width           =   1935
   End
   Begin VB.TextBox txtIndividual 
      Height          =   615
      Left            =   5160
      TabIndex        =   16
      Top             =   8400
      Width           =   2175
   End
   Begin VB.CommandButton cmdNaturalInstincts 
      Caption         =   "Compute"
      Height          =   495
      Left            =   9120
      TabIndex        =   14
      Top             =   2880
      Width           =   1815
   End
   Begin VB.CommandButton cmdPassingAbility 
      Caption         =   "Compute"
      Height          =   495
      Left            =   6840
      TabIndex        =   13
      Top             =   2880
      Width           =   1815
   End
   Begin VB.CommandButton cmdRunningAbility 
      Caption         =   "Compute"
      Height          =   495
      Left            =   4560
      TabIndex        =   12
      Top             =   2880
      Width           =   1695
   End
   Begin VB.PictureBox picOutputBoxQB 
      Height          =   1455
      Left            =   6240
      ScaleHeight     =   1395
      ScaleWidth      =   3195
      TabIndex        =   11
      Top             =   3960
      Width           =   3255
   End
   Begin VB.PictureBox Picture1 
      Height          =   1455
      Left            =   4800
      Picture         =   "frmQuarterBack.frx":11FF2C
      ScaleHeight     =   1395
      ScaleWidth      =   5835
      TabIndex        =   1
      Top             =   480
      Width           =   5895
   End
   Begin VB.CommandButton cmdReturnQuarterback 
      Caption         =   "Return to Main Menu"
      Height          =   975
      Left            =   12240
      TabIndex        =   0
      Top             =   9360
      Width           =   1935
   End
   Begin VB.Label lblName 
      Caption         =   "Designer: John Cloeter"
      Height          =   255
      Left            =   9960
      TabIndex        =   24
      Top             =   10200
      Width           =   1695
   End
   Begin VB.Label lblPlayer 
      Caption         =   "Player Rankings and Overall Average"
      Height          =   375
      Left            =   8280
      TabIndex        =   19
      Top             =   7920
      Width           =   1575
   End
   Begin VB.Label lblIndividual 
      Caption         =   "Input Players Full Name to See Average of His Rankings"
      Height          =   495
      Left            =   5280
      TabIndex        =   15
      Top             =   7800
      Width           =   2295
   End
   Begin VB.Label lblBest 
      Caption         =   "Quarterback's from Best to Worst for :"
      Height          =   255
      Left            =   6480
      TabIndex        =   10
      Top             =   3600
      Width           =   2775
   End
   Begin VB.Label lblInstincts 
      Caption         =   "Natural Instincts (5-1)"
      Height          =   255
      Left            =   9240
      TabIndex        =   9
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Label lblPassing 
      Caption         =   "Passing Ability (5-1)"
      Height          =   255
      Left            =   6960
      TabIndex        =   8
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Label lblRunning 
      Caption         =   "Running Ability (5-1)"
      Height          =   255
      Left            =   4680
      TabIndex        =   7
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Label lblqb5 
      Caption         =   "Donovan McNabb"
      Height          =   255
      Left            =   6960
      TabIndex        =   6
      Top             =   7200
      Width           =   1575
   End
   Begin VB.Label lblqb4 
      Caption         =   "Brett Favre"
      Height          =   375
      Left            =   6960
      TabIndex        =   5
      Top             =   6840
      Width           =   1575
   End
   Begin VB.Label lblqb3 
      Caption         =   "Daunte Culpepper"
      Height          =   495
      Left            =   6960
      TabIndex        =   4
      Top             =   6480
      Width           =   1575
   End
   Begin VB.Label lblqb2 
      Caption         =   "Michael Vick"
      Height          =   495
      Left            =   6960
      TabIndex        =   3
      Top             =   6120
      Width           =   1575
   End
   Begin VB.Label lblqb1 
      Caption         =   " Payton Manning"
      Height          =   375
      Left            =   6960
      TabIndex        =   2
      Top             =   5760
      Width           =   1695
   End
End
Attribute VB_Name = "frmQuarterBack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : FantasyFootball (Project1.vhp)
'Form Name : Frm Quarterback (frmQuarterback.frm)
'Author : John Cloeter
'Date : October 23, 2005
'Purpose of the Form : To take 5 quaterbacks names along with their passing, running, and instincts based on a scale of 1 to 5.  In this form, you can sort them from best to worst for each category, or type in a players name to see their average of their trait rankings.
Option Explicit
    Dim Player(1 To 5) As String
    Dim Numbers(1 To 5) As Single
    Dim Temp As Single
    Dim Guy As String
    Dim I As Integer
    Dim Pass As Integer

Private Sub cmdNaturalInstincts_Click() 'ranks quarterbacks from best to worst for natural instincts from a file list and prints the results.
    Open App.Path & "\QuarterbackNaturalInstincts.txt" For Input As #1
    For I = 1 To 5
        Input #1, Player(I), Numbers(I)
    Next I
    Close #1
    picOutputBoxQB.Cls
    picOutputBoxQB.Print "Natural Instincts"
    picOutputBoxQB.Print "--------------------------------------------------------"
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
        
        picOutputBoxQB.Print Numbers(I), Player(I)
    Next Pass
End Sub

Private Sub cmdPassingAbility_Click() 'ranks quarterbacks from best to worst for passing ability from a file list and prints the results.
    Open App.Path & "\QuarterbackPassing.txt" For Input As #1
    For I = 1 To 5
        Input #1, Player(I), Numbers(I)
    Next I
    Close #1
    picOutputBoxQB.Cls
    picOutputBoxQB.Print "Passing Ability"
    picOutputBoxQB.Print "--------------------------------------------------------"
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
        
        picOutputBoxQB.Print Numbers(I), Player(I)
    Next Pass
End Sub

Private Sub cmdReturnQuarterback_Click() 'sends user back to frmMain
    frmQuarterBack.Hide
    frmMain.Show
End Sub

Private Sub cmdRunningAbility_Click() 'ranks quarterbacks from best to worst for running ability from a file list and prints the results.
    Open App.Path & "\QuarterbackSpeed.txt" For Input As #1
    For I = 1 To 5
        Input #1, Player(I), Numbers(I)
    Next I
    Close #1
    picOutputBoxQB.Cls
    picOutputBoxQB.Print "Running Abilities"
    picOutputBoxQB.Print "--------------------------------------------------------"
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
        
        picOutputBoxQB.Print Numbers(I), Player(I)
    Next Pass
End Sub

Private Sub cmdShowIndividual_Click() 'takes quarterbacks name from input box, and prints out the average of their trait points.
    Open App.Path & "\Quarterback.txt" For Input As #1
    Dim NotFound As Boolean
    Dim Passing(1 To 5) As Integer
    Dim Instincts(1 To 5) As Integer
    Dim Average As Single
    Dim N As String
    picOutputBox2.Cls
    For I = 1 To 5
        Input #1, Player(I), Numbers(I), Passing(I), Instincts(I)
    Next I
    N = txtIndividual.Text
    I = 0
    NotFound = True
    Do While I < 5 And NotFound = True
        I = I + 1
        If Player(I) = N Then
            NotFound = False
            Average = (Numbers(I) + Passing(I) + Instincts(I)) / 3
        End If
    Loop
    If NotFound Then
            MsgBox "Sorry, but you must enter the correct name of your desired player.", , "Error"
        Else
            
            picOutputBox2.Print "Average Ranking ="; Average
        
    End If
    Close #1
    
End Sub
