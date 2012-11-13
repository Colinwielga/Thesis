VERSION 5.00
Begin VB.Form frmMillionare 
   BackColor       =   &H00000000&
   Caption         =   "Who wants to be a Millionaire..."
   ClientHeight    =   11010
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   13515
   LinkTopic       =   "Form1"
   ScaleHeight     =   11010
   ScaleWidth      =   13515
   Begin VB.PictureBox picMoney 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2295
      Left            =   120
      Picture         =   "Millionare.frx":0000
      ScaleHeight     =   2265
      ScaleWidth      =   2145
      TabIndex        =   31
      Top             =   7200
      Width           =   2175
   End
   Begin VB.PictureBox picCenter 
      Height          =   3375
      Left            =   2640
      Picture         =   "Millionare.frx":2243
      ScaleHeight     =   3315
      ScaleWidth      =   7035
      TabIndex        =   30
      Top             =   2160
      Width           =   7095
   End
   Begin VB.CommandButton cmdstart 
      BackColor       =   &H80000008&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   4080
      MaskColor       =   &H008080FF&
      TabIndex        =   28
      Top             =   5760
      Width           =   3855
   End
   Begin VB.TextBox txtanswer1 
      Height          =   495
      Left            =   360
      TabIndex        =   26
      Top             =   720
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtanswer 
      Height          =   495
      Left            =   360
      TabIndex        =   25
      Top             =   120
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdleave 
      Caption         =   "Leave with current winnings"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   10080
      TabIndex        =   23
      Top             =   8640
      Width           =   3255
   End
   Begin VB.CommandButton cmdd 
      Caption         =   "D"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   6840
      TabIndex        =   22
      Top             =   9480
      Width           =   2655
   End
   Begin VB.CommandButton cmdc 
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2760
      TabIndex        =   21
      Top             =   9480
      Width           =   2655
   End
   Begin VB.CommandButton cmdb 
      BackColor       =   &H8000000D&
      Caption         =   "B"
      DownPicture     =   "Millionare.frx":7D3F
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   6840
      MaskColor       =   &H0080FF80&
      Picture         =   "Millionare.frx":8A8E
      TabIndex        =   20
      Top             =   8280
      Width           =   2655
   End
   Begin VB.CommandButton cmda 
      BackColor       =   &H00C0FFC0&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2760
      TabIndex        =   19
      Top             =   8280
      Width           =   2655
   End
   Begin VB.CommandButton cmdaud 
      Caption         =   "Ask the Audience"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   480
      TabIndex        =   2
      Top             =   5280
      Width           =   1935
   End
   Begin VB.CommandButton cmdPhone 
      Caption         =   "Phone a friend"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   480
      TabIndex        =   1
      Top             =   3720
      Width           =   1935
   End
   Begin VB.CommandButton cmdfifty 
      BackColor       =   &H80000014&
      Caption         =   "50:50"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   480
      MaskColor       =   &H80000018&
      TabIndex        =   0
      Top             =   2160
      Width           =   1935
   End
   Begin VB.Label lblMarquee 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Who wants to be a Millionaire"
      BeginProperty Font 
         Name            =   "Palace Script MT"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   855
      Left            =   2640
      TabIndex        =   29
      Top             =   1440
      Width           =   6975
   End
   Begin VB.Label lblDesigner 
      BackColor       =   &H00000000&
      Caption         =   "Pradeep de Noronha"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   375
      Left            =   240
      TabIndex        =   27
      Top             =   10080
      Width           =   1695
   End
   Begin VB.Label lblheader 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   1215
      Left            =   2760
      TabIndex        =   24
      Top             =   240
      Width           =   7095
   End
   Begin VB.Label lblquestion 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   855
      Left            =   2400
      TabIndex        =   18
      Top             =   7200
      Width           =   7695
   End
   Begin VB.Label lblmillion 
      BackColor       =   &H0000C000&
      Caption         =   "$1,000,000"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   345
      Left            =   10080
      TabIndex        =   17
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label lbl500000 
      BackColor       =   &H0000C000&
      Caption         =   "$500,000"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   10080
      TabIndex        =   16
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label lbl250000 
      BackColor       =   &H0000C000&
      Caption         =   "$250,000"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   10080
      TabIndex        =   15
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label lbl100000 
      BackColor       =   &H0000C000&
      Caption         =   "$100,000"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   10080
      TabIndex        =   14
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Label lbl50000 
      BackColor       =   &H0000C000&
      Caption         =   "$50,000"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   10080
      TabIndex        =   13
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Label lbl25000 
      BackColor       =   &H0000C000&
      Caption         =   "$25,000"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   345
      Left            =   10080
      TabIndex        =   12
      Top             =   3240
      Width           =   1455
   End
   Begin VB.Label lbl16000 
      BackColor       =   &H0000C000&
      Caption         =   "$16,000"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   10080
      TabIndex        =   11
      Top             =   3600
      Width           =   1455
   End
   Begin VB.Label lbl8000 
      BackColor       =   &H0000C000&
      Caption         =   "$8,000"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   10080
      TabIndex        =   10
      Top             =   3960
      Width           =   1455
   End
   Begin VB.Label lbl4000 
      BackColor       =   &H0000C000&
      Caption         =   "$4,000"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   10080
      TabIndex        =   9
      Top             =   4320
      Width           =   1455
   End
   Begin VB.Label lbl2000 
      BackColor       =   &H0000C000&
      Caption         =   "$2,000"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   10080
      TabIndex        =   8
      Top             =   4680
      Width           =   1455
   End
   Begin VB.Label lbl1000 
      BackColor       =   &H0000C000&
      Caption         =   "$1,000"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   345
      Left            =   10080
      TabIndex        =   7
      Top             =   5040
      Width           =   1455
   End
   Begin VB.Label lbl500 
      BackColor       =   &H0000C000&
      Caption         =   "$500"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   10080
      TabIndex        =   6
      Top             =   5400
      Width           =   1455
   End
   Begin VB.Label lbl300 
      BackColor       =   &H0000C000&
      Caption         =   "$300"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   10080
      TabIndex        =   5
      Top             =   5760
      Width           =   1455
   End
   Begin VB.Label lbl200 
      BackColor       =   &H0000C000&
      Caption         =   "$200"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   10080
      TabIndex        =   4
      Top             =   6120
      Width           =   1455
   End
   Begin VB.Label lbl100 
      BackColor       =   &H0000C000&
      Caption         =   "$100"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   10080
      TabIndex        =   3
      Top             =   6480
      Width           =   1455
   End
End
Attribute VB_Name = "frmmillionare"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Who wants to be a Millionaire.(millionaire1.vbp)

'Form name: frmMillionaire; Form caption: Who wants to be a Millionaire...

'Author: Pradeep de Noronha

'Date written: 15th March, 2006

'Form Objective: This form displays the questions that are going to be asked,
'                the amount the questions are worth, the lifelines available to the user
'                and gives the user an option to leave with his or her winning at any
'                point during the game. Once the user gets all the questions right,
'                gets a question wrong or quits he or she is directed to the next form
'                (frmCheque) where their names and winning are displayed on a check.
                 

Dim counter, i, fiftycount, audcount, phonecount As Integer

Private Sub cmda_Click()
' If clicked txtanswer1.Text is set to A and compared with txtanswer.Text
' which subsiquintly decides whether to Call the "Right subroutine" or the
' "Wrong subroutine".
    
    txtanswer1.Text = "a"
    
        If txtanswer.Text = txtanswer1.Text Then
            Call Right
        Else
            Call Wrong
        End If
        
    Call Buttons1
    
End Sub

Private Sub cmdaud_Click()
' Asks the audience for assistance on a question. The audiences responses are stored in a
' text file. The answers are then transfered to an array and the appropriate answer is
' displayed. The command button is then disabled.

    audcount = audcount + 1
    Dim audience(1 To 15) As String
    'Dim ans As String
    Dim pos As Integer
    pos = 0
    Open App.Path & "\audienceresponses.txt" For Input As #1
    
        Do Until EOF(1)
            pos = pos + 1
            Input #1, audience(pos)
        Loop

        If counter = 1 Then
            frmaudience.Show
        Else
            lblheader.Caption = audience(i)
        End If

    cmdaud.Enabled = False
End Sub

Private Sub cmdb_Click()
' If clicked txtanswer1.Text is set to B and compared with txtanswer.Text
' which subsiquintly decides whether to Call the "Right subroutine" or the
' "Wrong subroutine".

    txtanswer1.Text = "b"
    
        If txtanswer.Text = txtanswer1.Text Then
            Call Right
        Else
            Call Wrong
        End If
        
    Call Buttons1
    
End Sub

Private Sub cmdc_Click()
' If clicked txtanswer1.Text is set to C and compared with txtanswer.Text
' which subsiquintly decides whether to Call the "Right subroutine" or the
' "Wrong subroutine".

    txtanswer1.Text = "c"
    
        If txtanswer.Text = txtanswer1.Text Then
            Call Right
        Else
            Call Wrong
        End If
        
    Call Buttons1
    
End Sub

Private Sub cmdd_Click()
' If clicked txtanswer1.Text is set to D and compared with txtanswer.Text
' which subsiquintly decides whether to Call the "Right subroutine" or the
' "Wrong subroutine".

    txtanswer1.Text = "d"
    
        If txtanswer.Text = txtanswer1.Text Then
            Call Right
        Else
            Call Wrong
        End If
        
    Call Buttons1
    
End Sub

Private Sub cmdfifty_Click()
'When clicked the program removes 2 of the wrong answers by setting Visibility
'equal to False. The command button is then disabled.

    fiftycount = fiftycount + 1
    lblheader.Caption = "Computer please take away 2 wrong answers.You now have a 50% chance of getting the answer right."

        Select Case counter
        
        Case 1
            cmda.Visible = False
            cmdc.Visible = False
        Case 2
            cmdb.Visible = False
            cmdc.Visible = False
        Case 3
            cmdd.Visible = False
            cmdc.Visible = False
        Case 4
            cmda.Visible = False
            cmdd.Visible = False
        Case 5
            cmdc.Visible = False
            cmdb.Visible = False
        Case 6
            cmda.Visible = False
            cmdb.Visible = False
        Case 7
            cmdd.Visible = False
            cmdc.Visible = False
        Case 8
            cmda.Visible = False
            cmdc.Visible = False
        Case 9
            cmdd.Visible = False
            cmda.Visible = False
        Case 10
            cmdb.Visible = False
            cmdc.Visible = False
        Case 11
            cmdb.Visible = False
            cmdd.Visible = False
        Case 12
            cmda.Visible = False
            cmdb.Visible = False
        Case 13
            cmdc.Visible = False
            cmdd.Visible = False
        Case 14
            cmda.Visible = False
            cmdb.Visible = False
        Case 15
            cmdb.Visible = False
            cmdc.Visible = False
            
        End Select
        
    cmdfifty.Enabled = False
    
End Sub

Private Sub cmdleave_Click()
    
    If counter = 1 Then
        amount = 0
        frmmillionare.Hide
        frmcheque.Show
    ElseIf counter = 2 Then
        amount = 100
        frmmillionare.Hide
        frmcheque.Show
    ElseIf counter = 3 Then
        amount = 200
        frmmillionare.Hide
        frmcheque.Show
    ElseIf counter = 4 Then
        amount = 300
        frmmillionare.Hide
        frmcheque.Show
    ElseIf counter = 5 Then
        amount = 500
        frmmillionare.Hide
        frmcheque.Show
    ElseIf counter = 6 Then
        amount = 1000
        frmmillionare.Hide
        frmcheque.Show
    ElseIf counter = 7 Then
        amount = 2000
        frmmillionare.Hide
        frmcheque.Show
    ElseIf counter = 8 Then
        amount = 4000
        frmmillionare.Hide
        frmcheque.Show
    ElseIf counter = 9 Then
        amount = 8000
        frmmillionare.Hide
        frmcheque.Show
    ElseIf counter = 10 Then
        amount = 16000
        frmmillionare.Hide
        frmcheque.Show
    ElseIf counter = 11 Then
        amount = 25000
        frmmillionare.Hide
        frmcheque.Show
    ElseIf counter = 12 Then
        amount = 50000
        frmmillionare.Hide
        frmcheque.Show
    ElseIf counter = 13 Then
        amount = 100000
        frmmillionare.Hide
        frmcheque.Show
    ElseIf counter = 14 Then
        amount = 250000
        frmmillionare.Hide
        frmcheque.Show
    ElseIf counter = 15 Then
        amount = 500000
        frmmillionare.Hide
        frmcheque.Show
End If
Call Buttons
End Sub



Private Sub cmdPhone_Click()
    phonecount = phonecount + 1
    Dim friendsname As String
    friendsname = InputBox("Please enter your friend's name", "Input")
     
        If counter = 1 Then
            lblheader.Caption = "Hi, " & friendsname & ". I have your friend " & username & " here, and he is stuck on the $100 question and he needs your help. So your friend thinks it is 'France'. He seems pretty sure."
        
        ElseIf counter = 2 Then
            lblheader.Caption = "Hi, " & friendsname & ". I have your friend " & username & " here, and he is stuck on the $200 question. So your friend thinks it is 'Your Back'. He seems pretty sure."
            
        ElseIf counter = 3 Then
            lblheader.Caption = "Hi, " & friendsname & ". I have your friend " & username & " here, he is stuck on the $300 question. So your friend thinks it is 'Santa Claus'. He seems pretty sure."
        
        ElseIf counter = 4 Then
            lblheader.Caption = "Hi, " & friendsname & ". I have your friend " & username & " here, he is stuck on the $500 question. So your friend thinks it is 'Home'. He seems pretty sure."
        
        ElseIf counter = 5 Then
            lblheader.Caption = "Hi, " & friendsname & ". I have your friend " & username & " here, he is stuck on the $1000 question. So your friend thinks it is 'Hempton'. He seems pretty sure."
        
        ElseIf counter = 6 Then
            lblheader.Caption = "Hi, " & friendsname & ". I have your friend " & username & " here, he is stuck on the $2000 question. So your friend think it is 'Sony'.He does not seems pretty sure."
        
        ElseIf counter = 7 Then
            lblheader.Caption = "Hi, " & friendsname & ". I have your friend " & username & " here, he is stuck on the $4000 question. So your friend is a 100% sure it is 'The Big Apple'."
        
        ElseIf counter = 8 Then
            lblheader.Caption = "Hi, " & friendsname & ". I have your friend " & username & " here, he is stuck on the $8000 question. So your friend is a 100% sure it is 'India'."
        
        ElseIf counter = 9 Then
            lblheader.Caption = "Hi, " & friendsname & ". I have your friend " & username & " here, he is stuck on the $16,000 question. So your friend thinks it's 'Irish'. But he isn't a 100% sure."
        
        ElseIf counter = 10 Then
            lblheader.Caption = "Hi, " & friendsname & ". I have your friend " & username & " here, he is stuck on the $25,000 question. So your friend is 75% sure that the answer is 'Cow'."
        
        ElseIf counter = 11 Then
            lblheader.Caption = "Hi, " & friendsname & ". I have your friend " & username & " here, he is stuck on the $50,000 question. So your friend is 95% sure it is 'Chlorine'."
        
        ElseIf counter = 12 Then
            lblheader.Caption = "Hi, " & friendsname & ". I have your friend " & username & " here, he is stuck on the $100,000 question. So your friend is pretty sure it is 'Polygamy'."
        
        ElseIf counter = 13 Then
            lblheader.Caption = "Hi, " & friendsname & ". I have your friend " & username & " here, he is stuck on the $250,000 question. So your friend is 90% sure it is '125'."
        
        ElseIf counter = 14 Then
            lblheader.Caption = "Hi, " & friendsname & ". I have your friend " & username & " here, he is stuck on the $500,000 question. So your friend is 90% sure it is 'Balloons'."
        
        ElseIf counter = 15 Then
            lblheader.Caption = "Hi, " & friendsname & ". I have your friend " & username & " here, he is stuck on the $1,000,000 question. So your friend is 60% sure it is a 'Liger'."
        
        End If

    cmdPhone.Enabled = False

End Sub

Private Sub cmdstart_Click()
    i = i + 1
        If (fiftycount >= 1) Then
            cmdfifty.Enabled = False
        Else
            cmdfifty.Enabled = True
        End If
        
        If (audcount >= 1) Then
            cmdaud.Enabled = False
        Else
            cmdaud.Enabled = True
        End If
        
        If (phonecount >= 1) Then
            cmdPhone.Enabled = False
        Else
            cmdPhone.Enabled = True
        End If
        
    cmdstart.Caption = "Next Question"
    cmdstart.Enabled = False
    counter = counter + 1
    lblheader.Caption = ""
    lblquestion.Visible = True
    cmda.Visible = True
    cmdb.Visible = True
    cmdc.Visible = True
    cmdd.Visible = True
    cmda.Enabled = True
    cmdb.Enabled = True
    cmdc.Enabled = True
    cmdd.Enabled = True
    cmdleave.Enabled = True
           
        
    If counter = 1 Then
        lblheader.Caption = "For $100 here is your first question."
    ElseIf counter = 2 Then
        lblheader.Caption = "For $200 here is your second question."
    ElseIf counter = 3 Then
        lblheader.Caption = "For $300 here is your third question."
    ElseIf counter = 4 Then
        lblheader.Caption = "For $500 here is your fourth question."
    ElseIf counter = 5 Then
        lblheader.Caption = "If you get this answer right the least your leaving with is a $1,000."
    ElseIf counter = 6 Then
        lblheader.Caption = "You are just nine questions away from a Million Dollars!"
    ElseIf counter = 7 Then
        lblheader.Caption = "For $4,000 heres your next question."
    ElseIf counter = 8 Then
        lblheader.Caption = "Keep it rolling. Remember you have lifelines to assist you."
    ElseIf counter = 9 Then
        lblheader.Caption = "Get the next 2 questions right and your walking away with atleast $25,000. "
    ElseIf counter = 10 Then
        lblheader.Caption = "You get this one you're $25,000 richer."
    ElseIf counter = 11 Then
        lblheader.Caption = "A free guess on this question. Here's your $50,000 question."
    ElseIf counter = 12 Then
        lblheader.Caption = "Big money. A $100,000 on the line."
    ElseIf counter = 13 Then
        lblheader.Caption = "Going for $250,000. If you get this question wrong you walk away with only $25,000 so weigh your odds."
    ElseIf counter = 14 Then
        lblheader.Caption = "$500,000. This is no joke. You have come so far and I would hate to see you leave without the million."
    ElseIf counter = 15 Then
        lblheader.Caption = "The final question. $1,000,000 on the line. Best of luck."
    End If

    If counter = 1 Then
        lblquestion.Caption = "Which country first created the metric system?"
        cmda.Caption = "A. England"
        cmdb.Caption = "B. Germany"
        cmdc.Caption = "C. Rome"
        cmdd.Caption = "D. France"
        txtanswer.Text = "d"
    
    ElseIf counter = 2 Then
        lblquestion.Caption = "Where in Your body would you find your spine?"
        cmda.Caption = "A. Your back"
        cmdb.Caption = "B. Your leg"
        cmdc.Caption = "C. Your nose"
        cmdd.Caption = "D. Your throat"
        txtanswer.Text = "a"
        
    ElseIf counter = 3 Then
        lblquestion.Caption = "Who delivers presents during Christmas?"
        cmda.Caption = "A. Santa Claw"
        cmdb.Caption = "B. Santa Claus"
        cmdc.Caption = "C. Santa Clone"
        cmdd.Caption = "D. Santa Cone"
        txtanswer.Text = "b"
        
    ElseIf counter = 4 Then
        lblquestion.Caption = "According to a common phrase, There is no place like..."
        cmda.Caption = "A. Heaven"
        cmdb.Caption = "B. Home"
        cmdc.Caption = "C. A candy store"
        cmdd.Caption = "D. Armageddon"
        txtanswer.Text = "b"
        
    ElseIf counter = 5 Then
        lblquestion.Caption = "Which is not a type of soda"
        cmda.Caption = "A. Hempton"
        cmdb.Caption = "B. Sprite"
        cmdc.Caption = "C. Coke"
        cmdd.Caption = "D. Pepsi"
        txtanswer.Text = "a"
        
    ElseIf counter = 6 Then
        lblquestion.Caption = "Which company make the Playstation game console? "
        cmda.Caption = "A. Sega"
        cmdb.Caption = "B. Nintendo"
        cmdc.Caption = "C. Microsoft"
        cmdd.Caption = "D. Sony"
        txtanswer.Text = "d"
        
    ElseIf counter = 7 Then
        lblquestion.Caption = "New York City is also known by what nickname?"
        cmda.Caption = "A. The Big Apple"
        cmdb.Caption = "B. The Big Easy"
        cmdc.Caption = "C. The Big Orange"
        cmdd.Caption = "D. The Big Traffic Jam"
        txtanswer.Text = "a"
        
    ElseIf counter = 8 Then
        lblquestion.Caption = "Where is the city of Calcutta located?"
        cmda.Caption = "A. Russia"
        cmdb.Caption = "B. India"
        cmdc.Caption = "C. Australia"
        cmdd.Caption = "D. Iraq"
        txtanswer.Text = "b"
        
    ElseIf counter = 9 Then
        lblquestion.Caption = "What nationality is celebrated on St. Patrick's Day?"
        cmda.Caption = "A. Mexican"
        cmdb.Caption = "B. Welsh"
        cmdc.Caption = "C. Irish"
        cmdd.Caption = "D. American Indian"
        txtanswer.Text = "c"
        
    ElseIf counter = 10 Then
        lblquestion.Caption = "Which of these is a Quadruped?"
        cmda.Caption = "A. Cow"
        cmdb.Caption = "B. Rectangle"
        cmdc.Caption = "C. Person over 50"
        cmdd.Caption = "D. Car"
        txtanswer.Text = "a"
        
    ElseIf counter = 11 Then
        lblquestion.Caption = "What chemical is added to the water in swimming pools?"
        cmda.Caption = "A. Chlorine"
        cmdb.Caption = "B. Sulfur"
        cmdc.Caption = "C. Salt"
        cmdd.Caption = "D. Ammonia"
        txtanswer.Text = "a"
        
    ElseIf counter = 12 Then
        lblquestion.Caption = "The act of having several spouses at one time is called what?"
        cmda.Caption = "A. Endogamy"
        cmdb.Caption = "B. Apogamy"
        cmdc.Caption = "C. Polygamy"
        cmdd.Caption = "D. Monogamy"
        txtanswer.Text = "c"
        
    ElseIf counter = 13 Then
        lblquestion.Caption = "What is 100 percent of 50 percent of 50 percent of 500?"
        cmda.Caption = "A. 125"
        cmdb.Caption = "B. 175"
        cmdc.Caption = "C. 200"
        cmdd.Caption = "D. 500"
        txtanswer.Text = "a"
        
    ElseIf counter = 14 Then
        lblquestion.Caption = "What is the leading cause of non-food choking deaths in children under age 14?"
        cmda.Caption = "A. Pen Caps"
        cmdb.Caption = "B. Coins"
        cmdc.Caption = "C. Buttons"
        cmdd.Caption = "D. Balloons"
        txtanswer.Text = "d"
        
    ElseIf counter = 15 Then
        lblquestion.Caption = "What is the offspring of a lion and a tigress called?"
        cmda.Caption = "A. Liger"
        cmdb.Caption = "B. Ligron"
        cmdc.Caption = "C. Tigon"
        cmdd.Caption = "D. Tigelon"
        txtanswer.Text = "a"
        
End If

End Sub

Private Sub Form_Load()
    i = -1
    amount = 0
    fiftycount = 0
    audcount = 0
    phonecount = 0
    lblheader.Caption = "Hi " & username & ". Welcome to Who Wants To Be a Millionaire."
    cmdstart.Caption = "First Question"
    Call Buttons
    cmdleave.Enabled = False
    
End Sub

Private Sub Right()

    MsgBox "Correct answer", , "Output"
    lblquestion.Caption = ""
    cmdfifty.Enabled = False
    cmdPhone.Enabled = False
    cmdaud.Enabled = False
    cmdleave.Enabled = False
    
        If counter = 1 Then
            lbl100.ForeColor = vbRed
            lbl100.BackColor = vbBlue
        ElseIf counter = 2 Then
            lbl200.ForeColor = vbRed
            lbl200.BackColor = vbBlue
        ElseIf counter = 3 Then
            lbl300.ForeColor = vbRed
            lbl300.BackColor = vbBlue
        ElseIf counter = 4 Then
            lbl500.ForeColor = vbRed
            lbl500.BackColor = vbBlue
        ElseIf counter = 5 Then
            lbl1000.ForeColor = vbRed
            lbl1000.BackColor = vbBlue
        ElseIf counter = 6 Then
            lbl2000.ForeColor = vbRed
            lbl2000.BackColor = vbBlue
        ElseIf counter = 7 Then
            lbl4000.ForeColor = vbRed
            lbl4000.BackColor = vbBlue
        ElseIf counter = 8 Then
            lbl8000.ForeColor = vbRed
            lbl8000.BackColor = vbBlue
        ElseIf counter = 9 Then
            lbl16000.ForeColor = vbRed
            lbl16000.BackColor = vbBlue
        ElseIf counter = 10 Then
            lbl25000.ForeColor = vbRed
            lbl25000.BackColor = vbBlue
        ElseIf counter = 11 Then
            lbl50000.ForeColor = vbRed
            lbl50000.BackColor = vbBlue
        ElseIf counter = 12 Then
            lbl100000.ForeColor = vbRed
            lbl100000.BackColor = vbBlue
        ElseIf counter = 13 Then
            lbl250000.ForeColor = vbRed
            lbl250000.BackColor = vbBlue
        ElseIf counter = 14 Then
            lbl500000.ForeColor = vbRed
            lbl500000.BackColor = vbBlue
        ElseIf counter = 15 Then
            lblmillion.ForeColor = vbRed
            lblmillion.BackColor = vbBlue
        End If
        
            If counter = 5 Then
                lblheader.Caption = "You have just guaranteed yourself $1000!!"
                
            ElseIf counter = 10 Then
                lblheader.Caption = "WOW! you are now guaranteed to leave here with at least $25,000"
                
            ElseIf counter = 15 Then
                MsgBox "Congratulations!!!You are now officially a MILLIONAIRE! Click OK to get your cheque"
                amount = 1000000
                frmmillionare.Hide
                frmcheque.Show
                                
            End If

End Sub

Private Sub Wrong()

    If counter = 1 Then
        MsgBox "Oh I am sorry but that that is the wrong answer. The right answer was: " & txtanswer.Text
        amount = 0
        frmmillionare.Hide
        frmcheque.Show
    End If
    
    If counter = 2 Then
        MsgBox "Oh I am sorry but that that is the wrong answer. The right answer was: " & txtanswer.Text
        amount = 0
        frmmillionare.Hide
        frmcheque.Show
    End If
    
    If counter = 3 Then
        MsgBox "Oh I am sorry but that that is the wrong answer. The right answer was: " & txtanswer.Text
        amount = 0
        frmmillionare.Hide
        frmcheque.Show
   End If
    
    If counter = 4 Then
        MsgBox "Oh I am sorry but that that is the wrong answer. The right answer was: " & txtanswer.Text
        amount = 0
        frmmillionare.Hide
        frmcheque.Show
    End If
    
    If counter = 5 Then
        MsgBox "Oh I am sorry but that that is the wrong answer. The right answer was: " & txtanswer.Text
        amount = 0
        frmmillionare.Hide
        frmcheque.Show
    End If
    
    If counter = 6 Then
        MsgBox "Oh I am sorry but that that is the wrong answer. The right answer was: " & txtanswer.Text
        amount = 1000
        frmmillionare.Hide
        frmcheque.Show
    End If
    
    If counter = 7 Then
        MsgBox "Oh I am sorry but that that is the wrong answer. The right answer was: " & txtanswer.Text
        amount = 1000
        frmmillionare.Hide
        frmcheque.Show
    End If
    
    If counter = 8 Then
        MsgBox "Oh I am sorry but that that is the wrong answer. The right answer was: " & txtanswer.Text
        amount = 1000
        frmmillionare.Hide
        frmcheque.Show
    End If
    
    If counter = 9 Then
        MsgBox "Oh I am sorry but that that is the wrong answer. The right answer was: " & txtanswer.Text
        amount = 1000
        frmmillionare.Hide
        frmcheque.Show
    End If
    
    If counter = 10 Then
        MsgBox "Oh I am sorry but that that is the wrong answer. The right answer was: " & txtanswer.Text
        amount = 1000
        frmmillionare.Hide
        frmcheque.Show
    End If
    
    If counter = 11 Then
        MsgBox "Oh I am sorry but that that is the wrong answer. The right answer was: " & txtanswer.Text
        amount = 25000
        frmmillionare.Hide
        frmcheque.Show
    End If
    
    If counter = 12 Then
        MsgBox "Oh I am sorry but that that is the wrong answer. The right answer was: " & txtanswer.Text
        amount = 25000
        frmmillionare.Hide
        frmcheque.Show
    End If
    
    If counter = 13 Then
        MsgBox "Oh I am sorry but that that is the wrong answer. The right answer was: " & txtanswer.Text
        amount = 25000
        frmmillionare.Hide
        frmcheque.Show
    End If
    
    If counter = 14 Then
        MsgBox "Oh I am sorry but that that is the wrong answer. The right answer was: " & txtanswer.Text
        amount = 25000
        frmmillionare.Hide
        frmcheque.Show
    End If
    
    If counter = 15 Then
        MsgBox "Oh I am sorry but that that is the wrong answer. The right answer was: " & txtanswer.Text
        amount = 25000
        frmmillionare.Hide
        frmcheque.Show
    End If
    
End Sub

Private Sub Buttons()
        
    cmdfifty.Enabled = False
    cmdPhone.Enabled = False
    cmdaud.Enabled = False
    cmda.Enabled = False
    cmdb.Enabled = False
    cmdc.Enabled = False
    cmdd.Enabled = False
        
End Sub

Private Sub Buttons1()
    
    cmda.Visible = False
    cmdb.Visible = False
    cmdc.Visible = False
    cmdd.Visible = False
    cmdstart.Enabled = True
    
End Sub




