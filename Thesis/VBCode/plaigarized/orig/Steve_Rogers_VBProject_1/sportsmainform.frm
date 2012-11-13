VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H8000000D&
   Caption         =   "Form1"
   ClientHeight    =   11685
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   19080
   LinkTopic       =   "Form1"
   Picture         =   "sportsmainform.frx":0000
   ScaleHeight     =   11685
   ScaleWidth      =   19080
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdStart 
      BackColor       =   &H80000013&
      Caption         =   "Click here to start the game!"
      Height          =   375
      Left            =   15120
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   1680
      Width           =   2775
   End
   Begin VB.CommandButton cmdReset 
      BackColor       =   &H80000013&
      Caption         =   "Click here to start over"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6240
      MaskColor       =   &H00FFFF80&
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   10320
      Width           =   2535
   End
   Begin VB.CommandButton cmdCheck 
      BackColor       =   &H80000013&
      Caption         =   "Click here to check your score"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   10560
      MaskColor       =   &H00FFFF80&
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   10320
      Width           =   2535
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H80000013&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1920
      MaskColor       =   &H00FFFF80&
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   10320
      Width           =   2535
   End
   Begin VB.CommandButton cmdFinal 
      BackColor       =   &H80000013&
      Caption         =   "Click here for the Final Question!"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   14880
      MaskColor       =   &H00FFFF80&
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   10320
      Width           =   2535
   End
   Begin VB.CommandButton cmd500Foot 
      BackColor       =   &H8000000C&
      Caption         =   "500"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   14880
      MaskColor       =   &H00FFFF80&
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   8280
      Width           =   2535
   End
   Begin VB.CommandButton cmd400Foot 
      BackColor       =   &H8000000C&
      Caption         =   "400"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   14880
      MaskColor       =   &H00FFFF80&
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   6960
      Width           =   2535
   End
   Begin VB.CommandButton cmd300Foot 
      BackColor       =   &H8000000C&
      Caption         =   "300"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   14880
      MaskColor       =   &H00FFFF80&
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   5640
      Width           =   2535
   End
   Begin VB.CommandButton cmd200Foot 
      BackColor       =   &H8000000C&
      Caption         =   "200"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   14880
      MaskColor       =   &H00FFFF80&
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   4320
      Width           =   2535
   End
   Begin VB.CommandButton cmd500Soccer 
      BackColor       =   &H8000000C&
      Caption         =   "500"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   10560
      MaskColor       =   &H00FFFF80&
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   8280
      Width           =   2535
   End
   Begin VB.CommandButton cmd400Soccer 
      BackColor       =   &H8000000C&
      Caption         =   "400"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   10560
      MaskColor       =   &H00FFFF80&
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   6960
      Width           =   2535
   End
   Begin VB.CommandButton cmd300Soccer 
      BackColor       =   &H8000000C&
      Caption         =   "300"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   10560
      MaskColor       =   &H00FFFF80&
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   5640
      Width           =   2535
   End
   Begin VB.CommandButton cmd200Soccer 
      BackColor       =   &H8000000C&
      Caption         =   "200"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   10560
      MaskColor       =   &H00FFFF80&
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   4320
      Width           =   2535
   End
   Begin VB.CommandButton cmd500Base 
      BackColor       =   &H8000000C&
      Caption         =   "500"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6240
      MaskColor       =   &H00FFFF80&
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   8280
      Width           =   2535
   End
   Begin VB.CommandButton cmd400Base 
      BackColor       =   &H8000000C&
      Caption         =   "400"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6240
      MaskColor       =   &H00FFFF80&
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   6960
      Width           =   2535
   End
   Begin VB.CommandButton cmd300Base 
      BackColor       =   &H8000000C&
      Caption         =   "300"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6240
      MaskColor       =   &H00FFFF80&
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   5640
      Width           =   2535
   End
   Begin VB.CommandButton cmd200Base 
      BackColor       =   &H8000000C&
      Caption         =   "200"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6240
      MaskColor       =   &H00FFFF80&
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   4320
      Width           =   2535
   End
   Begin VB.CommandButton cmd100Foot 
      BackColor       =   &H8000000C&
      Caption         =   "100"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   14880
      MaskColor       =   &H00FFFF80&
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3120
      Width           =   2535
   End
   Begin VB.CommandButton cmd100Soccer 
      BackColor       =   &H8000000C&
      Caption         =   "100"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   10560
      MaskColor       =   &H00FFFF80&
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3000
      Width           =   2535
   End
   Begin VB.CommandButton cmd100Base 
      BackColor       =   &H8000000C&
      Caption         =   "100"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6240
      MaskColor       =   &H00FFFF80&
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3000
      Width           =   2535
   End
   Begin VB.CommandButton cmd500BBall 
      BackColor       =   &H8000000C&
      Caption         =   "500"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1920
      MaskColor       =   &H00FFFF80&
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   8280
      Width           =   2535
   End
   Begin VB.CommandButton cmd400BBall 
      BackColor       =   &H8000000C&
      Caption         =   "400"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1920
      MaskColor       =   &H00FFFF80&
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6960
      Width           =   2535
   End
   Begin VB.CommandButton cmd300BBall 
      BackColor       =   &H8000000C&
      Caption         =   "300"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1920
      MaskColor       =   &H00FFFF80&
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5640
      Width           =   2535
   End
   Begin VB.CommandButton cmd200BBall 
      BackColor       =   &H8000000C&
      Caption         =   "200"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1920
      MaskColor       =   &H00FFFF80&
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4320
      Width           =   2535
   End
   Begin VB.CommandButton cmd100BBall 
      BackColor       =   &H8000000C&
      Caption         =   "100"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1920
      MaskColor       =   &H00FFFF80&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3000
      Width           =   2535
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H80000017&
      Caption         =   "American Football"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   495
      Left            =   14280
      TabIndex        =   4
      Top             =   2160
      Width           =   3735
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H80000017&
      Caption         =   "Soccer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   495
      Left            =   9960
      TabIndex        =   3
      Top             =   2160
      Width           =   3735
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H80000017&
      Caption         =   "Baseball"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   495
      Left            =   5640
      TabIndex        =   2
      Top             =   2160
      Width           =   3735
   End
   Begin VB.Label lblBasketball 
      Alignment       =   2  'Center
      BackColor       =   &H80000017&
      Caption         =   "Basketball"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   495
      Left            =   1320
      TabIndex        =   1
      Top             =   2160
      Width           =   3735
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackColor       =   &H80000004&
      Caption         =   $"sportsmainform.frx":0C42
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   18615
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This is the main form for the project. It is the first form to show up, and most of the user interaction
'will be in this form.
Option Explicit
Private Sub cmd100Base_Click()          'baseball 100 pts
Dim Answer As String, Answer1 As String
Answer = InputBox("There were two 1995 World Series teams that were picketed by the American Indian Movement. Name one of them.")
If Answer = "atlanta braves" Then
        runningTotal = runningTotal + 100
        MsgBox ("That's right! You win 100 points! The other team was the Cleveland Indians.")
    ElseIf Answer = "cleveland indians" Then
        runningTotal = runningTotal + 100
        MsgBox (" That's right! You win 100 points! The other team was the Atlanta Braves.")
    Else: MsgBox ("I'm sorry, the correct answer was either the Cleveland Indians or the Atlanta Braves.")
End If
CTR = CTR + 1
cmd100Base.Enabled = False
If CTR = 20 Then
    cmdFinal.Enabled = True
End If
End Sub

Private Sub cmd100BBall_Click()         'basketball 100 pts
Dim Answer As String
Answer = InputBox("Who was the first hoopster to win eight NBA scoring titles? (Hint: He also played for UNC in college, and for the Bulls and Wizards during his pro career)")
If Answer = "michael jordan" Then
        runningTotal = runningTotal + 100
        MsgBox ("That's right! You win 100 points!")
    Else: MsgBox ("I'm sorry, the answer is Michael Jordan.")
End If
CTR = CTR + 1
cmd100BBall.Enabled = False
If CTR = 20 Then
    cmdFinal.Enabled = True
End If
End Sub

Private Sub cmd100Foot_Click()          'picture question football 100pts
frmfootPicture.Show
frmMain.Hide
cmd100Foot.Enabled = False
If CTR = 20 Then
    cmdFinal.Enabled = True
End If
End Sub

Private Sub cmd100Soccer_Click()        'soccer 100 pts
Dim Answer As String
Answer = InputBox("This sport is only called Soccer in the United States. What is the other name for it?")
If Answer = "futbol" Then
        runningTotal = runningTotal + 100
        MsgBox ("That's right! You win 100 points!")
    Else: MsgBox ("I'm sorry, the correct answer is Futbol.")
End If
CTR = CTR + 1
cmd100Soccer.Enabled = False
If CTR = 20 Then
    cmdFinal.Enabled = True
End If
End Sub

Private Sub cmd200Base_Click()         'baseball 200 pts
Dim Answer As Single
Answer = InputBox("What number position is the catcher (give the actual number)?")
If Answer = "2" Then
        runningTotal = runningTotal + 200
        MsgBox ("That's right! You win 200 points!")
    Else: MsgBox ("I'm sorry, the correct answer is 2.")
End If
CTR = CTR + 1
cmd200Base.Enabled = False
If CTR = 20 Then
    cmdFinal.Enabled = True
End If
End Sub

Private Sub cmd200BBall_Click()         'basketball 200 pts
Dim Answer As String
Answer = InputBox("What NBA team is know in China as 'the Red Oxen'?")
If Answer = "chicago bulls" Then
        runningTotal = runningTotal + 200
        MsgBox ("That's right! You win 200 points!")
    Else: MsgBox ("I'm sorry, the correct answer is Chicago Bulls.")
End If
CTR = CTR + 1
cmd200BBall.Enabled = False
If CTR = 20 Then
    cmdFinal.Enabled = True
End If
End Sub

Private Sub cmd200Foot_Click()          'football 200 pts
Dim Answer As String
Answer = InputBox("There are two leagues in the NFL. Name one of them (Give the abbreviation).")
If Answer = "afc" Then
        runningTotal = runningTotal + 200
        MsgBox ("That's right! You win 200 points!")
    ElseIf Answer = "nfc" Then
        runningTotal = runningTotal + 200
        MsgBox ("That's right! You win 200 points!")
    Else: MsgBox ("I'm sorry, the correct answer is either AFC or NFC.")
End If
CTR = CTR + 1
cmd200Foot.Enabled = False
If CTR = 20 Then
    cmdFinal.Enabled = True
End If
End Sub

Private Sub cmd200Soccer_Click()        'soccer 200 pts
Dim Answer As String
Answer = InputBox("What International Team does David Beckham play for?")
If Answer = "england" Then
        runningTotal = runningTotal + 200
        MsgBox ("That's right! You win 200 points!")
    Else: MsgBox ("I'm sorry, the correct answer is England.")
End If
CTR = CTR + 1
cmd200Soccer.Enabled = False
If CTR = 20 Then
    cmdFinal.Enabled = True
End If
End Sub

Private Sub cmd300Base_Click()          'baseball 300 pts
Dim Answer As String
Answer = InputBox("What player held the record for total career runs? (before Barry Bonds broke the record)")
If Answer = "hank aaron" Then
        runningTotal = runningTotal + 300
        MsgBox ("That's right! You win 300 points!")
    Else: MsgBox ("I'm sorry, the correct answer is Hank Aaron.")
End If
CTR = CTR + 1
cmd300Base.Enabled = False
If CTR = 20 Then
    cmdFinal.Enabled = True
End If
End Sub

Private Sub cmd300BBall_Click()         ' Special Bonus Question Basketball
Dim Answer As String, Wager As Single
Wager = InputBox("Please enter a wager that is less than or equal to your total. If you have less than 200 points, then the most you may wager is 200. Your points thus far are: " & runningTotal & ".")
If runningTotal <= 200 Then
    Do While Wager > 200
        Wager = InputBox("Since you have less than 200 points, the most you may wager is 200.")
    Loop
        Answer = InputBox("What NBA team plays home games at a facility nicknamed 'The O-rena'?")
                If Answer = "orlando magic" Then
                    runningTotal = runningTotal + Wager
                    MsgBox ("That's right! You win " & Wager & " points!")
                Else:
                    runningTotal = runningTotal - Wager
                    MsgBox ("I'm sorry, the answer is Orlando Magic.")
                End If
ElseIf runningTotal > 200 Then
            Do While Wager > runningTotal
                Wager = InputBox("Please enter a wager that is less than or equal to your total.")
            Loop
                Answer = InputBox("What NBA team plays home games at a facility nicknamed 'The O-rena'?")
                If Answer = "orlando magic" Then
                    runningTotal = runningTotal + Wager
                    MsgBox ("That's right! You win " & Wager & " points!")
                Else:
                    runningTotal = runningTotal - Wager
                    MsgBox ("I'm sorry, the answer is Orlando Magic.")
                End If
Else: Wager = InputBox("I'm sorry. Please enter a wager that is less than or equal to your total. If you have less than 200 points, then the most you may wager is 200.")
End If
CTR = CTR + 1
cmd300BBall.Enabled = False
If CTR = 20 Then
    cmdFinal.Enabled = True
End If
End Sub

Private Sub cmd300Foot_Click()          'football 300 pts
Dim Answer As String
Answer = InputBox("On the offensive part of a football team, who generally wears the least pads?")
If Answer = "quarterback" Then
        runningTotal = runningTotal + 300
        MsgBox ("That's right! You win 300 points!")
    ElseIf Answer = "the quarterback" Then
        runningTotal = runningTotal + 300
        MsgBox ("That's right! You win 300 points!")
    Else: MsgBox ("I'm sorry, the correct answer is the Quarterback.")
End If
CTR = CTR + 1
cmd300Foot.Enabled = False
End Sub

Private Sub cmd300Soccer_Click()        'soccer 300 pts
Dim Answer As String
Answer = InputBox("How much time is there between each World Cup?")
If Answer = "4 years" Then
        runningTotal = runningTotal + 300
        MsgBox ("That's right! You win 300 points!")
    ElseIf Answer = "four years" Then
        runningTotal = runningTotal + 300
        MsgBox ("That's right! You win 300 points!")
    Else: MsgBox ("I'm sorry, the correct answer is 4 years.")
End If
CTR = CTR + 1
cmd300Soccer.Enabled = False
If CTR = 20 Then
    cmdFinal.Enabled = True
End If
End Sub

Private Sub cmd400Base_Click()          'baseball 400 pts
Dim Answer As String
Answer = InputBox("Whose 1996 single to center, in his first regular season at-bat, was the first hit by a Red Sox pitcher in 24 years? (Hint: he was nicknamed 'The Rocket')")
If Answer = "roger clemens" Then
        runningTotal = runningTotal + 400
        MsgBox ("That's right! You win 400 points!")
    Else: MsgBox ("I'm sorry, the correct answer is Roger Clemens.")
End If
CTR = CTR + 1
cmd400Base.Enabled = False
If CTR = 20 Then
    cmdFinal.Enabled = True
End If
End Sub

Private Sub cmd400BBall_Click()         'basketball 400 pts
Dim Answer As String
Answer = InputBox("What Chicago Bulls coach has checked into hotels under the pseudonym 'Mr. Red Cloud'?")
If Answer = "phil jackson" Then
        runningTotal = runningTotal + 500
        MsgBox ("That's right! You win 500 points!")
    Else: MsgBox ("I'm sorry, the correct answer is Phil Jackson.")
End If
CTR = CTR + 1
cmd400BBall.Enabled = False
If CTR = 20 Then
    cmdFinal.Enabled = True
End If
End Sub

Private Sub cmd400Foot_Click()          'football 400 pts
Dim Answer As String
Answer = InputBox("Who was the 49ers wide reciever that played with the 49ers most of his career, played for the Oakland Raiders, then finished his career with the 49ers?")
If Answer = "jerry rice" Then
        runningTotal = runningTotal + 400
        MsgBox ("That's right! You win 400 points!")
    Else: MsgBox ("I'm sorry, the correct answer is Jerry Rice.")
End If
CTR = CTR + 1
cmd400Foot.Enabled = False
If CTR = 20 Then
    cmdFinal.Enabled = True
End If
End Sub

Private Sub cmd400Soccer_Click()        'soccer 400 pts
Dim Answer As String
Answer = InputBox("Where is the 2010 World Cup being held?")
If Answer = "south africa" Then
        runningTotal = runningTotal + 400
        MsgBox ("That's right! You win 400 points!")
    Else: MsgBox ("I'm sorry, the correct answer is South Africa.")
End If
CTR = CTR + 1
cmd400Soccer.Enabled = False
If CTR = 20 Then
    cmdFinal.Enabled = True
End If
End Sub

Private Sub cmd500Base_Click()          'baseball 500 pts
Dim Answer As String
Answer = InputBox("Which finger on a pitcher's throwing hand controls a curve ball and slider?")
If Answer = "middle finger" Then
        runningTotal = runningTotal + 500
        MsgBox ("That's right! You win 500 points!")
    ElseIf Answer = "middle" Then
        runningTotal = runningTotal + 500
        MsgBox ("That's right! You win 500 points!")
    Else: MsgBox ("I'm sorry, the correct answer is the Middle Finger.")
End If
cmd500Base.Enabled = False
If CTR = 20 Then
    cmdFinal.Enabled = True
End If
End Sub

Private Sub cmd500BBall_Click()         'picture question Basketball 500 pts
frmbballPicture.Show
frmMain.Hide
CTR = CTR + 1
cmd500BBall.Enabled = False
If CTR = 20 Then
    cmdFinal.Enabled = True
End If
End Sub

Private Sub cmd500Foot_Click()          'football 500 pts
Dim Answer As String
Answer = InputBox("Who is currently the commissioner of the NFL?")
If Answer = "roger goodell" Then
        runningTotal = runningTotal + 500
        MsgBox ("That's right! You win 500 points!")
    Else: MsgBox ("I'm sorry, the correct answer is Roger Goodell.")
End If
CTR = CTR + 1
cmd500Foot.Enabled = False
If CTR = 20 Then
    cmdFinal.Enabled = True
End If
End Sub

Private Sub cmd500Soccer_Click()        'special bonus question soccer
Dim Answer As String, Wager As Single
Wager = InputBox("Please enter a wager that is less than or equal to your total. If you have less than 200 points, then the most you may wager is 200. Your points thus far are: " & runningTotal & ".")
If runningTotal <= 200 Then
    Do While Wager > 200
        Wager = InputBox("Since you have less than 200 points, the most you may wager is 200.")
    Loop
        Answer = InputBox("There are 4 teams in a specific group for this year's World Cup. This group has been nicknamed the 'Group of Death'. 3 of the teams are Brazil, North Korea, Ivory Coast. What is the last team?")
                If Answer = "portugal" Then
                    runningTotal = runningTotal + Wager
                    MsgBox ("That's right! You win " & Wager & " points!")
                Else:
                    runningTotal = runningTotal - Wager
                    MsgBox ("I'm sorry, the answer is Portugal.")
                End If
ElseIf runningTotal > 200 Then
            Do While Wager > runningTotal
                Wager = InputBox("Please enter a wager that is less than or equal to your total.")
            Loop
                Answer = InputBox("There are 4 teams in a specific group for this year's World Cup. This group has been nicknamed the 'Group of Death'. 3 of the teams are Brazil, North Korea, Ivory Coast. What is the last team?")
                If Answer = "portugal" Then
                    runningTotal = runningTotal + Wager
                    MsgBox ("That's right! You win " & Wager & " points!")
                Else:
                    runningTotal = runningTotal - Wager
                    MsgBox ("I'm sorry, the answer is Portugal.")
                End If
Else: Wager = InputBox("I'm sorry. Please enter a wager that is less than or equal to your total. If you have less than 200 points, then the most you may wager is 200.")
End If
CTR = CTR + 1
cmd500Soccer.Enabled = False
If CTR = 20 Then
    cmdFinal.Enabled = True
End If
End Sub

Private Sub cmdCheck_Click()            'check total points
MsgBox ("You have " & runningTotal & " points so far!")
End Sub


Private Sub cmdFinal_Click()            'click to go to the final question form
frmFinalQuestion.Show
frmMain.Hide
End Sub

Private Sub cmdQuit_Click()             'quit
End
End Sub

Private Sub cmdReset_Click()            'start over
runningTotal = 0
cmd100BBall.Enabled = True
cmd200BBall.Enabled = True
cmd300BBall.Enabled = True
cmd400BBall.Enabled = True
cmd500BBall.Enabled = True
cmd100Base.Enabled = True
cmd200Base.Enabled = True
cmd300Base.Enabled = True
cmd400Base.Enabled = True
cmd500Base.Enabled = True
cmd100Soccer.Enabled = True
cmd200Soccer.Enabled = True
cmd300Soccer.Enabled = True
cmd400Soccer.Enabled = True
cmd500Soccer.Enabled = True
cmd100Foot.Enabled = True
cmd200Foot.Enabled = True
cmd300Foot.Enabled = True
cmd400Foot.Enabled = True
cmd500Foot.Enabled = True
cmdFinal.Enabled = False
CTR = 0
End Sub

Private Sub cmdStart_Click()            'click this button to start the game
cmd100BBall.Enabled = True
cmd200BBall.Enabled = True
cmd300BBall.Enabled = True
cmd400BBall.Enabled = True
cmd500BBall.Enabled = True
cmd100Base.Enabled = True
cmd200Base.Enabled = True
cmd300Base.Enabled = True
cmd400Base.Enabled = True
cmd500Base.Enabled = True
cmd100Soccer.Enabled = True
cmd200Soccer.Enabled = True
cmd300Soccer.Enabled = True
cmd400Soccer.Enabled = True
cmd500Soccer.Enabled = True
cmd100Foot.Enabled = True
cmd200Foot.Enabled = True
cmd300Foot.Enabled = True
cmd400Foot.Enabled = True
cmd500Foot.Enabled = True
cmdCheck.Enabled = True
cmdFinal.Enabled = False
cmdReset.Enabled = True
CTR = 0
runningTotal = 0

End Sub

