VERSION 5.00
Begin VB.Form frmgamescreen 
   BackColor       =   &H00000000&
   Caption         =   "Form1"
   ClientHeight    =   8550
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11505
   LinkTopic       =   "Form1"
   ScaleHeight     =   8550
   ScaleWidth      =   11505
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FFFF&
      Caption         =   "quit"
      Height          =   735
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   56
      Top             =   7560
      Width           =   1575
   End
   Begin VB.TextBox txt1000000 
      BackColor       =   &H0000FFFF&
      Height          =   285
      Left            =   9600
      TabIndex        =   54
      Text            =   "1000000"
      Top             =   4080
      Width           =   2295
   End
   Begin VB.TextBox txt750000 
      BackColor       =   &H0000FFFF&
      Height          =   285
      Left            =   9600
      TabIndex        =   53
      Text            =   "750000"
      Top             =   3720
      Width           =   2295
   End
   Begin VB.TextBox txt500000 
      BackColor       =   &H0000FFFF&
      Height          =   285
      Left            =   9600
      TabIndex        =   52
      Text            =   "500000"
      Top             =   3360
      Width           =   2295
   End
   Begin VB.TextBox txt400000 
      BackColor       =   &H0000FFFF&
      Height          =   285
      Left            =   9600
      TabIndex        =   51
      Text            =   "400000"
      Top             =   3000
      Width           =   2295
   End
   Begin VB.TextBox txt300000 
      BackColor       =   &H0000FFFF&
      Height          =   285
      Left            =   9600
      TabIndex        =   50
      Text            =   "300000"
      Top             =   2640
      Width           =   2295
   End
   Begin VB.TextBox txt200000 
      BackColor       =   &H0000FFFF&
      Height          =   285
      Left            =   9600
      TabIndex        =   49
      Text            =   "200000"
      Top             =   2280
      Width           =   2295
   End
   Begin VB.TextBox txt100000 
      BackColor       =   &H0000FFFF&
      Height          =   285
      Left            =   9600
      TabIndex        =   48
      Text            =   "100000"
      Top             =   1920
      Width           =   2295
   End
   Begin VB.TextBox txt75000 
      BackColor       =   &H0000FFFF&
      Height          =   285
      Left            =   9600
      TabIndex        =   47
      Text            =   "75000"
      Top             =   1560
      Width           =   2295
   End
   Begin VB.TextBox txt50000 
      BackColor       =   &H0000FFFF&
      Height          =   285
      Left            =   9600
      TabIndex        =   46
      Text            =   "50000"
      Top             =   1200
      Width           =   2295
   End
   Begin VB.TextBox txt25000 
      BackColor       =   &H0000FFFF&
      Height          =   285
      Left            =   9600
      TabIndex        =   45
      Text            =   "25000"
      Top             =   840
      Width           =   2295
   End
   Begin VB.TextBox txt10000 
      BackColor       =   &H0000FFFF&
      Height          =   285
      Left            =   9600
      TabIndex        =   44
      Text            =   "10000"
      Top             =   480
      Width           =   2295
   End
   Begin VB.TextBox txt5000 
      BackColor       =   &H0000FFFF&
      Height          =   285
      Left            =   9600
      TabIndex        =   43
      Text            =   "5000"
      Top             =   120
      Width           =   2295
   End
   Begin VB.TextBox txt1000 
      BackColor       =   &H0000FFFF&
      Height          =   285
      Left            =   120
      TabIndex        =   42
      Text            =   "1000"
      Top             =   4440
      Width           =   2055
   End
   Begin VB.TextBox txt750 
      BackColor       =   &H0000FFFF&
      Height          =   285
      Left            =   120
      TabIndex        =   41
      Text            =   "750"
      Top             =   4080
      Width           =   2055
   End
   Begin VB.TextBox txt500 
      BackColor       =   &H0000FFFF&
      Height          =   285
      Left            =   120
      TabIndex        =   40
      Text            =   "500"
      Top             =   3720
      Width           =   2055
   End
   Begin VB.TextBox txt400 
      BackColor       =   &H0000FFFF&
      Height          =   285
      Left            =   120
      TabIndex        =   39
      Text            =   "400"
      Top             =   3360
      Width           =   2055
   End
   Begin VB.TextBox txt300 
      BackColor       =   &H0000FFFF&
      Height          =   285
      Left            =   120
      TabIndex        =   38
      Text            =   "300"
      Top             =   3000
      Width           =   2055
   End
   Begin VB.TextBox txt200 
      BackColor       =   &H0000FFFF&
      Height          =   285
      Left            =   120
      TabIndex        =   37
      Text            =   "200"
      Top             =   2640
      Width           =   2055
   End
   Begin VB.TextBox txt100 
      BackColor       =   &H0000FFFF&
      Height          =   285
      Left            =   120
      TabIndex        =   36
      Text            =   "100"
      Top             =   2280
      Width           =   2055
   End
   Begin VB.TextBox txt75 
      BackColor       =   &H0000FFFF&
      Height          =   285
      Left            =   120
      TabIndex        =   35
      Text            =   "75"
      Top             =   1920
      Width           =   2055
   End
   Begin VB.TextBox txt50 
      BackColor       =   &H0000FFFF&
      Height          =   285
      Left            =   120
      TabIndex        =   34
      Text            =   "50"
      Top             =   1560
      Width           =   2055
   End
   Begin VB.TextBox txt25 
      BackColor       =   &H0000FFFF&
      Height          =   285
      Left            =   120
      TabIndex        =   33
      Text            =   "25"
      Top             =   1200
      Width           =   2055
   End
   Begin VB.TextBox txt10 
      BackColor       =   &H0000FFFF&
      Height          =   285
      Left            =   120
      TabIndex        =   32
      Text            =   "10"
      Top             =   840
      Width           =   2055
   End
   Begin VB.TextBox txt5 
      BackColor       =   &H0000FFFF&
      Height          =   285
      Left            =   120
      TabIndex        =   31
      Text            =   "5"
      Top             =   480
      Width           =   2055
   End
   Begin VB.TextBox txt1 
      BackColor       =   &H0000FFFF&
      Height          =   285
      Left            =   120
      TabIndex        =   30
      Text            =   "1"
      Top             =   120
      Width           =   2055
   End
   Begin VB.TextBox txtpickcase 
      BackColor       =   &H0000FFFF&
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3960
      TabIndex        =   29
      Text            =   "Pick another case!!"
      Top             =   6240
      Width           =   3255
   End
   Begin VB.TextBox txtpickfirst 
      BackColor       =   &H0000FFFF&
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1800
      TabIndex        =   28
      Text            =   "Pick YOUR suitcase, which you'll keep all game!"
      Top             =   5160
      Width           =   7815
   End
   Begin VB.CommandButton cmdNODEAL 
      BackColor       =   &H000000FF&
      Caption         =   "NO DEAL"
      Height          =   495
      Left            =   10200
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   6360
      Width           =   735
   End
   Begin VB.PictureBox picoffer 
      BackColor       =   &H0000FFFF&
      Height          =   975
      Left            =   9000
      ScaleHeight     =   915
      ScaleWidth      =   1995
      TabIndex        =   26
      Top             =   7320
      Width           =   2055
   End
   Begin VB.CommandButton cmdDEAL 
      BackColor       =   &H0000FF00&
      Caption         =   "DEAL"
      Height          =   495
      Left            =   9000
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   6360
      Width           =   855
   End
   Begin VB.CommandButton cmdCase25 
      Caption         =   "Case 25"
      Height          =   735
      Left            =   7440
      TabIndex        =   24
      Top             =   4200
      Width           =   975
   End
   Begin VB.CommandButton cmdCase24 
      Caption         =   "Case 24"
      Height          =   735
      Left            =   6360
      TabIndex        =   23
      Top             =   4200
      Width           =   975
   End
   Begin VB.CommandButton cmdCase23 
      Caption         =   "Case 23"
      Height          =   735
      Left            =   5280
      TabIndex        =   22
      Top             =   4200
      Width           =   975
   End
   Begin VB.CommandButton cmdCase22 
      Caption         =   "Case 22"
      Height          =   735
      Left            =   4200
      TabIndex        =   21
      Top             =   4200
      Width           =   975
   End
   Begin VB.CommandButton cmdCase21 
      Caption         =   "Case 21"
      Height          =   735
      Left            =   3120
      TabIndex        =   20
      Top             =   4200
      Width           =   975
   End
   Begin VB.CommandButton cmdCase20 
      Caption         =   "Case 20"
      Height          =   735
      Left            =   7440
      TabIndex        =   19
      Top             =   3360
      Width           =   975
   End
   Begin VB.CommandButton cmdCase19 
      Caption         =   "Case 19"
      Height          =   735
      Left            =   6360
      TabIndex        =   18
      Top             =   3360
      Width           =   975
   End
   Begin VB.CommandButton cmdCase15 
      Caption         =   "Case 15"
      Height          =   735
      Left            =   7440
      TabIndex        =   17
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton cmdCase14 
      Caption         =   "Case 14"
      Height          =   735
      Left            =   6360
      TabIndex        =   16
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton cmdCase13 
      Caption         =   "Case 13"
      Height          =   735
      Left            =   5280
      TabIndex        =   15
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton cmdCase12 
      Caption         =   "Case 12"
      Height          =   735
      Left            =   4200
      TabIndex        =   14
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton cmdCase18 
      Caption         =   "Case 18"
      Height          =   735
      Left            =   5280
      TabIndex        =   13
      Top             =   3360
      Width           =   975
   End
   Begin VB.CommandButton cmdCase17 
      Caption         =   "Case 17"
      Height          =   735
      Left            =   4200
      TabIndex        =   12
      Top             =   3360
      Width           =   975
   End
   Begin VB.CommandButton cmdCase16 
      Caption         =   "Case 16"
      Height          =   735
      Left            =   3120
      TabIndex        =   11
      Top             =   3360
      Width           =   975
   End
   Begin VB.CommandButton cmdCase11 
      Caption         =   "Case 11"
      Height          =   735
      Left            =   3120
      TabIndex        =   10
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton cmdCase1 
      Caption         =   "Case 1"
      Height          =   735
      Left            =   3120
      TabIndex        =   9
      Top             =   840
      Width           =   975
   End
   Begin VB.CommandButton cmdCase5 
      Caption         =   "Case 5"
      Height          =   735
      Left            =   7440
      TabIndex        =   8
      Top             =   840
      Width           =   975
   End
   Begin VB.CommandButton cmdCase4 
      Caption         =   "Case 4"
      Height          =   735
      Left            =   6360
      TabIndex        =   7
      Top             =   840
      Width           =   975
   End
   Begin VB.CommandButton cmdCase10 
      Caption         =   "Case 10"
      Height          =   735
      Left            =   7440
      TabIndex        =   6
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton cmdCase9 
      Caption         =   "Case 9"
      Height          =   735
      Left            =   6360
      TabIndex        =   5
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton cmdCase8 
      Caption         =   "Case 8"
      Height          =   735
      Left            =   5280
      TabIndex        =   4
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton cmdCase7 
      Caption         =   "Case 7"
      Height          =   735
      Left            =   4200
      TabIndex        =   3
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton cmdCase6 
      Caption         =   "Case 6"
      Height          =   735
      Left            =   3120
      TabIndex        =   2
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton cmdCase3 
      Caption         =   "Case 3"
      Height          =   735
      Left            =   5280
      TabIndex        =   1
      Top             =   840
      Width           =   975
   End
   Begin VB.CommandButton cmdCase2 
      Caption         =   "Case 2"
      Height          =   735
      Left            =   4200
      TabIndex        =   0
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FFFF&
      Caption         =   "Bankers offer"
      Height          =   255
      Left            =   9480
      TabIndex        =   55
      Top             =   6960
      Width           =   1095
   End
End
Attribute VB_Name = "frmgamescreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit                  'makes programmer claim variables

Dim sum As Long                  'sets the variables
Dim ctr As Integer






Private Sub cmdNODEAL_Click()
picoffer.Cls                             'clears the bankersoffer from the text box
End Sub

Private Sub Command1_Click()
End                                    'ends program

End Sub

Private Sub Form_Load()

ctr = 0                              'sets variables
sum = 3418416

playerscase = 0
txtpickcase.Visible = False            'causes text box to appear only when button is pressd


End Sub

Private Sub cmdCase1_Click()
Dim case1 As Long
cmdCase1.Visible = False                'causes case/cmdbutton to dissapear
txtpickfirst.Visible = False
txtpickcase.Visible = True
                          

case1 = value(1)                        'sets values of variables
sum = sum - value(1)                    'subtracts value represented by case from total sum of all cases
ctr = ctr + 1
bankersoffer = sum / (25 - ctr)

If case1 = 1 And ctr > 1 Then                   'Makes the coinciding txt box turn red, coinciding with the value represented by the case
    txt1.BackColor = &HFF&
ElseIf case1 = 5 And ctr > 1 Then
    txt5.BackColor = &HFF&
ElseIf case1 = 10 And ctr > 1 Then
    txt10.BackColor = &HFF&
ElseIf case1 = 25 And ctr > 1 Then
    txt25.BackColor = &HFF&
ElseIf case1 = 50 And ctr > 1 Then
    txt50.BackColor = &HFF&
ElseIf case1 = 75 And ctr > 1 Then
    txt75.BackColor = &HFF&
ElseIf case1 = 100 And ctr > 1 Then
    txt100.BackColor = &HFF&
ElseIf case1 = 200 And ctr > 1 Then
    txt200.BackColor = &HFF&
ElseIf case1 = 300 And ctr > 1 Then
    txt300.BackColor = &HFF&
ElseIf case1 = 400 And ctr > 1 Then
    txt400.BackColor = &HFF&
ElseIf case1 = 500 And ctr > 1 Then
    txt500.BackColor = &HFF&
ElseIf case1 = 750 And ctr > 1 Then
    txt750.BackColor = &HFF&
ElseIf case1 = 1000 And ctr > 1 Then
    txt1000.BackColor = &HFF&
ElseIf case1 = 5000 And ctr > 1 Then
    txt5000.BackColor = &HFF&
ElseIf case1 = 10000 And ctr > 1 Then
    txt10000.BackColor = &HFF&
ElseIf case1 = 25000 And ctr > 1 Then
    txt25000.BackColor = &HFF&
ElseIf case1 = 50000 And ctr > 1 Then
    txt50000.BackColor = &HFF&
ElseIf case1 = 75000 And ctr > 1 Then
    txt75000.BackColor = &HFF&
ElseIf case1 = 100000 And ctr > 1 Then
    txt100000.BackColor = &HFF&
ElseIf case1 = 200000 And ctr > 1 Then
    txt200000.BackColor = &HFF&
ElseIf case1 = 300000 And ctr > 1 Then
    txt300000.BackColor = &HFF&
ElseIf case1 = 400000 And ctr > 1 Then
    txt400000.BackColor = &HFF&
ElseIf case1 = 500000 And ctr > 1 Then
    txt500000.BackColor = &HFF&
ElseIf case1 = 750000 And ctr > 1 Then
    txt750000.BackColor = &HFF&
ElseIf case1 = 1000000 And ctr > 1 Then
    txt1000000.BackColor = &HFF&
End If

    

    

If ctr = 1 Then                                'sets first case as variable "playerscase"
playerscase = value(1)
ElseIf ctr = 6 Then                                  'choose 5 boxes to recieve bankers offer
MsgBox "The Banker has offered" & bankersoffer       'Breaks the game into segments allowing the banker to make offers to the player
picoffer.Print ; bankersoffer;
ElseIf ctr = 10 Then
MsgBox "The Banker has offered" & bankersoffer       'choose 4 more boxes for next offer
picoffer.Print ; bankersoffer;
ElseIf ctr = 14 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 17 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 20 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 21 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 22 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 23 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 24 Then                                 'causes the form to switch to the final form in which the player can choose to keep their own case or pick the final remaining case on the board
frmgamescreen.Hide
frmgamefinal.Show
casefinal = value(1)                                 'sets the value of casefinal for use in the next form
End If



End Sub
Private Sub cmdCase2_Click()
Dim case2 As Long
cmdCase2.Visible = False                            'causes case/cmdbutton to dissapear
txtpickfirst.Visible = False
txtpickcase.Visible = True


case2 = value(2)                                    'sets values of variables
sum = sum - value(2)                                'subtracts value represented by case from total sum of all cases
ctr = ctr + 1
bankersoffer = sum / (25 - ctr)

If case2 = 1 And ctr > 1 Then
    txt1.BackColor = &HFF&                           'Makes the coinciding txt box turn red, coinciding with the value represented by the case
ElseIf case2 = 5 And ctr > 1 Then
    txt5.BackColor = &HFF&
ElseIf case2 = 10 And ctr > 1 Then
    txt10.BackColor = &HFF&
ElseIf case2 = 25 And ctr > 1 Then
    txt25.BackColor = &HFF&
ElseIf case2 = 50 And ctr > 1 Then
    txt50.BackColor = &HFF&
ElseIf case2 = 75 And ctr > 1 Then
    txt75.BackColor = &HFF&
ElseIf case2 = 100 And ctr > 1 Then
    txt100.BackColor = &HFF&
ElseIf case2 = 200 And ctr > 1 Then
    txt200.BackColor = &HFF&
ElseIf case2 = 300 And ctr > 1 Then
    txt300.BackColor = &HFF&
ElseIf case2 = 400 And ctr > 1 Then
    txt400.BackColor = &HFF&
ElseIf case2 = 500 And ctr > 1 Then
    txt500.BackColor = &HFF&
ElseIf case2 = 750 And ctr > 1 Then
    txt750.BackColor = &HFF&
ElseIf case2 = 1000 And ctr > 1 Then
    txt1000.BackColor = &HFF&
ElseIf case2 = 5000 And ctr > 1 Then
    txt5000.BackColor = &HFF&
ElseIf case2 = 10000 And ctr > 1 Then
    txt10000.BackColor = &HFF&
ElseIf case2 = 25000 And ctr > 1 Then
    txt25000.BackColor = &HFF&
ElseIf case2 = 50000 And ctr > 1 Then
    txt50000.BackColor = &HFF&
ElseIf case2 = 75000 And ctr > 1 Then
    txt75000.BackColor = &HFF&
ElseIf case2 = 100000 And ctr > 1 Then
    txt100000.BackColor = &HFF&
ElseIf case2 = 200000 And ctr > 1 Then
    txt200000.BackColor = &HFF&
ElseIf case2 = 300000 And ctr > 1 Then
    txt300000.BackColor = &HFF&
ElseIf case2 = 400000 And ctr > 1 Then
    txt400000.BackColor = &HFF&
ElseIf case2 = 500000 And ctr > 1 Then
    txt500000.BackColor = &HFF&
ElseIf case2 = 750000 And ctr > 1 Then
    txt750000.BackColor = &HFF&
ElseIf case2 = 1000000 And ctr > 1 Then
    txt1000000.BackColor = &HFF&
End If
If ctr = 1 Then                                     'sets first case as variable "playerscase"
playerscase = value(2)
ElseIf ctr = 6 Then                                     'choose 5 boxes to recieve bankers offer
MsgBox "The Banker has offered " & bankersoffer             'Breaks the game into segments allowing the banker to make offers to the player
picoffer.Print ; bankersoffer;
ElseIf ctr = 10 Then                                    'choose 4 more boxes for next offer
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 14 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 17 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 20 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 21 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 22 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 23 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 24 Then                            'causes the form to switch to the final form in which the player can choose to keep their own case or pick the final remaining case on the board
frmgamescreen.Hide
frmgamefinal.Show
casefinal = value(2)                            'sets the value of casefinal for use in the next form
End If
End Sub
Private Sub cmdCase3_Click()
Dim case3 As Long
cmdCase3.Visible = False                         'causes case/cmdbutton to dissapear
txtpickfirst.Visible = False
txtpickcase.Visible = True


case3 = value(3)                                 'sets values of variables
sum = sum - value(3)                             'subtracts value represented by case from total sum of all cases
ctr = ctr + 1
bankersoffer = sum / (25 - ctr)
If case3 = 1 And ctr > 1 Then                               'Makes the coinciding txt box turn red, coinciding with the value represented by the case
    txt1.BackColor = &HFF&
ElseIf case3 = 5 And ctr > 1 Then
    txt5.BackColor = &HFF&
ElseIf case3 = 10 And ctr > 1 Then
    txt10.BackColor = &HFF&
ElseIf case3 = 25 And ctr > 1 Then
    txt25.BackColor = &HFF&
ElseIf case3 = 50 And ctr > 1 Then
    txt50.BackColor = &HFF&
ElseIf case3 = 75 And ctr > 1 Then
    txt75.BackColor = &HFF&
ElseIf case3 = 100 And ctr > 1 Then
    txt100.BackColor = &HFF&
ElseIf case3 = 200 And ctr > 1 Then
    txt200.BackColor = &HFF&
ElseIf case3 = 300 And ctr > 1 Then
    txt300.BackColor = &HFF&
ElseIf case3 = 400 And ctr > 1 Then
    txt400.BackColor = &HFF&
ElseIf case3 = 500 And ctr > 1 Then
    txt500.BackColor = &HFF&
ElseIf case3 = 750 And ctr > 1 Then
    txt750.BackColor = &HFF&
ElseIf case3 = 1000 And ctr > 1 Then
    txt1000.BackColor = &HFF&
ElseIf case3 = 5000 And ctr > 1 Then
    txt5000.BackColor = &HFF&
ElseIf case3 = 10000 And ctr > 1 Then
    txt10000.BackColor = &HFF&
ElseIf case3 = 25000 And ctr > 1 Then
    txt25000.BackColor = &HFF&
ElseIf case3 = 50000 And ctr > 1 Then
    txt50000.BackColor = &HFF&
ElseIf case3 = 75000 And ctr > 1 Then
    txt75000.BackColor = &HFF&
ElseIf case3 = 100000 And ctr > 1 Then
    txt100000.BackColor = &HFF&
ElseIf case3 = 200000 And ctr > 1 Then
    txt200000.BackColor = &HFF&
ElseIf case3 = 300000 And ctr > 1 Then
    txt300000.BackColor = &HFF&
ElseIf case3 = 400000 And ctr > 1 Then
    txt400000.BackColor = &HFF&
ElseIf case3 = 500000 And ctr > 1 Then
    txt500000.BackColor = &HFF&
ElseIf case3 = 750000 Then
    txt750000.BackColor = &HFF&
ElseIf case3 = 1000000 And ctr > 1 Then
    txt1000000.BackColor = &HFF&
End If
If ctr = 1 Then                                         'sets first case as variable "playerscase"
playerscase = value(3)
ElseIf ctr = 6 Then                                     'choose 5 boxes to recieve bankers offer
MsgBox "The Banker has offered" & bankersoffer          'Breaks the game into segments allowing the banker to make offers to the player
picoffer.Print ; bankersoffer;
ElseIf ctr = 10 Then                                       'choose 4 more boxes for next offer
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 14 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 17 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 20 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 21 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 22 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 23 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 24 Then                        'causes the form to switch to the final form in which the player can choose to keep their own case or pick the final remaining case on the board
frmgamescreen.Hide
frmgamefinal.Show
casefinal = value(3)                        'sets the value of casefinal for use in the next form
End If
End Sub

Private Sub cmdCase4_Click()
Dim case4 As Long
cmdCase4.Visible = False                          'causes case/cmdbutton to dissapear
txtpickfirst.Visible = False
txtpickcase.Visible = True


case4 = value(4)                                  'sets values of variables
sum = sum - value(4)                              'subtracts value represented by case from total sum of all cases
ctr = ctr + 1
bankersoffer = sum / (25 - ctr)
If case4 = 1 And ctr > 1 Then                              'Makes the coinciding txt box turn red, coinciding with the value represented by the case
    txt1.BackColor = &HFF&
ElseIf case4 = 5 And ctr > 1 Then
    txt5.BackColor = &HFF&
ElseIf case4 = 10 And ctr > 1 Then
    txt10.BackColor = &HFF&
ElseIf case4 = 25 And ctr > 1 Then
    txt25.BackColor = &HFF&
ElseIf case4 = 50 And ctr > 1 Then
    txt50.BackColor = &HFF&
ElseIf case4 = 75 And ctr > 1 Then
    txt75.BackColor = &HFF&
ElseIf case4 = 100 And ctr > 1 Then
    txt100.BackColor = &HFF&
ElseIf case4 = 200 And ctr > 1 Then
    txt200.BackColor = &HFF&
ElseIf case4 = 300 And ctr > 1 Then
    txt300.BackColor = &HFF&
ElseIf case4 = 400 And ctr > 1 Then
    txt400.BackColor = &HFF&
ElseIf case4 = 500 And ctr > 1 Then
    txt500.BackColor = &HFF&
ElseIf case4 = 750 And ctr > 1 Then
    txt750.BackColor = &HFF&
ElseIf case4 = 1000 And ctr > 1 Then
    txt1000.BackColor = &HFF&
ElseIf case4 = 5000 And ctr > 1 Then
    txt5000.BackColor = &HFF&
ElseIf case4 = 10000 And ctr > 1 Then
    txt10000.BackColor = &HFF&
ElseIf case4 = 25000 And ctr > 1 Then
    txt25000.BackColor = &HFF&
ElseIf case4 = 50000 And ctr > 1 Then
    txt50000.BackColor = &HFF&
ElseIf case4 = 75000 And ctr > 1 Then
    txt75000.BackColor = &HFF&
ElseIf case4 = 100000 And ctr > 1 Then
    txt100000.BackColor = &HFF&
ElseIf case4 = 200000 And ctr > 1 Then
    txt200000.BackColor = &HFF&
ElseIf case4 = 300000 And ctr > 1 Then
    txt300000.BackColor = &HFF&
ElseIf case4 = 400000 And ctr > 1 Then
    txt400000.BackColor = &HFF&
ElseIf case4 = 500000 And ctr > 1 Then
    txt500000.BackColor = &HFF&
ElseIf case4 = 750000 Then
    txt750000.BackColor = &HFF&
ElseIf case4 = 1000000 And ctr > 1 Then
    txt1000000.BackColor = &HFF&
End If
If ctr = 1 Then                             'sets first case as variable "playerscase"
playerscase = value(4)
ElseIf ctr = 6 Then                             'choose 5 boxes to recieve bankers offer
MsgBox "The Banker has offered" & bankersoffer      'Breaks the game into segments allowing the banker to make offers to the player
picoffer.Print ; bankersoffer;
ElseIf ctr = 10 Then                                'choose 4 more boxes for next offer
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 14 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer
ElseIf ctr = 17 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 20 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 21 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 22 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 23 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 24 Then                            'causes the form to switch to the final form in which the player can choose to keep their own case or pick the final remaining case on the board
frmgamescreen.Hide
frmgamefinal.Show
casefinal = value(4)                            'sets the value of casefinal for use in the next form
End If
End Sub

Private Sub cmdCase5_Click()
Dim case5 As Long
cmdCase5.Visible = False                             'causes case/cmdbutton to dissapear
txtpickfirst.Visible = False
txtpickcase.Visible = True


case5 = value(5)                                     'sets values of variables
sum = sum - value(5)                                 'subtracts value represented by case from total sum of all cases
ctr = ctr + 1
bankersoffer = sum / (25 - ctr)
If case5 = 1 And ctr > 1 Then                                   'Makes the coinciding txt box turn red, coinciding with the value represented by the case
    txt1.BackColor = &HFF&
ElseIf case5 = 5 And ctr > 1 Then
    txt5.BackColor = &HFF&
ElseIf case5 = 10 And ctr > 1 Then
    txt10.BackColor = &HFF&
ElseIf case5 = 25 And ctr > 1 Then
    txt25.BackColor = &HFF&
ElseIf case5 = 50 And ctr > 1 Then
    txt50.BackColor = &HFF&
ElseIf case5 = 75 And ctr > 1 Then
    txt75.BackColor = &HFF&
ElseIf case5 = 100 And ctr > 1 Then
    txt100.BackColor = &HFF&
ElseIf case5 = 200 And ctr > 1 Then
    txt200.BackColor = &HFF&
ElseIf case5 = 300 And ctr > 1 Then
    txt300.BackColor = &HFF&
ElseIf case5 = 400 And ctr > 1 Then
    txt400.BackColor = &HFF&
ElseIf case5 = 500 And ctr > 1 Then
    txt500.BackColor = &HFF&
ElseIf case5 = 750 And ctr > 1 Then
    txt750.BackColor = &HFF&
ElseIf case5 = 1000 And ctr > 1 Then
    txt1000.BackColor = &HFF&
ElseIf case5 = 5000 And ctr > 1 Then
    txt5000.BackColor = &HFF&
ElseIf case5 = 10000 And ctr > 1 Then
    txt10000.BackColor = &HFF&
ElseIf case5 = 25000 And ctr > 1 Then
    txt25000.BackColor = &HFF&
ElseIf case5 = 50000 And ctr > 1 Then
    txt50000.BackColor = &HFF&
ElseIf case5 = 75000 And ctr > 1 Then
    txt75000.BackColor = &HFF&
ElseIf case5 = 100000 And ctr > 1 Then
    txt100000.BackColor = &HFF&
ElseIf case5 = 200000 And ctr > 1 Then
    txt200000.BackColor = &HFF&
ElseIf case5 = 300000 And ctr > 1 Then
    txt300000.BackColor = &HFF&
ElseIf case5 = 400000 And ctr > 1 Then
    txt400000.BackColor = &HFF&
ElseIf case5 = 500000 And ctr > 1 Then
    txt500000.BackColor = &HFF&
ElseIf case5 = 750000 Then
    txt750000.BackColor = &HFF&
ElseIf case5 = 1000000 And ctr > 1 Then
    txt1000000.BackColor = &HFF&
End If
If ctr = 1 Then                                 'sets first case as variable "playerscase"
playerscase = value(5)
ElseIf ctr = 6 Then                                 'choose 5 boxes to recieve bankers offer
MsgBox "The Banker has offered" & bankersoffer          'Breaks the game into segments allowing the banker to make offers to the player
picoffer.Print ; bankersoffer;
ElseIf ctr = 10 Then                                'choose 4 more boxes for next offer
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 14 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 17 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 20 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 21 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 22 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 23 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 24 Then                            'causes the form to switch to the final form in which the player can choose to keep their own case or pick the final remaining case on the board
frmgamescreen.Hide
frmgamefinal.Show
casefinal = value(5)                            'sets the value of casefinal for use in the next form
End If

End Sub

Private Sub cmdCase6_Click()
Dim case6 As Long
cmdCase6.Visible = False                          'causes case/cmdbutton to dissapear
txtpickfirst.Visible = False
txtpickcase.Visible = True


case6 = value(6)                                  'sets values of variables
sum = sum - value(6)                              'subtracts value represented by case from total sum of all cases
ctr = ctr + 1
bankersoffer = sum / (25 - ctr)
If case6 = 1 And ctr > 1 Then                                'Makes the coinciding txt box turn red, coinciding with the value represented by the case
    txt1.BackColor = &HFF&
ElseIf case6 = 5 And ctr > 1 Then
    txt5.BackColor = &HFF&
ElseIf case6 = 10 And ctr > 1 Then
    txt10.BackColor = &HFF&
ElseIf case6 = 25 And ctr > 1 Then
    txt25.BackColor = &HFF&
ElseIf case6 = 50 And ctr > 1 Then
    txt50.BackColor = &HFF&
ElseIf case6 = 75 And ctr > 1 Then
    txt75.BackColor = &HFF&
ElseIf case6 = 100 And ctr > 1 Then
    txt100.BackColor = &HFF&
ElseIf case6 = 200 And ctr > 1 Then
    txt200.BackColor = &HFF&
ElseIf case6 = 300 And ctr > 1 Then
    txt300.BackColor = &HFF&
ElseIf case6 = 400 And ctr > 1 Then
    txt400.BackColor = &HFF&
ElseIf case6 = 500 And ctr > 1 Then
    txt500.BackColor = &HFF&
ElseIf case6 = 750 And ctr > 1 Then
    txt750.BackColor = &HFF&
ElseIf case6 = 1000 And ctr > 1 Then
    txt1000.BackColor = &HFF&
ElseIf case6 = 5000 And ctr > 1 Then
    txt5000.BackColor = &HFF&
ElseIf case6 = 10000 And ctr > 1 Then
    txt10000.BackColor = &HFF&
ElseIf case6 = 25000 And ctr > 1 Then
    txt25000.BackColor = &HFF&
ElseIf case6 = 50000 And ctr > 1 Then
    txt50000.BackColor = &HFF&
ElseIf case6 = 75000 And ctr > 1 Then
    txt75000.BackColor = &HFF&
ElseIf case6 = 100000 And ctr > 1 Then
    txt100000.BackColor = &HFF&
ElseIf case6 = 200000 And ctr > 1 Then
    txt200000.BackColor = &HFF&
ElseIf case6 = 300000 And ctr > 1 Then
    txt300000.BackColor = &HFF&
ElseIf case6 = 400000 And ctr > 1 Then
    txt400000.BackColor = &HFF&
ElseIf case6 = 500000 And ctr > 1 Then
    txt500000.BackColor = &HFF&
ElseIf case6 = 750000 Then
    txt750000.BackColor = &HFF&
ElseIf case6 = 1000000 And ctr > 1 Then
    txt1000000.BackColor = &HFF&
End If

If ctr = 1 Then                                     'sets first case as variable "playerscase"
playerscase = value(6)
ElseIf ctr = 6 Then                                 'choose 5 boxes to recieve bankers offer
MsgBox "The Banker has offered" & bankersoffer          'Breaks the game into segments allowing the banker to make offers to the player
picoffer.Print ; bankersoffer;
ElseIf ctr = 10 Then                                'choose 4 more boxes for next offer
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 14 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 17 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 20 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 21 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 22 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 23 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 25 And ctr > 1 Then
frmgamescreen.Hide                              'causes the form to switch to the final form in which the player can choose to keep their own case or pick the final remaining case on the board
frmgamefinal.Show
casefinal = value(6)                           'sets the value of casefinal for use in the next form
End If
End Sub

Private Sub cmdCase7_Click()
Dim case7 As Long
cmdCase7.Visible = False                     'causes case/cmdbutton to dissapear
txtpickfirst.Visible = False
txtpickcase.Visible = True


case7 = value(7)                             'sets values of variables
sum = sum - value(7)                         'subtracts value represented by case from total sum of all cases
ctr = ctr + 1
bankersoffer = sum / (25 - ctr)
If case7 = 1 And ctr > 1 Then                            'Makes the coinciding txt box turn red, coinciding with the value represented by the case
    txt1.BackColor = &HFF&
ElseIf case7 = 5 And ctr > 1 Then
    txt5.BackColor = &HFF&
ElseIf case7 = 10 And ctr > 1 Then
    txt10.BackColor = &HFF&
ElseIf case7 = 25 And ctr > 1 Then
    txt25.BackColor = &HFF&
ElseIf case7 = 50 And ctr > 1 Then
    txt50.BackColor = &HFF&
ElseIf case7 = 75 And ctr > 1 Then
    txt75.BackColor = &HFF&
ElseIf case7 = 100 And ctr > 1 Then
    txt100.BackColor = &HFF&
ElseIf case7 = 200 And ctr > 1 Then
    txt200.BackColor = &HFF&
ElseIf case7 = 300 And ctr > 1 Then
    txt300.BackColor = &HFF&
ElseIf case7 = 400 And ctr > 1 Then
    txt400.BackColor = &HFF&
ElseIf case7 = 500 And ctr > 1 Then
    txt500.BackColor = &HFF&
ElseIf case7 = 750 And ctr > 1 Then
    txt750.BackColor = &HFF&
ElseIf case7 = 1000 And ctr > 1 Then
    txt1000.BackColor = &HFF&
ElseIf case7 = 5000 And ctr > 1 Then
    txt5000.BackColor = &HFF&
ElseIf case7 = 10000 And ctr > 1 Then
    txt10000.BackColor = &HFF&
ElseIf case7 = 25000 And ctr > 1 Then
    txt25000.BackColor = &HFF&
ElseIf case7 = 50000 And ctr > 1 Then
    txt50000.BackColor = &HFF&
ElseIf case7 = 75000 And ctr > 1 Then
    txt75000.BackColor = &HFF&
ElseIf case7 = 100000 And ctr > 1 Then
    txt100000.BackColor = &HFF&
ElseIf case7 = 200000 And ctr > 1 Then
    txt200000.BackColor = &HFF&
ElseIf case7 = 300000 And ctr > 1 Then
    txt300000.BackColor = &HFF&
ElseIf case7 = 400000 And ctr > 1 Then
    txt400000.BackColor = &HFF&
ElseIf case7 = 500000 And ctr > 1 Then
    txt500000.BackColor = &HFF&
ElseIf case7 = 750000 Then
    txt750000.BackColor = &HFF&
ElseIf case7 = 1000000 And ctr > 1 Then
    txt1000000.BackColor = &HFF&
End If
If ctr = 1 Then                                 'sets first case as variable "playerscase"
playerscase = value(7)
ElseIf ctr = 6 Then                             'choose 5 boxes to recieve bankers offer
MsgBox "The Banker has offered," & bankersoffer     'Breaks the game into segments allowing the banker to make offers to the player
picoffer.Print ; bankersoffer;
ElseIf ctr = 10 Then                                'choose 4 more boxes for next offer
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 14 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 17 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 20 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 21 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 22 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 23 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 24 Then                        'causes the form to switch to the final form in which the player can choose to keep their own case or pick the final remaining case on the board
frmgamescreen.Hide
frmgamefinal.Show
casefinal = value(7)                            'sets the value of casefinal for use in the next form
End If
End Sub

Private Sub cmdCase8_Click()
Dim case8 As Long
cmdCase8.Visible = False                      'causes case/cmdbutton to dissapear
txtpickfirst.Visible = False
txtpickcase.Visible = True


case8 = value(8)                              'sets values of variables
sum = sum - value(8)                          'subtracts value represented by case from total sum of all cases
ctr = ctr + 1
bankersoffer = sum / (25 - ctr)
If case8 = 1 And ctr > 1 Then                           'Makes the coinciding txt box turn red, coinciding with the value represented by the case
    txt1.BackColor = &HFF&
ElseIf case8 = 5 And ctr > 1 Then
    txt5.BackColor = &HFF&
ElseIf case8 = 10 And ctr > 1 Then
    txt10.BackColor = &HFF&
ElseIf case8 = 25 And ctr > 1 Then
    txt25.BackColor = &HFF&
ElseIf case8 = 50 And ctr > 1 Then
    txt50.BackColor = &HFF&
ElseIf case8 = 75 And ctr > 1 Then
    txt75.BackColor = &HFF&
ElseIf case8 = 100 And ctr > 1 Then
    txt100.BackColor = &HFF&
ElseIf case8 = 200 And ctr > 1 Then
    txt200.BackColor = &HFF&
ElseIf case8 = 300 And ctr > 1 Then
    txt300.BackColor = &HFF&
ElseIf case8 = 400 And ctr > 1 Then
    txt400.BackColor = &HFF&
ElseIf case8 = 500 And ctr > 1 Then
    txt500.BackColor = &HFF&
ElseIf case8 = 750 And ctr > 1 Then
    txt750.BackColor = &HFF&
ElseIf case8 = 1000 And ctr > 1 Then
    txt1000.BackColor = &HFF&
ElseIf case8 = 5000 And ctr > 1 Then
    txt5000.BackColor = &HFF&
ElseIf case8 = 10000 And ctr > 1 Then
    txt10000.BackColor = &HFF&
ElseIf case8 = 25000 And ctr > 1 Then
    txt25000.BackColor = &HFF&
ElseIf case8 = 50000 And ctr > 1 Then
    txt50000.BackColor = &HFF&
ElseIf case8 = 75000 And ctr > 1 Then
    txt75000.BackColor = &HFF&
ElseIf case8 = 100000 And ctr > 1 Then
    txt100000.BackColor = &HFF&
ElseIf case8 = 200000 And ctr > 1 Then
    txt200000.BackColor = &HFF&
ElseIf case8 = 300000 And ctr > 1 Then
    txt300000.BackColor = &HFF&
ElseIf case8 = 400000 And ctr > 1 Then
    txt400000.BackColor = &HFF&
ElseIf case8 = 500000 And ctr > 1 Then
    txt500000.BackColor = &HFF&
ElseIf case8 = 750000 Then
    txt750000.BackColor = &HFF&
ElseIf case8 = 1000000 And ctr > 1 Then
    txt1000000.BackColor = &HFF&
End If
If ctr = 1 Then                                     'sets first case as variable "playerscase"
playerscase = value(8)
ElseIf ctr = 6 Then                                 'choose 5 boxes to recieve bankers offer
MsgBox "The Banker has offered" & bankersoffer          'Breaks the game into segments allowing the banker to make offers to the player
picoffer.Print ; bankersoffer;
ElseIf ctr = 10 Then                                'choose 4 more boxes for next offer
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 14 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 17 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 20 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 21 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 22 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 23 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 24 Then                            'causes the form to switch to the final form in which the player can choose to keep their own case or pick the final remaining case on the board
frmgamescreen.Hide
frmgamefinal.Show
casefinal = value(8)                         'sets the value of casefinal for use in the next form
End If
End Sub

Private Sub cmdCase9_Click()
Dim case9 As Long
cmdCase9.Visible = False                           'causes case/cmdbutton to dissapear
txtpickfirst.Visible = False
txtpickcase.Visible = True


case9 = value(9)                                   'sets values of variables
sum = sum - value(9)                               'subtracts value represented by case from total sum of all cases
ctr = ctr + 1
bankersoffer = sum / (25 - ctr)
If case9 = 1 And ctr > 1 Then                                'Makes the coinciding txt box turn red, coinciding with the value represented by the case
    txt1.BackColor = &HFF&
ElseIf case9 = 5 And ctr > 1 Then
    txt5.BackColor = &HFF&
ElseIf case9 = 10 And ctr > 1 Then
    txt10.BackColor = &HFF&
ElseIf case9 = 25 And ctr > 1 Then
    txt25.BackColor = &HFF&
ElseIf case9 = 50 And ctr > 1 Then
    txt50.BackColor = &HFF&
ElseIf case9 = 75 And ctr > 1 Then
    txt75.BackColor = &HFF&
ElseIf case9 = 100 And ctr > 1 Then
    txt100.BackColor = &HFF&
ElseIf case9 = 200 And ctr > 1 Then
    txt200.BackColor = &HFF&
ElseIf case9 = 300 And ctr > 1 Then
    txt300.BackColor = &HFF&
ElseIf case9 = 400 And ctr > 1 Then
    txt400.BackColor = &HFF&
ElseIf case9 = 500 And ctr > 1 Then
    txt500.BackColor = &HFF&
ElseIf case9 = 750 And ctr > 1 Then
    txt750.BackColor = &HFF&
ElseIf case9 = 1000 And ctr > 1 Then
    txt1000.BackColor = &HFF&
ElseIf case9 = 5000 And ctr > 1 Then
    txt5000.BackColor = &HFF&
ElseIf case9 = 10000 And ctr > 1 Then
    txt10000.BackColor = &HFF&
ElseIf case9 = 25000 And ctr > 1 Then
    txt25000.BackColor = &HFF&
ElseIf case9 = 50000 And ctr > 1 Then
    txt50000.BackColor = &HFF&
ElseIf case9 = 75000 And ctr > 1 Then
    txt75000.BackColor = &HFF&
ElseIf case9 = 100000 And ctr > 1 Then
    txt100000.BackColor = &HFF&
ElseIf case9 = 200000 And ctr > 1 Then
    txt200000.BackColor = &HFF&
ElseIf case9 = 300000 And ctr > 1 Then
    txt300000.BackColor = &HFF&
ElseIf case9 = 400000 And ctr > 1 Then
    txt400000.BackColor = &HFF&
ElseIf case9 = 500000 And ctr > 1 Then
    txt500000.BackColor = &HFF&
ElseIf case9 = 750000 Then
    txt750000.BackColor = &HFF&
ElseIf case9 = 1000000 And ctr > 1 Then
    txt1000000.BackColor = &HFF&
End If
If ctr = 1 Then                             'sets first case as variable "playerscase"
playerscase = value(9)
ElseIf ctr = 6 Then                            'choose 5 boxes to recieve bankers offer
MsgBox "The Banker has offered" & bankersoffer      'Breaks the game into segments allowing the banker to make offers to the player
picoffer.Print ; bankersoffer;
ElseIf ctr = 10 Then                                'choose 4 more boxes for next offer
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 14 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 17 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 20 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 21 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 22 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 23 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 24 Then                                   'causes the form to switch to the final form in which the player can choose to keep their own case or pick the final remaining case on the board
frmgamescreen.Hide
frmgamefinal.Show
casefinal = value(9)                                  'sets the value of casefinal for use in the next form
End If
End Sub
Private Sub cmdCase10_Click()

Dim case10 As Long
cmdCase10.Visible = False                            'causes case/cmdbutton to dissapear
txtpickfirst.Visible = False
txtpickcase.Visible = True


case10 = value(10)                                   'sets values of variables
sum = sum - value(10)                                 'subtracts value represented by case from total sum of all cases
ctr = ctr + 1
bankersoffer = sum / (25 - ctr)
If case10 = 10 Then                                    'Makes the coinciding txt box turn red, coinciding with the value represented by the case
    txt1.BackColor = &HFF&
ElseIf case10 = 5 And ctr > 1 Then
    txt5.BackColor = &HFF&
ElseIf case10 = 10 And ctr > 1 Then
    txt10.BackColor = &HFF&
ElseIf case10 = 25 And ctr > 1 Then
    txt25.BackColor = &HFF&
ElseIf case10 = 50 And ctr > 1 Then
    txt50.BackColor = &HFF&
ElseIf case10 = 75 And ctr > 1 Then
    txt75.BackColor = &HFF&
ElseIf case10 = 100 And ctr > 1 Then
    txt100.BackColor = &HFF&
ElseIf case10 = 200 And ctr > 1 Then
    txt200.BackColor = &HFF&
ElseIf case10 = 300 And ctr > 1 Then
    txt300.BackColor = &HFF&
ElseIf case10 = 400 And ctr > 1 Then
    txt400.BackColor = &HFF&
ElseIf case10 = 500 And ctr > 1 Then
    txt500.BackColor = &HFF&
ElseIf case10 = 750 And ctr > 1 Then
    txt750.BackColor = &HFF&
ElseIf case10 = 1000 And ctr > 1 Then
    txt1000.BackColor = &HFF&
ElseIf case10 = 5000 And ctr > 1 Then
    txt5000.BackColor = &HFF&
ElseIf case10 = 10000 And ctr > 1 Then
    txt10000.BackColor = &HFF&
ElseIf case10 = 25000 And ctr > 1 Then
    txt25000.BackColor = &HFF&
ElseIf case10 = 50000 And ctr > 1 Then
    txt50000.BackColor = &HFF&
ElseIf case10 = 75000 And ctr > 1 Then
    txt75000.BackColor = &HFF&
ElseIf case10 = 100000 And ctr > 1 Then
    txt100000.BackColor = &HFF&
ElseIf case10 = 200000 And ctr > 1 Then
    txt200000.BackColor = &HFF&
ElseIf case10 = 300000 And ctr > 1 Then
    txt300000.BackColor = &HFF&
ElseIf case10 = 400000 And ctr > 1 Then
    txt400000.BackColor = &HFF&
ElseIf case10 = 500000 And ctr > 1 Then
    txt500000.BackColor = &HFF&
ElseIf case10 = 750000 Then
    txt750000.BackColor = &HFF&
ElseIf case10 = 1000000 And ctr > 1 Then
    txt1000000.BackColor = &HFF&
End If
If ctr = 1 Then                             'sets first case as variable "playerscase"
playerscase = value(10)
ElseIf ctr = 6 Then                             'choose 5 boxes to recieve bankers offer
MsgBox "The Banker has offered" & bankersoffer          'Breaks the game into segments allowing the banker to make offers to the player
picoffer.Print ; bankersoffer;
ElseIf ctr = 10 Then                                'choose 4 more boxes for next offer
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 14 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 17 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 20 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 21 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 22 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 23 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 24 Then                           'causes the form to switch to the final form in which the player can choose to keep their own case or pick the final remaining case on the board
frmgamescreen.Hide
frmgamefinal.Show
casefinal = value(10)                        'sets the value of casefinal for use in the next form

End If
End Sub

Private Sub cmdCase11_Click()
Dim case11 As Long
cmdCase11.Visible = False                    'causes case/cmdbutton to dissapear
txtpickfirst.Visible = False
txtpickcase.Visible = True


case11 = value(11)                           'sets values of variables
sum = sum - value(11)                         'subtracts value represented by case from total sum of all cases
ctr = ctr + 1
bankersoffer = sum / (25 - ctr)
If case11 = 1 And ctr > 1 Then                             'Makes the coinciding txt box turn red, coinciding with the value represented by the case
    txt1.BackColor = &HFF&
ElseIf case11 = 5 And ctr > 1 Then
    txt5.BackColor = &HFF&
ElseIf case11 = 10 And ctr > 1 Then
    txt10.BackColor = &HFF&
ElseIf case11 = 25 And ctr > 1 Then
    txt25.BackColor = &HFF&
ElseIf case11 = 50 And ctr > 1 Then
    txt50.BackColor = &HFF&
ElseIf case11 = 75 And ctr > 1 Then
    txt75.BackColor = &HFF&
ElseIf case11 = 100 And ctr > 1 Then
    txt100.BackColor = &HFF&
ElseIf case11 = 200 And ctr > 1 Then
    txt200.BackColor = &HFF&
ElseIf case11 = 300 And ctr > 1 Then
    txt300.BackColor = &HFF&
ElseIf case11 = 400 And ctr > 1 Then
    txt400.BackColor = &HFF&
ElseIf case11 = 500 And ctr > 1 Then
    txt500.BackColor = &HFF&
ElseIf case11 = 750 And ctr > 1 Then
    txt750.BackColor = &HFF&
ElseIf case11 = 1000 And ctr > 1 Then
    txt1000.BackColor = &HFF&
ElseIf case11 = 5000 And ctr > 1 Then
    txt5000.BackColor = &HFF&
ElseIf case11 = 10000 And ctr > 1 Then
    txt10000.BackColor = &HFF&
ElseIf case11 = 25000 And ctr > 1 Then
    txt25000.BackColor = &HFF&
ElseIf case11 = 50000 And ctr > 1 Then
    txt50000.BackColor = &HFF&
ElseIf case11 = 75000 And ctr > 1 Then
    txt75000.BackColor = &HFF&
ElseIf case11 = 100000 And ctr > 1 Then
    txt100000.BackColor = &HFF&
ElseIf case11 = 200000 And ctr > 1 Then
    txt200000.BackColor = &HFF&
ElseIf case11 = 300000 And ctr > 1 Then
    txt300000.BackColor = &HFF&
ElseIf case11 = 400000 And ctr > 1 Then
    txt400000.BackColor = &HFF&
ElseIf case11 = 500000 And ctr > 1 Then
    txt500000.BackColor = &HFF&
ElseIf case11 = 750000 Then
    txt750000.BackColor = &HFF&
ElseIf case11 = 1000000 And ctr > 1 Then
    txt1000000.BackColor = &HFF&
End If
If ctr = 1 Then                                 'sets first case as variable "playerscase"
playerscase = value(11)
ElseIf ctr = 6 Then                                 'choose 5 boxes to recieve bankers offer
MsgBox "The Banker has offered" & bankersoffer          'Breaks the game into segments allowing the banker to make offers to the player
picoffer.Print ; bankersoffer;
ElseIf ctr = 10 Then                                'choose 4 more boxes for next offer
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 14 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 17 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 20 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 21 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 22 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 23 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 24 Then                            'causes the form to switch to the final form in which the player can choose to keep their own case or pick the final remaining case on the board
frmgamescreen.Hide
frmgamefinal.Show
casefinal = value(11)                                'sets the value of casefinal for use in the next form
End If
End Sub

Private Sub cmdCase12_Click()
Dim case12 As Long
cmdCase12.Visible = False                           'causes case/cmdbutton to dissapear
txtpickfirst.Visible = False
txtpickcase.Visible = True


case12 = value(12)                                  'sets values of variables
sum = sum - value(12)                               'subtracts value represented by case from total sum of all cases
ctr = ctr + 1
bankersoffer = sum / (25 - ctr)
If case12 = 1 And ctr > 1 Then                                  'Makes the coinciding txt box turn red, coinciding with the value represented by the case
    txt1.BackColor = &HFF&
ElseIf case12 = 5 And ctr > 1 Then
    txt5.BackColor = &HFF&
ElseIf case12 = 10 And ctr > 1 Then
    txt10.BackColor = &HFF&
ElseIf case12 = 25 And ctr > 1 Then
    txt25.BackColor = &HFF&
ElseIf case12 = 50 And ctr > 1 Then
    txt50.BackColor = &HFF&
ElseIf case12 = 75 And ctr > 1 Then
    txt75.BackColor = &HFF&
ElseIf case12 = 100 And ctr > 1 Then
    txt100.BackColor = &HFF&
ElseIf case12 = 200 And ctr > 1 Then
    txt200.BackColor = &HFF&
ElseIf case12 = 300 And ctr > 1 Then
    txt300.BackColor = &HFF&
ElseIf case12 = 400 And ctr > 1 Then
    txt400.BackColor = &HFF&
ElseIf case12 = 500 And ctr > 1 Then
    txt500.BackColor = &HFF&
ElseIf case12 = 750 And ctr > 1 Then
    txt750.BackColor = &HFF&
ElseIf case12 = 1000 And ctr > 1 Then
    txt1000.BackColor = &HFF&
ElseIf case12 = 5000 And ctr > 1 Then
    txt5000.BackColor = &HFF&
ElseIf case12 = 10000 And ctr > 1 Then
    txt10000.BackColor = &HFF&
ElseIf case12 = 25000 And ctr > 1 Then
    txt25000.BackColor = &HFF&
ElseIf case12 = 50000 And ctr > 1 Then
    txt50000.BackColor = &HFF&
ElseIf case12 = 75000 And ctr > 1 Then
    txt75000.BackColor = &HFF&
ElseIf case12 = 100000 And ctr > 1 Then
    txt100000.BackColor = &HFF&
ElseIf case12 = 200000 And ctr > 1 Then
    txt200000.BackColor = &HFF&
ElseIf case12 = 300000 And ctr > 1 Then
    txt300000.BackColor = &HFF&
ElseIf case12 = 400000 And ctr > 1 Then
    txt400000.BackColor = &HFF&
ElseIf case12 = 500000 And ctr > 1 Then
    txt500000.BackColor = &HFF&
ElseIf case12 = 750000 Then
    txt750000.BackColor = &HFF&
ElseIf case12 = 1000000 And ctr > 1 Then
    txt1000000.BackColor = &HFF&
End If
If ctr = 1 Then                                 'sets first case as variable "playerscase"
playerscase = value(12)
ElseIf ctr = 6 Then                                 'choose 5 boxes to recieve bankers offer
MsgBox "The Banker has offered" & bankersoffer          'Breaks the game into segments allowing the banker to make offers to the player
picoffer.Print ; bankersoffer;
ElseIf ctr = 10 Then                                'choose 4 more boxes for next offer
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 14 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 17 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 20 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 21 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 22 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 23 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 24 Then                                    'causes the form to switch to the final form in which the player can choose to keep their own case or pick the final remaining case on the board
frmgamescreen.Hide
frmgamefinal.Show
casefinal = value(12)                                     'sets the value of casefinal for use in the next form
End If
End Sub

Private Sub cmdCase13_Click()
Dim case13 As Long
cmdCase13.Visible = False                                'causes case/cmdbutton to dissapear
txtpickfirst.Visible = False
txtpickcase.Visible = True


case13 = value(13)                                       'sets values of variables
sum = sum - value(13)                                    'subtracts value represented by case from total sum of all cases
ctr = ctr + 1
bankersoffer = sum / (25 - ctr)
If case13 = 1 And ctr > 1 Then                                       'Makes the coinciding txt box turn red, coinciding with the value represented by the case
    txt1.BackColor = &HFF&
ElseIf case13 = 5 And ctr > 1 Then
    txt5.BackColor = &HFF&
ElseIf case13 = 10 And ctr > 1 Then
    txt10.BackColor = &HFF&
ElseIf case13 = 25 And ctr > 1 Then
    txt25.BackColor = &HFF&
ElseIf case13 = 50 And ctr > 1 Then
    txt50.BackColor = &HFF&
ElseIf case13 = 75 And ctr > 1 Then
    txt75.BackColor = &HFF&
ElseIf case13 = 100 And ctr > 1 Then
    txt100.BackColor = &HFF&
ElseIf case13 = 200 And ctr > 1 Then
    txt200.BackColor = &HFF&
ElseIf case13 = 300 And ctr > 1 Then
    txt300.BackColor = &HFF&
ElseIf case13 = 400 And ctr > 1 Then
    txt400.BackColor = &HFF&
ElseIf case13 = 500 And ctr > 1 Then
    txt500.BackColor = &HFF&
ElseIf case13 = 750 And ctr > 1 Then
    txt750.BackColor = &HFF&
ElseIf case13 = 1000 And ctr > 1 Then
    txt1000.BackColor = &HFF&
ElseIf case13 = 5000 And ctr > 1 Then
    txt5000.BackColor = &HFF&
ElseIf case13 = 10000 And ctr > 1 Then
    txt10000.BackColor = &HFF&
ElseIf case13 = 25000 And ctr > 1 Then
    txt25000.BackColor = &HFF&
ElseIf case13 = 50000 And ctr > 1 Then
    txt50000.BackColor = &HFF&
ElseIf case13 = 75000 And ctr > 1 Then
    txt75000.BackColor = &HFF&
ElseIf case13 = 100000 And ctr > 1 Then
    txt100000.BackColor = &HFF&
ElseIf case13 = 200000 And ctr > 1 Then
    txt200000.BackColor = &HFF&
ElseIf case13 = 300000 And ctr > 1 Then
    txt300000.BackColor = &HFF&
ElseIf case13 = 400000 And ctr > 1 Then
    txt400000.BackColor = &HFF&
ElseIf case13 = 500000 And ctr > 1 Then
    txt500000.BackColor = &HFF&
ElseIf case13 = 750000 Then
    txt750000.BackColor = &HFF&
ElseIf case13 = 1000000 And ctr > 1 Then
    txt1000000.BackColor = &HFF&
End If
If ctr = 1 Then                                 'sets first case as variable "playerscase"
playerscase = value(13)
ElseIf ctr = 6 Then                                 'choose 5 boxes to recieve bankers offer
MsgBox "The Banker has offered" & bankersoffer          'Breaks the game into segments allowing the banker to make offers to the player
picoffer.Print ; bankersoffer;
ElseIf ctr = 10 Then                                'choose 4 more boxes for next offer
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 14 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 17 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 20 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 21 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 22 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 23 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 24 Then                            'causes the form to switch to the final form in which the player can choose to keep their own case or pick the final remaining case on the board
frmgamescreen.Hide
frmgamefinal.Show
casefinal = value(13)                           'sets the value of casefinal for use in the next form
End If
End Sub

Private Sub cmdCase14_Click()
Dim case14 As Long
cmdCase14.Visible = False                      'causes case/cmdbutton to dissapear
txtpickfirst.Visible = False
txtpickcase.Visible = True


case14 = value(14)                             'sets values of variables
sum = sum - value(14)                          'subtracts value represented by case from total sum of all cases
ctr = ctr + 1
bankersoffer = sum / (25 - ctr)
If case14 = 1 And ctr > 1 Then                            'Makes the coinciding txt box turn red, coinciding with the value represented by the case
    txt1.BackColor = &HFF&
ElseIf case14 = 5 And ctr > 1 Then
    txt5.BackColor = &HFF&
ElseIf case14 = 10 And ctr > 1 Then
    txt10.BackColor = &HFF&
ElseIf case14 = 25 And ctr > 1 Then
    txt25.BackColor = &HFF&
ElseIf case14 = 50 And ctr > 1 Then
    txt50.BackColor = &HFF&
ElseIf case14 = 75 And ctr > 1 Then
    txt75.BackColor = &HFF&
ElseIf case14 = 100 And ctr > 1 Then
    txt100.BackColor = &HFF&
ElseIf case14 = 200 And ctr > 1 Then
    txt200.BackColor = &HFF&
ElseIf case14 = 300 And ctr > 1 Then
    txt300.BackColor = &HFF&
ElseIf case14 = 400 And ctr > 1 Then
    txt400.BackColor = &HFF&
ElseIf case14 = 500 And ctr > 1 Then
    txt500.BackColor = &HFF&
ElseIf case14 = 750 And ctr > 1 Then
    txt750.BackColor = &HFF&
ElseIf case14 = 1000 And ctr > 1 Then
    txt1000.BackColor = &HFF&
ElseIf case14 = 5000 And ctr > 1 Then
    txt5000.BackColor = &HFF&
ElseIf case14 = 10000 And ctr > 1 Then
    txt10000.BackColor = &HFF&
ElseIf case14 = 25000 And ctr > 1 Then
    txt25000.BackColor = &HFF&
ElseIf case14 = 50000 And ctr > 1 Then
    txt50000.BackColor = &HFF&
ElseIf case14 = 75000 And ctr > 1 Then
    txt75000.BackColor = &HFF&
ElseIf case14 = 100000 And ctr > 1 Then
    txt100000.BackColor = &HFF&
ElseIf case14 = 200000 And ctr > 1 Then
    txt200000.BackColor = &HFF&
ElseIf case14 = 300000 And ctr > 1 Then
    txt300000.BackColor = &HFF&
ElseIf case14 = 400000 And ctr > 1 Then
    txt400000.BackColor = &HFF&
ElseIf case14 = 500000 And ctr > 1 Then
    txt500000.BackColor = &HFF&
ElseIf case14 = 750000 Then
    txt750000.BackColor = &HFF&
ElseIf case14 = 1000000 And ctr > 1 Then
    txt1000000.BackColor = &HFF&
End If
If ctr = 1 Then                             'sets first case as variable "playerscase"
playerscase = value(14)
ElseIf ctr = 6 Then                             'choose 5 boxes to recieve bankers offer
MsgBox "The Banker has offered" & bankersoffer          'Breaks the game into segments allowing the banker to make offers to the player
picoffer.Print ; bankersoffer;
ElseIf ctr = 10 Then                                'choose 4 more boxes for next offer
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 14 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 17 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 20 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 21 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 22 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 23 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 24 Then                                'causes the form to switch to the final form in which the player can choose to keep their own case or pick the final remaining case on the board
frmgamescreen.Hide
frmgamefinal.Show
casefinal = value(14)                               'sets the value of casefinal for use in the next form
End If
End Sub

Private Sub cmdCase15_Click()
Dim case15 As Long
cmdCase15.Visible = False                         'causes case/cmdbutton to dissapear
txtpickfirst.Visible = False
txtpickcase.Visible = True


case15 = value(15)                                'sets values of variables
sum = sum - value(15)                             'subtracts value represented by case from total sum of all cases
ctr = ctr + 1
bankersoffer = sum / (25 - ctr)
If case15 = 1 And ctr > 1 Then                              'Makes the coinciding txt box turn red, coinciding with the value represented by the case
    txt1.BackColor = &HFF&
ElseIf case15 = 5 And ctr > 1 Then
    txt5.BackColor = &HFF&
ElseIf case15 = 10 And ctr > 1 Then
    txt10.BackColor = &HFF&
ElseIf case15 = 25 And ctr > 1 Then
    txt25.BackColor = &HFF&
ElseIf case15 = 50 And ctr > 1 Then
    txt50.BackColor = &HFF&
ElseIf case15 = 75 And ctr > 1 Then
    txt75.BackColor = &HFF&
ElseIf case15 = 100 And ctr > 1 Then
    txt100.BackColor = &HFF&
ElseIf case15 = 200 And ctr > 1 Then
    txt200.BackColor = &HFF&
ElseIf case15 = 300 And ctr > 1 Then
    txt300.BackColor = &HFF&
ElseIf case15 = 400 And ctr > 1 Then
    txt400.BackColor = &HFF&
ElseIf case15 = 500 And ctr > 1 Then
    txt500.BackColor = &HFF&
ElseIf case15 = 750 And ctr > 1 Then
    txt750.BackColor = &HFF&
ElseIf case15 = 1000 And ctr > 1 Then
    txt1000.BackColor = &HFF&
ElseIf case15 = 5000 And ctr > 1 Then
    txt5000.BackColor = &HFF&
ElseIf case15 = 10000 And ctr > 1 Then
    txt10000.BackColor = &HFF&
ElseIf case15 = 25000 And ctr > 1 Then
    txt25000.BackColor = &HFF&
ElseIf case15 = 50000 And ctr > 1 Then
    txt50000.BackColor = &HFF&
ElseIf case15 = 75000 And ctr > 1 Then
    txt75000.BackColor = &HFF&
ElseIf case15 = 100000 And ctr > 1 Then
    txt100000.BackColor = &HFF&
ElseIf case15 = 200000 And ctr > 1 Then
    txt200000.BackColor = &HFF&
ElseIf case15 = 300000 And ctr > 1 Then
    txt300000.BackColor = &HFF&
ElseIf case15 = 400000 And ctr > 1 Then
    txt400000.BackColor = &HFF&
ElseIf case15 = 500000 And ctr > 1 Then
    txt500000.BackColor = &HFF&
ElseIf case15 = 750000 Then
    txt750000.BackColor = &HFF&
ElseIf case15 = 1000000 And ctr > 1 Then
    txt1000000.BackColor = &HFF&
End If
If ctr = 1 Then                                 'sets first case as variable "playerscase"
playerscase = value(15)
ElseIf ctr = 6 Then                             'choose 5 boxes to recieve bankers offer
MsgBox "The Banker has offered" & bankersoffer          'Breaks the game into segments allowing the banker to make offers to the player
picoffer.Print ; bankersoffer;
ElseIf ctr = 10 Then                                'choose 4 more boxes for next offer
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 14 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 17 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 20 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 21 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 22 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 23 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 24 Then                            'causes the form to switch to the final form in which the player can choose to keep their own case or pick the final remaining case on the board
frmgamescreen.Hide
frmgamefinal.Show
casefinal = value(15)                           'sets the value of casefinal for use in the next form
End If
End Sub

Private Sub cmdCase16_Click()
Dim case16 As Long
cmdCase16.Visible = False                      'causes case/cmdbutton to dissapear
txtpickfirst.Visible = False
txtpickcase.Visible = True


case16 = value(16)                             'sets values of variables
sum = sum - value(16)                          'subtracts value represented by case from total sum of all cases
ctr = ctr + 1
bankersoffer = sum / (25 - ctr)
If case16 = 1 And ctr > 1 Then                            'Makes the coinciding txt box turn red, coinciding with the value represented by the case
    txt1.BackColor = &HFF&
ElseIf case16 = 5 And ctr > 1 Then
    txt5.BackColor = &HFF&
ElseIf case16 = 10 And ctr > 1 Then
    txt10.BackColor = &HFF&
ElseIf case16 = 25 And ctr > 1 Then
    txt25.BackColor = &HFF&
ElseIf case16 = 50 And ctr > 1 Then
    txt50.BackColor = &HFF&
ElseIf case16 = 75 And ctr > 1 Then
    txt75.BackColor = &HFF&
ElseIf case16 = 100 And ctr > 1 Then
    txt100.BackColor = &HFF&
ElseIf case16 = 200 And ctr > 1 Then
    txt200.BackColor = &HFF&
ElseIf case16 = 300 And ctr > 1 Then
    txt300.BackColor = &HFF&
ElseIf case16 = 400 And ctr > 1 Then
    txt400.BackColor = &HFF&
ElseIf case16 = 500 And ctr > 1 Then
    txt500.BackColor = &HFF&
ElseIf case16 = 750 And ctr > 1 Then
    txt750.BackColor = &HFF&
ElseIf case16 = 1000 And ctr > 1 Then
    txt1000.BackColor = &HFF&
ElseIf case16 = 5000 And ctr > 1 Then
    txt5000.BackColor = &HFF&
ElseIf case16 = 10000 And ctr > 1 Then
    txt10000.BackColor = &HFF&
ElseIf case16 = 25000 And ctr > 1 Then
    txt25000.BackColor = &HFF&
ElseIf case16 = 50000 And ctr > 1 Then
    txt50000.BackColor = &HFF&
ElseIf case16 = 75000 And ctr > 1 Then
    txt75000.BackColor = &HFF&
ElseIf case16 = 100000 And ctr > 1 Then
    txt100000.BackColor = &HFF&
ElseIf case16 = 200000 And ctr > 1 Then
    txt200000.BackColor = &HFF&
ElseIf case16 = 300000 And ctr > 1 Then
    txt300000.BackColor = &HFF&
ElseIf case16 = 400000 And ctr > 1 Then
    txt400000.BackColor = &HFF&
ElseIf case16 = 500000 And ctr > 1 Then
    txt500000.BackColor = &HFF&
ElseIf case16 = 750000 Then
    txt750000.BackColor = &HFF&
ElseIf case16 = 1000000 And ctr > 1 Then
    txt1000000.BackColor = &HFF&
End If
If ctr = 1 Then                                 'sets first case as variable "playerscase"
playerscase = value(16)
ElseIf ctr = 6 Then                             'choose 5 boxes to recieve bankers offer
MsgBox "The Banker has offered" & bankersoffer          'Breaks the game into segments allowing the banker to make offers to the player
picoffer.Print ; bankersoffer;
ElseIf ctr = 10 Then                                'choose 4 more boxes for next offer
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 14 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 17 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 20 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 21 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 22 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 23 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 24 Then                        'causes the form to switch to the final form in which the player can choose to keep their own case or pick the final remaining case on the board
frmgamescreen.Hide
frmgamefinal.Show
casefinal = value(16)                       'sets the value of casefinal for use in the next form
End If
End Sub

Private Sub cmdCase17_Click()
Dim case17 As Long
cmdCase17.Visible = False                    'causes case/cmdbutton to dissapear
txtpickfirst.Visible = False
txtpickcase.Visible = True


case17 = value(17)                           'sets value of variables
sum = sum - value(17)                        'subtracts value represented by case from total sum of all cases
ctr = ctr + 1
bankersoffer = sum / (25 - ctr)
If case17 = 1 And ctr > 1 Then                                'Makes the coinciding txt box turn red, coinciding with the value represented by the case
    txt1.BackColor = &HFF&
ElseIf case17 = 5 And ctr > 1 Then
    txt5.BackColor = &HFF&
ElseIf case17 = 10 And ctr > 1 Then
    txt10.BackColor = &HFF&
ElseIf case17 = 25 And ctr > 1 Then
    txt25.BackColor = &HFF&
ElseIf case17 = 50 And ctr > 1 Then
    txt50.BackColor = &HFF&
ElseIf case17 = 75 And ctr > 1 Then
    txt75.BackColor = &HFF&
ElseIf case17 = 100 And ctr > 1 Then
    txt100.BackColor = &HFF&
ElseIf case17 = 200 And ctr > 1 Then
    txt200.BackColor = &HFF&
ElseIf case17 = 300 And ctr > 1 Then
    txt300.BackColor = &HFF&
ElseIf case17 = 400 And ctr > 1 Then
    txt400.BackColor = &HFF&
ElseIf case17 = 500 And ctr > 1 Then
    txt500.BackColor = &HFF&
ElseIf case17 = 750 And ctr > 1 Then
    txt750.BackColor = &HFF&
ElseIf case17 = 1000 And ctr > 1 Then
    txt1000.BackColor = &HFF&
ElseIf case17 = 5000 And ctr > 1 Then
    txt5000.BackColor = &HFF&
ElseIf case17 = 10000 And ctr > 1 Then
    txt10000.BackColor = &HFF&
ElseIf case17 = 25000 And ctr > 1 Then
    txt25000.BackColor = &HFF&
ElseIf case17 = 50000 And ctr > 1 Then
    txt50000.BackColor = &HFF&
ElseIf case17 = 75000 And ctr > 1 Then
    txt75000.BackColor = &HFF&
ElseIf case17 = 100000 And ctr > 1 Then
    txt100000.BackColor = &HFF&
ElseIf case17 = 200000 And ctr > 1 Then
    txt200000.BackColor = &HFF&
ElseIf case17 = 300000 And ctr > 1 Then
    txt300000.BackColor = &HFF&
ElseIf case17 = 400000 And ctr > 1 Then
    txt400000.BackColor = &HFF&
ElseIf case17 = 500000 And ctr > 1 Then
    txt500000.BackColor = &HFF&
ElseIf case17 = 750000 Then
    txt750000.BackColor = &HFF&
ElseIf case17 = 1000000 And ctr > 1 Then
    txt1000000.BackColor = &HFF&
End If
If ctr = 1 Then                                 'sets first case as variable "playerscase"
playerscase = value(17)
ElseIf ctr = 6 Then                             'choose 5 boxes to recieve bankers offer
MsgBox "The Banker has offered" & bankersoffer      'Breaks the game into segments allowing the banker to make offers to the player
picoffer.Print ; bankersoffer;
ElseIf ctr = 10 Then                                'choose 4 more boxes for next offer
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 14 Then
MsgBox "The Banker has offered& bankersoffer"
picoffer.Print ; bankersoffer;
ElseIf ctr = 17 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 20 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 21 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 22 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 23 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 24 Then                            'causes the form to switch to the final form in which the player can choose to keep their own case or pick the final remaining case on the board
frmgamescreen.Hide
frmgamefinal.Show
casefinal = value(17)                          'sets the value of casefinal for use in the next form
End If
End Sub

Private Sub cmdCase18_Click()
Dim case18 As Long
cmdCase18.Visible = False                     'causes case/cmdbutton to dissapear
txtpickfirst.Visible = False
txtpickcase.Visible = True


case18 = value(18)                            'sets values of variables
sum = sum - value(18)                         'subtracts value represented by case from total sum of all cases
ctr = ctr + 1
bankersoffer = sum / (25 - ctr)
If case18 = 1 And ctr > 1 Then                            'Makes the coinciding txt box turn red, coinciding with the value represented by the case
    txt1.BackColor = &HFF&
ElseIf case18 = 5 And ctr > 1 Then
    txt5.BackColor = &HFF&
ElseIf case18 = 10 And ctr > 1 Then
    txt10.BackColor = &HFF&
ElseIf case18 = 25 And ctr > 1 Then
    txt25.BackColor = &HFF&
ElseIf case18 = 50 And ctr > 1 Then
    txt50.BackColor = &HFF&
ElseIf case18 = 75 And ctr > 1 Then
    txt75.BackColor = &HFF&
ElseIf case18 = 100 And ctr > 1 Then
    txt100.BackColor = &HFF&
ElseIf case18 = 200 And ctr > 1 Then
    txt200.BackColor = &HFF&
ElseIf case18 = 300 And ctr > 1 Then
    txt300.BackColor = &HFF&
ElseIf case18 = 400 And ctr > 1 Then
    txt400.BackColor = &HFF&
ElseIf case18 = 500 And ctr > 1 Then
    txt500.BackColor = &HFF&
ElseIf case18 = 750 And ctr > 1 Then
    txt750.BackColor = &HFF&
ElseIf case18 = 1000 And ctr > 1 Then
    txt1000.BackColor = &HFF&
ElseIf case18 = 5000 And ctr > 1 Then
    txt5000.BackColor = &HFF&
ElseIf case18 = 10000 And ctr > 1 Then
    txt10000.BackColor = &HFF&
ElseIf case18 = 25000 And ctr > 1 Then
    txt25000.BackColor = &HFF&
ElseIf case18 = 50000 And ctr > 1 Then
    txt50000.BackColor = &HFF&
ElseIf case18 = 75000 And ctr > 1 Then
    txt75000.BackColor = &HFF&
ElseIf case18 = 100000 And ctr > 1 Then
    txt100000.BackColor = &HFF&
ElseIf case18 = 200000 And ctr > 1 Then
    txt200000.BackColor = &HFF&
ElseIf case18 = 300000 And ctr > 1 Then
    txt300000.BackColor = &HFF&
ElseIf case18 = 400000 And ctr > 1 Then
    txt400000.BackColor = &HFF&
ElseIf case18 = 500000 And ctr > 1 Then
    txt500000.BackColor = &HFF&
ElseIf case18 = 750000 Then
    txt750000.BackColor = &HFF&
ElseIf case18 = 1000000 And ctr > 1 Then
    txt1000000.BackColor = &HFF&
End If
If ctr = 1 Then                                 'sets first case as variable "playerscase"
playerscase = value(18)
ElseIf ctr = 6 Then                                'choose 5 boxes to recieve bankers offer
MsgBox "The Banker has offered" & bankersoffer          'Breaks the game into segments allowing the banker to make offers to the player
picoffer.Print ; bankersoffer;
ElseIf ctr = 10 Then                                'choose 4 more boxes for next offer
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 14 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 17 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 20 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 21 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 22 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 23 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 24 Then                        'causes the form to switch to the final form in which the player can choose to keep their own case or pick the final remaining case on the board
frmgamescreen.Hide
frmgamefinal.Show
casefinal = value(18)                       'sets the value of casefinal for use in the next form
End If
End Sub

Private Sub cmdCase19_Click()
Dim case19 As Long
cmdCase19.Visible = False                  'causes case/cmdbutton to dissapear
txtpickfirst.Visible = False
txtpickcase.Visible = True


case19 = value(19)                         'sets values of variables
sum = sum - value(19)                      'subtracts value represented by case from total sum of all cases
ctr = ctr + 1
bankersoffer = sum / (25 - ctr)
If case19 = 1 And ctr > 1 Then                         'Makes the coinciding txt box turn red, coinciding with the value represented by the case
    txt1.BackColor = &HFF&
ElseIf case19 = 5 And ctr > 1 Then
    txt5.BackColor = &HFF&
ElseIf case19 = 10 And ctr > 1 Then
    txt10.BackColor = &HFF&
ElseIf case19 = 25 And ctr > 1 Then
    txt25.BackColor = &HFF&
ElseIf case19 = 50 And ctr > 1 Then
    txt50.BackColor = &HFF&
ElseIf case19 = 75 And ctr > 1 Then
    txt75.BackColor = &HFF&
ElseIf case19 = 100 And ctr > 1 Then
    txt100.BackColor = &HFF&
ElseIf case19 = 200 And ctr > 1 Then
    txt200.BackColor = &HFF&
ElseIf case19 = 300 And ctr > 1 Then
    txt300.BackColor = &HFF&
ElseIf case19 = 400 And ctr > 1 Then
    txt400.BackColor = &HFF&
ElseIf case19 = 500 And ctr > 1 Then
    txt500.BackColor = &HFF&
ElseIf case19 = 750 And ctr > 1 Then
    txt750.BackColor = &HFF&
ElseIf case19 = 1000 And ctr > 1 Then
    txt1000.BackColor = &HFF&
ElseIf case19 = 5000 And ctr > 1 Then
    txt5000.BackColor = &HFF&
ElseIf case19 = 10000 And ctr > 1 Then
    txt10000.BackColor = &HFF&
ElseIf case19 = 25000 And ctr > 1 Then
    txt25000.BackColor = &HFF&
ElseIf case19 = 50000 And ctr > 1 Then
    txt50000.BackColor = &HFF&
ElseIf case19 = 75000 And ctr > 1 Then
    txt75000.BackColor = &HFF&
ElseIf case19 = 100000 And ctr > 1 Then
    txt100000.BackColor = &HFF&
ElseIf case19 = 200000 And ctr > 1 Then
    txt200000.BackColor = &HFF&
ElseIf case19 = 300000 And ctr > 1 Then
    txt300000.BackColor = &HFF&
ElseIf case19 = 400000 And ctr > 1 Then
    txt400000.BackColor = &HFF&
ElseIf case19 = 500000 And ctr > 1 Then
    txt500000.BackColor = &HFF&
ElseIf case19 = 750000 Then
    txt750000.BackColor = &HFF&
ElseIf case19 = 1000000 And ctr > 1 Then
    txt1000000.BackColor = &HFF&
End If
If ctr = 1 Then                                 'sets first case as variable "playerscase"
playerscase = value(19)
ElseIf ctr = 6 Then                             'choose 5 boxes to recieve bankers offer
MsgBox "The Banker has offered" & bankersoffer          'Breaks the game into segments allowing the banker to make offers to the player
picoffer.Print ; bankersoffer;
ElseIf ctr = 10 Then                                'choose 4 more boxes for next offer
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 14 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 17 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 20 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 21 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 22 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 23 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 24 Then                                   'causes the form to switch to the final form in which the player can choose to keep their own case or pick the final remaining case on the board
frmgamescreen.Hide
frmgamefinal.Show
casefinal = value(19)                               'sets the value of casefinal for use in the next form
End If
End Sub
Private Sub cmdCase20_Click()
Dim case20 As Long
cmdCase20.Visible = False                  'causes case/cmdbutton to dissapear
txtpickfirst.Visible = False
txtpickcase.Visible = True


case20 = value(20)                         'sets values of variables
sum = sum - value(20)                      'subtracts value represented by case from total sum of all cases
ctr = ctr + 1
bankersoffer = sum / (25 - ctr)
If case20 = 1 And ctr > 1 Then                          'Makes the coinciding txt box turn red, coinciding with the value represented by the case
    txt1.BackColor = &HFF&
ElseIf case20 = 5 And ctr > 1 Then
    txt5.BackColor = &HFF&
ElseIf case20 = 10 And ctr > 1 Then
    txt10.BackColor = &HFF&
ElseIf case20 = 25 And ctr > 1 Then
    txt25.BackColor = &HFF&
ElseIf case20 = 50 And ctr > 1 Then
    txt50.BackColor = &HFF&
ElseIf case20 = 75 And ctr > 1 Then
    txt75.BackColor = &HFF&
ElseIf case20 = 100 And ctr > 1 Then
    txt100.BackColor = &HFF&
ElseIf case20 = 200 And ctr > 1 Then
    txt200.BackColor = &HFF&
ElseIf case20 = 300 And ctr > 1 Then
    txt300.BackColor = &HFF&
ElseIf case20 = 400 And ctr > 1 Then
    txt400.BackColor = &HFF&
ElseIf case20 = 500 And ctr > 1 Then
    txt500.BackColor = &HFF&
ElseIf case20 = 750 And ctr > 1 Then
    txt750.BackColor = &HFF&
ElseIf case20 = 1000 And ctr > 1 Then
    txt1000.BackColor = &HFF&
ElseIf case20 = 5000 And ctr > 1 Then
    txt5000.BackColor = &HFF&
ElseIf case20 = 10000 And ctr > 1 Then
    txt10000.BackColor = &HFF&
ElseIf case20 = 25000 And ctr > 1 Then
    txt25000.BackColor = &HFF&
ElseIf case20 = 50000 And ctr > 1 Then
    txt50000.BackColor = &HFF&
ElseIf case20 = 75000 And ctr > 1 Then
    txt75000.BackColor = &HFF&
ElseIf case20 = 100000 And ctr > 1 Then
    txt100000.BackColor = &HFF&
ElseIf case20 = 200000 And ctr > 1 Then
    txt200000.BackColor = &HFF&
ElseIf case20 = 300000 And ctr > 1 Then
    txt300000.BackColor = &HFF&
ElseIf case20 = 400000 And ctr > 1 Then
    txt400000.BackColor = &HFF&
ElseIf case20 = 500000 And ctr > 1 Then
    txt500000.BackColor = &HFF&
ElseIf case20 = 750000 Then
    txt750000.BackColor = &HFF&
ElseIf case20 = 1000000 And ctr > 1 Then
    txt1000000.BackColor = &HFF&
End If
If ctr = 1 Then                                         'sets first case as variable "playerscase"
playerscase = value(20)
ElseIf ctr = 6 Then                                 'choose 5 boxes to recieve bankers offer
MsgBox "The Banker has offered" & bankersoffer              'Breaks the game into segments allowing the banker to make offers to the player
picoffer.Print ; bankersoffer;
ElseIf ctr = 10 Then                                    'choose 4 more boxes for next offer
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 14 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 17 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 20 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 21 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 22 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 23 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 24 Then                                'causes the form to switch to the final form in which the player can choose to keep their own case or pick the final remaining case on the board
frmgamescreen.Hide
frmgamefinal.Show
casefinal = value(20)                               'sets the value of casefinal for use in the next form
End If
End Sub

Private Sub cmdCase21_Click()
Dim case21 As Long
cmdCase21.Visible = False                        'causes case/cmdbutton to dissapear
txtpickfirst.Visible = False
txtpickcase.Visible = True


case21 = value(21)                               'sets values of variables
sum = sum - value(21)                             'subtracts value represented by case from total sum of all cases
ctr = ctr + 1
bankersoffer = sum / (25 - ctr)
If case21 = 1 And ctr > 1 Then                              'Makes the coinciding txt box turn red, coinciding with the value represented by the case
    txt1.BackColor = &HFF&
ElseIf case21 = 5 And ctr > 1 Then
    txt5.BackColor = &HFF&
ElseIf case21 = 10 And ctr > 1 Then
    txt10.BackColor = &HFF&
ElseIf case21 = 25 And ctr > 1 Then
    txt25.BackColor = &HFF&
ElseIf case21 = 50 And ctr > 1 Then
    txt50.BackColor = &HFF&
ElseIf case21 = 75 And ctr > 1 Then
    txt75.BackColor = &HFF&
ElseIf case21 = 100 And ctr > 1 Then
    txt100.BackColor = &HFF&
ElseIf case21 = 200 And ctr > 1 Then
    txt200.BackColor = &HFF&
ElseIf case21 = 300 And ctr > 1 Then
    txt300.BackColor = &HFF&
ElseIf case21 = 400 And ctr > 1 Then
    txt400.BackColor = &HFF&
ElseIf case21 = 500 And ctr > 1 Then
    txt500.BackColor = &HFF&
ElseIf case21 = 750 And ctr > 1 Then
    txt750.BackColor = &HFF&
ElseIf case21 = 1000 And ctr > 1 Then
    txt1000.BackColor = &HFF&
ElseIf case21 = 5000 And ctr > 1 Then
    txt5000.BackColor = &HFF&
ElseIf case21 = 10000 And ctr > 1 Then
    txt10000.BackColor = &HFF&
ElseIf case21 = 25000 And ctr > 1 Then
    txt25000.BackColor = &HFF&
ElseIf case21 = 50000 And ctr > 1 Then
    txt50000.BackColor = &HFF&
ElseIf case21 = 75000 And ctr > 1 Then
    txt75000.BackColor = &HFF&
ElseIf case21 = 100000 And ctr > 1 Then
    txt100000.BackColor = &HFF&
ElseIf case21 = 200000 And ctr > 1 Then
    txt200000.BackColor = &HFF&
ElseIf case21 = 300000 And ctr > 1 Then
    txt300000.BackColor = &HFF&
ElseIf case21 = 400000 And ctr > 1 Then
    txt400000.BackColor = &HFF&
ElseIf case21 = 500000 And ctr > 1 Then
    txt500000.BackColor = &HFF&
ElseIf case21 = 750000 Then
    txt750000.BackColor = &HFF&
ElseIf case21 = 1000000 And ctr > 1 Then
    txt1000000.BackColor = &HFF&
End If
If ctr = 1 Then                                 'sets first case as variable "playerscase"
playerscase = value(21)
ElseIf ctr = 6 Then                             'choose 5 boxes to recieve bankers offer
MsgBox "The Banker has offered" & bankersoffer          'Breaks the game into segments allowing the banker to make offers to the player
picoffer.Print ; bankersoffer;
ElseIf ctr = 10 Then                                        'choose 4 more boxes for next offer
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 14 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 17 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 20 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 21 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 22 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 23 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 24 Then                               'causes the form to switch to the final form in which the player can choose to keep their own case or pick the final remaining case on the board
frmgamescreen.Hide
frmgamefinal.Show
casefinal = value(21)                             'sets the value of casefinal for use in the next form
End If
End Sub

Private Sub cmdCase22_Click()
Dim case22 As Long
cmdCase22.Visible = False                         'causes case/cmdbutton to dissapear
txtpickfirst.Visible = False
txtpickcase.Visible = True


case22 = value(22)                                'sets values of variables
sum = sum - value(22)                             'subtracts value represented by case from total sum of all cases
ctr = ctr + 1
bankersoffer = sum / (25 - ctr)
If case22 = 1 And ctr > 1 Then                                 'Makes the coinciding txt box turn red, coinciding with the value represented by the case
    txt1.BackColor = &HFF&
ElseIf case22 = 5 And ctr > 1 Then
    txt5.BackColor = &HFF&
ElseIf case22 = 10 And ctr > 1 Then
    txt10.BackColor = &HFF&
ElseIf case22 = 25 And ctr > 1 Then
    txt25.BackColor = &HFF&
ElseIf case22 = 50 And ctr > 1 Then
    txt50.BackColor = &HFF&
ElseIf case22 = 75 And ctr > 1 Then
    txt75.BackColor = &HFF&
ElseIf case22 = 100 And ctr > 1 Then
    txt100.BackColor = &HFF&
ElseIf case22 = 200 And ctr > 1 Then
    txt200.BackColor = &HFF&
ElseIf case22 = 300 And ctr > 1 Then
    txt300.BackColor = &HFF&
ElseIf case22 = 400 And ctr > 1 Then
    txt400.BackColor = &HFF&
ElseIf case22 = 500 And ctr > 1 Then
    txt500.BackColor = &HFF&
ElseIf case22 = 750 And ctr > 1 Then
    txt750.BackColor = &HFF&
ElseIf case22 = 1000 And ctr > 1 Then
    txt1000.BackColor = &HFF&
ElseIf case22 = 5000 And ctr > 1 Then
    txt5000.BackColor = &HFF&
ElseIf case22 = 10000 And ctr > 1 Then
    txt10000.BackColor = &HFF&
ElseIf case22 = 25000 And ctr > 1 Then
    txt25000.BackColor = &HFF&
ElseIf case22 = 50000 And ctr > 1 Then
    txt50000.BackColor = &HFF&
ElseIf case22 = 75000 And ctr > 1 Then
    txt75000.BackColor = &HFF&
ElseIf case22 = 100000 And ctr > 1 Then
    txt100000.BackColor = &HFF&
ElseIf case22 = 200000 And ctr > 1 Then
    txt200000.BackColor = &HFF&
ElseIf case22 = 300000 And ctr > 1 Then
    txt300000.BackColor = &HFF&
ElseIf case22 = 400000 And ctr > 1 Then
    txt400000.BackColor = &HFF&
ElseIf case22 = 500000 And ctr > 1 Then
    txt500000.BackColor = &HFF&
ElseIf case22 = 750000 Then
    txt750000.BackColor = &HFF&
ElseIf case22 = 1000000 And ctr > 1 Then
    txt1000000.BackColor = &HFF&
End If
If ctr = 1 Then                                 'sets first case as variable "playerscase"
playerscase = value(22)
ElseIf ctr = 6 Then                             'choose 5 boxes to recieve bankers offer
MsgBox "The Banker has offered" & bankersoffer          'Breaks the game into segments allowing the banker to make offers to the player
picoffer.Print ; bankersoffer;
ElseIf ctr = 10 Then                                    'choose 4 more boxes for next offer
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 14 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 17 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 20 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 21 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 22 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 23 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 24 Then                            'causes the form to switch to the final form in which the player can choose to keep their own case or pick the final remaining case on the board
frmgamescreen.Hide
frmgamefinal.Show
casefinal = value(22)                               'sets the value of casefinal for use in the next form
End If
End Sub

Private Sub cmdCase23_Click()
Dim case23 As Long
cmdCase23.Visible = False                        'causes case/cmdbutton to dissapear
txtpickfirst.Visible = False
txtpickcase.Visible = True


case23 = value(23)                               'sets values of variables
sum = sum - value(23)                             'subtracts value represented by case from total sum of all cases
ctr = ctr + 1
bankersoffer = sum / (25 - ctr)
If case23 = 1 And ctr > 1 Then                               'Makes the coinciding txt box turn red, coinciding with the value represented by the case
    txt1.BackColor = &HFF&
ElseIf case23 = 5 And ctr > 1 Then
    txt5.BackColor = &HFF&
ElseIf case23 = 10 And ctr > 1 Then
    txt10.BackColor = &HFF&
ElseIf case23 = 25 And ctr > 1 Then
    txt25.BackColor = &HFF&
ElseIf case23 = 50 And ctr > 1 Then
    txt50.BackColor = &HFF&
ElseIf case23 = 75 And ctr > 1 Then
    txt75.BackColor = &HFF&
ElseIf case23 = 100 And ctr > 1 Then
    txt100.BackColor = &HFF&
ElseIf case23 = 200 And ctr > 1 Then
    txt200.BackColor = &HFF&
ElseIf case23 = 300 And ctr > 1 Then
    txt300.BackColor = &HFF&
ElseIf case23 = 400 And ctr > 1 Then
    txt400.BackColor = &HFF&
ElseIf case23 = 500 And ctr > 1 Then
    txt500.BackColor = &HFF&
ElseIf case23 = 750 And ctr > 1 Then
    txt750.BackColor = &HFF&
ElseIf case23 = 1000 And ctr > 1 Then
    txt1000.BackColor = &HFF&
ElseIf case23 = 5000 And ctr > 1 Then
    txt5000.BackColor = &HFF&
ElseIf case23 = 10000 And ctr > 1 Then
    txt10000.BackColor = &HFF&
ElseIf case23 = 25000 And ctr > 1 Then
    txt25000.BackColor = &HFF&
ElseIf case23 = 50000 And ctr > 1 Then
    txt50000.BackColor = &HFF&
ElseIf case23 = 75000 And ctr > 1 Then
    txt75000.BackColor = &HFF&
ElseIf case23 = 100000 And ctr > 1 Then
    txt100000.BackColor = &HFF&
ElseIf case23 = 200000 And ctr > 1 Then
    txt200000.BackColor = &HFF&
ElseIf case23 = 300000 And ctr > 1 Then
    txt300000.BackColor = &HFF&
ElseIf case23 = 400000 And ctr > 1 Then
    txt400000.BackColor = &HFF&
ElseIf case23 = 500000 And ctr > 1 Then
    txt500000.BackColor = &HFF&
ElseIf case23 = 750000 Then
    txt750000.BackColor = &HFF&
ElseIf case23 = 1000000 And ctr > 1 Then
    txt1000000.BackColor = &HFF&
End If
If ctr = 1 Then                                 'sets first case as variable "playerscase"
playerscase = value(23)
ElseIf ctr = 6 Then                             'choose 5 boxes to recieve bankers offer
MsgBox "The Banker has offered" & bankersoffer          'Breaks the game into segments allowing the banker to make offers to the player
picoffer.Print ; bankersoffer;
ElseIf ctr = 10 Then                                    'choose 4 more boxes for next offer
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 14 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 17 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 20 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 21 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 22 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 23 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 24 Then                                    'causes the form to switch to the final form in which the player can choose to keep their own case or pick the final remaining case on the board
frmgamescreen.Hide
frmgamefinal.Show
casefinal = value(23)                                  'sets the value of casefinal for use in the next form
End If
End Sub

Private Sub cmdCase24_Click()
Dim case24 As Long
cmdCase24.Visible = False                              'causes case/cmdbutton to dissapear
txtpickfirst.Visible = False
txtpickcase.Visible = True


case24 = value(24)                                     'sets values of variables
sum = sum - value(24)                                  'subtracts value represented by case from total sum of all cases
ctr = ctr + 1
bankersoffer = sum / (25 - ctr)
If case24 = 1 And ctr > 1 Then                                  'Makes the coinciding txt box turn red, coinciding with the value represented by the case
    txt1.BackColor = &HFF&
ElseIf case24 = 5 And ctr > 1 Then
    txt5.BackColor = &HFF&
ElseIf case24 = 10 And ctr > 1 Then
    txt10.BackColor = &HFF&
ElseIf case24 = 25 And ctr > 1 Then
    txt25.BackColor = &HFF&
ElseIf case24 = 50 And ctr > 1 Then
    txt50.BackColor = &HFF&
ElseIf case24 = 75 And ctr > 1 Then
    txt75.BackColor = &HFF&
ElseIf case24 = 100 And ctr > 1 Then
    txt100.BackColor = &HFF&
ElseIf case24 = 200 And ctr > 1 Then
    txt200.BackColor = &HFF&
ElseIf case24 = 300 And ctr > 1 Then
    txt300.BackColor = &HFF&
ElseIf case24 = 400 And ctr > 1 Then
    txt400.BackColor = &HFF&
ElseIf case24 = 500 And ctr > 1 Then
    txt500.BackColor = &HFF&
ElseIf case24 = 750 And ctr > 1 Then
    txt750.BackColor = &HFF&
ElseIf case24 = 1000 And ctr > 1 Then
    txt1000.BackColor = &HFF&
ElseIf case24 = 5000 And ctr > 1 Then
    txt5000.BackColor = &HFF&
ElseIf case24 = 10000 And ctr > 1 Then
    txt10000.BackColor = &HFF&
ElseIf case24 = 25000 And ctr > 1 Then
    txt25000.BackColor = &HFF&
ElseIf case24 = 50000 And ctr > 1 Then
    txt50000.BackColor = &HFF&
ElseIf case24 = 75000 And ctr > 1 Then
    txt75000.BackColor = &HFF&
ElseIf case24 = 100000 And ctr > 1 Then
    txt100000.BackColor = &HFF&
ElseIf case24 = 200000 And ctr > 1 Then
    txt200000.BackColor = &HFF&
ElseIf case24 = 300000 And ctr > 1 Then
    txt300000.BackColor = &HFF&
ElseIf case24 = 400000 And ctr > 1 Then
    txt400000.BackColor = &HFF&
ElseIf case24 = 500000 And ctr > 1 Then
    txt500000.BackColor = &HFF&
ElseIf case24 = 750000 Then
    txt750000.BackColor = &HFF&
ElseIf case24 = 1000000 And ctr > 1 Then
    txt1000000.BackColor = &HFF&
End If
If ctr = 1 Then                                 'sets first case as variable "playerscase"
playerscase = value(24)
ElseIf ctr = 6 Then                             'choose 5 boxes to recieve bankers offer
MsgBox "The Banker has offered" & bankersoffer          'Breaks the game into segments allowing the banker to make offers to the player
picoffer.Print ; bankersoffer;
ElseIf ctr = 10 Then                                    'choose 4 more boxes for next offer
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 14 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 17 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 20 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 21 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 22 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 23 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 24 Then                                    'causes the form to switch to the final form in which the player can choose to keep their own case or pick the final remaining case on the board
frmgamescreen.Hide
frmgamefinal.Show
casefinal = value(24)                                   'sets the value of casefinal for use in the next form
End If
End Sub

Private Sub cmdCase25_Click()
Dim case25 As Long
cmdCase25.Visible = False                                'causes case/cmdbutton to dissapear
txtpickfirst.Visible = False
txtpickcase.Visible = True


case25 = value(25)                                       'sets values of variables
sum = sum - value(25)                                    'subtracts value represented by case from total sum of all cases
ctr = ctr + 1
bankersoffer = sum / (25 - ctr)
If case25 = 25 And ctr > 1 Then                                        'Makes the coinciding txt box turn red, coinciding with the value represented by the case
    txt1.BackColor = &HFF&
ElseIf case25 = 5 And ctr > 1 Then
    txt5.BackColor = &HFF&
ElseIf case25 = 10 And ctr > 1 Then
    txt10.BackColor = &HFF&
ElseIf case25 = 25 And ctr > 1 Then
    txt25.BackColor = &HFF&
ElseIf case25 = 50 And ctr > 1 Then
    txt50.BackColor = &HFF&
ElseIf case25 = 75 And ctr > 1 Then
    txt75.BackColor = &HFF&
ElseIf case25 = 100 And ctr > 1 Then
    txt100.BackColor = &HFF&
ElseIf case25 = 200 And ctr > 1 Then
    txt200.BackColor = &HFF&
ElseIf case25 = 300 And ctr > 1 Then
    txt300.BackColor = &HFF&
ElseIf case25 = 400 And ctr > 1 Then
    txt400.BackColor = &HFF&
ElseIf case25 = 500 And ctr > 1 Then
    txt500.BackColor = &HFF&
ElseIf case25 = 750 And ctr > 1 Then
    txt750.BackColor = &HFF&
ElseIf case25 = 1000 And ctr > 1 Then
    txt1000.BackColor = &HFF&
ElseIf case25 = 5000 And ctr > 1 Then
    txt5000.BackColor = &HFF&
ElseIf case25 = 10000 And ctr > 1 Then
    txt10000.BackColor = &HFF&
ElseIf case25 = 25000 And ctr > 1 Then
    txt25000.BackColor = &HFF&
ElseIf case25 = 50000 And ctr > 1 Then
    txt50000.BackColor = &HFF&
ElseIf case25 = 75000 And ctr > 1 Then
    txt75000.BackColor = &HFF&
ElseIf case25 = 100000 And ctr > 1 Then
    txt100000.BackColor = &HFF&
ElseIf case25 = 200000 And ctr > 1 Then
    txt200000.BackColor = &HFF&
ElseIf case25 = 300000 And ctr > 1 Then
    txt300000.BackColor = &HFF&
ElseIf case25 = 400000 And ctr > 1 Then
    txt400000.BackColor = &HFF&
ElseIf case25 = 500000 And ctr > 1 Then
    txt500000.BackColor = &HFF&
ElseIf case25 = 750000 Then
    txt750000.BackColor = &HFF&
ElseIf case25 = 1000000 And ctr > 1 Then
    txt1000000.BackColor = &HFF&
End If
If ctr = 1 Then                                 'sets first case as variable "playerscase"
playerscase = value(25)
ElseIf ctr = 6 Then                             'choose 5 boxes to recieve bankers offer
MsgBox "The Banker has offered" & bankersoffer          'Breaks the game into segments allowing the banker to make offers to the player
picoffer.Print ; bankersoffer;
ElseIf ctr = 10 Then                                'choose 4 more boxes for next offer
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 14 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 17 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 20 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 21 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 22 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 23 Then
MsgBox "The Banker has offered" & bankersoffer
picoffer.Print ; bankersoffer;
ElseIf ctr = 24 Then                        'causes the form to switch to the final form in which the player can choose to keep their own case or pick the final remaining case on the board
frmgamescreen.Hide
frmgamefinal.Show
casefinal = value(25)                       'sets the value of casefinal for use in the next form
End If
End Sub

Private Sub cmdDEAL_Click()
frmgamescreen.Hide
frmdeal.Show

End Sub



