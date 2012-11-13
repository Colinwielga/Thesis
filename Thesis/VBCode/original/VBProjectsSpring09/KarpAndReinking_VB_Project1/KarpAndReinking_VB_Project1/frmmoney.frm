VERSION 5.00
Begin VB.Form frmmoney 
   BackColor       =   &H00000000&
   Caption         =   "Money"
   ClientHeight    =   9885
   ClientLeft      =   2400
   ClientTop       =   840
   ClientWidth     =   11310
   LinkTopic       =   "Form1"
   Picture         =   "frmmoney.frx":0000
   ScaleHeight     =   9885
   ScaleWidth      =   11310
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   495
      Left            =   10080
      TabIndex        =   28
      Top             =   9120
      Width           =   975
   End
   Begin VB.CommandButton cmdBanker 
      Caption         =   "THE BANKER IS CALLING"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4440
      TabIndex        =   27
      Top             =   8640
      Width           =   1935
   End
   Begin VB.CommandButton cmdReturn 
      BackColor       =   &H000000FF&
      Caption         =   "Return to Cases"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4440
      TabIndex        =   26
      Top             =   240
      Width           =   1935
   End
   Begin VB.CommandButton cmdmoney26 
      BackColor       =   &H0000C0C0&
      Caption         =   "$1,000,000.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7080
      MaskColor       =   &H0000C0C0&
      TabIndex        =   25
      Top             =   7320
      Width           =   2535
   End
   Begin VB.CommandButton cmdmoney13 
      Caption         =   "$750.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   24
      Top             =   7320
      Width           =   2535
   End
   Begin VB.CommandButton cmdmoney25 
      Caption         =   "$750,000.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7080
      TabIndex        =   23
      Top             =   6720
      Width           =   2535
   End
   Begin VB.CommandButton cmdmoney24 
      Caption         =   "$500,000.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7080
      TabIndex        =   22
      Top             =   6120
      Width           =   2535
   End
   Begin VB.CommandButton cmdmoney23 
      Caption         =   "$400,000.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7080
      TabIndex        =   21
      Top             =   5520
      Width           =   2535
   End
   Begin VB.CommandButton cmdmoney22 
      Caption         =   "$300,000.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7080
      TabIndex        =   20
      Top             =   4920
      Width           =   2535
   End
   Begin VB.CommandButton cmdmoney21 
      Caption         =   "$200,000.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7080
      TabIndex        =   19
      Top             =   4320
      Width           =   2535
   End
   Begin VB.CommandButton cmdmoney20 
      Caption         =   "$100,000.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7080
      TabIndex        =   18
      Top             =   3720
      Width           =   2535
   End
   Begin VB.CommandButton cmdmoney19 
      Caption         =   "$75,000.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7080
      TabIndex        =   17
      Top             =   3120
      Width           =   2535
   End
   Begin VB.CommandButton cmdmoney18 
      Caption         =   "$50,000.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7080
      TabIndex        =   16
      Top             =   2520
      Width           =   2535
   End
   Begin VB.CommandButton cmdmoney14 
      Caption         =   "$1,000.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7080
      TabIndex        =   15
      Top             =   120
      Width           =   2535
   End
   Begin VB.CommandButton cmdmoney17 
      Caption         =   "$25,000.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7080
      TabIndex        =   14
      Top             =   1920
      Width           =   2535
   End
   Begin VB.CommandButton cmdmoney16 
      Caption         =   "$10,000.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7080
      TabIndex        =   13
      Top             =   1320
      Width           =   2535
   End
   Begin VB.CommandButton cmdmoney15 
      Caption         =   "$5,000.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7080
      TabIndex        =   12
      Top             =   720
      Width           =   2535
   End
   Begin VB.CommandButton cmdmoney12 
      Caption         =   "$500.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   11
      Top             =   6720
      Width           =   2535
   End
   Begin VB.CommandButton cmdmoney11 
      Caption         =   "$400.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   10
      Top             =   6120
      Width           =   2535
   End
   Begin VB.CommandButton cmdmoney10 
      Caption         =   "$300.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   9
      Top             =   5520
      Width           =   2535
   End
   Begin VB.CommandButton cmdmoney9 
      Caption         =   "$200.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   8
      Top             =   4920
      Width           =   2535
   End
   Begin VB.CommandButton cmdmoney8 
      Caption         =   "$100.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   7
      Top             =   4320
      Width           =   2535
   End
   Begin VB.CommandButton cmdmoney7 
      Caption         =   "$75.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   6
      Top             =   3720
      Width           =   2535
   End
   Begin VB.CommandButton cmdmoney6 
      Caption         =   "$50.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   5
      Top             =   3120
      Width           =   2535
   End
   Begin VB.CommandButton cmdmoney5 
      Caption         =   "$25.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   4
      Top             =   2520
      Width           =   2535
   End
   Begin VB.CommandButton cmdmoney4 
      Caption         =   "$10.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   3
      Top             =   1920
      Width           =   2535
   End
   Begin VB.CommandButton cmdmoney3 
      Caption         =   "$5.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   2
      Top             =   1320
      Width           =   2535
   End
   Begin VB.CommandButton cmdmoney2 
      Caption         =   "$1.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   1
      Top             =   720
      Width           =   2535
   End
   Begin VB.CommandButton cmdmoney1 
      Caption         =   "$0.01"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   0
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label lblbanker 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Click below to hear the sounds of the banker calling"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   615
      Left            =   4560
      TabIndex        =   30
      Top             =   7080
      Width           =   1695
   End
   Begin VB.OLE OLE1 
      Class           =   "Package"
      Height          =   615
      Left            =   4920
      OleObjectBlob   =   "frmmoney.frx":2CEE8
      SourceDoc       =   "M:\CS130\Project\the.mp3"
      TabIndex        =   29
      Top             =   7800
      Width           =   855
   End
End
Attribute VB_Name = "frmmoney"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Avg As Single
'Project:Deal or No Deal
'frmmoney
'Holly Reinking and Danielle Karp
'Written 3/15/09
'Purpose: To show the user what money amounts they have left, to "call" the banker, and to figure out whether or not they should take the Deal


Private Sub cmdBanker_Click()           'Decides whether or not the user accepts or rejects the banker's offer, proposes an amount of money
Dim Accept As String


MsgBox ("The Banker is calling.... "), , "Ring, Ring, Ring"
MsgBox ("The offer is " & FormatCurrency(Avg, 2)), , "Are you going to accept it?"      'To have the banker call and make an offer
Accept = InputBox("Type in, Deal or No Deal ", "What's your answer?")                   'Prompt the player to select Deal or No Deal

frmmoney.cmdBanker.Enabled = False                                                      'Enable and Disable some buttons
frmmoney.cmdReturn.Enabled = True

If Accept = "Deal" Then
    MsgBox "Congratulations " & id & "!" & " Your winnings are " & FormatCurrency(Avg, 2) & " ! ", , " You have WON! " 'If user entered "Deal" We let them know what money they won
    MsgBox "Thank you for playing Deal or No Deal, please click QUIT now."
Else
    If K = 24 Then
        MsgBox "Let's check your first choice..."
        MsgBox "Congratulations " & id & "!" & " Your winnings are " & FormatCurrency(Good, 2) & " ! ", , "Go buy something new!"
        MsgBox "Thank you for playing Deal or No Deal, Please click QUIT now."
    End If                                                                                               'If the user entered "No Deal" we let them keep playing
    If K <> 24 Then                                                                                      'If it is their last case we then tell them what they have won
        MsgBox "Let's continue the game.", , "You're going to win"
            frmdealornodeal.Show
            frmmoney.Hide
    End If
End If

    
End Sub

Private Sub cmdmoney1_Click()       'Assigns an amount to a specific case
cmdmoney1 = 0.01
End Sub

Private Sub cmdmoney10_Click()      'Assigns an amount to a specific case
cmdmoney10 = 300
End Sub

Private Sub cmdmoney11_Click()      'Assigns an amount to a specific case
cmdmoney11 = 400
End Sub

Private Sub cmdmoney12_Click()      'Assigns an amount to a specific case
cmdmoney12 = 500
End Sub

Private Sub cmdmoney13_Click()      'Assigns an amount to a specific case
cmdmoney13 = 750
End Sub

Private Sub cmdmoney14_Click()      'Assigns an amount to a specific case
cmdmoney14 = 1000
End Sub
    
Private Sub cmdmoney15_Click()      'Assigns an amount to a specific case
cmdmoney15 = 5000
End Sub

Private Sub cmdmoney16_Click()      'Assigns an amount to a specific case
cmdmoney16 = 10000
End Sub

Private Sub cmdmoney17_Click()      'Assigns an amount to a specific case
cmdmoney17 = 25000
End Sub

Private Sub cmdmoney18_Click()      'Assigns an amount to a specific case
cmdmoney18 = 50000
End Sub

Private Sub cmdmoney19_Click()      'Assigns an amount to a specific case
cmdmoney19 = 75000
End Sub

Private Sub cmdmoney2_Click()       'Assigns an amount to a specific case
cmdmoney2 = 1
End Sub

Private Sub cmdmoney20_Click()      'Assigns an amount to a specific case
cmdmoney20 = 100000
End Sub

Private Sub cmdmoney21_Click()      'Assigns an amount to a specific case
cmdmoney21 = 200000
End Sub

Private Sub cmdmoney22_Click()      'Assigns an amount to a specific case
cmdmoney22 = 300000
End Sub

Private Sub cmdmoney23_Click()      'Assigns an amount to a specific case
cmdmoney23 = 400000
End Sub

Private Sub cmdmoney24_Click()      'Assigns an amount to a specific case
cmdmoney24 = 500000
End Sub

Private Sub cmdmoney25_Click()      'Assigns an amount to a specific case
cmdmoney25 = 750000
End Sub

Private Sub cmdmoney26_Click()      'Assigns an amount to a specific case
cmdmoney26 = 1000000
End Sub

Private Sub cmdmoney3_Click()       'Assigns an amount to a specific case
cmdmoney3 = 5
End Sub

Private Sub cmdmoney4_Click()       'Assigns an amount to a specific case
cmdmoney4 = 10
End Sub

Private Sub cmdmoney5_Click()       'Assigns an amount to a specific case
cmdmoney5 = 25
End Sub

Private Sub cmdmoney6_Click()       'Assigns an amount to a specific case
cmdmoney6 = 50
End Sub

Private Sub cmdmoney7_Click()       'Assigns an amount to a specific case
cmdmoney7 = 75
End Sub

Private Sub cmdmoney8_Click()       'Assigns an amount to a specific case
cmdmoney8 = 100
End Sub

Private Sub cmdmoney9_Click()        'Assigns an amount to a specific case
cmdmoney9 = 200
End Sub

Private Sub cmdQuit_Click()
    End                               'To Quit the Program
End Sub

Private Sub cmdReturn_Click()          'Holds an equation, with a counter, that decides how much money to offer the player. Also returns to forms

frmmoney.Hide
frmdealornodeal.Show                   'To clear our picbox, Hide one form and show another and to print in our picbox
frmdealornodeal.piccasenumber.Cls
frmdealornodeal.piccasenumber.Print Num



If K = 6 Then
        Avg = Int(Sum / (26 - (0.5 * K)))
        MsgBox "You have picked 6 cases.  Please click on 'THE BANKER IS CALLING'."     'To calculate the amount offered by the banker after 6 cases are picked
            frmmoney.cmdBanker.Enabled = True
            frmmoney.cmdReturn.Enabled = False
                frmdealornodeal.Hide
                frmmoney.Show
        
        
ElseIf K = 11 Then
        Avg = Int(Sum / (26 - (0.5 * K)))
        MsgBox "You have picked 5 cases.  Please click on, 'THE BANKER IS CALLING'. "     'To calculate the amount offered by the banker after 5 more cases are picked
            cmdBanker.Enabled = True
            cmdReturn.Enabled = False
                frmdealornodeal.Hide
                frmmoney.Show
    
        
ElseIf K = 15 Then
        Avg = Int(Sum / (26 - (0.5 * K)))
        MsgBox "You have picked 4 cases.  Please click on, 'THE BANKER IS CALLING'. "       'To calculate the amount offered by the banker after 4 more cases are picked
            cmdBanker.Enabled = True
            cmdReturn.Enabled = False
                frmdealornodeal.Hide
                frmmoney.Show
        
ElseIf K = 18 Then
        Avg = Int(Sum / (26 - (0.5 * K)))
        MsgBox "You have picked 3 cases.  Please click on, 'THE BANKER IS CALLING'. "       'To calculate the amount offered by the banker after 3 more cases are picked
            cmdBanker.Enabled = True
            cmdReturn.Enabled = False
                frmdealornodeal.Hide
                frmmoney.Show
        
        
ElseIf K = 20 Then
        Avg = Int(Sum / (26 - (0.5 * K)))
        MsgBox "You have picked 2 cases.  Please click on, 'THE BANKER IS CALLING'. "       'To calculate the amount offered by the banker after 2 more cases are picked
            cmdBanker.Enabled = True
            cmdReturn.Enabled = False
                frmdealornodeal.Hide
                frmmoney.Show
        
        
ElseIf K = 21 Then
        Avg = Int(Sum / (26 - (0.5 * K)))
        MsgBox "You have picked 1 case.  Please click on, 'THE BANKER IS CALLING'. "        'To calculate the amount offered by the banker after 1 more case is picked
            cmdBanker.Enabled = True
            cmdReturn.Enabled = False
                frmdealornodeal.Hide
                frmmoney.Show
       
ElseIf K = 22 Then
        Avg = Int(Sum / (26 - (0.5 * K)))
        MsgBox "You have picked 1 case.  Please click on, 'THE BANKER IS CALLING'. "        'To calculate the amount offered by the banker after 1 more case is picked
            cmdBanker.Enabled = True
            cmdReturn.Enabled = False
                frmdealornodeal.Hide
                frmmoney.Show
        
ElseIf K = 23 Then
        Avg = Int(Sum / (26 - (0.5 * K)))
        MsgBox "You have picked 1 case.  Please click on, 'THE BANKER IS CALLING'. "        'To calculate the amount offered by the banker after 1 more case is picked
            cmdBanker.Enabled = True
            cmdReturn.Enabled = False
                frmdealornodeal.Hide
                frmmoney.Show
                
ElseIf K = 24 Then
        Avg = Int(Sum / (26 - (0.5 * K)))
        MsgBox "You have picked 1 case.  Please click on, 'THE BANKER IS CALLING'. "        'To calculate the amount offered by the banker after 1 more case is picked
            cmdBanker.Enabled = True
            cmdReturn.Enabled = False
                frmdealornodeal.Hide
                frmmoney.Show
        

End If
 
        
End Sub

