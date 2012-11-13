VERSION 5.00
Begin VB.Form frmWest 
   BackColor       =   &H000080FF&
   Caption         =   "West Regional"
   ClientHeight    =   8745
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11100
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   8745
   ScaleWidth      =   11100
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdWestWinner 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   4680
      Width           =   1455
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00FF0000&
      Height          =   1575
      Left            =   6720
      ScaleHeight     =   1515
      ScaleWidth      =   3075
      TabIndex        =   38
      Top             =   2880
      Width           =   3135
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "PLEASE CLICK!!! ON WINNER OF REGION TO SUBMIT THAT WINNER TO FINAL FOUR BRACKET"
         BeginProperty Font 
            Name            =   "Papyrus"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   0
         TabIndex        =   39
         Top             =   0
         Width           =   3015
      End
   End
   Begin VB.CommandButton cmdGoToEast 
      Caption         =   "Go To East Bracket"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8880
      TabIndex        =   36
      Top             =   6120
      Width           =   1575
   End
   Begin VB.CommandButton cmdGoToMidwest 
      Caption         =   "Go To Midwest Bracket"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7200
      TabIndex        =   35
      Top             =   6840
      Width           =   1575
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00000000&
      Height          =   2415
      Left            =   7560
      Picture         =   "frmWest.frx":0000
      ScaleHeight     =   2355
      ScaleWidth      =   2355
      TabIndex        =   33
      Top             =   240
      Width           =   2415
   End
   Begin VB.CommandButton cmdCompute 
      Caption         =   "Compute Section Total"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8040
      TabIndex        =   32
      Top             =   7560
      Width           =   1575
   End
   Begin VB.CommandButton cmdGoToFinalFour 
      Caption         =   "Go To Final Four"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8880
      TabIndex        =   31
      Top             =   6840
      Width           =   1575
   End
   Begin VB.CommandButton cmdGoToSouth 
      Caption         =   "Go To South Bracket"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7200
      TabIndex        =   30
      Top             =   6120
      Width           =   1575
   End
   Begin VB.CommandButton cmdwinner13 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   2520
      Width           =   1455
   End
   Begin VB.CommandButton cmdwinner14 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5040
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   6000
      Width           =   1455
   End
   Begin VB.CommandButton cmdwinner12 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   7320
      Width           =   1455
   End
   Begin VB.CommandButton cmdwinner11 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   5400
      Width           =   1455
   End
   Begin VB.CommandButton cmdwinner10 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   3480
      Width           =   1455
   End
   Begin VB.CommandButton cmdwinner8 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   7680
      Width           =   1455
   End
   Begin VB.CommandButton cmdwinner7 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   6960
      Width           =   1455
   End
   Begin VB.CommandButton cmdwinner6 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   5760
      Width           =   1455
   End
   Begin VB.CommandButton cmdwinner5 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   5040
      Width           =   1455
   End
   Begin VB.CommandButton cmdwinner4 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   3840
      Width           =   1455
   End
   Begin VB.CommandButton cmdwinner3 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   3120
      Width           =   1455
   End
   Begin VB.CommandButton cmdWeberSt 
      BackColor       =   &H00FFFFFF&
      Caption         =   "15 Weber St"
      Height          =   255
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   7800
      Width           =   1455
   End
   Begin VB.CommandButton cmdUCLA 
      BackColor       =   &H00FFFFFF&
      Caption         =   "2 UCLA"
      Height          =   255
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   7560
      Width           =   1455
   End
   Begin VB.CommandButton cmdGonzaga 
      BackColor       =   &H00FFFFFF&
      Caption         =   "10 Gonzaga"
      Height          =   255
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   7080
      Width           =   1455
   End
   Begin VB.CommandButton cmdIndiana 
      BackColor       =   &H00FFFFFF&
      Caption         =   "7 Indiana"
      Height          =   255
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   6840
      Width           =   1455
   End
   Begin VB.CommandButton cmdWrightSt 
      BackColor       =   &H00FFFFFF&
      Caption         =   "14 Wright St"
      Height          =   255
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   5880
      Width           =   1455
   End
   Begin VB.CommandButton cmdPittsburgh 
      BackColor       =   &H00FFFFFF&
      Caption         =   "3 Pittsburgh"
      Height          =   255
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   5640
      Width           =   1455
   End
   Begin VB.CommandButton cmdVCU 
      BackColor       =   &H00FFFFFF&
      Caption         =   "11 VCU"
      Height          =   255
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5160
      Width           =   1455
   End
   Begin VB.CommandButton cmdDuke 
      BackColor       =   &H00FFFFFF&
      Caption         =   "6 Duke"
      Height          =   255
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4920
      Width           =   1455
   End
   Begin VB.CommandButton cmdHolyCross 
      BackColor       =   &H00FFFFFF&
      Caption         =   "13 HolyCross"
      Height          =   255
      Left            =   240
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3960
      Width           =   1455
   End
   Begin VB.CommandButton cmdSIllinois 
      BackColor       =   &H00FFFFFF&
      Caption         =   "4 S Illinois"
      Height          =   255
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3720
      Width           =   1455
   End
   Begin VB.CommandButton cmdIllinois 
      BackColor       =   &H00FFFFFF&
      Caption         =   "12 Illinois"
      Height          =   255
      Left            =   240
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3240
      Width           =   1455
   End
   Begin VB.CommandButton cmdVirginiaTech 
      BackColor       =   &H00FFFFFF&
      Caption         =   "5 Virginia Tech"
      Height          =   255
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3000
      Width           =   1455
   End
   Begin VB.CommandButton cmdwinner9 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1560
      Width           =   1455
   End
   Begin VB.CommandButton cmdwinner2 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1920
      Width           =   1455
   End
   Begin VB.CommandButton cmdVillanova 
      BackColor       =   &H00FFFFFF&
      Caption         =   "9 Villanova"
      Height          =   255
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2040
      Width           =   1455
   End
   Begin VB.CommandButton cmdKentucky 
      BackColor       =   &H00FFFFFF&
      Caption         =   "8 Kentucky"
      Height          =   255
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1800
      Width           =   1455
   End
   Begin VB.CommandButton cmdwinner1 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1200
      Width           =   1455
   End
   Begin VB.CommandButton cmdNiagra 
      BackColor       =   &H00FFFFFF&
      Caption         =   "16 Niagra"
      Height          =   255
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1320
      Width           =   1455
   End
   Begin VB.CommandButton cmdKansas 
      BackColor       =   &H00FFFFFF&
      Caption         =   "1 Kansas"
      Height          =   255
      Left            =   240
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1080
      Width           =   1455
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000012&
      Height          =   2295
      Left            =   7080
      ScaleHeight     =   2235
      ScaleWidth      =   3435
      TabIndex        =   37
      Top             =   6000
      Width           =   3495
   End
   Begin VB.Image Image1 
      Height          =   1500
      Left            =   8520
      Picture         =   "frmWest.frx":914D
      Top             =   4200
      Width           =   1500
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TO SEE YOUR SCORES CLICK!! ON SECTION TOTAL TO TRANSFER SCORE TO TOTALS PAGE AND SUBMIT BRACKET"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   4920
      TabIndex        =   40
      Top             =   6360
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "West Regional Bracket"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   34
      Top             =   240
      Width           =   7455
   End
End
Attribute VB_Name = "frmWest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'This form is where the user will begin to start and select who they believe will win each round
' this form specifically is for the west bracket


Private Sub cmdFinalFour_Click()
    frmWest.Hide                        'button to allow user to go from west bracket to
    frmFinals.Show
End Sub

Private Sub cmdCompute_Click()
    Dim WestR1(1 To 8) As String                        'array for winners of west bracket round 1
    Dim WestR1Pos(1 To 8) As Integer                    'this is all of the west first round winners rankings in case we need an array of rankings to calculate upsets for example
    Dim WestR2(1 To 4) As String                        'the other arrays are the same except for round 2 and 3 winners in the west
    Dim WestR2Pos(1 To 4) As Integer
    Dim WestR3(1 To 2) As String
    Dim WestR3Pos(1 To 2) As Integer
    Dim Ctr As Integer, CTR2 As Integer, CTR3 As Integer
    
    Open App.Path & "\WestR1.txt" For Input As #1           'notepade with round 1 west winners
    Ctr = 0
    Do Until EOF(1)
        Ctr = Ctr + 1
        Input #1, WestR1Pos(Ctr), WestR1(Ctr)               'setting ctr to rankings and names of winners
    Loop
    Close #1
    
    WestR1Sum = 0                                           'setting round 1 west sum to 0
    If cmdwinner1.Caption = WestR1(1) Then                  'if the if statment is true it will add to total
        WestR1Sum = WestR1Sum + 1
    End If
    If cmdwinner2.Caption = WestR1(2) Then
        WestR1Sum = WestR1Sum + 1
    End If
    If cmdwinner3.Caption = WestR1(3) Then
        WestR1Sum = WestR1Sum + 1
    End If
    If cmdwinner4.Caption = WestR1(4) Then
        WestR1Sum = WestR1Sum + 1
    End If
    If cmdwinner5.Caption = WestR1(5) Then
        WestR1Sum = WestR1Sum + 1
    End If
    If cmdwinner6.Caption = WestR1(6) Then
        WestR1Sum = WestR1Sum + 1
    End If
    If cmdwinner7.Caption = WestR1(7) Then
        WestR1Sum = WestR1Sum + 1
    End If
    If cmdwinner8.Caption = WestR1(8) Then
        WestR1Sum = WestR1Sum + 1
    End If
    
    'this part as well as the next is the same concept as above except for round 2 and 3 winners
    Open App.Path & "\WestR2.txt" For Input As #2
    CTR2 = 0
    Do Until EOF(2)
        CTR2 = CTR2 + 1
        Input #2, WestR2Pos(CTR2), WestR2(CTR2)
    Loop
    Close #2
    
    WestR2Sum = 0
    If cmdwinner9.Caption = WestR2(1) Then
        WestR2Sum = WestR2Sum + 2
    End If
    If cmdwinner10.Caption = WestR2(2) Then
        WestR2Sum = WestR2Sum + 2
    End If
    If cmdwinner11.Caption = WestR2(3) Then
        WestR2Sum = WestR2Sum + 2
    End If
    If cmdwinner12.Caption = WestR2(4) Then
        WestR2Sum = WestR2Sum + 2
    End If
    
    Open App.Path & "\WestR3.txt" For Input As #3
    CTR3 = 0
    Do Until EOF(3)
        CTR3 = CTR3 + 1
        Input #3, WestR3Pos(CTR3), WestR3(CTR3)
    Loop
    Close #3
    
    WestR3Sum = 0
    If cmdwinner13.Caption = WestR3(1) Then
        WestR3Sum = WestR3Sum + 4
    End If
    If cmdwinner14.Caption = WestR3(2) Then
        WestR3Sum = WestR3Sum + 4
    End If
    
    WestTotal = WestR1Sum + WestR2Sum + WestR3Sum           'total for west bracket
   
    
End Sub

'buttons used to allow user to free go to east, midwest, south, and finals forms
Private Sub cmdGoToEast_Click()
    frmWest.Hide
    frmEast.Show
End Sub

Private Sub cmdGoToFinalFour_Click()
    frmWest.Hide
    frmFinals.Show
End Sub

Private Sub cmdGoToMidwest_Click()
    frmWest.Hide
    frmMidwest.Show
End Sub
Private Sub cmdGoToSouth_Click()
    frmWest.Hide
    frmSouth.Show
End Sub

'The rest of the buttons are used to make captions from the clicked on winner transfer to the next rounds button
'this is done by setting the caption equal to the caption of the next rounds button
Private Sub cmdKansas_Click()
    cmdwinner1.Caption = "1 Kansas"
End Sub
Private Sub cmdNiagra_Click()
    cmdwinner1.Caption = "16 Jackson St"
End Sub
Private Sub cmdKentucky_Click()
    cmdwinner2.Caption = "8 Kentucky"
End Sub
Private Sub cmdVillanova_Click()
    cmdwinner2.Caption = "9 Villanova"
End Sub
Private Sub cmdVirginiaTech_Click()
    cmdwinner3.Caption = "5 Virginia Tech"
End Sub
Private Sub cmdIllinois_Click()
    cmdwinner3.Caption = "12 Old Illinois"
End Sub
Private Sub cmdSIllinois_Click()
    cmdwinner4.Caption = "4 S Illinois"
End Sub
Private Sub cmdHolyCross_Click()
    cmdwinner4.Caption = "13 Holy Cross"
End Sub
Private Sub cmdDuke_Click()
    cmdwinner5.Caption = "6 Duke"
End Sub
Private Sub cmdVCU_Click()
    cmdwinner5.Caption = "11 VCU"
End Sub
Private Sub cmdPittsburgh_Click()
    cmdwinner6.Caption = "3 Pittsburgh"
End Sub
'this button will allow the program to take winner of west bracket and transfer the caption to button on finals form
Private Sub cmdWestWinner_Click()
    WestWinner = cmdWestWinner.Caption
End Sub

Private Sub cmdWrightSt_Click()
    cmdwinner6.Caption = "14 Wright St"
End Sub
Private Sub cmdIndiana_Click()
    cmdwinner7.Caption = "7 Indiana"
End Sub
Private Sub cmdGonzaga_Click()
    cmdwinner7.Caption = "10 Gonzaga"
End Sub
Private Sub cmdUCLA_Click()
    cmdwinner8.Caption = "2 UCLA"
End Sub
Private Sub cmdTexasAMCC_Click()
    cmdwinner8.Caption = "15 Weber St"
End Sub

Private Sub cmdwinner1_Click()
    cmdwinner9.Caption = cmdwinner1.Caption
End Sub
Private Sub cmdwinner2_Click()
    cmdwinner9.Caption = cmdwinner2.Caption
End Sub
Private Sub cmdwinner3_Click()
    cmdwinner10.Caption = cmdwinner3.Caption
End Sub
Private Sub cmdwinner4_Click()
    cmdwinner10.Caption = cmdwinner4.Caption
End Sub
Private Sub cmdwinner5_Click()
    cmdwinner11.Caption = cmdwinner5.Caption
End Sub
Private Sub cmdwinner6_Click()
    cmdwinner11.Caption = cmdwinner6.Caption
End Sub
Private Sub cmdwinner7_Click()
    cmdwinner12.Caption = cmdwinner7.Caption
End Sub
Private Sub cmdwinner8_Click()
    cmdwinner12.Caption = cmdwinner8.Caption
End Sub
Private Sub cmdwinner9_Click()
   cmdwinner13.Caption = cmdwinner9.Caption
End Sub
Private Sub cmdwinner10_Click()
    cmdwinner13.Caption = cmdwinner10.Caption
End Sub
Private Sub cmdwinner11_Click()
    cmdwinner14.Caption = cmdwinner11.Caption
End Sub
Private Sub cmdwinner12_Click()
    cmdwinner14.Caption = cmdwinner12.Caption
End Sub
Private Sub cmdwinner13_Click()
    cmdWestWinner.Caption = cmdwinner13.Caption
End Sub
Private Sub cmdwinner14_Click()
    cmdWestWinner.Caption = cmdwinner14.Caption
End Sub

