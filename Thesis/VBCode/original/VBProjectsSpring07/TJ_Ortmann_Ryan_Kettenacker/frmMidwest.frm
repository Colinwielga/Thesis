VERSION 5.00
Begin VB.Form frmMidwest 
   BackColor       =   &H000080FF&
   Caption         =   "Midwest Regional"
   ClientHeight    =   8745
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11100
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   8745
   ScaleWidth      =   11100
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00FF0000&
      Height          =   1695
      Left            =   6960
      ScaleHeight     =   1635
      ScaleWidth      =   2955
      TabIndex        =   39
      Top             =   2760
      Width           =   3015
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   1575
         Left            =   -120
         TabIndex        =   40
         Top             =   0
         Width           =   3015
      End
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
      Left            =   9000
      TabIndex        =   37
      Top             =   6960
      Width           =   1695
   End
   Begin VB.CommandButton cmdGoToWest 
      Caption         =   "Go To West Bracket"
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
      Left            =   9000
      TabIndex        =   36
      Top             =   6240
      Width           =   1695
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      Height          =   2415
      Left            =   7800
      Picture         =   "frmMidwest.frx":0000
      ScaleHeight     =   2355
      ScaleWidth      =   2355
      TabIndex        =   35
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
      Left            =   8160
      TabIndex        =   33
      Top             =   7680
      Width           =   1695
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
      TabIndex        =   32
      Top             =   6960
      Width           =   1695
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
      Left            =   7200
      TabIndex        =   31
      Top             =   6240
      Width           =   1695
   End
   Begin VB.CommandButton cmdMidwestWinner 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   4680
      Width           =   1455
   End
   Begin VB.CommandButton cmdwinner13 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   2640
      Width           =   1455
   End
   Begin VB.CommandButton cmdwinner14 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5040
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   6120
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
      Top             =   3600
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
      Top             =   5160
      Width           =   1455
   End
   Begin VB.CommandButton cmdwinner4 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   3960
      Width           =   1455
   End
   Begin VB.CommandButton cmdwinner3 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   3240
      Width           =   1455
   End
   Begin VB.CommandButton cmdTexasAMCC 
      BackColor       =   &H00FFFFFF&
      Caption         =   "15 Tex AM CC"
      Height          =   255
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   7800
      Width           =   1455
   End
   Begin VB.CommandButton cmdWisconsin 
      BackColor       =   &H00FFFFFF&
      Caption         =   "2 Wisconsin"
      Height          =   255
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   7560
      Width           =   1455
   End
   Begin VB.CommandButton cmdGeorgiaTech 
      BackColor       =   &H00FFFFFF&
      Caption         =   "10 Georgia Tech"
      Height          =   255
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   7080
      Width           =   1455
   End
   Begin VB.CommandButton cmdUNLV 
      BackColor       =   &H00FFFFFF&
      Caption         =   "7 UNLV"
      Height          =   255
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   6840
      Width           =   1455
   End
   Begin VB.CommandButton cmdMiamiOh 
      BackColor       =   &H00FFFFFF&
      Caption         =   "14 Miami (Oh)"
      Height          =   255
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   5880
      Width           =   1455
   End
   Begin VB.CommandButton cmdOregon 
      BackColor       =   &H00FFFFFF&
      Caption         =   "3 Oregon"
      Height          =   255
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   5640
      Width           =   1455
   End
   Begin VB.CommandButton cmdWinthrop 
      BackColor       =   &H00FFFFFF&
      Caption         =   "11 Winthrop"
      Height          =   255
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5280
      Width           =   1455
   End
   Begin VB.CommandButton cmdNotreDame 
      BackColor       =   &H00FFFFFF&
      Caption         =   "6 Notre Dame"
      Height          =   255
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5040
      Width           =   1455
   End
   Begin VB.CommandButton cmdDavidson 
      BackColor       =   &H00FFFFFF&
      Caption         =   "13 Davidson"
      Height          =   255
      Left            =   240
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4080
      Width           =   1455
   End
   Begin VB.CommandButton cmdMaryland 
      BackColor       =   &H00FFFFFF&
      Caption         =   "4 Maryland"
      Height          =   255
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3840
      Width           =   1455
   End
   Begin VB.CommandButton cmdOldDominion 
      BackColor       =   &H00FFFFFF&
      Caption         =   "12 Old Dominion"
      Height          =   255
      Left            =   240
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3360
      Width           =   1455
   End
   Begin VB.CommandButton cmdButler 
      BackColor       =   &H00FFFFFF&
      Caption         =   "5 Butler"
      Height          =   255
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3120
      Width           =   1455
   End
   Begin VB.CommandButton cmdwinner9 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1680
      Width           =   1455
   End
   Begin VB.CommandButton cmdwinner2 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2040
      Width           =   1455
   End
   Begin VB.CommandButton cmdPurdue 
      BackColor       =   &H00FFFFFF&
      Caption         =   "9 Purdue"
      Height          =   255
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2160
      Width           =   1455
   End
   Begin VB.CommandButton cmdArizona 
      BackColor       =   &H00FFFFFF&
      Caption         =   "8 Arizona"
      Height          =   255
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1920
      Width           =   1455
   End
   Begin VB.CommandButton cmdwinner1 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1320
      Width           =   1455
   End
   Begin VB.CommandButton cmdJacksonSt 
      BackColor       =   &H00FFFFFF&
      Caption         =   "16 Jackson St."
      Height          =   255
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1440
      Width           =   1455
   End
   Begin VB.CommandButton cmdFlorida 
      BackColor       =   &H00FFFFFF&
      Caption         =   "1 Florida"
      Height          =   255
      Left            =   240
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1200
      Width           =   1455
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00000000&
      Height          =   2295
      Left            =   7080
      ScaleHeight     =   2235
      ScaleWidth      =   3675
      TabIndex        =   38
      Top             =   6120
      Width           =   3735
   End
   Begin VB.Image Image1 
      Height          =   1500
      Left            =   8520
      Picture         =   "frmMidwest.frx":914D
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
      TabIndex        =   41
      Top             =   6480
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Midwest Regional Bracket"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   34
      Top             =   240
      Width           =   7575
   End
End
Attribute VB_Name = "frmMidwest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'This form is where the user will begin to start and select who they believe will win each round
' this form specifically is for the midwest bracket

'Below this button is used to compute the totals of each round in the the midwest bracket
Private Sub cmdCompute_Click()
    Dim MidwestR1(1 To 8) As String                         'making an array of midwest round 1 winners
    Dim MidwestR1Pos(1 To 8) As Integer                     'this is for if we want to do different scoring and have upsets be worth more points, in this way we already have all rankings of teams in an array
    Dim MidwestR2(1 To 4) As String                         'the following arrays are the same except for round 2 and 3
    Dim MidwestR2Pos(1 To 4) As Integer
    Dim MidwestR3(1 To 2) As String
    Dim MidwestR3Pos(1 To 2) As Integer
    Dim Ctr As Integer, CTR2 As Integer, CTR3 As Integer       'counters for each array
    
    Open App.Path & "\MidwestR1.txt" For Input As #1            'notepad for midwest round 1 winners
    Ctr = 0
    Do Until EOF(1)
        Ctr = Ctr + 1
        Input #1, MidwestR1Pos(Ctr), MidwestR1(Ctr)
    Loop
    Close #1
    
    MidwestR1Sum = 0                                        'set sum = to 0
    If cmdwinner1.Caption = MidwestR1(1) Then               'this section uses if statements to add points to sum if the if statement is proven to be true
        MidwestR1Sum = MidwestR1Sum + 1
    End If
    If cmdwinner2.Caption = MidwestR1(2) Then
        MidwestR1Sum = MidwestR1Sum + 1
    End If
    If cmdwinner3.Caption = MidwestR1(3) Then
        MidwestR1Sum = MidwestR1Sum + 1
    End If
    If cmdwinner4.Caption = MidwestR1(4) Then
        MidwestR1Sum = MidwestR1Sum + 1
    End If
    If cmdwinner5.Caption = MidwestR1(5) Then
        MidwestR1Sum = MidwestR1Sum + 1
    End If
    If cmdwinner6.Caption = MidwestR1(6) Then
        MidwestR1Sum = MidwestR1Sum + 1
    End If
    If cmdwinner7.Caption = MidwestR1(7) Then
        MidwestR1Sum = MidwestR1Sum + 1
    End If
    If cmdwinner8.Caption = MidwestR1(8) Then
        MidwestR1Sum = MidwestR1Sum + 1
    End If
    
    Open App.Path & "\MidwestR2.txt" For Input As #2        'opening round 2 winners form midwest
    CTR2 = 0
    Do Until EOF(2)
        CTR2 = CTR2 + 1
        Input #2, MidwestR2Pos(CTR2), MidwestR2(CTR2)
    Loop
    Close #2
    
    MidwestR2Sum = 0                                           'same as above setting round 2 midwest sum to 0
    If cmdwinner9.Caption = MidwestR2(1) Then                   'if the if statement is proven true then points are added to sum
        MidwestR2Sum = MidwestR2Sum + 2
    End If
    If cmdwinner10.Caption = MidwestR2(2) Then
        MidwestR2Sum = MidwestR2Sum + 2
    End If
    If cmdwinner11.Caption = MidwestR2(3) Then
        MidwestR2Sum = MidwestR2Sum + 2
    End If
    If cmdwinner12.Caption = MidwestR2(4) Then
        MidwestR2Sum = MidwestR2Sum + 2
    End If
    
    Open App.Path & "\MidwestR3.txt" For Input As #3
    CTR3 = 0
    Do Until EOF(3)
        CTR3 = CTR3 + 1
        Input #3, MidwestR3Pos(CTR3), MidwestR3(CTR3)
    Loop
    Close #3
    
    MidwestR3Sum = 0
    If cmdwinner13.Caption = MidwestR3(1) Then
        MidwestR3Sum = MidwestR3Sum + 4
    End If
    If cmdwinner14.Caption = MidwestR3(2) Then
        MidwestR3Sum = MidwestR3Sum + 4
    End If
    
    MidwestTotal = MidwestR1Sum + MidwestR2Sum + MidwestR3Sum           'final total for midwest
    
End Sub

'The following buttons are to allow the user to go from the midwest form to any of the follwing east, west, south, finals forms
Private Sub cmdGoToEast_Click()
    frmMidwest.Hide
    frmEast.Show
End Sub
Private Sub cmdGoToFinalFour_Click()
    frmMidwest.Hide
    frmFinals.Show
End Sub
Private Sub cmdGoToSouth_Click()
    frmMidwest.Hide
    frmSouth.Show
End Sub
Private Sub cmdGoToWest_Click()
    frmMidwest.Hide
    frmWest.Show
End Sub

'The rest of the buttons are used to make captions from the clicked on winner transfer to the next rounds button
'this is done by setting the caption equal to the caption of the next rounds button
Private Sub cmdFlorida_Click()
    cmdwinner1.Caption = "1 Florida"
End Sub
Private Sub cmdJacksonSt_Click()
    cmdwinner1.Caption = "16 Jackson St"
End Sub
Private Sub cmdArizona_Click()
    cmdwinner2.Caption = "8 Arizona"
End Sub

'this button is special because it will allow us to transfer midwest winner caption to the finals form
Private Sub cmdMidwestWinner_Click()
    MidwestWinner = cmdMidwestWinner.Caption
End Sub

Private Sub cmdPurdue_Click()
    cmdwinner2.Caption = "9 Purdue"
End Sub
Private Sub cmdButler_Click()
    cmdwinner3.Caption = "5 Butler"
End Sub
Private Sub cmdOldDominion_Click()
    cmdwinner3.Caption = "12 Old Dominion"
End Sub
Private Sub cmdMaryland_Click()
    cmdwinner4.Caption = "4 Maryland"
End Sub
Private Sub cmdDavidson_Click()
    cmdwinner4.Caption = "13 Davidson"
End Sub
Private Sub cmdNotreDame_Click()
    cmdwinner5.Caption = "6 Notre Dame"
End Sub

Private Sub cmdwinner1_Click()
    cmdwinner9.Caption = cmdwinner1.Caption
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
    cmdMidwestWinner.Caption = cmdwinner13.Caption
End Sub

Private Sub cmdwinner14_Click()
    cmdMidwestWinner.Caption = cmdwinner14.Caption
End Sub

Private Sub cmdwinner2_Click()
    cmdwinner9.Caption = cmdwinner1.Caption
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

Private Sub cmdWinthrop_Click()
    cmdwinner5.Caption = "11 Winthrop"
End Sub
Private Sub cmdOregon_Click()
    cmdwinner6.Caption = "3 Oregon"
End Sub
Private Sub cmdMiamiOh_Click()
    cmdwinner6.Caption = "14 Miami (Oh)"
End Sub
Private Sub cmdUNLV_Click()
    cmdwinner7.Caption = "7 UNLV"
End Sub
Private Sub cmdGeorgiaTech_Click()
    cmdwinner7.Caption = "10 Geogia Tech"
End Sub
Private Sub cmdWisconsin_Click()
    cmdwinner8.Caption = "2 Wisconsin"
End Sub
Private Sub cmdTexasAMCC_Click()
    cmdwinner8.Caption = "15 Texas AM CC"
End Sub




