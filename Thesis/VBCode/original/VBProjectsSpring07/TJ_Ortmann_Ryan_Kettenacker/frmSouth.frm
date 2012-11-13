VERSION 5.00
Begin VB.Form frmSouth 
   BackColor       =   &H000080FF&
   Caption         =   "South Regional"
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
      Height          =   1575
      Left            =   6720
      ScaleHeight     =   1515
      ScaleWidth      =   3075
      TabIndex        =   39
      Top             =   2880
      Width           =   3135
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
         Height          =   1455
         Left            =   0
         TabIndex        =   40
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
      Left            =   9000
      TabIndex        =   37
      Top             =   6000
      Width           =   1695
   End
   Begin VB.CommandButton GoToFinalFour 
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
      TabIndex        =   36
      Top             =   6720
      Width           =   1695
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      Height          =   2415
      Left            =   7800
      Picture         =   "frmSouth.frx":0000
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
      Left            =   8280
      TabIndex        =   33
      Top             =   7440
      Width           =   1695
   End
   Begin VB.CommandButton cmdGoToMidwest 
      Caption         =   "Go To Midwest Bracket"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7200
      TabIndex        =   32
      Top             =   6000
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
      Left            =   7200
      TabIndex        =   31
      Top             =   6720
      Width           =   1695
   End
   Begin VB.CommandButton cmdSouthWinner 
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
      Top             =   7440
      Width           =   1455
   End
   Begin VB.CommandButton cmdwinner11 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   5520
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
      Top             =   7800
      Width           =   1455
   End
   Begin VB.CommandButton cmdwinner7 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   7080
      Width           =   1455
   End
   Begin VB.CommandButton cmdwinner6 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   5880
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
   Begin VB.CommandButton cmdNorthTexas 
      BackColor       =   &H00FFFFFF&
      Caption         =   "15 North Texas"
      Height          =   255
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   7920
      Width           =   1455
   End
   Begin VB.CommandButton cmdMemphis 
      BackColor       =   &H00FFFFFF&
      Caption         =   "2 Memphis"
      Height          =   255
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   7680
      Width           =   1455
   End
   Begin VB.CommandButton cmdCreighton 
      BackColor       =   &H00FFFFFF&
      Caption         =   "10 Creighton"
      Height          =   255
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   7200
      Width           =   1455
   End
   Begin VB.CommandButton cmdNevada 
      BackColor       =   &H00FFFFFF&
      Caption         =   "7 Nevada"
      Height          =   255
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   6960
      Width           =   1455
   End
   Begin VB.CommandButton cmdPenn 
      BackColor       =   &H00FFFFFF&
      Caption         =   "14 Penn"
      Height          =   255
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   6000
      Width           =   1455
   End
   Begin VB.CommandButton cmdTexasAM 
      BackColor       =   &H00FFFFFF&
      Caption         =   "3 Texas AM"
      Height          =   255
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   5760
      Width           =   1455
   End
   Begin VB.CommandButton cmdStanford 
      BackColor       =   &H00FFFFFF&
      Caption         =   "11 Stanford"
      Height          =   255
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5280
      Width           =   1455
   End
   Begin VB.CommandButton cmdLouisville 
      BackColor       =   &H00FFFFFF&
      Caption         =   "6 Louisville"
      Height          =   255
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5040
      Width           =   1455
   End
   Begin VB.CommandButton cmdAlbany 
      BackColor       =   &H00FFFFFF&
      Caption         =   "13 Albany"
      Height          =   255
      Left            =   240
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4080
      Width           =   1455
   End
   Begin VB.CommandButton cmdVirginia 
      BackColor       =   &H00FFFFFF&
      Caption         =   "4 Virginia"
      Height          =   255
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3840
      Width           =   1455
   End
   Begin VB.CommandButton cmdLongBeach 
      BackColor       =   &H00FFFFFF&
      Caption         =   "12 Long Beach"
      Height          =   255
      Left            =   240
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3360
      Width           =   1455
   End
   Begin VB.CommandButton cmdTennessee 
      BackColor       =   &H00FFFFFF&
      Caption         =   "5 Tennessee"
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
   Begin VB.CommandButton cmdXavier 
      BackColor       =   &H00FFFFFF&
      Caption         =   "9 Xavier"
      Height          =   255
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2160
      Width           =   1455
   End
   Begin VB.CommandButton cmdBYU 
      BackColor       =   &H00FFFFFF&
      Caption         =   "8 BYU"
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
   Begin VB.CommandButton cmdCConnSt 
      BackColor       =   &H00FFFFFF&
      Caption         =   "16 C Conn St"
      Height          =   255
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1440
      Width           =   1455
   End
   Begin VB.CommandButton cmdOhioSt 
      BackColor       =   &H00FFFFFF&
      Caption         =   "1 Ohio St"
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
      Top             =   5880
      Width           =   3735
   End
   Begin VB.Image Image1 
      Height          =   1500
      Left            =   8760
      Picture         =   "frmSouth.frx":914D
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
      BorderStyle     =   1  'Fixed Single
      Caption         =   "South Regional Bracket"
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
Attribute VB_Name = "frmSouth"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'This form is where the user will begin to start and select who they believe will win each round
' this form specifically is for the south bracket


Private Sub cmdCompute_Click()
    Dim SouthR1(1 To 8) As String                   'this array is for the names of all first round winners from south bracket
    Dim SouthR2(1 To 4) As String                   'this array is the rankings of all first round winners, this will be used if we decide later to do a scoring systems where rankings are needed to calculate upsets
    Dim SouthR3(1 To 2) As String                   'the other arrays are the same except for they are for round 2 and 3 winners of south region
    Dim SouthR1Pos(1 To 8) As Integer
    Dim SouthR2Pos(1 To 4) As Integer
    Dim SouthR3Pos(1 To 2) As Integer
    Dim Ctr As Integer, CTR2 As Integer, CTR3 As Integer        'setting counter for each round
    
    
    Open App.Path & "\SouthR1.txt" For Input As #1              'open notepad with first round winner
    Ctr = 0
    Do Until EOF(1)
        Ctr = Ctr + 1
        Input #1, SouthR1Pos(Ctr), SouthR1(Ctr)                 'setting counter to round 1 winners
    Loop
    Close #1
    
    SouthR1Sum = 0                                              'set sum to 0
    If cmdwinner1.Caption = SouthR1(1) Then                     'use if statements to add to total if the if statement is proven to be true
        SouthR1Sum = SouthR1Sum + 1
    End If
    If cmdwinner2.Caption = SouthR1(2) Then
        SouthR1Sum = SouthR1Sum + 1
    End If
    If cmdwinner3.Caption = SouthR1(3) Then
        SouthR1Sum = SouthR1Sum + 1
    End If
    If cmdwinner4.Caption = SouthR1(4) Then
        SouthR1Sum = SouthR1Sum + 1
    End If
    If cmdwinner5.Caption = SouthR1(5) Then
        SouthR1Sum = SouthR1Sum + 1
    End If
    If cmdwinner6.Caption = SouthR1(6) Then
        SouthR1Sum = SouthR1Sum + 1
    End If
    If cmdwinner7.Caption = SouthR1(7) Then
        SouthR1Sum = SouthR1Sum + 1
    End If
    If cmdwinner8.Caption = SouthR1(8) Then
        SouthR1Sum = SouthR1Sum + 1
    End If
    
    'the next to are the same as above except for round 2 and 3
   Open App.Path & "\SouthR2.txt" For Input As #2
        CTR2 = 0
        Do Until EOF(2)
            CTR2 = CTR2 + 1
            Input #2, SouthR2Pos(CTR2), SouthR2(CTR2)
        Loop
        Close #2
    
    SouthR2Sum = 0
    If cmdwinner9.Caption = SouthR2(1) Then
        SouthR2Sum = SouthR2Sum + 2
    End If
    If cmdwinner10.Caption = SouthR2(2) Then
        SouthR2Sum = SouthR2Sum + 2
    End If
    If cmdwinner11.Caption = SouthR2(3) Then
        SouthR2Sum = SouthR2Sum + 2
    End If
    If cmdwinner12.Caption = SouthR2(4) Then
        SouthR2Sum = SouthR2Sum + 2
    End If
    
    Open App.Path & "\SouthR3.txt" For Input As #3
        CTR3 = 0
        Do Until EOF(3)
            CTR3 = CTR3 + 1
            Input #3, SouthR3Pos(CTR3), SouthR3(CTR3)
        Loop
    Close #3
    
    SouthR3Sum = 0
    
    If cmdwinner13.Caption = SouthR3(1) Then
        SouthR3Sum = SouthR3Sum + 4
    End If
    If cmdwinner14.Caption = SouthR3(2) Then
        SouthR3Sum = SouthR3Sum + 4
    End If
    
    SouthTotal = SouthR1Sum + SouthR2Sum + SouthR3Sum       'total score for south bracket
   
End Sub
'will take caption from winner of south region and transfer to finals form
Private Sub cmdSouthWinner_Click()
    SouthWinner = cmdSouthWinner.Caption
End Sub

'these buttons will allow the user to freely go to finals, east, west, and midwest brackets
Private Sub GoToFinalFour_Click()
    frmSouth.Hide
    frmFinals.Show
End Sub
Private Sub cmdGoToEast_Click()
    frmSouth.Hide
    frmEast.Show
End Sub
Private Sub cmdGoToMidwest_Click()
    frmSouth.Hide
    frmMidwest.Show
End Sub
Private Sub cmdGoToWest_Click()
    frmSouth.Hide
    frmWest.Show
End Sub

'The rest of the buttons are used to make captions from the clicked on winner transfer to the next rounds button
'this is done by setting the caption equal to the caption of the next rounds button
Private Sub cmdOhioSt_Click()
    cmdwinner1.Caption = "1 Ohio St"
End Sub
Private Sub cmdCConnSt_Click()
    cmdwinner1.Caption = "16 C Conn St"
End Sub
Private Sub cmdBYU_Click()
    cmdwinner2.Caption = "8 BYU"
End Sub
Private Sub cmdXavier_Click()
    cmdwinner2.Caption = "9 Xavier"
End Sub
Private Sub cmdTennessee_Click()
    cmdwinner3.Caption = "5 Tennessee"
End Sub
Private Sub cmdLongBeach_Click()
    cmdwinner3.Caption = "12 LongBeach"
End Sub
Private Sub cmdVirginia_Click()
    cmdwinner4.Caption = "4 Virginia"
End Sub
Private Sub cmdAlbany_Click()
    cmdwinner4.Caption = "13 Albany"
End Sub
Private Sub cmdLouisville_Click()
    cmdwinner5.Caption = "6 Louisville"
End Sub
Private Sub cmdStanford_Click()
    cmdwinner5.Caption = "11 Stanford"
End Sub
Private Sub cmdTexasAM_Click()
    cmdwinner6.Caption = "3 Texas AM"
End Sub
Private Sub cmdPenn_Click()
    cmdwinner6.Caption = "14 Penn"
End Sub
Private Sub cmdNevada_Click()
    cmdwinner7.Caption = "7 Nevada"
End Sub
Private Sub cmdCreighton_Click()
    cmdwinner7.Caption = "10 Creighton"
End Sub
Private Sub cmdMemphis_Click()
    cmdwinner8.Caption = "2 Memphis"
End Sub
Private Sub cmdNorthTexas_Click()
    cmdwinner8.Caption = "15 North Texas"
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
    cmdSouthWinner.Caption = cmdwinner13.Caption
End Sub
Private Sub cmdwinner14_Click()
    cmdSouthWinner.Caption = cmdwinner14.Caption
End Sub

