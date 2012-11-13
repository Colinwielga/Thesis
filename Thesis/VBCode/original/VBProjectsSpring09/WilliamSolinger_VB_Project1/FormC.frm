VERSION 5.00
Begin VB.Form frmCalculations 
   BackColor       =   &H00FFFFC0&
   Caption         =   "Form1"
   ClientHeight    =   9570
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   14490
   LinkTopic       =   "Form1"
   ScaleHeight     =   9570
   ScaleWidth      =   14490
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCalculateAVG 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Calculate Average and Show Calculation"
      Height          =   855
      Left            =   9240
      Style           =   1  'Graphical
      TabIndex        =   47
      Top             =   6480
      Width           =   2655
   End
   Begin VB.PictureBox picAVGResults 
      BackColor       =   &H00C0C0FF&
      Height          =   1695
      Left            =   9240
      ScaleHeight     =   1635
      ScaleWidth      =   2595
      TabIndex        =   46
      Top             =   7560
      Width           =   2655
   End
   Begin VB.TextBox txtAtBats3 
      Height          =   495
      Left            =   3960
      TabIndex        =   41
      Top             =   7680
      Width           =   855
   End
   Begin VB.TextBox txtHitSauce 
      Height          =   495
      Left            =   2880
      TabIndex        =   39
      Top             =   7680
      Width           =   855
   End
   Begin VB.TextBox txtAtBats2 
      Height          =   495
      Left            =   7200
      TabIndex        =   35
      Top             =   4560
      Width           =   855
   End
   Begin VB.CommandButton cmdCalculateSLG 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Calculate Slugging Percentage and Show Calculation"
      Height          =   855
      Left            =   9240
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   3360
      Width           =   2655
   End
   Begin VB.PictureBox picSLGResults 
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   8520
      ScaleHeight     =   1635
      ScaleWidth      =   4035
      TabIndex        =   33
      Top             =   4440
      Width           =   4095
   End
   Begin VB.TextBox txtHomeRuns 
      Height          =   495
      Left            =   6120
      TabIndex        =   24
      Top             =   4560
      Width           =   855
   End
   Begin VB.TextBox txtTriples 
      Height          =   495
      Left            =   5040
      TabIndex        =   23
      Top             =   4560
      Width           =   855
   End
   Begin VB.TextBox txtDoubles 
      Height          =   495
      Left            =   3960
      TabIndex        =   22
      Top             =   4560
      Width           =   855
   End
   Begin VB.TextBox txtSingles 
      Height          =   495
      Left            =   2880
      TabIndex        =   21
      Top             =   4560
      Width           =   855
   End
   Begin VB.TextBox txtHBP 
      Height          =   495
      Left            =   5040
      TabIndex        =   19
      Top             =   1320
      Width           =   855
   End
   Begin VB.CommandButton cmdfrmSearch 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Go To the Searching Form"
      Height          =   2055
      Left            =   12840
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   3840
      Width           =   1455
   End
   Begin VB.CommandButton cmdfrmSort 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Go To the Sorting Form"
      Height          =   2055
      Left            =   12840
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   1560
      Width           =   1455
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Quit"
      Height          =   2055
      Left            =   12840
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   6120
      Width           =   1455
   End
   Begin VB.CommandButton cmdCalculateOBP 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Calculate On Base Percentage and Show Calculation"
      Height          =   855
      Left            =   9240
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   240
      Width           =   2655
   End
   Begin VB.PictureBox picOBPResults 
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   9240
      ScaleHeight     =   1635
      ScaleWidth      =   2595
      TabIndex        =   13
      Top             =   1320
      Width           =   2655
   End
   Begin VB.TextBox txtSacFlies 
      Height          =   495
      Left            =   7200
      TabIndex        =   7
      Top             =   1320
      Width           =   855
   End
   Begin VB.TextBox txtAtBats 
      Height          =   495
      Left            =   6120
      TabIndex        =   5
      Top             =   1320
      Width           =   855
   End
   Begin VB.TextBox txtWalks 
      Height          =   495
      Left            =   3960
      TabIndex        =   4
      Top             =   1320
      Width           =   855
   End
   Begin VB.TextBox txtHits 
      Height          =   495
      Left            =   2880
      TabIndex        =   0
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label lblEqualsSLG 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFC0&
      Caption         =   "="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   8040
      TabIndex        =   45
      Top             =   7320
      Width           =   285
   End
   Begin VB.Label lblLineC 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFC0&
      Caption         =   "---------------------"
      Height          =   195
      Left            =   6000
      TabIndex        =   44
      Top             =   7320
      Width           =   945
   End
   Begin VB.Label lblAVGAtBats 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFC0&
      Caption         =   "At Bats"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6000
      TabIndex        =   43
      Top             =   7560
      Width           =   855
   End
   Begin VB.Label lblAVGHits 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFC0&
      Caption         =   "Hits"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6240
      TabIndex        =   42
      Top             =   6960
      Width           =   465
   End
   Begin VB.Label txtAtBatz 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Input Number of At Bats"
      Height          =   855
      Left            =   3960
      TabIndex        =   40
      Top             =   6600
      Width           =   735
   End
   Begin VB.Label lblHitsauce 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Input Number of Hits"
      Height          =   855
      Left            =   2880
      TabIndex        =   38
      Top             =   6600
      Width           =   855
   End
   Begin VB.Label lblToCalcAVG 
      BackColor       =   &H00FFFFC0&
      Caption         =   "To Calculate Average                                            ---------------->"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   240
      TabIndex        =   37
      Top             =   6600
      Width           =   2415
   End
   Begin VB.Label lblAtBatz 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Input Number of At Bats"
      Height          =   735
      Left            =   7200
      TabIndex        =   36
      Top             =   3600
      Width           =   855
   End
   Begin VB.Label lblEqualsAVG 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFC0&
      Caption         =   "="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   8040
      TabIndex        =   32
      Top             =   5160
      Width           =   285
   End
   Begin VB.Label lblSLGCalc2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFC0&
      Caption         =   "At Bats"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4920
      TabIndex        =   31
      Top             =   5880
      Width           =   630
   End
   Begin VB.Label lblLineA 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFC0&
      Caption         =   "------------------------------------------------------------------------------------------------------------"
      Height          =   195
      Left            =   2880
      TabIndex        =   30
      Top             =   5640
      Width           =   4860
   End
   Begin VB.Label lblSLGCalc 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFC0&
      Caption         =   "(Singles) + (2 * Doubles) + (3 * Triples) + (4 * Home Runs)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2880
      TabIndex        =   29
      Top             =   5280
      Width           =   5010
   End
   Begin VB.Label lblHomeRuns 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Input Number of Home Runs"
      Height          =   735
      Left            =   6120
      TabIndex        =   28
      Top             =   3600
      Width           =   855
   End
   Begin VB.Label lblTriples 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Input Number of Triples"
      Height          =   735
      Left            =   5040
      TabIndex        =   27
      Top             =   3600
      Width           =   855
   End
   Begin VB.Label lblDoubles 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Input Number of Doubles"
      Height          =   735
      Left            =   3960
      TabIndex        =   26
      Top             =   3600
      Width           =   855
   End
   Begin VB.Label lblSingles 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Input Number of Singles"
      Height          =   735
      Left            =   2880
      TabIndex        =   25
      Top             =   3600
      Width           =   855
   End
   Begin VB.Label lblRawr 
      BackColor       =   &H00FFFFC0&
      Caption         =   "To Calculate Slugging Percentage             -------------->"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   360
      TabIndex        =   20
      Top             =   3600
      Width           =   2295
   End
   Begin VB.Label lblHBP 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Input Number of Hit By Pitches"
      Height          =   855
      Left            =   5040
      TabIndex        =   18
      Top             =   360
      Width           =   855
   End
   Begin VB.Label lblEqualsOBP 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFC0&
      Caption         =   "="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   8280
      TabIndex        =   12
      Top             =   2160
      Width           =   285
   End
   Begin VB.Label lblLineB 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFC0&
      Caption         =   "--------------------------------------------------------------------------------------------"
      Height          =   195
      Left            =   3480
      TabIndex        =   11
      Top             =   2160
      Width           =   4140
   End
   Begin VB.Label lblCalculationOBPBottom 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFC0&
      Caption         =   "(At Bats + Walks + HBP + Sacrifice Flies)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3360
      TabIndex        =   10
      Top             =   2400
      Width           =   4320
   End
   Begin VB.Label lblCalculationOBP 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFC0&
      Caption         =   "(Hits + Walks + HBP) "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4320
      TabIndex        =   9
      Top             =   1920
      Width           =   2280
   End
   Begin VB.Label lblE 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Input Number of Sacrifice Flies"
      Height          =   855
      Left            =   7200
      TabIndex        =   8
      Top             =   360
      Width           =   855
   End
   Begin VB.Label lblD 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Input Number of At Bats"
      Height          =   615
      Left            =   6120
      TabIndex        =   6
      Top             =   480
      Width           =   855
   End
   Begin VB.Label lblC 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Input Number of Walks"
      Height          =   615
      Left            =   3960
      TabIndex        =   3
      Top             =   480
      Width           =   855
   End
   Begin VB.Label lblB 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Input Number of  Hits"
      Height          =   615
      Left            =   2880
      TabIndex        =   2
      Top             =   480
      Width           =   855
   End
   Begin VB.Label lblOBP 
      BackColor       =   &H00FFFFC0&
      Caption         =   "To Calculate On Base Percentage              -------------->"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2235
      Left            =   360
      TabIndex        =   1
      Top             =   600
      Width           =   2295
   End
End
Attribute VB_Name = "frmCalculations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'2008 Minnesota Twins
'How are the calculations done?
'Bill Solinger
'March 24, 2009
'This form will show the user how to do the calculations to find OBP, SLG, and AVG.
'The user will input statistics into text boxes and calculate each statistic.

Private Sub cmdCalculateAVG_Click()
    'This form will take the values that the user puts into a text box and calculate the statistic Average.
    'The results box will also show the steps in how the calculation is done.
    
    Dim Hitz As Integer, AtBatz As Integer
    Dim AVGz As Single
    
    Hitz = txtHitSauce.Text
    AtBatz = txtAtBats3.Text
    
    AVGz = Hitz / AtBatz
    
    picAVGResults.Cls
    picAVGResults.Print Hitz
    picAVGResults.Print "---------"
    picAVGResults.Print AtBatz
    picAVGResults.Print
    picAVGResults.Print "AVG = "; FormatNumber(AVGz, 3)
    
End Sub

Private Sub cmdCalculateOBP_Click()
    'This form will take the values that the user puts into a text box and calculate the statistic On Base Percentage.
    'The results box will also show the steps in how the calculation is done.
    
    Dim Hitz As Integer, Walks As Integer, AtBats As Integer, SacFlies As Integer, HBP As Integer
    Dim OnBP As Single
    Hitz = txtHits.Text
    Walks = txtWalks.Text
    AtBats = txtAtBats.Text
    SacFlies = txtSacFlies.Text
    HBP = txtHBP.Text
    
    OnBP = (Hitz + Walks + HBP) / (AtBats + Walks + HBP + SacFlies)
    picOBPResults.Cls
    picOBPResults.Print "(" & Hitz & " + " & Walks & " + " & HBP & ")"
    picOBPResults.Print "------------------------"
    picOBPResults.Print "(" & AtBats & " + " & Walks & " + " & HBP & " + " & SacFlies & ")"
    picOBPResults.Print
    picOBPResults.Print "OBP = "; FormatNumber(OnBP, 3)
End Sub

Private Sub cmdCalculateSLG_Click()
    'This form will take the values that the user puts into a text box and calculate the statistic Average.
    'The results box will also show the steps in how the calculation is done.
    
    Dim Singles As Integer, Doubles As Integer, AtBats As Integer, Triples As Integer, HomeR As Integer
    Dim SluggingP As Single
    Singles = txtSingles.Text
    Doubles = txtDoubles.Text
    Triples = txtTriples.Text
    HomeR = txtHomeRuns.Text
    AtBats = txtAtBats2.Text
    
    SluggingP = ((Singles) + (2 * Doubles) + (3 * Triples) + (4 * HomeR)) / AtBats
    
    picSLGResults.Cls
    picSLGResults.Print "((" & Singles & ") + (2 * " & Doubles & ") + (3 * " & Triples & ") + (4 * " & HomeR & "))"
    picSLGResults.Print "-------------------------------------------------"
    picSLGResults.Print AtBats
    picSLGResults.Print
    picSLGResults.Print "SLG = "; FormatNumber(SluggingP, 3)
End Sub

Private Sub cmdfrmSearch_Click()
    'This button will bring the user to the Searching form.
    frmSearch.Visible = True
    frmCalculations.Visible = False
End Sub

Private Sub cmdfrmSort_Click()
    'This form will bring the user to the Sorting form.
    frmSort.Visible = True
    frmCalculations.Visible = False
End Sub

Private Sub cmdQuit_Click()
    End
End Sub
