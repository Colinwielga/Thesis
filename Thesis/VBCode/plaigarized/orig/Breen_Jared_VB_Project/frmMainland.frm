VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Main Screen"
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7740
   LinkTopic       =   "Form1"
   ScaleHeight     =   8490
   ScaleWidth      =   7740
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdNextYear 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Next Year"
      Height          =   975
      Left            =   5160
      MaskColor       =   &H00C0C0FF&
      TabIndex        =   23
      Top             =   6120
      Width           =   2295
   End
   Begin VB.PictureBox picProfits 
      Height          =   375
      Left            =   2640
      ScaleHeight     =   315
      ScaleWidth      =   2115
      TabIndex        =   22
      Top             =   7560
      Width           =   2175
   End
   Begin VB.PictureBox picExpenses 
      Height          =   375
      Left            =   2640
      ScaleHeight     =   315
      ScaleWidth      =   2115
      TabIndex        =   20
      Top             =   6840
      Width           =   2175
   End
   Begin VB.PictureBox picRevenue 
      Height          =   375
      Left            =   2640
      ScaleHeight     =   315
      ScaleWidth      =   2115
      TabIndex        =   18
      Top             =   6120
      Width           =   2175
   End
   Begin VB.PictureBox picSatisfaction 
      Height          =   375
      Left            =   3480
      ScaleHeight     =   315
      ScaleWidth      =   1275
      TabIndex        =   16
      Top             =   4320
      Width           =   1335
   End
   Begin VB.PictureBox picBalance 
      Height          =   375
      Left            =   3480
      ScaleHeight     =   315
      ScaleWidth      =   1275
      TabIndex        =   14
      Top             =   3240
      Width           =   1335
   End
   Begin VB.CommandButton cmdRates 
      Caption         =   "Rates"
      Height          =   975
      Left            =   5160
      TabIndex        =   12
      Top             =   3240
      Width           =   2295
   End
   Begin VB.PictureBox picProduction 
      Height          =   375
      Left            =   3480
      ScaleHeight     =   315
      ScaleWidth      =   1275
      TabIndex        =   11
      Top             =   1800
      Width           =   1335
   End
   Begin VB.PictureBox picDemand 
      Height          =   375
      Left            =   3480
      ScaleHeight     =   315
      ScaleWidth      =   1275
      TabIndex        =   8
      Top             =   2520
      Width           =   1335
   End
   Begin VB.CommandButton cmdRnD 
      Caption         =   "Research and Development"
      Height          =   975
      Left            =   5160
      TabIndex        =   7
      Top             =   4680
      Width           =   2295
   End
   Begin VB.CommandButton cmdProduction 
      Caption         =   "Production"
      Height          =   975
      Left            =   5160
      TabIndex        =   6
      Top             =   1800
      Width           =   2295
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   495
      Left            =   6240
      TabIndex        =   5
      Top             =   7800
      Width           =   1335
   End
   Begin VB.PictureBox picFunds 
      Height          =   375
      Left            =   2640
      ScaleHeight     =   315
      ScaleWidth      =   2115
      TabIndex        =   2
      Top             =   5400
      Width           =   2175
   End
   Begin VB.PictureBox picYear 
      Height          =   375
      Left            =   6960
      ScaleHeight     =   315
      ScaleWidth      =   555
      TabIndex        =   0
      Top             =   120
      Width           =   615
   End
   Begin VB.Label lblProfit 
      Caption         =   "Annual Profits"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   21
      Top             =   7560
      Width           =   2175
   End
   Begin VB.Label lblCosts 
      Caption         =   "Annual Expenses"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   19
      Top             =   6840
      Width           =   2175
   End
   Begin VB.Label lblRevenue 
      Caption         =   "Annual Revenue"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   17
      Top             =   6120
      Width           =   2055
   End
   Begin VB.Label lblSatisfaction 
      Caption         =   "Popular Satisfaction"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   15
      Top             =   4320
      Width           =   2655
   End
   Begin VB.Label lblBalance 
      Caption         =   "Power Balance (MW/h)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   13
      Top             =   3240
      Width           =   2775
   End
   Begin VB.Label lblProduction 
      Caption         =   "Power Produced (MW/h)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   10
      Top             =   1800
      Width           =   2775
   End
   Begin VB.Label lblDemand 
      Caption         =   "Power Demand (MW/h)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Top             =   2520
      Width           =   2655
   End
   Begin VB.Label lblMain 
      Caption         =   "Operations"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   840
      TabIndex        =   4
      Top             =   240
      Width           =   2535
   End
   Begin VB.Label lblFunds 
      Caption         =   "Available Funds"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   5400
      Width           =   1935
   End
   Begin VB.Label lbl1 
      Caption         =   "Year:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6240
      TabIndex        =   1
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: PowerSim2010
'Form Name: frmMain
'Author: Jared Breen
'Date Written: February 13, 2010
'Objective:
'This form is the central hub of the simulation.  From here, the player gets access to the core statistics around which
'   his or her decisions will be based.  The year is provided to keep the end of the game in perspective, the demand vs.
'   production numbers are provided to allow the player to determine what (if anything) to build, popular satisfaction
'   services and rates are given (largely for flavor), available money and the balance sheet are given as well.
'   Buttons are provided for transitioning between the various forms in the game.
'
'Of special note is the presence of the Next Year algorithm, which adjusts demand from year to year, contains the game's
'   events, checks for nuclear meltdowns and scores the game once it is over.
Option Explicit

Private Sub cmdNextYear_Click()
    'The heart and soul of the program, this button handles the evolution of the market and a variety of other tasks
    'The first part checks for various game-ending criteria (time limit, no money, nuclear meltdown, etc.
    Dim HighEnd As Single, Random As Single, Meltdown As Boolean, EndGame As Boolean, Score As Integer
    Meltdown = False
    EndGame = False
    Year = Year + 1
    Funds = Funds + Revenue - Expenses
    If Year = 2010 Then
        MsgBox ("It is now the year 2010, and the simulation is over.")
        EndGame = True
    End If
    If Satisfaction <= 0 Then
        MsgBox ("Your public support has bottomed out, and the state government has decided to replace your company with a new one.  Your simulation is at an end.")
        EndGame = True
    End If
    If Funds < 0 Then
        MsgBox ("Your company is completely broke.  While deficit spending may be tolerated in many companies, when you have complete control of a market like you do, there is no excuse for incurring debt if you have the proper long-term mindset.  The board of directors agrees.  Your simulation is at an end.")
        EndGame = True
    End If
    If PlantOwned(5) > 0 Then
        HighEnd = 750 / PlantOwned(5)
        Random = Int(HighEnd * Rnd)
        If Random <= 1 Then
            Meltdown = True
        End If
        MsgBox ("Meltdown check: " & Random)
    End If
    'Ensuring that the player is heavily penalized in the event of a meltdown
    If Meltdown = True Then
        EnvironmentTotal = EnvironmentTotal + 1000000
        Funds = 0
        Satisfaction = 0
        PlantOwned(5) = PlantOwned(5) - 1
        MsgBox ("One of your nuclear power plants has experienced a catastrophic meltdown.  Technicians were unable to contain the radiation within the plant, and it has begun leaking outside.  Tens of thousands of people have been exposed to heavy radiation, and the area around the plant has been designated a national disaster area.  Your building has been seized by the FBI who are now conducting an investigation into operations at your plant.  You have officially been removed from your position by the board of trustees, and your game is over.")
        EndGame = True
    End If
    'Checking to see if the game is over
    'If it is not, the game adjusts demand and executes events
    If EndGame = False Then
            'The only two events I ended up including in the program, more as a proof of concept than anything else
            If Year = 1973 Then
                MsgBox ("This year, OPEC established an oil embargo against the United States for its continuing support of Israel.  For the next year, the maintenance of all oil power plants will be doubled.")
                PlantMaint(3) = 2 * PlantMaint(3)
            End If
            If Year = 1974 Then
                MsgBox ("The oil crisis has ended.  Oil prices stabilized this year at their old levels.")
                PlantMaint(3) = 0.5 * PlantMaint(3)
            End If
            'If the player satisfied demand in the year, they get a satisfaction bonus, if not, then a huge penalty
            If Balance >= 0 Then
                Satisfaction = Satisfaction + 1
            Else
                Satisfaction = Satisfaction - 5
                MsgBox ("Rolling brownouts are taking place all throughout California due to inadequate power generation.  Popular satisfaction has been noticeably affected by this.")
            End If
            'Scaling energy demand increases by the decade may be arbitrary, but it helps to keep the game from getting
            '   out of control with the restrictions placed on the player (one plant per year, etc.)  Brief playtesting
            '   showed that these numbers were fairly reasonable, if not perfectly balanced
            Select Case Year
                Case 1950 To 1960
                    Demand = Demand * 1.11
                Case 1960 To 1970
                    Demand = Demand * 1.085
                Case 1970 To 1980
                    Demand = Demand * 1.06
                Case 1980 To 2000
                    Demand = Demand * 1.03
                Case 2000 To 2010
                    Demand = Demand * 1.015
            End Select
            'Tracking environmental degradation and adjusting the balance and revenue with new demand levels
            EnvironmentTotal = EnvironmentTotal + PlantEnv(1) * PlantOwned(1) + PlantEnv(2) * PlantOwned(2) + PlantEnv(3) * PlantOwned(3) + PlantEnv(4) * PlantOwned(4) + PlantEnv(5) * PlantOwned(5) + PlantEnv(6) * PlantOwned(6) + PlantEnv(7) * PlantOwned(7)
            Balance = Production - Demand
            If Balance >= 0 Then
                Revenue = Demand * Rate
            Else
                Revenue = (Demand + Balance) * Rate
            End If
            Expenses = PlantOwned(1) * PlantMaint(1) + PlantOwned(2) * PlantMaint(2) + PlantOwned(3) * PlantMaint(3) + PlantOwned(4) * PlantMaint(4) + PlantOwned(5) * PlantMaint(5) + PlantOwned(6) * PlantMaint(6) + PlantOwned(7) * PlantMaint(7)
            Profit = Revenue - Expenses
            'Refreshing the numbers on the main page to reflect changes from the year
            frmMain.picYear.Cls
            frmMain.picDemand.Cls
            frmMain.picProduction.Cls
            frmMain.picBalance.Cls
            frmMain.picSatisfaction.Cls
            frmMain.picFunds.Cls
            frmMain.picRevenue.Cls
            frmMain.picExpenses.Cls
            frmMain.picProfits.Cls
            frmMain.picYear.Print Year
            frmMain.picDemand.Print Demand
            frmMain.picProduction.Print Production
            frmMain.picBalance.Print Balance
            frmMain.picSatisfaction.Print Satisfaction
            frmMain.picFunds.Print FormatCurrency(Funds)
            frmMain.picRevenue.Print FormatCurrency(Revenue)
            frmMain.picExpenses.Print FormatCurrency(Expenses)
            frmMain.picProfits.Print FormatCurrency(Profit)
            'Reenabling production and rate changes
            frmProduction.cmdBuild.Enabled = True
            frmRates.cmdAdjust.Enabled = True
    'If the game is over, the player is scored and it ends.
    Else
        Score = 100
        MsgBox ("Your final score will be based on three criteria: your available funds, your popular satisfaction, and the cumulative environmental impact of your power plants.")
        MsgBox ("Your base score is 100 points.")
        'Multipliers for popular satisfaction
        Select Case Satisfaction
            Case Is < 0
                Score = Score * 0.01
            Case 0 To 20
                Score = Score * 0.25
            Case 20 To 40
                Score = Score * 0.6
            Case 40 To 60
                Score = Score * 1
            Case 60 To 80
                Score = Score * 1.4
            Case 80 To 100
                Score = Score * 1.9
            Case Is > 100
                Score = Score * 2.5
        End Select
        MsgBox ("With your public satisfaction multiplier, your score is " & Score)
        'Multipliers for money on hand
        Select Case Funds
            Case Is < 0
                Score = Score * 0.1
            Case 0 To 50000
                Score = Score * 0.5
            Case 50000 To 100000
                Score = Score * 0.8
            Case 100000 To 300000
                Score = Score * 1
            Case 300000 To 700000
                Score = Score * 1.3
            Case 700000 To 1250000
                Score = Score * 1.7
            Case Is > 1250000
                Score = Score * 2.25
        End Select
        MsgBox ("With your available money modifier, your score is " & Score)
        'Multipliers for environment (could use more balancing work)
        Select Case EnvironmentTotal
            Case Is < 6000
                Score = Score * 3
            Case 6000 To 8000
                Score = Score * 2
            Case 8000 To 10000
                Score = Score * 1.5
            Case 10000 To 15000
                Score = Score * 1
            Case 17500 To 22500
                Score = Score * 0.5
            Case 22500 To 30000
                Score = Score * 0.25
            Case 30000 To 60000
                Score = Score * 0.1
            Case Is > 60000
                Score = Score * 0
        End Select
        MsgBox ("With your environmental modifier, your final score is " & Score & ".  Thank you for playing.")
        End
    End If
End Sub

Private Sub cmdProduction_Click()
    'Switches screen to production and fills its picture boxes with up-to-date numbers
    frmProduction.Show
    frmMain.Hide
    frmProduction.picYear.Cls
    frmProduction.picFunds.Cls
    frmProduction.picProfits.Cls
    frmProduction.picBalance.Cls
    frmProduction.picYear.Print Year
    frmProduction.picFunds.Print FormatCurrency(Funds)
    frmProduction.picProfits.Print FormatCurrency(Profit)
    frmProduction.picBalance.Print Balance
End Sub

Private Sub cmdQuit_Click()
    End
End Sub

Private Sub cmdRates_Click()
    'Switches screen to rates and fills its picture boxes with up-to-date numbers
    frmMain.Hide
    frmRates.Show
    frmRates.picYear.Cls
    frmRates.picDemand.Cls
    frmRates.picSatisfaction.Cls
    frmRates.picRevenue.Cls
    frmRates.picExpenses.Cls
    frmRates.picProfits.Cls
    frmRates.picRate.Cls
    frmRates.picYear.Print Year
    frmRates.picDemand.Print Demand
    frmRates.picSatisfaction.Print Satisfaction
    frmRates.picRevenue.Print FormatCurrency(Revenue)
    frmRates.picExpenses.Print FormatCurrency(Expenses)
    frmRates.picProfits.Print FormatCurrency(Profit)
    frmRates.picRate.Print FormatCurrency(Rate)
End Sub

Private Sub cmdRnD_Click()
    'Switches screen to research and fills its picture boxes with up-to-date numbers
    frmRnD.Show
    frmMain.Hide
    frmRnD.picYear.Cls
    frmRnD.picFunds.Cls
    frmRnD.picYear.Print Year
    frmRnD.picFunds.Print FormatCurrency(Funds)
End Sub
