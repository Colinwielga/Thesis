VERSION 5.00
Begin VB.Form frmProduction
   Caption         =   "Production"
   ClientHeight    =   6165
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8235
   LinkTopic       =   "Form1"
   ScaleHeight     =   6165
   ScaleWidth      =   8235
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picFunds
      Height          =   375
      Left            =   6000
      ScaleHeight     =   315
      ScaleWidth      =   2115
      TabIndex        =   18
      Top             =   3240
      Width           =   2175
   End
   Begin VB.CommandButton cmdInfo
      Caption         =   "More Information"
      Height          =   855
      Left            =   3120
      TabIndex        =   17
      Top             =   5040
      Width           =   2055
   End
   Begin VB.PictureBox picProfits
      Height          =   375
      Left            =   6000
      ScaleHeight     =   315
      ScaleWidth      =   2115
      TabIndex        =   15
      Top             =   2520
      Width           =   2175
   End
   Begin VB.PictureBox picBalance
      Height          =   375
      Left            =   6840
      ScaleHeight     =   315
      ScaleWidth      =   1275
      TabIndex        =   13
      Top             =   1560
      Width           =   1335
   End
   Begin VB.CommandButton cmdMain
      Caption         =   "Return to Operations"
      Height          =   855
      Left            =   5880
      TabIndex        =   12
      Top             =   5040
      Width           =   1935
   End
   Begin VB.CommandButton cmdBuild
      Caption         =   "Authorize Construction"
      Height          =   855
      Left            =   360
      TabIndex        =   11
      Top             =   5040
      Width           =   2055
   End
   Begin VB.Frame fmePlantSelect
      Caption         =   "Build a Power Plant"
      Height          =   3135
      Left            =   240
      TabIndex        =   3
      Top             =   1440
      Width           =   2895
      Begin VB.OptionButton optWind
         Caption         =   "Wind Turbine Farm"
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   2520
         Width           =   2175
      End
      Begin VB.OptionButton optSolar
         Caption         =   "Solar Power Plant"
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   2160
         Width           =   2175
      End
      Begin VB.OptionButton optNuclear
         Caption         =   "Nuclear Power Plant"
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   1800
         Width           =   2175
      End
      Begin VB.OptionButton optGas
         Caption         =   "Natural Gas Power Plant"
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   1440
         Width           =   2175
      End
      Begin VB.OptionButton optOil
         Caption         =   "Oil Power Plant"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   1080
         Width           =   1815
      End
      Begin VB.OptionButton optHydroelectric
         Caption         =   "Hydroelectric Dam"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   1935
      End
      Begin VB.OptionButton optCoal
         Caption         =   "Coal Power Plant"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.PictureBox picYear
      Height          =   375
      Left            =   7440
      ScaleHeight     =   315
      ScaleWidth      =   555
      TabIndex        =   2
      Top             =   120
      Width           =   615
   End
   Begin VB.Label lblNote
      Caption         =   "Note: Only one plant may be constructed per year."
      Height          =   255
      Left            =   360
      TabIndex        =   20
      Top             =   4680
      Width           =   3615
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
      Left            =   3600
      TabIndex        =   19
      Top             =   3240
      Width           =   1935
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
      Left            =   3600
      TabIndex        =   16
      Top             =   2520
      Width           =   2175
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
      Left            =   3600
      TabIndex        =   14
      Top             =   1560
      Width           =   2775
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
      Left            =   6600
      TabIndex        =   1
      Top             =   120
      Width           =   615
   End
   Begin VB.Label lblProd
      Caption         =   "Production"
      BeginProperty Font
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1440
      TabIndex        =   0
      Top             =   240
      Width           =   2655
   End
End
Attribute VB_Name = "frmProduction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project name: PowerSim2010
'Form name: frmProduction
'Author: Jared Breen
'Date Written: February 13, 2010
'Objective:
'This form is where the player constructs new power plants to fulfill the electricity needs of the populace.  At the beginning
'   of the simulation, only a limited number of plant types are available, but this number can be expanded over time via
'   expenditure in research and development.  There is a strong incentive to do so largely for access to nuclear power and the
'   forms of clean power, depending on what path the player wishes to take.  There is nothing stopping the player from simply
'   building huge numbers of coal plants, but this is not the most cost-effective long-term solution for winning the game and
'   will tank the player's environmental responsibility score.  Hydroelectric damming offers the player relatively clean power
'   early in the game, but is capped at 1, requiring the player to look elsewhere.  Information on each plant can be acquired via
'   the "More Information" button, and purchases are authorized via the Construction button, and the player is provided with a
'   few key statistics on this screen to aid in decision making.  Only one plant can be constructed per year, limiting overall
'   options.  One may return to the main screen at any time via the final button.

Option Explicit

Private Sub cmdBuild_Click()
If True Then
End If
    'Each routine here does several things: checks to ensure the player has the money to buy the plant, adds one to the player's
    '   number of owned plants, adds its expenses and production to the appropriate figures, recalculates the balance and revenue,
    '   generates a new profit number, displays the plant was produced and disables the build button for that year.
    '   Exception to this rule is the Dam, which has a hard cap of 1 owned in any particular game, so it disables itself when
    '   purchased.
    If optCoal.Value = True Then
        If Funds >= PlantCost(1) Then
            PlantOwned(1) = PlantOwned(1) + 1
            Funds = Funds - PlantCost(1)
            Expenses = Expenses + PlantMaint(1)
            Production = Production + PlantOutput(1)
            Balance = Production - Demand
            If Balance >= 0 Then
                Revenue = Demand * Rate
            Else
                Revenue = (Demand + Balance) * Rate
            End If
            Profit = Revenue - Expenses
            MsgBox ("A new coal power plant has been built.  It has now been factored in to our estimates in Operations.")
            cmdBuild.Enabled = False
        Else
            MsgBox ("Insufficient funds.")
        End If
    ElseIf optHydroelectric.Value = True Then
        If Funds >= PlantCost(2) Then
            PlantOwned(2) = 1
            Funds = Funds - PlantCost(2)
            Expenses = Expenses + PlantMaint(2)
            Production = Production + PlantOutput(2)
            Balance = Production - Demand
            If Balance >= 0 Then
                Revenue = Demand * Rate
            Else
                Revenue = (Demand + Balance) * Rate
            End If
            Profit = Revenue - Expenses
            optHydroelectric.Enabled = False
            MsgBox ("A new hydroelectric dam has been built.  It has now been factored in to our estimates in Operations.")
            cmdBuild.Enabled = False
        Else
            MsgBox ("Insufficient funds.")
        End If
    ElseIf optOil.Value = True Then
        If Funds >= PlantCost(3) Then
            PlantOwned(3) = PlantOwned(3) + 1
            Funds = Funds - PlantCost(3)
            Expenses = Expenses + PlantMaint(3)
            Production = Production + PlantOutput(3)
            Balance = Production - Demand
            If Balance >= 0 Then
                Revenue = Demand * Rate
            Else
                Revenue = (Demand + Balance) * Rate
            End If
            Profit = Revenue - Expenses
            MsgBox ("A new oil power plant has been built.  It has now been factored in to our estimates in Operations.")
            cmdBuild.Enabled = False
        Else
            MsgBox ("Insufficient funds.")
        End If
    ElseIf optGas.Value = True Then
        If Funds >= PlantCost(4) Then
            PlantOwned(4) = PlantOwned(4) + 1
            Funds = Funds - PlantCost(4)
            Expenses = Expenses + PlantMaint(4)
            Production = Production + PlantOutput(4)
            Balance = Production - Demand
            If Balance >= 0 Then
                Revenue = Demand * Rate
            Else
                Revenue = (Demand + Balance) * Rate
            End If
            Profit = Revenue - Expenses
            MsgBox ("A new natural gas power plant has been built.  It has now been factored in to our estimates in Operations.")
            cmdBuild.Enabled = False
        Else
            MsgBox ("Insufficient funds.")
        End If
    ElseIf optNuclear.Value = True Then
        If Funds >= PlantCost(5) Then
            PlantOwned(5) = PlantOwned(5) + 1
            Funds = Funds - PlantCost(5)
            Expenses = Expenses + PlantMaint(5)
            Production = Production + PlantOutput(5)
            Balance = Production - Demand
            If Balance >= 0 Then
                Revenue = Demand * Rate
            Else
                Revenue = (Demand + Balance) * Rate
            End If
            Profit = Revenue - Expenses
            MsgBox ("A new nuclear power plant has been built.  It has now been factored in to our estimates in Operations.")
            cmdBuild.Enabled = False
        Else
            MsgBox ("Insufficient funds.")
        End If
    ElseIf optSolar.Value = True Then
        If Funds >= PlantCost(6) Then
            PlantOwned(6) = PlantOwned(6) + 1
            Funds = Funds - PlantCost(6)
            Expenses = Expenses + PlantMaint(6)
            Production = Production + PlantOutput(6)
            Balance = Production - Demand
            If Balance >= 0 Then
                Revenue = Demand * Rate
            Else
                Revenue = (Demand + Balance) * Rate
            End If
            Profit = Revenue - Expenses
            MsgBox ("A new solar power plant has been built.  It has now been factored in to our estimates in Operations.")
            cmdBuild.Enabled = False
        Else
            MsgBox ("Insufficient funds.")
        End If
    ElseIf optWind.Value = True Then
        If Funds >= PlantCost(7) Then
            PlantOwned(7) = PlantOwned(7) + 1
            Funds = Funds - PlantCost(7)
            Expenses = Expenses + PlantMaint(7)
            Production = Production + PlantOutput(7)
            Balance = Production - Demand
            If Balance >= 0 Then
                Revenue = Demand * Rate
            Else
                Revenue = (Demand + Balance) * Rate
            End If
            Profit = Revenue - Expenses
            MsgBox ("A new wind power plant has been built.  It has now been factored in to our estimates in Operations.")
            cmdBuild.Enabled = False
        Else
            MsgBox ("Insufficient funds.")
        End If
    Else
        'Displayed if the player does not toggle the option button.
        MsgBox ("Please select a type of plant.")
    End If
    'Updates the stats on the production page
    frmProduction.picYear.Cls
    frmProduction.picFunds.Cls
    frmProduction.picProfits.Cls
    frmProduction.picBalance.Cls
    frmProduction.picYear.Print Year
    frmProduction.picFunds.Print FormatCurrency(Funds)
    frmProduction.picProfits.Print FormatCurrency(Profit)
    frmProduction.picBalance.Print Balance
End Sub

Private Sub cmdInfo_Click()
    'Displays a message box with stats on the given plant.  These stats evolve with the player's research.
    If optCoal.Value = True Then
        MsgBox ("Cost: " & FormatCurrency(PlantCost(1)) & vbNewLine & "Annual Maintenance: " & FormatCurrency(PlantMaint(1)) & vbNewLine & "Annual Output: " & PlantOutput(1) & " MW/h")
    ElseIf optHydroelectric.Value = True Then
        MsgBox ("Cost: " & FormatCurrency(PlantCost(2)) & vbNewLine & "Annual Maintenance: " & FormatCurrency(PlantMaint(2)) & vbNewLine & "Annual Output: " & PlantOutput(2) & " MW/h" & vbNewLine & "Only one Hydroelectric Dam may be built.")
    ElseIf optOil.Value = True Then
        MsgBox ("Cost: " & FormatCurrency(PlantCost(3)) & vbNewLine & "Annual Maintenance: " & FormatCurrency(PlantMaint(3)) & vbNewLine & "Annual Output: " & PlantOutput(3) & " MW/h")
    ElseIf optGas.Value = True Then
        MsgBox ("Cost: " & FormatCurrency(PlantCost(4)) & vbNewLine & "Annual Maintenance: " & FormatCurrency(PlantMaint(4)) & vbNewLine & "Annual Output: " & PlantOutput(4) & " MW/h")
    ElseIf optNuclear.Value = True Then
        MsgBox ("Cost: " & FormatCurrency(PlantCost(5)) & vbNewLine & "Annual Maintenance: " & FormatCurrency(PlantMaint(5)) & vbNewLine & "Annual Output: " & PlantOutput(5) & " MW/h" & vbNewLine & "Slight chance of meltdown.")
    ElseIf optSolar.Value = True Then
        MsgBox ("Cost: " & FormatCurrency(PlantCost(6)) & vbNewLine & "Annual Maintenance: " & FormatCurrency(PlantMaint(6)) & vbNewLine & "Annual Output: " & PlantOutput(6) & " MW/h")
    ElseIf optWind.Value = True Then
        MsgBox ("Cost: " & FormatCurrency(PlantCost(7)) & vbNewLine & "Annual Maintenance: " & FormatCurrency(PlantMaint(7)) & vbNewLine & "Annual Output: " & PlantOutput(7) & " MW/h")
    Else
        MsgBox ("Please select a type of plant.")
    End If
End Sub

Private Sub cmdMain_Click()
    'Switches forms and refreshes the main page
    frmProduction.Hide
    frmMain.Show
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
If True Then
End If
End Sub
