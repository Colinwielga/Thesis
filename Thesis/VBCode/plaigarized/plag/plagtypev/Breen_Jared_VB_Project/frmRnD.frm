VERSION 5.00
Begin VB.Form frmRnD
   Caption         =   "Research and Development"
   ClientHeight    =   7275
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13455
   LinkTopic       =   "Form1"
   ScaleHeight     =   7275
   ScaleWidth      =   13455
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picAdvWind
      Height          =   1455
      Left            =   6840
      Picture         =   "frmRnD.frx":0000
      ScaleHeight     =   1395
      ScaleWidth      =   1395
      TabIndex        =   19
      Top             =   4560
      Width           =   1455
   End
   Begin VB.PictureBox picAdvSolar
      Height          =   1455
      Left            =   6840
      Picture         =   "frmRnD.frx":62DC
      ScaleHeight     =   1395
      ScaleWidth      =   1395
      TabIndex        =   17
      Top             =   2880
      Width           =   1455
   End
   Begin VB.PictureBox picWind
      Height          =   1455
      Left            =   6840
      Picture         =   "frmRnD.frx":9756
      ScaleHeight     =   1395
      ScaleWidth      =   1395
      TabIndex        =   15
      Top             =   1200
      Width           =   1455
   End
   Begin VB.PictureBox picSolar
      Height          =   1455
      Left            =   240
      Picture         =   "frmRnD.frx":B350
      ScaleHeight     =   1395
      ScaleWidth      =   1395
      TabIndex        =   13
      Top             =   4560
      Width           =   1455
   End
   Begin VB.PictureBox picNuclear
      Height          =   1455
      Left            =   240
      Picture         =   "frmRnD.frx":E2AD
      ScaleHeight     =   1395
      ScaleWidth      =   1395
      TabIndex        =   11
      Top             =   2880
      Width           =   1455
   End
   Begin VB.PictureBox picGas
      Height          =   1455
      Left            =   240
      Picture         =   "frmRnD.frx":116AF
      ScaleHeight     =   1395
      ScaleWidth      =   1395
      TabIndex        =   9
      Top             =   1200
      Width           =   1455
   End
   Begin VB.CommandButton cmdMain
      Caption         =   "Return to Operations"
      Height          =   735
      Left            =   10440
      TabIndex        =   7
      Top             =   6360
      Width           =   2775
   End
   Begin VB.CommandButton cmdResearch
      Caption         =   "Authorize Research Project"
      Height          =   735
      Left            =   6840
      TabIndex        =   6
      Top             =   6360
      Width           =   2895
   End
   Begin VB.TextBox txtResearch
      Height          =   375
      Left            =   5640
      TabIndex        =   5
      Top             =   6480
      Width           =   375
   End
   Begin VB.PictureBox picFunds
      Height          =   375
      Left            =   11160
      ScaleHeight     =   315
      ScaleWidth      =   2115
      TabIndex        =   3
      Top             =   720
      Width           =   2175
   End
   Begin VB.PictureBox picYear
      Height          =   375
      Left            =   12720
      ScaleHeight     =   315
      ScaleWidth      =   555
      TabIndex        =   1
      Top             =   120
      Width           =   615
   End
   Begin VB.Label lblAdvWind
      Caption         =   "6. Advanced Turbine Manufacturing (1995) - $7,500"
      Height          =   375
      Left            =   8520
      TabIndex        =   20
      Top             =   5160
      Width           =   5055
   End
   Begin VB.Label lblAdvSolar
      Caption         =   "5. Improved Solar Cells (1990) - $5,000 + $2,500 per existing Solar Power Plant"
      Height          =   495
      Left            =   8520
      TabIndex        =   18
      Top             =   3480
      Width           =   5055
   End
   Begin VB.Label lblWind
      Caption         =   "4. Wind Turbines (1980) - $7,500"
      Height          =   375
      Left            =   8520
      TabIndex        =   16
      Top             =   1800
      Width           =   2535
   End
   Begin VB.Label lblSolar
      Caption         =   "3. Solar Power Plant (1970) - $10,000"
      Height          =   375
      Left            =   1920
      TabIndex        =   14
      Top             =   5160
      Width           =   3495
   End
   Begin VB.Label lblNuclear
      Caption         =   "2. Nuclear Power Plant (1970) - $40,000"
      Height          =   375
      Left            =   1920
      TabIndex        =   12
      Top             =   3480
      Width           =   3375
   End
   Begin VB.Label lblGas
      Caption         =   "1. Natural Gas Power Plant (1960) - $5,000"
      Height          =   375
      Left            =   1920
      TabIndex        =   10
      Top             =   1800
      Width           =   3495
   End
   Begin VB.Label lblResearch
      Caption         =   "Enter the number of the project to research:"
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
      Left            =   720
      TabIndex        =   8
      Top             =   6480
      Width           =   4935
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
      Left            =   8760
      TabIndex        =   4
      Top             =   720
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
      Left            =   12000
      TabIndex        =   2
      Top             =   120
      Width           =   615
   End
   Begin VB.Label lblRnD
      Caption         =   "Research and Development"
      BeginProperty Font
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1560
      TabIndex        =   0
      Top             =   360
      Width           =   6615
   End
End
Attribute VB_Name = "frmRnD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project name: PowerSim2010
'Form name: frmRnD
'Author: Jared Breen
'Date Written: February 13, 2010
'Objective:
'This form allows the player to pour his/her profits into a number of research projects designed to aid in power generation.
'   New projects become available over the course of time and offer varying benefits, either in the form of more advanced types
'   of generation or efficiency bonuses on the normally less cost-effective 'clean' forms of power.  The last two projects in
'   particular reward the player for thinking ahead and building the more advanced, environmentally-conscious generators, and the
'   last one makes the mass-purchase of wind turbines a viable end-game strategy.  The numbers on the last two research projects
'   are intentionally hidden to make it seem like something of a gamble, just like real clean energy technology.

Option Explicit

Private Sub cmdMain_Click()
    'Switches back to the main form and refreshes its numbers
    frmRnD.Hide
    frmMain.Show
    frmMain.picYear.Cls
    frmMain.picDemand.Cls
    frmMain.picProduction.Cls
    frmMain.picBalance.Cls
    frmMain.picSatisfaction.Cls
    frmMain.picFunds.Cls
    frmMain.picRevenue.Cls
    frmMain.picExpenses.Cls
Dim bbb as Integer
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
End Sub

Private Sub cmdResearch_Click()
    'Set up variables and take input from the user
    Dim Project As Integer
    Dim Decision
    Project = txtResearch.Text
    'Choosing project (only first one is commented due to extreme similarity)
    Select Case Project
        'Checks for which project was chosen
        Case Is = 1
            'Checks to make sure the project is available
            If Year < 1960 Then
                MsgBox ("You cannot research this project yet.")
            Else
                'Checks if the player already researched the project
                If Research(1) = True Then
                    MsgBox ("You have already researched this project.")
                Else
                    'Checks that the player has enough money to research the project
                    If Funds < 5000 Then
                        MsgBox ("Insufficient funds.")
                    Else
                        'Asks the player for confirmation using a yes/no message box
                        Decision = MsgBox("Are you sure you wish to research Natural Gas Plants?  This will cost $7,500.", vbYesNo)
                        If Decision = vbYes Then
                            'If approved, the project is executed, marked researched, disabled on the
                            'research screen, paid for, and the funds variable is updated
                            frmProduction.optGas.Enabled = True
                            Research(1) = True
                            lblGas.Enabled = False
                            picGas.Enabled = False
                            Funds = Funds - 5000
                            picFunds.Cls
                            picFunds.Print FormatCurrency(Funds)
                        End If
                    End If
                End If
            End If
        Case Is = 2
            If Year < 1970 Then
                MsgBox ("You cannot research this project yet.")
            Else
                If Research(2) = True Then
                    MsgBox ("You have already researched this project.")
                Else
Dim ccc as Integer
                    If Funds < 40000 Then
                        MsgBox ("Insufficient funds.")
                    Else
                        Decision = MsgBox("Are you sure you wish to research Nuclear Power Plants?  This will cost $30,000.", vbYesNo)
                        If Decision = vbYes Then
                            frmProduction.optNuclear.Enabled = True
                            Research(2) = True
                            lblNuclear.Enabled = False
                            picNuclear.Enabled = False
                            Funds = Funds - 40000
                            picFunds.Cls
                            picFunds.Print FormatCurrency(Funds)
                        End If
                    End If
                End If
            End If
        Case Is = 3
            If Year < 1970 Then
                MsgBox ("You cannot research this project yet.")
            Else
                If Research(3) = True Then
                    MsgBox ("You have already researched this project.")
                Else
                    If Funds < 10000 Then
                        MsgBox ("Insufficient funds.")
                    Else
                        Decision = MsgBox("Are you sure you wish to research Solar Power Plants?  This will cost $10,000.", vbYesNo)
                        If Decision = vbYes Then
                            frmProduction.optSolar.Enabled = True
                            Research(3) = True
                            lblGas.Enabled = False
                            picGas.Enabled = False
                            Funds = Funds - 10000
                            picFunds.Cls
                            picFunds.Print FormatCurrency(Funds)
                        End If
                    End If
                End If
            End If
        Case Is = 4
            If Year < 1980 Then
                MsgBox ("You cannot research this project yet.")
            Else
                If Research(4) = True Then
                    MsgBox ("You have already researched this project.")
                Else
                    If Funds < 7500 Then
                        MsgBox ("Insufficient funds.")
                    Else
                        Decision = MsgBox("Are you sure you wish to research Wind Turbines?  This will cost $7,500.", vbYesNo)
                        If Decision = vbYes Then
                            frmProduction.optWind.Enabled = True
Dim dddd as Integer
                            Research(4) = True
                            lblWind.Enabled = False
                            picWind.Enabled = False
                            Funds = Funds - 7500
                            picFunds.Cls
                            picFunds.Print FormatCurrency(Funds)
                        End If
                    End If
                End If
            End If
        Case Is = 5
            If Year < 1990 Then
                MsgBox ("You cannot research this project yet.")
            Else
                If Research(5) = True Then
                    MsgBox ("You have already researched this project.")
                Else
                    If Funds < (5000 + (2500 * PlantOwned(6))) Then
                        MsgBox ("Insufficient funds.")
                    Else
                        Dim SolarCost As Single
                        SolarCost = (5000 + (2500 * PlantOwned(6)))
                        Decision = MsgBox("Are you sure you wish to research Improved Solar Cells?  This will cost " & FormatCurrency(SolarCost), vbYesNo)
                        If Decision = vbYes Then
                            PlantMaint(6) = PlantMaint(6) - 1500
                            PlantOutput(6) = PlantOutput(6) + 1000
                            Research(5) = True
                            lblAdvSolar.Enabled = False
                            picAdvSolar.Enabled = False
                            Funds = Funds - SolarCost
                            picFunds.Cls
                            picFunds.Print FormatCurrency(Funds)
                            Production = Production + (1000 * PlantOwned(6))
                            Balance = Production - Demand
                            Expenses = Expenses - (1000 * PlantOwned(6))
                            If Balance >= 0 Then
                                Revenue = Demand * Rate
                            Else
                                Revenue = (Demand + Balance) * Rate
                            End If
                            Profit = Revenue - Expenses
                        End If
                    End If
                End If
            End If
        Case Is = 6
            If Year < 1995 Then
                MsgBox ("You cannot research this project yet.")
            Else
                If Research(6) = True Then
                    MsgBox ("You have already researched this project.")
                Else
                    If Funds < 7500 Then
                        MsgBox ("Insufficient funds.")
                    Else
                        Decision = MsgBox("Are you sure you wish to research Advanced Turbine Manufacturing?  This will cost $7,500.", vbYesNo)
                        If Decision = vbYes Then
                            PlantCost(7) = PlantCost(7) - 2000
                            PlantMaint(7) = PlantMaint(7) - 1000
                            PlantOutput(7) = PlantOutput(7) + 1000
                            Research(6) = True
                            lblAdvWind.Enabled = False
Dim eeee as Integer
                            picAdvWind.Enabled = False
                            Funds = Funds - 7500
                            picFunds.Cls
                            picFunds.Print FormatCurrency(Funds)
                            Production = Production + (500 * PlantOwned(7))
                            Expenses = Expenses - (500 * PlantOwned(7))
                            Balance = Production - Demand
                            If Balance >= 0 Then
                                Revenue = Demand * Rate
                            Else
                                Revenue = (Demand + Balance) * Rate
                            End If
                            Profit = Revenue - Expenses
                        End If
                    End If
                End If
            End If
        Case Else
            MsgBox ("That is not a valid project.")
    End Select
End Sub
