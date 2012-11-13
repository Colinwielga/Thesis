VERSION 5.00
Begin VB.Form frmIntro 
   Caption         =   "Introduction"
   ClientHeight    =   5370
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7410
   LinkTopic       =   "Form1"
   ScaleHeight     =   5370
   ScaleWidth      =   7410
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Height          =   615
      Left            =   2640
      TabIndex        =   0
      Top             =   4560
      Width           =   2055
   End
   Begin VB.Label lblIntro3 
      Caption         =   "Good luck."
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   4080
      Width           =   6855
   End
   Begin VB.Label lblIntro2 
      Caption         =   $"frmIntro.frx":0000
      Height          =   1575
      Left            =   240
      TabIndex        =   3
      Top             =   2280
      Width           =   6855
   End
   Begin VB.Label lblIntro 
      Caption         =   $"frmIntro.frx":02E8
      Height          =   1095
      Left            =   240
      TabIndex        =   2
      Top             =   1080
      Width           =   6855
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      Caption         =   "PowerSim 2010"
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
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   6855
   End
End
Attribute VB_Name = "frmIntro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project name: PowerSim2010
'Form name: frmIntro
'Author: Jared Breen
'Date Written: February 16, 2010
'Project Objective:
'The goal of this project is to create an extremely basic business simulation game dealing with electricity generation.
'   The player is tasked with generating electricity for an arbitrary region (defined as California in the introduction)
'   and given only a minimum amount of direction toward completing this task.  Many of the numbers that represent the core
'   of the game code are hidden from the player, such as the existence of certain events and the possiblity of accidents,
'   as well as the environmental impact of each additional power plant.  In a highly underhanded move, the player is judged
'   based on their environmental impact without being told before the game is complete, with the express purpose of making a
'   point: the short-term decisions that a person makes in a modern business climate can have tangible repercussions in the
'   long-term that were unforeseen or unaccounted for.  The game is designed so that the most cost-effective means of
'   completion is to build huge numbers of coal and oil power plants, but the player is heavily penalized.  Clean power is
'   harder to get and more expensive, but adds a bonus to score in the end.  Nuclear power is highly efficient when it
'   doesn't melt down and effectively end the game prematurely.  The player is merely told to keep people happy and earn
'   profits for the company, belying the true number of calculations going on under the hood.  It's up to the player to
'   determine what is more important: profits and power generation or environmental responsibility.  There is no one right
'   answer.
'
'Form Objective:
'Originally intended to be a mere help/reference screen, this form took on a life of its own in response to my inability
'   to make the module execute any actions before the game started (read data, initialize variables, print onto the main
'   form, etc.)  This form serves as a brief introduction to the scenario, and then performs all of these tasks upon the
'   activation of the Start button.

Option Explicit
Private Sub cmdStart_Click()
    'Preparing for and executing the data reading process for plant data
    Dim Ctr As Single
    Ctr = 0
    Open App.Path & "\plantdata.txt" For Input As #1
        Do Until EOF(1)
            Ctr = Ctr + 1
            Input #1, PlantType(Ctr), PlantCost(Ctr), PlantMaint(Ctr), PlantEnv(Ctr), PlantOutput(Ctr)
        Loop
    Close #1
    'Setting initial values of core variables before transitioning to main screen
    Year = 1950
    Funds = 50000
    Demand = 3000
    Revenue = 0
    Expenses = 0
    Profit = Revenue - Expenses
    Production = 0
    Balance = Production - Demand
    Satisfaction = 40
    Rate = 1.5
    'Switching forms and filling the main form's picture boxes
    frmIntro.Hide
    frmMain.Show
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
