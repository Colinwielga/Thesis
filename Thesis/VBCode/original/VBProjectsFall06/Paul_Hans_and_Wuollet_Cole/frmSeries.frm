VERSION 5.00
Begin VB.Form frmSeries 
   Caption         =   "1987 World Series Champs"
   ClientHeight    =   6810
   ClientLeft      =   2055
   ClientTop       =   3090
   ClientWidth     =   9600
   LinkTopic       =   "Form2"
   ScaleHeight     =   6810
   ScaleWidth      =   9600
   Visible         =   0   'False
   Begin VB.CommandButton cmdContinue 
      Caption         =   "Continue"
      Height          =   855
      Left            =   6120
      TabIndex        =   1
      Top             =   5880
      Width           =   1575
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return to Main Screen"
      Height          =   855
      Left            =   7800
      TabIndex        =   0
      Top             =   5880
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   6765
      Left            =   0
      Picture         =   "frmSeries.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9600
   End
End
Attribute VB_Name = "frmSeries"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project name: 1987 World Series
'Form name: frmSeries
'Authors: Hans Paul and Cole Wuollet
'Date Written: Tuesday October 31, 2006
'Objective: A welcome screen to the World Series segment of the Project.
Option Explicit

Private Sub cmdContinue_Click() 'Hides Current Form and Goes to The Series Stats Form
    frmSeries.Hide
    frmSeriesStats.Show
End Sub

Private Sub cmdReturn_Click() 'Hides Current Form and Loads frontpage From
    frmSeries.Hide
    frmTwins.Show
End Sub

