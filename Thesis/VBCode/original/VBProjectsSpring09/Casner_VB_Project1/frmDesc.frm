VERSION 5.00
Begin VB.Form frmDesc 
   BackColor       =   &H8000000A&
   Caption         =   "Descriptive Statistics"
   ClientHeight    =   1095
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4155
   LinkTopic       =   "Form1"
   ScaleHeight     =   1095
   ScaleWidth      =   4155
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBoth 
      Caption         =   "Both Sets"
      Height          =   615
      Left            =   2760
      TabIndex        =   2
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton cmdSecondary 
      Caption         =   "Secondary Set"
      Height          =   615
      Left            =   1440
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton cmdPrimary 
      Caption         =   "Primary Set"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frmDesc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Visual Basic T-test
'Benjamin Casner
'March 13th, 2009
'.frmDesc
'this form displays the basic descriptive statistics for either
'the primary, secondary, or both data sets
Private Sub cmdBoth_Click()
    'displays descriptive statistics for both sets
    frmDesc.Hide
    frmStats.picResults.Cls
    frmStats.picResults.Print Tab(5); "Mean"; Tab(25); "Variance"; Tab(45); "Standard Deviation"
    frmStats.picResults.Print ; Tab(5); "*************************************************************************"
    frmStats.picResults.Print "x"; Tab(5); Round(xBar, 2); Tab(25); Round(Sx, 2); Tab(45); Round(Sx ^ (1 / 2), 2)
    frmStats.picResults.Print "y"; Tab(5); Round(yBar, 2); Tab(25); Round(Sy, 2); Tab(45); Round(Sy ^ (1 / 2), 2)
End Sub

Private Sub cmdPrimary_Click()
    'displays descriptive statistics for the primary set
    frmDesc.Hide
    frmStats.picResults.Cls
    frmStats.picResults.Print Tab(5); "Mean"; Tab(25); "Variance"; Tab(45); "Standard Deviation"
    frmStats.picResults.Print ; Tab(5); "*************************************************************************"
    frmStats.picResults.Print Tab(5); Round(xBar, 2); Tab(25); Round(Sx, 2); Tab(45); Round(Sx ^ (1 / 2), 2)
End Sub

Private Sub cmdSecondary_Click()
    'displays descriptive statistics for secondary set
    frmDesc.Hide
    frmStats.picResults.Cls
    frmStats.picResults.Print Tab(5); "Mean"; Tab(25); "Variance"; Tab(45); "Standard Deviation"
    frmStats.picResults.Print ; Tab(5); "*************************************************************************"
    frmStats.picResults.Print Tab(5); Round(yBar, 2); Tab(25); Round(Sy, 2); Tab(45); Round(Sy ^ (1 / 2), 2)
End Sub

Private Sub Form_Load()
    frmDataDisplay.Hide
End Sub
