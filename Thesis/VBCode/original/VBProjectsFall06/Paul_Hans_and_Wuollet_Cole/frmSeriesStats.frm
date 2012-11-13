VERSION 5.00
Begin VB.Form frmSeriesStats 
   BackColor       =   &H80000001&
   Caption         =   "World Series Stats"
   ClientHeight    =   6705
   ClientLeft      =   3330
   ClientTop       =   2745
   ClientWidth     =   7785
   LinkTopic       =   "Form1"
   Picture         =   "frmSeriesStats.frx":0000
   ScaleHeight     =   6705
   ScaleWidth      =   7785
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back to Main Page"
      Height          =   615
      Left            =   4320
      TabIndex        =   10
      Top             =   6000
      Width           =   1695
   End
   Begin VB.CommandButton cmdPitching 
      Caption         =   "View Team Pitching Stats for The World Series"
      Height          =   615
      Left            =   2160
      TabIndex        =   9
      Top             =   6000
      Width           =   1695
   End
   Begin VB.CommandButton cmdHitting 
      Caption         =   "View Team Hitting Stats for The World Series"
      Height          =   615
      Left            =   240
      TabIndex        =   8
      Top             =   6000
      Width           =   1695
   End
   Begin VB.CommandButton cmdBox1 
      Caption         =   "View Box Score for Game 1"
      Height          =   615
      Left            =   6480
      TabIndex        =   7
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   615
      Left            =   6480
      TabIndex        =   6
      Top             =   6000
      Width           =   1215
   End
   Begin VB.CommandButton cmdBox5 
      Caption         =   "View Box Score for Game 5"
      Height          =   615
      Left            =   6480
      TabIndex        =   5
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton cmdBox6 
      Caption         =   "View Box Score for Game 6"
      Height          =   615
      Left            =   6480
      TabIndex        =   4
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton cmdBox7 
      Caption         =   "View Box Score for Game 7"
      Height          =   615
      Left            =   6480
      TabIndex        =   3
      Top             =   5160
      Width           =   1215
   End
   Begin VB.CommandButton cmdBox2 
      Caption         =   "View Box Score for Game 2"
      Height          =   615
      Left            =   6480
      TabIndex        =   2
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton cmdBox3 
      Caption         =   "View Box Score for Game 3"
      Height          =   615
      Left            =   6480
      TabIndex        =   1
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton cmdBox4 
      Caption         =   "View Box Score For Game 4"
      Height          =   615
      Left            =   6480
      TabIndex        =   0
      Top             =   2640
      Width           =   1215
   End
End
Attribute VB_Name = "frmSeriesStats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project name: 1987 World Series
'Form name: frmSeriesStats
'Authors: Hans Paul and Cole Wuollet
'Date Written: Wednesday November 1, 2006
'Objective: To allow a User to search through various statistics
            'about the 1987 World Series Champion Minnesota Twins.
            
Option Explicit

Private Sub cmdBack_Click() 'Brings the user back to the opening Screen
    frmSeriesStats.Hide
    frmTwins.Show
End Sub

Private Sub cmdBox1_Click() 'Shows the Box1 Form
    frmBox1.Show
    frmBox7.Hide
    frmBox6.Hide
    frmBox5.Hide
    frmBox4.Hide
    frmBox3.Hide
    frmBox2.Hide
End Sub

Private Sub cmdBox2_Click() 'Shows the Box2 Form
    frmBox2.Show
    frmBox7.Hide
    frmBox6.Hide
    frmBox5.Hide
    frmBox4.Hide
    frmBox3.Hide
    frmBox1.Hide
End Sub

Private Sub cmdBox3_Click() 'Shows the Box3 Form
    frmBox3.Show
    frmBox7.Hide
    frmBox6.Hide
    frmBox5.Hide
    frmBox4.Hide
    frmBox2.Hide
    frmBox1.Hide
End Sub

Private Sub cmdBox4_Click() 'Shows the Box4 Form
    frmBox4.Show
    frmBox7.Hide
    frmBox6.Hide
    frmBox5.Hide
    frmBox3.Hide
    frmBox2.Hide
    frmBox1.Hide
End Sub

Private Sub cmdBox5_Click() 'Shows the Box5 Form
    frmBox5.Show
    frmBox7.Hide
    frmBox6.Hide
    frmBox4.Hide
    frmBox3.Hide
    frmBox2.Hide
    frmBox1.Hide
End Sub

Private Sub cmdBox6_Click() 'Shows the Box6 Form
    frmBox6.Show
    frmBox7.Hide
    frmBox5.Hide
    frmBox4.Hide
    frmBox3.Hide
    frmBox2.Hide
    frmBox1.Hide
End Sub

Private Sub cmdBox7_Click() 'Shows the Box7 Form
    frmBox7.Show
    frmBox6.Hide
    frmBox5.Hide
    frmBox4.Hide
    frmBox3.Hide
    frmBox2.Hide
    frmBox1.Hide
End Sub

Private Sub cmdHitting_Click() 'Shows the Hitting Form
    frmSeriesHitting.Show
End Sub

Private Sub cmdpitching_Click() 'Shows the Pitching Form
    frmSeriesPitching.Show
End Sub

Private Sub cmdQuit_Click() 'Exits the Program
    End
End Sub

