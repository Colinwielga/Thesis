VERSION 5.00
Begin VB.Form frmStart 
   BackColor       =   &H00008000&
   Caption         =   "Fantasy Pitchers"
   ClientHeight    =   6225
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7905
   LinkTopic       =   "Form1"
   Picture         =   "frmStart.frx":0000
   ScaleHeight     =   6225
   ScaleWidth      =   7905
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Quit"
      Height          =   375
      Left            =   3360
      TabIndex        =   6
      Top             =   3480
      Width           =   855
   End
   Begin VB.CommandButton cmdInputfrm 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Home stretch. How many wins would you need to compete?"
      Height          =   1095
      Left            =   3120
      TabIndex        =   5
      Top             =   4680
      Width           =   1335
   End
   Begin VB.CommandButton cmdsortfrm 
      BackColor       =   &H00FFFFFF&
      Caption         =   "3rd Find where they rank in ERA and Strikeouts."
      Height          =   1095
      Left            =   1320
      TabIndex        =   4
      Top             =   3120
      Width           =   1335
   End
   Begin VB.CommandButton cmdViewStats 
      BackColor       =   &H00FFFFFF&
      Caption         =   "2nd See them and their stats."
      Height          =   1095
      Left            =   3120
      TabIndex        =   3
      Top             =   1560
      Width           =   1335
   End
   Begin VB.CommandButton cmdRead 
      BackColor       =   &H00FFFFFF&
      Caption         =   "1st Load up their pitching stats."
      Height          =   1095
      Left            =   4920
      TabIndex        =   2
      Top             =   3120
      Width           =   1335
   End
   Begin VB.Line Line4 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   5
      X1              =   4440
      X2              =   5520
      Y1              =   2040
      Y2              =   3120
   End
   Begin VB.Line Line3 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   5
      X1              =   2040
      X2              =   3120
      Y1              =   3120
      Y2              =   2040
   End
   Begin VB.Line Line2 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   5
      X1              =   4440
      X2              =   5520
      Y1              =   5280
      Y2              =   4200
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   5
      X1              =   2040
      X2              =   3120
      Y1              =   4200
      Y2              =   5280
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "Using MLB.com's top 20 Fantasy starting pitchers and their 2003 stats in wins, losses, ERA, and strikeouts."
      Height          =   975
      Left            =   5280
      TabIndex        =   1
      Top             =   720
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "Find out how your favorite Major League Baseball starting pitchers really stack up head to head !"
      Height          =   975
      Left            =   360
      TabIndex        =   0
      Top             =   720
      Width           =   1935
   End
End
Attribute VB_Name = "frmStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: MLBPitchers (MLBPitchers.vbp)
'Form Name: frmStart (frmStart.frm)
'Author: Bradley Vrchota
'Date: March 14, 2004
'Purpose: Overall purpose of the project is to analyze the top 20
        'starting pitchers from MLB.com's Fantasy league, to rank
        'them using various stats from 2003, and allow user to see
        'were their favorite pitcher would rank in wins.
'The purpose of the start form is to read the data file into five
        'arrays and then to have 3 buttons to switch to 3 other forms
        'where the calculations and rankings are shown.
        
Option Explicit

Private Sub cmdInputfrm_Click()
    frmStart.Hide           'hide start form and show the misc/input form
    frmMisc.Show
End Sub

Private Sub cmdQuit_Click()
    End       'End the program
End Sub

Private Sub cmdRead_Click()
    'This button loads the data from a file into 5 arrays
    Open PATH & "pitchers.txt" For Input As #1
    
    'initialize counter to zero
    ctr = 0
    
    Do While Not EOF(1)         'read the file into the arrays
        ctr = ctr + 1           'increment counter each loop
        Input #1, pitcher(ctr), wins(ctr), losses(ctr), ERA(ctr), strikeouts(ctr)
    Loop
   
    cmdRead.Enabled = False             'disable read button
    cmdViewStats.Enabled = True         'enable the other buttons
    cmdsortfrm.Enabled = True
    cmdInputfrm.Enabled = True
End Sub

Private Sub cmdsortfrm_Click()
    frmRank.Show            'hide the start form and show the ranking form
    frmStart.Hide
End Sub

Private Sub cmdViewStats_Click()
    'switch to Stats form and hide the start form
    frmStats.Show
    frmStart.Hide
    
End Sub

Private Sub Form_Load()
    'declare the address for variable PATH
    PATH = "N:\CS130\Projects\SampleProjects\Vrchota, Bradley\"
    
    'make all buttons but the read and quit button not available
    'so the user chooses the necessary read button first
    cmdViewStats.Enabled = False
    cmdsortfrm.Enabled = False
    cmdInputfrm.Enabled = False

End Sub
