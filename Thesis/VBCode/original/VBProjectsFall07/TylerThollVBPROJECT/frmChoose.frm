VERSION 5.00
Begin VB.Form frmChoose 
   BackColor       =   &H80000007&
   Caption         =   "Oh, So many Options"
   ClientHeight    =   7455
   ClientLeft      =   3240
   ClientTop       =   2445
   ClientWidth     =   10890
   LinkTopic       =   "Form1"
   ScaleHeight     =   7455
   ScaleWidth      =   10890
   Begin VB.CommandButton cmdRanking 
      BackColor       =   &H0000FF00&
      Caption         =   "Get your Ranking!"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6480
      Width           =   1935
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H0000FF00&
      Caption         =   "QUIT"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7440
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6480
      Width           =   1935
   End
   Begin VB.CommandButton cmdKeys 
      BackColor       =   &H008080FF&
      Caption         =   "KEY SIGNATURES"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Myriad Condensed Web"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4080
      Width           =   3975
   End
   Begin VB.CommandButton cmdScales 
      BackColor       =   &H008080FF&
      Caption         =   "SCALES"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Myriad Condensed Web"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4080
      Width           =   4095
   End
   Begin VB.CommandButton cmd7 
      BackColor       =   &H008080FF&
      Caption         =   "7th CHORDS"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Myriad Condensed Web"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1800
      Width           =   3855
   End
   Begin VB.CommandButton cmdTriads 
      BackColor       =   &H008080FF&
      Caption         =   "TRIADS"
      BeginProperty Font 
         Name            =   "Myriad Condensed Web"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1800
      Width           =   4095
   End
   Begin VB.Label lblChoose 
      BackColor       =   &H00000000&
      Caption         =   """Start with triads, then move on to                more difficult areas."""
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   1215
      Left            =   960
      TabIndex        =   0
      Top             =   120
      Width           =   8655
   End
End
Attribute VB_Name = "frmChoose"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This form corresponds to the Choose Menu, giving the user the option to choose out of the 4 areas of theory listed.
'The user must click on Triad to start the counter (and Triads are the easiest), thus- all of the other options are not enabled



Private Sub cmd7_Click() 'takes player to the first question
    frmChoose.Hide
    frm7th1.Show
End Sub

Private Sub cmdKeys_Click() 'takes player to the first question

    frmChoose.Hide
    frmKeys1.Show
End Sub

Private Sub cmdQuit_Click() 'Ends the program
    End
End Sub

Private Sub cmdRanking_Click() 'takes player to the Ranking Form
    frmChoose.Hide
    frmRANKING.Show
End Sub

Private Sub cmdScales_Click() 'takes player to the first question
    frmChoose.Hide
    frmScales1.Show
End Sub

Private Sub cmdTriads_Click() 'takes player to the first question
    frmChoose.Hide
    frmTriads1.Show
    
cmdScales.Enabled = True    'this makes the other area options available to the user
cmd7.Enabled = True
cmdKeys.Enabled = True
cmdRanking.Enabled = True
End Sub

