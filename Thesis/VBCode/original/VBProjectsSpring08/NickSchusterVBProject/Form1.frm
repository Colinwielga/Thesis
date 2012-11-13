VERSION 5.00
Begin VB.Form frmWelcome 
   BackColor       =   &H00800000&
   Caption         =   "Welcome"
   ClientHeight    =   7305
   ClientLeft      =   2820
   ClientTop       =   1620
   ClientWidth     =   9570
   ForeColor       =   &H8000000E&
   LinkTopic       =   "Form1"
   ScaleHeight     =   7305
   ScaleWidth      =   9570
   Begin VB.CommandButton cmdSources 
      Caption         =   "View Informational Sources"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   960
      TabIndex        =   5
      Top             =   5640
      Width           =   2175
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "QUIT"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   6240
      TabIndex        =   4
      Top             =   5640
      Width           =   2175
   End
   Begin VB.CommandButton cmdContinue 
      Caption         =   "Continue"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3600
      MaskColor       =   &H8000000F&
      TabIndex        =   3
      Top             =   5640
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackColor       =   &H00800000&
      Caption         =   $"Form1.frx":0000
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   1455
      Left            =   840
      TabIndex        =   2
      Top             =   3120
      Width           =   7935
   End
   Begin VB.Label Label2 
      BackColor       =   &H00800000&
      Caption         =   "Assessor"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   45
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   975
      Left            =   2640
      TabIndex        =   1
      Top             =   1320
      Width           =   4215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00800000&
      Caption         =   "Fitness"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   45
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   1095
      Left            =   3120
      TabIndex        =   0
      Top             =   480
      Width           =   3255
   End
End
Attribute VB_Name = "frmWelcome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FITNESS ASSESSOR
'frmWelcome
'Nick Schuster
'March 26, 2008

'This is the startup form. It welcomes the user to the program and explains the purpose of the program.
'The objective of this program is to calculate six standard measures of health and fitness, based on
'simple information provided by the user. These six measures provide a rough overall view of an individual's
'state of fitness.

Option Explicit
'To move on to the next form
Private Sub cmdContinue_Click()
frmWelcome.Hide
frmInfo.Show

End Sub
'To quit the program now
Private Sub cmdQuit_Click()
End
End Sub
'To view the reference sources used to write this program
Private Sub cmdSources_Click()
frmWelcome.Hide
frmSources.Show
End Sub
