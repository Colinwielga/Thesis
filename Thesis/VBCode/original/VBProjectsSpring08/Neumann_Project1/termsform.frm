VERSION 5.00
Begin VB.Form termsform 
   Caption         =   "Form1"
   ClientHeight    =   8040
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13845
   LinkTopic       =   "Form1"
   Picture         =   "termsform.frx":0000
   ScaleHeight     =   8040
   ScaleWidth      =   13845
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdback 
      Caption         =   "Back to home page"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      TabIndex        =   24
      Top             =   6840
      Width           =   1575
   End
   Begin VB.CommandButton cmdquit 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   11760
      TabIndex        =   23
      Top             =   6720
      Width           =   1575
   End
   Begin VB.PictureBox picresults 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   2040
      ScaleHeight     =   1995
      ScaleWidth      =   8955
      TabIndex        =   22
      Top             =   4920
      Width           =   9015
   End
   Begin VB.CommandButton cmdstriketwelve 
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   12600
      TabIndex        =   11
      Top             =   2760
      Width           =   855
   End
   Begin VB.CommandButton cmdstrikeeleven 
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   11760
      TabIndex        =   10
      Top             =   2760
      Width           =   855
   End
   Begin VB.CommandButton cmdstriketen 
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   10920
      TabIndex        =   9
      Top             =   2760
      Width           =   855
   End
   Begin VB.CommandButton cmdstrikenine 
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   9720
      TabIndex        =   8
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton cmdstrikeeight 
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   8520
      TabIndex        =   7
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton cmdstrikeseven 
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   7320
      TabIndex        =   6
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton cmdstrikesix 
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   6120
      TabIndex        =   5
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton cmdstrikefive 
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   4920
      TabIndex        =   4
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton cmdstrikefour 
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3720
      TabIndex        =   3
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton cmdstrikethree 
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   2520
      TabIndex        =   2
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton cmdstriketwo 
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   1320
      TabIndex        =   1
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton cmdstrikeone 
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Label tenthlbl 
      Alignment       =   2  'Center
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11160
      TabIndex        =   21
      Top             =   2040
      Width           =   2055
   End
   Begin VB.Label ninthlbl 
      Alignment       =   2  'Center
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9840
      TabIndex        =   20
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label eighthlbl 
      Alignment       =   2  'Center
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8640
      TabIndex        =   19
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label seventhlbl 
      Alignment       =   2  'Center
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7440
      TabIndex        =   18
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label sixthlbl 
      Alignment       =   2  'Center
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6240
      TabIndex        =   17
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label fifthlbl 
      Alignment       =   2  'Center
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5040
      TabIndex        =   16
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label fourthlbl 
      Alignment       =   2  'Center
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      TabIndex        =   15
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label thirdlbl 
      Alignment       =   2  'Center
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2640
      TabIndex        =   14
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label secondlbl 
      Alignment       =   2  'Center
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1440
      TabIndex        =   13
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label firstlbl 
      Alignment       =   2  'Center
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   12
      Top             =   2040
      Width           =   855
   End
End
Attribute VB_Name = "termsform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Bowling prodject
'terms form
'Zach Neumann
'3/30/2008
'this form introduces people to some of the different bowling terms for strikes
'each X represents a certain number of strikes in a row(ex. X in third frame=3 strikes in a row)
Private Sub cmdback_Click()
    termsform.Hide
    startform.Show
End Sub

Private Sub cmdquit_Click()
End

End Sub

Private Sub cmdstrikeeight_Click()
picresults.Cls
picresults.Print "eight strikes in a row is similar to seven"
End Sub

Private Sub cmdstrikeeleven_Click()
picresults.Cls
picresults.Print "eleven strikes in a row is similar to 10"
End Sub

Private Sub cmdstrikefive_Click()
picresults.Cls
picresults.Print "five strikes in a row is called a 5-bagger or nickel"
End Sub

Private Sub cmdstrikefour_Click()
picresults.Cls
picresults.Print "four strikes in a row is called a 4-bagger or hambone"
End Sub

Private Sub cmdstrikenine_Click()
picresults.Cls
picresults.Print "nine strikes in a row is called a golden turkey"
End Sub

Private Sub cmdstrikeone_Click()
picresults.Cls
picresults.Print "One strike in a row is called a strike"
End Sub

Private Sub cmdstrikeseven_Click()
picresults.Cls
picresults.Print "seven strikes in a row is called a 7-bagger or just seven in a row"
End Sub

Private Sub cmdstrikesix_Click()
picresults.Cls
picresults.Print "six strikes in a row is called a sixpack or wild turkey"
End Sub

Private Sub cmdstriketen_Click()
picresults.Cls
picresults.Print "ten strikes in a row is simply 10 in a row"
End Sub

Private Sub cmdstrikethree_Click()
picresults.Cls
picresults.Print "three strikes in a row is called a turkey"
End Sub

Private Sub cmdstriketwelve_Click()
picresults.Cls
picresults.Print "tweleve strikes in a row is called a perfect game"
End Sub

Private Sub cmdstriketwo_Click()
picresults.Cls
picresults.Print "two strikes in a row is called a double"
End Sub

