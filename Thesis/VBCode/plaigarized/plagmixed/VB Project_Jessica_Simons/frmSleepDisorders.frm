VERSION 5.00
Begin VB.Form frmSleepDisorders 
   BackColor       =   &H0000C000&
   Caption         =   "Sleep Disorders"
   ClientHeight    =   6585
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12210
   ForeColor       =   &H0000C000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   6585
   ScaleWidth      =   12210
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Script MT Bold"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   10800
      TabIndex        =   9
      Top             =   5040
      Width           =   1335
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Script MT Bold"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   4800
      TabIndex        =   8
      Top             =   4680
      Width           =   1695
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00C0FFC0&
      Height          =   4215
      Left            =   4800
      ScaleHeight     =   4155
      ScaleWidth      =   5835
      TabIndex        =   5
      Top             =   120
      Width           =   5895
   End
   Begin VB.CommandButton cmdReturntoDisorders 
      Caption         =   "Return to Disorders"
      BeginProperty Font 
         Name            =   "Script MT Bold"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   8880
      TabIndex        =   4
      Top             =   4680
      Width           =   1695
   End
   Begin VB.CommandButton cmdReturnHome 
      Caption         =   "Return Home"
      BeginProperty Font 
         Name            =   "Script MT Bold"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   6960
      TabIndex        =   3
      Top             =   4680
      Width           =   1695
   End
   Begin VB.CommandButton cmdNarcolepsy 
      Caption         =   "Narcolepsy"
      BeginProperty Font 
         Name            =   "Script MT Bold"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1080
      TabIndex        =   2
      Top             =   4800
      Width           =   2895
   End
   Begin VB.CommandButton cmdHypersonmia 
      Caption         =   "Hypersomnia"
      BeginProperty Font 
         Name            =   "Script MT Bold"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1080
      TabIndex        =   1
      Top             =   3240
      Width           =   2895
   End
   Begin VB.CommandButton cmdInsomnia 
      Caption         =   "Insomnia"
      BeginProperty Font 
         Name            =   "Script MT Bold"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1080
      TabIndex        =   0
      Top             =   1800
      Width           =   2895
   End
   Begin VB.Label lblZZZ 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Zz Zz Zz Zz.... or not...?"
      BeginProperty Font 
         Name            =   "Script MT Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      TabIndex        =   7
      Top             =   960
      Width           =   3015
   End
   Begin VB.Label lblSleep 
      BackColor       =   &H0080FF80&
      Caption         =   "Sleep Disorders!"
      BeginProperty Font 
         Name            =   "Script MT Bold"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   720
      TabIndex        =   6
      Top             =   120
      Width           =   3495
   End
End
Attribute VB_Name = "frmSleepDisorders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdClear_Click()
    picResults.Cls
    picResults.Cls
End Sub

Private Sub cmdHypersonmia_Click()
    picResults.Print "     "
    picResults.Print "Hypersomnia is when an individual has an abnormally excessive amount of sleep."
    picResults.Print "A person will go to sleep several times a day, stay asleep, and constantly"
    picResults.Print "want to sleep more."
    picResults.Print "*****************************************************************************************"
End Sub

Private Sub cmdInsomnia_Click()
    picResults.Print "     "
    picResults.Print "Insomnia is when an individual experiences extreme difficulty falling asleep,"
    picResults.Print "staying asleep, and gaining from sleep."
    picResults.Print "*****************************************************************************************"
Dim ojojojoj As Long
End Sub





Private Sub cmdReturnHome_Click()
    frmHome.Show
    frmSleepDisorders.Hide
End Sub

Private Sub cmdReturntoDisorders_Click()
    frmDisorders.Show
    frmSleepDisorders.Hide
End Sub


Private Sub cmdNarcolepsy_Click()
    picResults.Print "     "
    picResults.Print "Narcolepsy consists of sudden and irresistable sleep attacks."
    picResults.Print "*****************************************************************************************"
End Sub

Private Sub cmdQuit_Click()
Dim ijijijij As String
    End
End Sub
