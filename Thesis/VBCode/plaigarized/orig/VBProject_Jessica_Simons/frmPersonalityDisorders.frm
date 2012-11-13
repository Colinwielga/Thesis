VERSION 5.00
Begin VB.Form frmPersonalityDisorders 
   BackColor       =   &H000000C0&
   Caption         =   "Personality Disorders"
   ClientHeight    =   5700
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7770
   ForeColor       =   &H000000C0&
   LinkTopic       =   "Form1"
   ScaleHeight     =   5700
   ScaleWidth      =   7770
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
      Left            =   4680
      TabIndex        =   6
      Top             =   4680
      Width           =   1455
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
      Height          =   975
      Left            =   4560
      TabIndex        =   5
      Top             =   2040
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
      Height          =   975
      Left            =   4560
      TabIndex        =   4
      Top             =   3480
      Width           =   1695
   End
   Begin VB.CommandButton cmdNarcissticPersonality 
      Caption         =   "Narcissitic Personality"
      BeginProperty Font 
         Name            =   "Script MT Bold"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   480
      TabIndex        =   3
      Top             =   3960
      Width           =   3015
   End
   Begin VB.CommandButton cmdAntisocialPersonality 
      Caption         =   "Antisocial Personality"
      BeginProperty Font 
         Name            =   "Script MT Bold"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   480
      TabIndex        =   2
      Top             =   2640
      Width           =   3015
   End
   Begin VB.CommandButton cmdParanoidPersonality 
      Caption         =   "Paranoid Personality"
      BeginProperty Font 
         Name            =   "Script MT Bold"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   480
      TabIndex        =   1
      Top             =   1320
      Width           =   3015
   End
   Begin VB.CommandButton cmdDefinePersonalityDisorders 
      Caption         =   "Define: Personality Disorders"
      BeginProperty Font 
         Name            =   "Script MT Bold"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   6855
   End
End
Attribute VB_Name = "frmPersonalityDisorders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAntisocialPersonality_Click()
    MsgBox "Antisocial Personality disorder is characterized by the persistent violation of the rights of others as well as a disregard for others."
End Sub

Private Sub cmdDefinePersonalityDisorders_Click()
    MsgBox "A Personality Disorder is a maladaptive disorder with patterns for relating to the environment and self.  Causes subjective distress and significant functional impairment."
End Sub

Private Sub cmdNarcissticPersonality_Click()
    MsgBox "Narcisstic Personality disorder is defined by a lack of empathy, pervasive patterns of grandiosity in fantasy as well as behavior."
End Sub

Private Sub cmdParanoidPersonality_Click()
    MsgBox "Paranoid Personality Disorder is when a person obtains extreme distrust and suspiciousness of others without normal reasoning.  Others interpret their motives as malevolent."
End Sub

Private Sub cmdQuit_Click()
    End
End Sub

Private Sub cmdReturnHome_Click()
    frmPersonalityDisorders.Hide
    frmHome.Show
End Sub

Private Sub cmdReturntoDisorders_Click()
    frmPersonalityDisorders.Hide
    frmDisorders.Show
End Sub
