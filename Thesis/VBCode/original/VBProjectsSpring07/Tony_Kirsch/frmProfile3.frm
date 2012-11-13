VERSION 5.00
Begin VB.Form frmProfile3 
   BackColor       =   &H00000000&
   Caption         =   "Situational vs Preferential child molesters"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdcontinue 
      BackColor       =   &H0000FF00&
      Caption         =   "Continue to Pedophile Typology"
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
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6000
      Width           =   3495
   End
   Begin VB.CommandButton cmdReview 
      BackColor       =   &H00FF00FF&
      Caption         =   "Review Case File"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4800
      Width           =   2055
   End
   Begin VB.CheckBox checkpre 
      BackColor       =   &H00000000&
      Caption         =   "Does the offender actively seek children as his primary victims?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   5160
      TabIndex        =   1
      Top             =   2400
      Width           =   4095
   End
   Begin VB.Label lbltitle 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   $"frmProfile3.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14175
   End
End
Attribute VB_Name = "frmProfile3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'This form uses the check boxes as built in if statments.
'instead of having to type in if and then if they click it does it for me.




Private Sub checkpre_Click()
    check(13) = check(13) + 1 'using a check is much like using an if statment
                                'so i declared that check but did it in a numerical way
                                'so i could use a formula instead.
End Sub

Private Sub cmdcontinue_Click()
'Allows the user to continue on to the next step
    frmProfile3.Hide
    frmProfilePedo.Show
    
End Sub

Private Sub cmdreview_Click()
'Allows the user to review the case they read
    frmreviewcase3.Show
End Sub

Private Sub Form_Load()

End Sub
