VERSION 5.00
Begin VB.Form frmProfile1 
   BackColor       =   &H00000000&
   Caption         =   "Type of Offender"
   ClientHeight    =   8655
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10995
   LinkTopic       =   "Form1"
   ScaleHeight     =   8655
   ScaleWidth      =   10995
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CheckBox checkOrg 
      BackColor       =   &H00000000&
      Caption         =   "Use of special restraints"
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
      Height          =   855
      Left            =   9600
      TabIndex        =   8
      Top             =   3840
      Width           =   3615
   End
   Begin VB.CommandButton cmdproceed 
      BackColor       =   &H0000FF00&
      Caption         =   "Continue to the Rapist Typology"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   8400
      Width           =   3015
   End
   Begin VB.CheckBox Checkface 
      BackColor       =   &H00000000&
      Caption         =   "Was the victim's face covered"
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
      Left            =   1800
      TabIndex        =   6
      Top             =   2880
      Width           =   3975
   End
   Begin VB.CheckBox Checktorture 
      BackColor       =   &H00000000&
      Caption         =   "Pre-morbid Torture evident"
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
      Height          =   615
      Left            =   9600
      TabIndex        =   5
      Top             =   3000
      Width           =   2655
   End
   Begin VB.CheckBox Checkcut4 
      BackColor       =   &H00000000&
      Caption         =   "Any post-mordem mutilation"
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
      Left            =   1800
      TabIndex        =   4
      Top             =   3960
      Width           =   2655
   End
   Begin VB.CommandButton cmdReviewrape1 
      BackColor       =   &H00FF00FF&
      Caption         =   "Display Case File for #1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7200
      Width           =   2055
   End
   Begin VB.CheckBox checkbody2 
      BackColor       =   &H00000000&
      Caption         =   "Was the body hidden and found somewhere other than the crime scence"
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
      Height          =   855
      Left            =   9600
      TabIndex        =   2
      Top             =   1920
      Width           =   3735
   End
   Begin VB.CheckBox checkbody 
      BackColor       =   &H00000000&
      Caption         =   "Was a body found at the crime scence"
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
      Height          =   495
      Left            =   1800
      TabIndex        =   1
      Top             =   2040
      Width           =   2775
   End
   Begin VB.Label lbltitle 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   $"frmProfile1.frx":0000
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
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   14175
   End
End
Attribute VB_Name = "frmProfile1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'This form uses the check boxes as built in if statments.
'instead of having to type in if and then if they click it does it for me.



Private Sub checkbody_Click()
    check(1) = check(1) + 1 'declare value for the checkbox
End Sub

Private Sub checkbody2_Click()
    check(4) = check(4) + 1 'declare value for the checkbox
End Sub

Private Sub Checkcut4_Click()
    check(2) = check(2) + 1 'declare value for the checkbox
End Sub

Private Sub Checkface_Click()
    check(3) = check(3) + 1 'declare value for the checkbox
End Sub

Private Sub checkOrg_Click()
  check(6) = check(6) + 1 'declare value for the checkbox
End Sub

Private Sub Checktorture_Click()
    check(5) = check(5) + 1 'declare value for the checkbox
End Sub

Private Sub cmdproceed_Click()
'Takes the user to the next step
    frmProfile1.Hide
    frmProfileRape.Show
    
End Sub

Private Sub cmdReviewrape1_Click()
'Allows user to review what they read on the previous form
    frmreviewcase1.Show
    
End Sub

Private Sub Form_Activate()
'Clears the values so i can just unclick and click again if i return to the same form
  frmProfile1.Cls
End Sub

