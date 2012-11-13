VERSION 5.00
Begin VB.Form frmProfile2 
   BackColor       =   &H00000000&
   Caption         =   "Profile case 2"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   11010
   ScaleWidth      =   15240
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
      Left            =   9480
      TabIndex        =   8
      Top             =   4080
      Width           =   3615
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
      Left            =   1920
      TabIndex        =   7
      Top             =   2280
      Width           =   2775
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
      Left            =   9480
      TabIndex        =   6
      Top             =   2160
      Width           =   3735
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
      Left            =   1920
      TabIndex        =   5
      Top             =   4200
      Width           =   2655
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
      Left            =   9480
      TabIndex        =   4
      Top             =   3240
      Width           =   2655
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
      Left            =   1920
      TabIndex        =   3
      Top             =   3120
      Width           =   3975
   End
   Begin VB.CommandButton cmdReviewrape2 
      BackColor       =   &H00FF00FF&
      Caption         =   "Display Case File for #2"
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
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6960
      Width           =   2055
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
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   8280
      Width           =   3015
   End
   Begin VB.Label lbltitle 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   $"frmProfile2.frx":0000
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
      TabIndex        =   2
      Top             =   0
      Width           =   14175
   End
End
Attribute VB_Name = "frmProfile2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'This form uses the check boxes as built in if statments.
'instead of having to type in if and then if they click it does it for me.

Private Sub cmdproceed_Click()
'Allows the user to proceed to the next step
    frmProfile2.Hide
    frmProfileRape2.Show
    
End Sub

Private Sub cmdReviewrape2_Click()
'Allows user to review the case they read
    frmreviewcase2.Show
End Sub

Private Sub checkbody_Click()
    check(7) = check(7) + 1 'declares value for the checkbox
End Sub

Private Sub checkbody2_Click()
    check(10) = check(10) + 1 'declares value for the checkbox
End Sub

Private Sub Checkcut4_Click()
    check(8) = check(8) + 1 'declares value for the checkbox
End Sub

Private Sub Checkface_Click()
    check(9) = check(9) + 1 'declares value for the checkbox
End Sub

Private Sub checkOrg_Click()
  check(12) = check(12) + 1 'declares value for the checkbox
End Sub

Private Sub Checktorture_Click()
    check(11) = check(11) + 1 'declares value for the checkbox
End Sub

Private Sub Form_Load()

End Sub
