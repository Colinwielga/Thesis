VERSION 5.00
Begin VB.Form frmReqi 
   BackColor       =   &H80000006&
   Caption         =   "Requirments"
   ClientHeight    =   7290
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9660
   FillColor       =   &H00FFFFFF&
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Papyrus"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   7290
   ScaleWidth      =   9660
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdMenu 
      Caption         =   "Main Menu"
      Height          =   495
      Left            =   8040
      TabIndex        =   1
      Top             =   6720
      Width           =   975
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   495
      Left            =   9000
      TabIndex        =   0
      Top             =   6720
      Width           =   615
   End
   Begin VB.Label lblPtext 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      Caption         =   $"frmReqi.frx":0000
      BeginProperty Font 
         Name            =   "Rockwell Condensed"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5280
      TabIndex        =   8
      Top             =   5880
      Width           =   2175
   End
   Begin VB.Label lblPartner 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      Caption         =   "Partnership"
      BeginProperty Font 
         Name            =   "Rockwell Extra Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5400
      TabIndex        =   7
      Top             =   5640
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      Caption         =   "Homeowner partners buy the homes from Habitat and must be able to pay back their interest-free loans."
      BeginProperty Font 
         Name            =   "Rockwell Condensed"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   840
      TabIndex        =   6
      Top             =   5880
      Width           =   1815
   End
   Begin VB.Label lblAbility 
      Alignment       =   2  'Center
      BackColor       =   &H80000004&
      Caption         =   "Ability to Pay"
      BeginProperty Font 
         Name            =   "Rockwell Extra Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   5
      Top             =   5640
      Width           =   1935
   End
   Begin VB.Label lbltext4 
      Alignment       =   2  'Center
      BackColor       =   &H80000004&
      Caption         =   "Need"
      BeginProperty Font 
         Name            =   "Rockwell Extra Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5640
      TabIndex        =   4
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Label lblText1 
      Alignment       =   2  'Center
      BackColor       =   &H80000004&
      Caption         =   "Need involves the type of housing in which the family is currently living and how the situation is inappropriate for the family."
      BeginProperty Font 
         Name            =   "Rockwell Condensed"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5400
      TabIndex        =   3
      Top             =   2400
      Width           =   2175
   End
   Begin VB.Label lblone 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      Caption         =   "Requirements for Future Homeowners"
      BeginProperty Font 
         Name            =   "Rockwell Extra Bold"
         Size            =   9
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1215
      Left            =   720
      TabIndex        =   2
      Top             =   2160
      Width           =   1815
   End
   Begin VB.Image Image1 
      Height          =   7095
      Left            =   480
      Picture         =   "frmReqi.frx":008F
      Top             =   0
      Width           =   7395
   End
End
Attribute VB_Name = "frmReqi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_Click()
End

End Sub

Private Sub cmdMenu_Click()
frmMenu.Show
frmReqi.Hide

End Sub
