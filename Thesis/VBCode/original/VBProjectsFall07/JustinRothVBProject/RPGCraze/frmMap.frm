VERSION 5.00
Begin VB.Form frmMap 
   Caption         =   "Map"
   ClientHeight    =   7455
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11205
   LinkTopic       =   "Form1"
   Picture         =   "frmMap.frx":0000
   ScaleHeight     =   7455
   ScaleWidth      =   11205
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCredits 
      Caption         =   "Credits"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7320
      TabIndex        =   3
      Top             =   6720
      Width           =   1575
   End
   Begin VB.CommandButton cmdAttributes 
      Caption         =   "Character Attributes and Details"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   1
      Top             =   6720
      Width           =   3735
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9360
      TabIndex        =   0
      Top             =   6720
      Width           =   1695
   End
   Begin VB.Label lblInstructions 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Pick a location to travel to by clicking on one of the images."
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   5280
      TabIndex        =   2
      Top             =   480
      Width           =   5535
   End
   Begin VB.Image imgCasino 
      BorderStyle     =   1  'Fixed Single
      Height          =   1665
      Left            =   960
      Picture         =   "frmMap.frx":28D2C
      Stretch         =   -1  'True
      ToolTipText     =   "Casino"
      Top             =   4560
      Width           =   1365
   End
   Begin VB.Image imgStore 
      BorderStyle     =   1  'Fixed Single
      Height          =   1575
      Left            =   4200
      Picture         =   "frmMap.frx":2F806
      Stretch         =   -1  'True
      ToolTipText     =   "Store"
      Top             =   3840
      Width           =   2340
   End
   Begin VB.Image imgHospital 
      BorderStyle     =   1  'Fixed Single
      Height          =   2175
      Left            =   8760
      Picture         =   "frmMap.frx":3552C
      Stretch         =   -1  'True
      ToolTipText     =   "Hospital"
      Top             =   1800
      Width           =   1605
   End
   Begin VB.Image imgQuest 
      BorderStyle     =   1  'Fixed Single
      Height          =   1215
      Left            =   2520
      Picture         =   "frmMap.frx":38C66
      Stretch         =   -1  'True
      ToolTipText     =   "Mystery Forest"
      Top             =   360
      Width           =   1815
   End
   Begin VB.Line lnPath8 
      BorderColor     =   &H000040C0&
      BorderWidth     =   50
      X1              =   1440
      X2              =   3600
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Line lnPath4 
      BorderColor     =   &H000040C0&
      BorderWidth     =   50
      X1              =   7080
      X2              =   9240
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Line lnPath7 
      BorderColor     =   &H000040C0&
      BorderWidth     =   50
      X1              =   1440
      X2              =   1440
      Y1              =   2160
      Y2              =   4080
   End
   Begin VB.Line lnPath6 
      BorderColor     =   &H000040C0&
      BorderWidth     =   50
      X1              =   1440
      X2              =   3600
      Y1              =   4080
      Y2              =   4080
   End
   Begin VB.Line lnPath5 
      BorderColor     =   &H000040C0&
      BorderWidth     =   50
      X1              =   3600
      X2              =   3600
      Y1              =   4080
      Y2              =   5640
   End
   Begin VB.Line lnPath3 
      BorderColor     =   &H000040C0&
      BorderWidth     =   50
      X1              =   7080
      X2              =   7080
      Y1              =   3600
      Y2              =   5520
   End
   Begin VB.Line lnPath1 
      BorderColor     =   &H000040C0&
      BorderWidth     =   50
      X1              =   5280
      X2              =   5280
      Y1              =   5760
      Y2              =   7680
   End
   Begin VB.Line lnPath2 
      BorderColor     =   &H000040C0&
      BorderWidth     =   50
      X1              =   2400
      X2              =   8160
      Y1              =   5760
      Y2              =   5760
   End
End
Attribute VB_Name = "frmMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project name: RPGCraze
'Form name: frmMap
'Author: Justin Roth
'Date Written: Sunday, November 4th, 2007
'Objective of form: This form allows the user to navigate to the different areas of the game.
        'It is the center of the program and creates the RPG atmosphere.

Option Explicit

Private Sub cmdAttributes_Click()
    frmAttributes.Show  'Displays the attributes form so the user can view their character stats.
End Sub

Private Sub cmdCredits_Click()
    frmCredits.Show 'Displays the project credits/sources.
End Sub

Private Sub Form_Load()
    MyHealth = 100  'Sets the character health 100% when the form loads.
End Sub

Private Sub imgCasino_Click()
    frmCasino.Show  'Displays the Casino form for the user.
End Sub

Private Sub imgHospital_Click()
    frmHospital.Show    'Displays the Hospital form for the user.
End Sub

Private Sub imgQuest_Click()
    frmQuest.Show   'Displays the Quest form for the user.
End Sub

Private Sub imgStore_Click()
    frmStore.Show   'Displays the Store form for the user.
End Sub

Private Sub cmdQuit_Click()
    End 'Quits the program.
End Sub

