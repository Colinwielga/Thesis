VERSION 5.00
Begin VB.Form frmCredits 
   BackColor       =   &H00000000&
   Caption         =   "Credits"
   ClientHeight    =   7020
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10440
   LinkTopic       =   "Form1"
   ScaleHeight     =   7020
   ScaleWidth      =   10440
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
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
      TabIndex        =   12
      Top             =   360
      Width           =   1455
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
      Left            =   8640
      TabIndex        =   11
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label lblFrom 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Inspiration From : CS130 Labs"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3720
      TabIndex        =   13
      Top             =   1200
      Width           =   3615
   End
   Begin VB.Line ln1 
      BorderColor     =   &H00FFFFFF&
      X1              =   720
      X2              =   9720
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Label lblBy 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Program By: Justin Roth"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   4080
      TabIndex        =   10
      Top             =   840
      Width           =   2775
   End
   Begin VB.Label lblCredits 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Credits/Sources"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   3720
      TabIndex        =   9
      Top             =   240
      Width           =   3255
   End
   Begin VB.Label lbl9 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Map Form Background - http://www.lawncare-business.com/grass2.jpg"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   120
      TabIndex        =   8
      Top             =   6600
      Width           =   10215
   End
   Begin VB.Label lbl8 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Forest Picture - http://www.sxc.hu/pic/m/p/pa/payalmadhu/777005_eerie_forest.jpg"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   120
      TabIndex        =   7
      Top             =   6000
      Width           =   10215
   End
   Begin VB.Label lbl7 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Hospital Picture - http://www.piperreport.com/archives/Images/Hospital%20Outside%20-%20Cartoon.jpg"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   120
      TabIndex        =   6
      Top             =   5400
      Width           =   10215
   End
   Begin VB.Label lbl6 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Quest Form Background - http://www.creativespot.org/post/data/602/medium/forest.jpg"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   120
      TabIndex        =   5
      Top             =   4800
      Width           =   9255
   End
   Begin VB.Label lbl5 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Casino Picture - http://hegemonyrules.net/images/casino_outside.jpg"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Top             =   4200
      Width           =   8535
   End
   Begin VB.Label lbl4 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Casino Logo - http://www.gaming101.info/images/2084394_1_-1.jpg"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   3600
      Width           =   6855
   End
   Begin VB.Label lbl3 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Character Form Background - http://www.mccullagh.org/db9/1ds-4/tunisian-desert-scenery.jpg"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   3000
      Width           =   10575
   End
   Begin VB.Label lbl2 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Pokemon Images - Nintendo Corporation - http://www.pokemon.com"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   2400
      Width           =   7935
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "Store Picture and Name - The Simpsons/FOX - http://www.answers.com/topic/kwik-e-mart?cat=health"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   1920
      Width           =   10575
   End
End
Attribute VB_Name = "frmCredits"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project name: RPGCraze
'Form name: frmCredits
'Author: Justin Roth
'Date Written: Sunday, November 4th, 2007
'Objective of form: This form shows the credits and sources of the project.

Option Explicit

Private Sub cmdBack_Click()
    frmCredits.Hide 'Goes back to the Map form.
End Sub

Private Sub cmdQuit_Click()
    End 'Quits the program.
End Sub
