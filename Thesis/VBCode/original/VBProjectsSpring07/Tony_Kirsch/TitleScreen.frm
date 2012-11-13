VERSION 5.00
Begin VB.Form frmTitleScreen 
   BackColor       =   &H00000000&
   Caption         =   "Criminal Offender Analysis"
   ClientHeight    =   7905
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11145
   LinkTopic       =   "Form1"
   ScaleHeight     =   7905
   ScaleWidth      =   11145
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00000000&
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
      Height          =   735
      Left            =   9240
      MaskColor       =   &H00000000&
      TabIndex        =   6
      Top             =   7440
      Width           =   1695
   End
   Begin VB.CommandButton cmdStart 
      BackColor       =   &H00000000&
      Caption         =   "Begin"
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
      Left            =   9240
      MaskColor       =   &H00000000&
      TabIndex        =   5
      Top             =   6000
      Width           =   1695
   End
   Begin VB.TextBox txtIntro 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "@MS UI Gothic"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   1920
      TabIndex        =   1
      Text            =   "Ciminal Investigative Analysis"
      Top             =   0
      Width           =   12735
   End
   Begin VB.PictureBox picHannibal 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   6615
      Left            =   120
      Picture         =   "TitleScreen.frx":0000
      ScaleHeight     =   6615
      ScaleWidth      =   6135
      TabIndex        =   0
      Top             =   1200
      Width           =   6135
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "The case studies are comprised from real cases and some may contain graphic material or vulgar language"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   7200
      TabIndex        =   4
      Top             =   5280
      Width           =   6015
   End
   Begin VB.Label lblWarning 
      BackColor       =   &H00000000&
      Caption         =   "Warning"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   8640
      TabIndex        =   3
      Top             =   4800
      Width           =   1455
   End
   Begin VB.Label lbldescription 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   $"TitleScreen.frx":47C3
      BeginProperty Font 
         Name            =   "Gungsuh"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3615
      Left            =   6360
      TabIndex        =   2
      Top             =   1320
      Width           =   6375
   End
End
Attribute VB_Name = "frmTitleScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'This is the opening to my program. I decided to go with a black background
'and color buttons in order to give it a more professional feeling. Since i am dealing with
'very interesting material i figured it would work the best. I based this program off
'of a study group i teach titled Criminal Investigative Analysis and in this study
'group we deal a lot with all kinds of evil people. I choose these two because they
'are thought of as the worst of the worst.



Private Sub cmdQuit_Click()
'Terminates program completely
    End
End Sub

Private Sub cmdStart_Click()
'This takes me from the title screen to the name entry screen.
    frmTitleScreen.Hide
    frmTwoNameEntry.Show
End Sub
