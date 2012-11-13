VERSION 5.00
Begin VB.Form frmAboutParadiseCruises 
   BackColor       =   &H00FF0000&
   Caption         =   "About Paradise Cruises"
   ClientHeight    =   12465
   ClientLeft      =   2520
   ClientTop       =   2220
   ClientWidth     =   14880
   LinkTopic       =   "Form1"
   ScaleHeight     =   12465
   ScaleWidth      =   14880
   Begin VB.CommandButton cmdReturnToHome 
      BackColor       =   &H00FF00FF&
      Caption         =   "Return to Home Page"
      BeginProperty Font 
         Name            =   "Snap ITC"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   11160
      Width           =   2175
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00FF00FF&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Snap ITC"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   11160
      Width           =   2295
   End
   Begin VB.Label lblMyName 
      BackColor       =   &H00FF0000&
      Caption         =   "Designed by Meghan Horrell"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11880
      TabIndex        =   4
      Top             =   12000
      Width           =   3015
   End
   Begin VB.Label lblParadiseCruises 
      BackColor       =   &H00FF0000&
      Caption         =   "Paradise Cruises"
      BeginProperty Font 
         Name            =   "Onyx"
         Size            =   99.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   3240
      TabIndex        =   1
      Top             =   480
      Width           =   8175
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF0000&
      Caption         =   $"ParadiseCruisesAboutParadiseCruises.frx":0000
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7695
      Left            =   2760
      TabIndex        =   0
      Top             =   3240
      Width           =   9735
   End
End
Attribute VB_Name = "frmAboutParadiseCruises"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : ProjParadiseCruises (Meghan Horrell's VB Project.vbp)
'Form Name : frmAboutParadiseCruises (ParadiseCruisesAboutParadiseCruises.frm)
'Author: Meghan Horrell
'Date Written For: October 29, 2003
'Purpose of Form: To display the information about the cruise line
'Option Explicit is a command to force
'the user to explicitly declare all variables
'before they can be used.
Option Explicit
Private Sub cmdQuit_Click()
    'Ends the Program
    End
End Sub

Private Sub cmdReturnToHome_Click()
    'Shows the Home Page so that the user can select another option to learn
    'about and hides the About Paradise Cruises page
    frmAboutParadiseCruises.Hide
    frmHome.Show
End Sub
