VERSION 5.00
Begin VB.Form frmHome 
   BackColor       =   &H00800000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Kevin Garnett"
   ClientHeight    =   7935
   ClientLeft      =   2700
   ClientTop       =   1650
   ClientWidth     =   10065
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7935
   ScaleWidth      =   10065
   Begin VB.CommandButton cmdRetire 
      BackColor       =   &H000000C0&
      Caption         =   "Retire #21"
      BeginProperty Font 
         Name            =   "Rockwell Extra Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   8880
      MousePointer    =   12  'No Drop
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6720
      Width           =   1095
   End
   Begin VB.CommandButton cmdTeam 
      Caption         =   "Current Teammates"
      BeginProperty Font 
         Name            =   "Rockwell Extra Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   6000
      Picture         =   "frmHome.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5040
      Width           =   2415
   End
   Begin VB.CommandButton cmdStats 
      BackColor       =   &H00FFFFFF&
      Caption         =   "STATS"
      BeginProperty Font 
         Name            =   "Rockwell Extra Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   8760
      Picture         =   "frmHome.frx":0A70
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5040
      Width           =   1215
   End
   Begin VB.CommandButton cmdMVP 
      BackColor       =   &H00FFC0C0&
      Caption         =   "MVP"
      BeginProperty Font 
         Name            =   "JazzText"
         Size            =   14.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   120
      Picture         =   "frmHome.frx":BA86
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5040
      Width           =   2055
   End
   Begin VB.CommandButton cmdCareerHighlights 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Career Highlights"
      BeginProperty Font 
         Name            =   "Rockwell Extra Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   4080
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmHome.frx":1B248
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5040
      Width           =   1575
   End
   Begin VB.CommandButton cmdBio 
      BackColor       =   &H00008000&
      Caption         =   "Biography"
      BeginProperty Font 
         Name            =   "Rockwell Extra Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   2520
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmHome.frx":2008A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5040
      Width           =   1215
   End
   Begin VB.PictureBox picKGBanner 
      BorderStyle     =   0  'None
      Height          =   3135
      Left            =   480
      Picture         =   "frmHome.frx":22F40
      ScaleHeight     =   3135
      ScaleWidth      =   9135
      TabIndex        =   0
      Top             =   120
      Width           =   9135
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   5
      X1              =   3120
      X2              =   3120
      Y1              =   3240
      Y2              =   5040
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   5
      X1              =   1200
      X2              =   1200
      Y1              =   3240
      Y2              =   5040
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   5
      X1              =   7200
      X2              =   7200
      Y1              =   3240
      Y2              =   5040
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   5
      X1              =   9360
      X2              =   9360
      Y1              =   3240
      Y2              =   5040
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   5
      X1              =   4800
      X2              =   4800
      Y1              =   3240
      Y2              =   5040
   End
End
Attribute VB_Name = "frmHome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'ProjectKG
'frmHome
'Jon Jerabek
'10-25-05 & 10-26-05
'Objective-Allows user to move to all forms in project
'Objective of Project- This project's sole purpose is to give information about
                    'Kevin Garnett.It highlights his accomplishments and also
                    'allows the user to interact with different aspects of his career.
                    
                    
Option Explicit
Dim I As Double
Dim Career(1 To 1) As String
Dim G(1 To 9) As Double, GS(1 To 9) As Double, MPG(1 To 9) As Double, FG(1 To 9) As Double, Three(1 To 9) As Double, FT(1 To 9) As Double, OFFReb(1 To 9) As Double, DEFReb(1 To 9) As Double, RPG(1 To 9) As Double, APG(1 To 9) As Double, SPG(1 To 9) As Double, BPG(1 To 9) As Double, TurnO(1 To 9) As Double, PF(1 To 9) As Double, PPG(1 To 9) As Double

Private Sub cmdBio_Click()
frmBio.Show
frmHome.Hide
End Sub

Private Sub cmdCareerHighlights_Click()
frmCareerHighlights.Show
frmHome.Hide
End Sub

Private Sub cmdMVP_Click()
frmMVP.Show
frmHome.Hide
End Sub

Private Sub cmdRetire_Click()
End
End Sub

Private Sub cmdStats_Click()
frmStatistics.Show
frmHome.Hide
End Sub

Private Sub cmdTeam_Click()
frmTeam.Show
frmHome.Hide
End Sub
