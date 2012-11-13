VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Main Menu"
   ClientHeight    =   6960
   ClientLeft      =   2280
   ClientTop       =   1755
   ClientWidth     =   10905
   LinkTopic       =   "Form1"
   Picture         =   "ChrisDonnelly.frx":0000
   ScaleHeight     =   6960
   ScaleWidth      =   10905
   Begin VB.CommandButton cmdWinter 
      BackColor       =   &H80000013&
      Caption         =   "Learn about the winter elements"
      BeginProperty Font 
         Name            =   "Gill Sans Ultra Bold Condensed"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   120
      MaskColor       =   &H80000000&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   240
      Width           =   2175
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H80000013&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Gill Sans Ultra Bold Condensed"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4680
      Width           =   2175
   End
   Begin VB.CommandButton cmdRide 
      BackColor       =   &H80000013&
      Caption         =   "Select Your Ride!"
      BeginProperty Font 
         Name            =   "Gill Sans Ultra Bold Condensed"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2520
      Width           =   2175
   End
   Begin VB.Label lblChriss 
      Alignment       =   2  'Center
      BackColor       =   &H80000013&
      Caption         =   "Chris Donnelly 10/31/2005"
      BeginProperty Font 
         Name            =   "Gill Sans MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   735
      Left            =   8040
      TabIndex        =   5
      Top             =   6120
      Width           =   2535
   End
   Begin VB.Label lblChris 
      Alignment       =   2  'Center
      BackColor       =   &H80000013&
      Caption         =   "Snowmobile Program Created by: "
      BeginProperty Font 
         Name            =   "Gill Sans MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   855
      Left            =   8040
      TabIndex        =   4
      Top             =   5280
      Width           =   2535
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000013&
      Caption         =   $"ChrisDonnelly.frx":5661E
      BeginProperty Font 
         Name            =   "Gill Sans MT"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   3
      Top             =   5760
      Width           =   7215
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'ends program
Private Sub cmdQuit_Click()
End
End Sub
'jumps to Ride form
Private Sub cmdRide_Click()
frmMain.Hide
frmRide.Show
End Sub
'jumps to Winter form
Private Sub cmdWinter_Click()
frmMain.Hide
frmWinter.Show
End Sub
