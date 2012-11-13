VERSION 5.00
Begin VB.Form FrmPlayers 
   BackColor       =   &H8000000D&
   Caption         =   "Form1"
   ClientHeight    =   8280
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10395
   LinkTopic       =   "Form1"
   ScaleHeight     =   8280
   ScaleWidth      =   10395
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000013&
      ForeColor       =   &H80000013&
      Height          =   7575
      Left            =   240
      ScaleHeight     =   7515
      ScaleWidth      =   9675
      TabIndex        =   0
      Top             =   240
      Width           =   9735
      Begin VB.CommandButton Command14 
         BackColor       =   &H80000010&
         Caption         =   "Quit Program"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   6840
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   6480
         Width           =   2055
      End
      Begin VB.CommandButton Command13 
         BackColor       =   &H80000010&
         Caption         =   "Return to Main"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   4680
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   6480
         Width           =   1935
      End
      Begin VB.CommandButton CmdSpan 
         BackColor       =   &H8000000E&
         Caption         =   "Denard Span"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   6240
         Width           =   3255
      End
      Begin VB.CommandButton CmdPunto 
         BackColor       =   &H8000000E&
         Caption         =   "Nick Punto"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   5640
         Width           =   3255
      End
      Begin VB.CommandButton CmdCasilla 
         BackColor       =   &H8000000E&
         Caption         =   "Alexi Casilla"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   5040
         Width           =   3255
      End
      Begin VB.CommandButton CmdGomez 
         BackColor       =   &H8000000E&
         Caption         =   "Carlos Gomez"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   4440
         Width           =   3255
      End
      Begin VB.CommandButton CmdYoung 
         BackColor       =   &H8000000E&
         Caption         =   "Delmon Young"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   3840
         Width           =   3255
      End
      Begin VB.CommandButton CmdCuddyer 
         BackColor       =   &H8000000E&
         Caption         =   "Michael Cuddyer"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   3240
         Width           =   3255
      End
      Begin VB.CommandButton CmdMorneau 
         BackColor       =   &H8000000E&
         Caption         =   "Justin Morneau"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   2640
         Width           =   3255
      End
      Begin VB.CommandButton CmdMauer 
         BackColor       =   &H80000009&
         Caption         =   "Joe Mauer"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   2040
         Width           =   3255
      End
      Begin VB.PictureBox picTeam 
         Height          =   1335
         Left            =   6360
         Picture         =   "FrmPlayers.frx":0000
         ScaleHeight     =   1275
         ScaleWidth      =   3195
         TabIndex        =   7
         Top             =   1080
         Width           =   3255
      End
      Begin VB.PictureBox picTeam3 
         Height          =   2295
         Left            =   5760
         Picture         =   "FrmPlayers.frx":18746
         ScaleHeight     =   2235
         ScaleWidth      =   3795
         TabIndex        =   5
         Top             =   4080
         Width           =   3855
      End
      Begin VB.PictureBox picTeam2 
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         BorderStyle     =   0  'None
         DrawWidth       =   3
         ForeColor       =   &H80000008&
         Height          =   2415
         Left            =   5160
         Picture         =   "FrmPlayers.frx":49E88
         ScaleHeight     =   2415
         ScaleWidth      =   3855
         TabIndex        =   3
         Top             =   2280
         Width           =   3855
         Begin VB.PictureBox Picture5 
            Height          =   495
            Left            =   480
            ScaleHeight     =   435
            ScaleWidth      =   1635
            TabIndex        =   6
            Top             =   2400
            Width           =   1695
         End
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2895
         Left            =   4800
         ScaleHeight     =   2895
         ScaleWidth      =   4455
         TabIndex        =   4
         Top             =   2040
         Width           =   4455
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H000000FF&
         Height          =   5295
         Left            =   360
         ScaleHeight     =   5235
         ScaleWidth      =   3915
         TabIndex        =   16
         Top             =   1680
         Width           =   3975
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H80000013&
         Caption         =   "Click for Player's Bio!"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   495
         Left            =   240
         TabIndex        =   2
         Top             =   1200
         Width           =   4335
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Meet the Players"
         BeginProperty Font 
            Name            =   "Cooper Black"
            Size            =   48
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   1215
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   9735
      End
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Project By: Steph Arel"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   375
      Left            =   7080
      TabIndex        =   19
      Top             =   7920
      Width           =   3255
   End
End
Attribute VB_Name = "FrmPlayers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Title: Minnesota Twins Fan
'Form Name: FrmPlayers
'Project By: Stephanie Arel
'Date Written: 3/12/2009
'Objective: The objective of this form is to allow the user to view information about a handful of the more famous twins players. Each button will take the user to just a brief page about the selected player.

Option Explicit


Private Sub CmdCasilla_Click()
'Allows user to see the profile of Alexi Casilla
FrmPlayers.Hide
FrmCasilla.Show
End Sub

Private Sub CmdCuddyer_Click()
'Allows user to see the profile of Michael Cuddyer
FrmPlayers.Hide
FrmCuddyer.Show
End Sub

Private Sub CmdGomez_Click()
'Allows user to see the profile for Carlos Gomez
FrmPlayers.Hide
FrmGomez.Show
End Sub

Private Sub CmdMauer_Click()
'Allows user to see the profile for Joe Mauer
FrmPlayers.Hide
FrmMauer.Show
End Sub

Private Sub CmdMorneau_Click()
'Allows user to see the profile for Justin Morneau
FrmPlayers.Hide
FrmMorneau.Show
End Sub

Private Sub CmdPunto_Click()
'Allows user to see the profile of Nick Punto
FrmPlayers.Hide
FrmPunto.Show
End Sub

Private Sub CmdSpan_Click()
'Allows user to see the profile of Denard Span
FrmPlayers.Hide
FrmSpan.Show
End Sub

Private Sub CmdYoung_Click()
'Allows user to see the profile for Delmon Young
FrmPlayers.Hide
FrmYoung.Show
End Sub

Private Sub Command13_Click()
'Allows user to return to the main choosing menu

FrmPlayers.Hide
FrmMain.Show
End Sub

Private Sub Command14_Click()
'Ends the Program
End
End Sub

Private Sub Form_Load()
Top = Screen.Height / 2 - Height / 2
Left = Screen.Width / 2 - Width / 2
End Sub
