VERSION 5.00
Begin VB.Form FrmPunto 
   BackColor       =   &H80000007&
   Caption         =   "Form1"
   ClientHeight    =   7545
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9735
   LinkTopic       =   "Form1"
   ScaleHeight     =   7545
   ScaleWidth      =   9735
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      ForeColor       =   &H80000008&
      Height          =   6375
      Left            =   0
      Picture         =   "FrmPunto.frx":0000
      ScaleHeight     =   6345
      ScaleWidth      =   9825
      TabIndex        =   0
      Top             =   1200
      Width           =   9855
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1695
         Index           =   1
         Left            =   7800
         Picture         =   "FrmPunto.frx":7B6DE
         ScaleHeight     =   1695
         ScaleWidth      =   1695
         TabIndex        =   12
         Top             =   4440
         Width           =   1695
      End
      Begin VB.CommandButton CmdReturn 
         BackColor       =   &H8000000E&
         Caption         =   "Return to Players"
         BeginProperty Font 
            Name            =   "Rockwell Extra Bold"
            Size            =   9
            Charset         =   0
            Weight          =   800
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   2880
         Picture         =   "FrmPunto.frx":855A8
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   5280
         Width           =   1335
      End
      Begin VB.CommandButton CmdMain 
         BackColor       =   &H8000000E&
         Caption         =   "Go Back to Main Menu"
         BeginProperty Font 
            Name            =   "Rockwell Extra Bold"
            Size            =   9
            Charset         =   0
            Weight          =   800
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   1320
         Picture         =   "FrmPunto.frx":86C76
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   5280
         Width           =   1335
      End
      Begin VB.CommandButton CmdQuit 
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "Rockwell Extra Bold"
            Size            =   9
            Charset         =   0
            Weight          =   800
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         Picture         =   "FrmPunto.frx":88344
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   5280
         Width           =   735
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackColor       =   &H80000012&
         BackStyle       =   0  'Transparent
         Caption         =   "*Punto is one of four Twins players nicknamed ""The Piranhas"" by White Sox manager, Ozzie Guillén."
         BeginProperty Font 
            Name            =   "Rockwell"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   1095
         Left            =   4800
         TabIndex        =   13
         Top             =   3600
         Width           =   4095
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Birthdate: November 8, 1977 "
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   5640
         TabIndex        =   8
         Top             =   720
         Width           =   3735
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Hometown:  San Diego, California"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   735
         Left            =   5400
         TabIndex        =   7
         Top             =   1200
         Width           =   4215
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H80000012&
         BackStyle       =   0  'Transparent
         Caption         =   "Bats: Right Throws: Right"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   1335
         Left            =   7680
         TabIndex        =   6
         Top             =   1800
         Width           =   1815
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Height: 5'9"" Weight: 190 lbs"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   1215
         Left            =   7560
         TabIndex        =   5
         Top             =   2640
         Width           =   2055
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Full Name: Nicholas Paul Punto "
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   4560
         TabIndex        =   4
         Top             =   240
         Width           =   4695
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Nick Punto"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   1215
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   6735
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ShortStop"
      BeginProperty Font 
         Name            =   "Rockwell Condensed"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   495
      Left            =   5400
      TabIndex        =   2
      Top             =   720
      Width           =   3855
   End
   Begin VB.Label Label8 
      BackColor       =   &H80000012&
      Caption         =   "#8"
      BeginProperty Font 
         Name            =   "Rockwell Extra Bold"
         Size            =   12
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   735
      Left            =   360
      TabIndex        =   3
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "FrmPunto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Title: Minnesota Twins Fan
'Form Name: FrmPunto
'Project By: Stephanie Arel
'Date Written: 3/12/2009
'This Form just lists the facts about Nick Punto.
Option Explicit

Private Sub CmdMain_Click()
'Takes user back to main menu
FrmPunto.Hide
FrmMain.Show
End Sub

Private Sub CmdQuit_Click()
'Ends the Program
End
End Sub

Private Sub CmdReturn_Click()
'Takes user back to the main players menu
FrmPunto.Hide
FrmPlayers.Show
End Sub

Private Sub Picture1_Click()
Top = Screen.Height / 2 - Height / 2
Left = Screen.Width / 2 - Width / 2
End Sub
