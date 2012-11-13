VERSION 5.00
Begin VB.Form Greenwich 
   BackColor       =   &H00C0C000&
   Caption         =   "Greenwich"
   ClientHeight    =   14850
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   19080
   LinkTopic       =   "Form1"
   ScaleHeight     =   14850
   ScaleWidth      =   19080
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdquit 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   16560
      TabIndex        =   8
      Top             =   11880
      Width           =   1695
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return to Map of London"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   13080
      TabIndex        =   7
      Top             =   11760
      Width           =   3015
   End
   Begin VB.PictureBox Picture2 
      Height          =   3015
      Left            =   120
      Picture         =   "Greenwich.frx":0000
      ScaleHeight     =   2955
      ScaleWidth      =   18795
      TabIndex        =   4
      Top             =   8160
      Width           =   18855
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H0080FFFF&
      Height          =   4575
      Left            =   6120
      Picture         =   "Greenwich.frx":54F1
      ScaleHeight     =   4515
      ScaleWidth      =   4635
      TabIndex        =   1
      Top             =   1440
      Width           =   4695
   End
   Begin VB.Label Label6 
      Caption         =   "Created by Chelsey Johnson"
      Height          =   255
      Left            =   4560
      TabIndex        =   9
      Top             =   12240
      Width           =   3375
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFC0C0&
      Caption         =   $"Greenwich.frx":AE35
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7440
      TabIndex        =   6
      Top             =   7440
      Width           =   10815
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Maritime Museum"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4800
      TabIndex        =   5
      Top             =   7560
      Width           =   2055
   End
   Begin VB.Label Label3 
      BackColor       =   &H0080FFFF&
      Caption         =   $"Greenwich.frx":AF28
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   11040
      TabIndex        =   3
      Top             =   2640
      Width           =   2295
   End
   Begin VB.Label Label2 
      BackColor       =   &H0080FFFF&
      Caption         =   "Entrance to the Foot Tunnel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11040
      TabIndex        =   2
      Top             =   1560
      Width           =   3015
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF80FF&
      Caption         =   "The Foot Tunnel and teh maitime Museum are two very important sites in the district of Greenwich."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4440
      TabIndex        =   0
      Top             =   240
      Width           =   11895
   End
End
Attribute VB_Name = "Greenwich"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Discovering London (Project1.vbp)
'Form Name: Greenwich (Greenwich.frm)
'Author: Chelsey Johnson
'Date Written: March 14, 2004
'Purpose of Form: This form lets the user read the history of both the foot tunnel and the Maritime Museum.
'Option Explicit is a command to force
'the user to explicitly declare all variables
'before they can be used.
Option Explicit
Private Sub cmdquit_Click()
End
End Sub

Private Sub cmdreturn_Click()
'Returns the user back to the Map of London page, so they are able to look at a different district and its' sites
Greenwich.Hide
MapLondon.Show
End Sub
