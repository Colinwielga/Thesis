VERSION 5.00
Begin VB.Form City 
   BackColor       =   &H000000FF&
   Caption         =   "City"
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
      Height          =   1215
      Left            =   17520
      TabIndex        =   14
      Top             =   12840
      Width           =   1215
   End
   Begin VB.CommandButton cmdreturn 
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
      Height          =   1215
      Left            =   15720
      TabIndex        =   13
      Top             =   12840
      Width           =   1455
   End
   Begin VB.PictureBox Picture2 
      Height          =   5295
      Left            =   12840
      Picture         =   "City.frx":0000
      ScaleHeight     =   5235
      ScaleWidth      =   5955
      TabIndex        =   9
      Top             =   6960
      Width           =   6015
   End
   Begin VB.PictureBox Picture1 
      Height          =   3975
      Left            =   9360
      Picture         =   "City.frx":705D
      ScaleHeight     =   3915
      ScaleWidth      =   2595
      TabIndex        =   8
      Top             =   6240
      Width           =   2655
   End
   Begin VB.PictureBox picstpauls 
      Height          =   5055
      Left            =   360
      Picture         =   "City.frx":1125E
      ScaleHeight     =   4995
      ScaleWidth      =   7995
      TabIndex        =   6
      Top             =   7200
      Width           =   8055
   End
   Begin VB.PictureBox picinside 
      Height          =   2295
      Left            =   14640
      Picture         =   "City.frx":1A4E3
      ScaleHeight     =   2235
      ScaleWidth      =   3915
      TabIndex        =   3
      Top             =   2640
      Width           =   3975
   End
   Begin VB.PictureBox picglobe 
      Height          =   4575
      Left            =   1320
      Picture         =   "City.frx":269E9
      ScaleHeight     =   4515
      ScaleWidth      =   6795
      TabIndex        =   1
      Top             =   1920
      Width           =   6855
   End
   Begin VB.Label Label9 
      Caption         =   "Created by Chelsey Johnson"
      Height          =   375
      Left            =   4680
      TabIndex        =   15
      Top             =   14280
      Width           =   2175
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFF80&
      Caption         =   $"City.frx":3028E
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7080
      TabIndex        =   12
      Top             =   12600
      Width           =   7455
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFF80&
      Caption         =   "Other side of St. Pauls"
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
      Left            =   14880
      TabIndex        =   11
      Top             =   6480
      Width           =   2895
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFF80&
      Caption         =   "Inside St. Pauls"
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
      Left            =   9840
      TabIndex        =   10
      Top             =   10320
      Width           =   1935
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFF00&
      Caption         =   "St. Pauls Cathedral"
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
      Left            =   2760
      TabIndex        =   7
      Top             =   6720
      Width           =   2055
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000FF00&
      Caption         =   $"City.frx":30394
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   8400
      TabIndex        =   5
      Top             =   2760
      Width           =   5895
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000FF00&
      Caption         =   "Inside the Global Theatre"
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
      Left            =   15360
      TabIndex        =   4
      Top             =   2040
      Width           =   2895
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000FF00&
      Caption         =   "Global Theatre"
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
      Left            =   3720
      TabIndex        =   2
      Top             =   1440
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FFFF&
      Caption         =   "The two main sites in the City District in London are St. Paul's Cathedral and The Global Theatre."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4320
      TabIndex        =   0
      Top             =   120
      Width           =   12015
   End
End
Attribute VB_Name = "City"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Discovering London (Project1.vbp)
'Form Name: City (City.frm)
'Author: Chelsey Johnson
'Date Written: March 14, 2004
'Purpose of Form: This form is to familiarize the user with the history behind the Global Theatre and St. Pauls Cathedral
'Option Explicit is a command to force
'the user to explicitly declare all variables
'before they can be used.
Option Explicit
Private Sub cmdquit_Click()
End
End Sub

Private Sub cmdreturn_Click()
'Returns the user back to the map of London page, so they can choose to view a new district and its' sites
City.Hide
MapLondon.Show
End Sub
