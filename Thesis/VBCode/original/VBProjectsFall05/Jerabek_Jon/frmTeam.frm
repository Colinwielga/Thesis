VERSION 5.00
Begin VB.Form frmTeam 
   BackColor       =   &H00800000&
   Caption         =   "Current Teammates"
   ClientHeight    =   7125
   ClientLeft      =   2925
   ClientTop       =   1860
   ClientWidth     =   9885
   LinkTopic       =   "Form1"
   Picture         =   "frmTeam.frx":0000
   ScaleHeight     =   7125
   ScaleWidth      =   9885
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   1335
      Left            =   6120
      MultiLine       =   -1  'True
      TabIndex        =   34
      Text            =   "frmTeam.frx":1070C6
      Top             =   4800
      Width           =   975
   End
   Begin VB.PictureBox Picture1 
      Height          =   1335
      Left            =   4680
      Picture         =   "frmTeam.frx":1070DF
      ScaleHeight     =   1275
      ScaleWidth      =   915
      TabIndex        =   32
      Top             =   4800
      Width           =   975
   End
   Begin VB.PictureBox Picture15 
      Height          =   1335
      Left            =   3240
      Picture         =   "frmTeam.frx":109F95
      ScaleHeight     =   1275
      ScaleWidth      =   915
      TabIndex        =   15
      Top             =   4800
      Width           =   975
   End
   Begin VB.PictureBox Picture14 
      Height          =   1335
      Left            =   1800
      Picture         =   "frmTeam.frx":10CE4B
      ScaleHeight     =   1275
      ScaleWidth      =   915
      TabIndex        =   14
      Top             =   4800
      Width           =   975
   End
   Begin VB.PictureBox Picture13 
      Height          =   1335
      Left            =   360
      Picture         =   "frmTeam.frx":10FD01
      ScaleHeight     =   1275
      ScaleWidth      =   915
      TabIndex        =   13
      Top             =   4800
      Width           =   975
   End
   Begin VB.PictureBox Picture12 
      Height          =   1335
      Left            =   7560
      Picture         =   "frmTeam.frx":112BB7
      ScaleHeight     =   1275
      ScaleWidth      =   915
      TabIndex        =   12
      Top             =   2520
      Width           =   975
   End
   Begin VB.PictureBox Picture11 
      Height          =   1335
      Left            =   6120
      Picture         =   "frmTeam.frx":115A6D
      ScaleHeight     =   1275
      ScaleWidth      =   915
      TabIndex        =   11
      Top             =   2520
      Width           =   975
   End
   Begin VB.PictureBox Picture10 
      Height          =   1335
      Left            =   4680
      Picture         =   "frmTeam.frx":118923
      ScaleHeight     =   1275
      ScaleWidth      =   915
      TabIndex        =   10
      Top             =   2520
      Width           =   975
   End
   Begin VB.PictureBox Picture9 
      Height          =   1335
      Left            =   3240
      Picture         =   "frmTeam.frx":11B7D9
      ScaleHeight     =   1275
      ScaleWidth      =   915
      TabIndex        =   9
      Top             =   2520
      Width           =   975
   End
   Begin VB.PictureBox Picture8 
      Height          =   1335
      Left            =   1800
      Picture         =   "frmTeam.frx":11E68F
      ScaleHeight     =   1275
      ScaleWidth      =   915
      TabIndex        =   8
      Top             =   2520
      Width           =   975
   End
   Begin VB.PictureBox Picture7 
      Height          =   1335
      Left            =   360
      Picture         =   "frmTeam.frx":121545
      ScaleHeight     =   1275
      ScaleWidth      =   915
      TabIndex        =   7
      Top             =   2520
      Width           =   975
   End
   Begin VB.PictureBox pic6 
      Height          =   1335
      Left            =   7560
      Picture         =   "frmTeam.frx":1243FB
      ScaleHeight     =   1275
      ScaleWidth      =   915
      TabIndex        =   6
      Top             =   240
      Width           =   975
   End
   Begin VB.PictureBox pic5 
      Height          =   1335
      Left            =   6120
      Picture         =   "frmTeam.frx":1272B1
      ScaleHeight     =   1275
      ScaleWidth      =   915
      TabIndex        =   5
      Top             =   240
      Width           =   975
   End
   Begin VB.PictureBox pic4 
      Height          =   1335
      Left            =   4680
      Picture         =   "frmTeam.frx":12A167
      ScaleHeight     =   1275
      ScaleWidth      =   915
      TabIndex        =   4
      Top             =   240
      Width           =   975
   End
   Begin VB.PictureBox pic3 
      Height          =   1335
      Left            =   3240
      Picture         =   "frmTeam.frx":12D01D
      ScaleHeight     =   1275
      ScaleWidth      =   915
      TabIndex        =   3
      Top             =   240
      Width           =   975
   End
   Begin VB.PictureBox pic2 
      Height          =   1335
      Left            =   1800
      Picture         =   "frmTeam.frx":12FED3
      ScaleHeight     =   1275
      ScaleWidth      =   915
      TabIndex        =   2
      Top             =   240
      Width           =   975
   End
   Begin VB.PictureBox pic1 
      Height          =   1335
      Left            =   360
      Picture         =   "frmTeam.frx":132D89
      ScaleHeight     =   1275
      ScaleWidth      =   915
      TabIndex        =   1
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton cmdMain5 
      Caption         =   "Main"
      BeginProperty Font 
         Name            =   "Rockwell Extra Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9000
      TabIndex        =   0
      Top             =   6240
      Width           =   735
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "John Lucas III"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   6120
      TabIndex        =   33
      Top             =   6240
      Width           =   975
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "Wally Szczerbiak"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   4680
      TabIndex        =   31
      Top             =   6240
      Width           =   975
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "Troy Hudson"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   3240
      TabIndex        =   30
      Top             =   6240
      Width           =   975
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "Trenton Hassell"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   1800
      TabIndex        =   29
      Top             =   6240
      Width           =   975
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "Ryan Humphrey"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   360
      TabIndex        =   28
      Top             =   6240
      Width           =   975
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "Richie Frahm"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   7560
      TabIndex        =   27
      Top             =   3960
      Width           =   975
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "Rashad McCants"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   6120
      TabIndex        =   26
      Top             =   3960
      Width           =   975
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "Nikoloz Tskitishvili "
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   4680
      TabIndex        =   25
      Top             =   3960
      Width           =   975
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "Ndudi  Ebi"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   3240
      TabIndex        =   24
      Top             =   3960
      Width           =   975
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "Michael Olowokandi"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   1680
      TabIndex        =   23
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "Marko Jaric"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   360
      TabIndex        =   22
      Top             =   3960
      Width           =   975
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "Mark Madsen"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   7560
      TabIndex        =   21
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "Lionel Chalmers"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   6120
      TabIndex        =   20
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "Eddie Griffin"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   4680
      TabIndex        =   19
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "Dwayne Jones"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   3240
      TabIndex        =   18
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "Bracey Wright"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   1800
      TabIndex        =   17
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "Anthony Carter"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   360
      TabIndex        =   16
      Top             =   1680
      Width           =   975
   End
End
Attribute VB_Name = "frmTeam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'ProjectKG
'frmTeam
'Jon Jerabek
'10-25-05
'Objective-Displays current teammates of Kevin Garnett

Private Sub cmdMain5_Click()    'Return to Main Menu
frmHome.Show
frmTeam.Hide
End Sub
