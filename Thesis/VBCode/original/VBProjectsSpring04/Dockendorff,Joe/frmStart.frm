VERSION 5.00
Begin VB.Form frmStart 
   BackColor       =   &H0000FFFF&
   Caption         =   "Start Page"
   ClientHeight    =   6615
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8115
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   24
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   6615
   ScaleWidth      =   8115
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
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
      Height          =   555
      Left            =   6480
      TabIndex        =   10
      Top             =   6000
      Width           =   1095
   End
   Begin VB.PictureBox picBestBuy 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   5040
      Picture         =   "frmStart.frx":0000
      ScaleHeight     =   1395
      ScaleWidth      =   2115
      TabIndex        =   9
      Top             =   960
      Width           =   2175
   End
   Begin VB.CommandButton cmdMP3 
      Caption         =   "MP3 Players"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6240
      TabIndex        =   7
      Top             =   2760
      Width           =   1575
   End
   Begin VB.CommandButton cmdTelevisions 
      Caption         =   "Televisions"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3480
      TabIndex        =   6
      Top             =   2760
      Width           =   1455
   End
   Begin VB.CommandButton cmdComputers 
      Caption         =   "Computers"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   600
      TabIndex        =   5
      Top             =   2760
      Width           =   1455
   End
   Begin VB.PictureBox picMP3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   6240
      Picture         =   "frmStart.frx":0844
      ScaleHeight     =   1755
      ScaleWidth      =   1515
      TabIndex        =   3
      Top             =   3960
      Width           =   1575
   End
   Begin VB.PictureBox picTV 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   3000
      Picture         =   "frmStart.frx":1349
      ScaleHeight     =   2355
      ScaleWidth      =   2595
      TabIndex        =   2
      Top             =   3600
      Width           =   2655
   End
   Begin VB.PictureBox picComp 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   240
      Picture         =   "frmStart.frx":2BB2
      ScaleHeight     =   2475
      ScaleWidth      =   2115
      TabIndex        =   1
      Top             =   3600
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000FFFF&
      Caption         =   "With prices lower than:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      TabIndex        =   8
      Top             =   960
      Width           =   3735
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000FFFF&
      Caption         =   "Please begin shopping:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   2040
      Width           =   3975
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FFFF&
      Caption         =   "Welcome to Electronics Plus+"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   6495
   End
End
Attribute VB_Name = "frmStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : ProjElectrPlus (Joe Dockendorff's VB Project.vbp)
'Form Name : frmStart (frmStart.frm)
'Author: Joe Dockendorff
'Date Written: March 13, 2004
'Purpose of Form: This form is the beginning of the project, the user
                 'is clicks the only button available and is taken
                 'to the next form.
                 
'Option Explicit is a command to force
'the user to declare all variables
'before they can be used.
Option Explicit

Private Sub cmdComputers_Click()
'This button disables the Computers button, enables the Televisions button and also
'hides the start form and brings up the ProdComp form so the user can look at the
'different computers available.

frmStart.Hide
frmProdComp.Show
cmdComputers.Enabled = False
cmdTelevisions.Enabled = True
End Sub

Private Sub cmdMP3_Click()
'This button disables cmdMP3 and enables cmdComputers so that the user has the ability
'to restart the program. This button also jumps to the prodmp3 form so the users can
'check out the mp3 players. It hides the start form too.

frmStart.Hide
frmProdMP3.Show
cmdMP3.Enabled = False
cmdComputers.Enabled = True
End Sub

Private Sub cmdQuit_Click()
'This button is placed so that the user can end the program.
End
End Sub

Private Sub cmdTelevisions_Click()
'This button enables cmdMP3 and disables cmdTelevisions. It also hides the start form
'and brings up the ProdTvs form.
frmStart.Hide
frmProdTVs.Show
cmdTelevisions.Enabled = False
cmdMP3.Enabled = True
End Sub

Private Sub Form_Load()
'Enables only the cmdComputers so the user has to pick this first.
cmdTelevisions.Enabled = False
cmdMP3.Enabled = False
cmdComputers.Enabled = True
End Sub
