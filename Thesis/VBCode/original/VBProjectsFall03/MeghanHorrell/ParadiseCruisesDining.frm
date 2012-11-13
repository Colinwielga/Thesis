VERSION 5.00
Begin VB.Form frmDining 
   BackColor       =   &H000040C0&
   Caption         =   "Form1"
   ClientHeight    =   12555
   ClientLeft      =   2175
   ClientTop       =   2040
   ClientWidth     =   14865
   LinkTopic       =   "Form1"
   ScaleHeight     =   12555
   ScaleWidth      =   14865
   Begin VB.CommandButton cmdHome 
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
      Height          =   1335
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   10560
      Width           =   1935
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
      Height          =   1335
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   10560
      Width           =   1935
   End
   Begin VB.PictureBox Picture3 
      Height          =   1815
      Left            =   10560
      Picture         =   "ParadiseCruisesDining.frx":0000
      ScaleHeight     =   1755
      ScaleWidth      =   3315
      TabIndex        =   5
      Top             =   9000
      Width           =   3375
   End
   Begin VB.PictureBox pbxResultsCrystal 
      Height          =   1935
      Left            =   10920
      Picture         =   "ParadiseCruisesDining.frx":5C4A
      ScaleHeight     =   1875
      ScaleWidth      =   2235
      TabIndex        =   4
      Top             =   3120
      Width           =   2295
   End
   Begin VB.PictureBox Picture1 
      Height          =   1815
      Left            =   10560
      Picture         =   "ParadiseCruisesDining.frx":6D8F
      ScaleHeight     =   1755
      ScaleWidth      =   3315
      TabIndex        =   3
      Top             =   6120
      Width           =   3375
   End
   Begin VB.Label lblMyName 
      BackColor       =   &H000040C0&
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
      Height          =   615
      Left            =   12960
      TabIndex        =   10
      Top             =   11760
      Width           =   1695
   End
   Begin VB.Label lblDiningOptions 
      BackColor       =   &H000040C0&
      Caption         =   "Dining Options"
      BeginProperty Font 
         Name            =   "Broadway"
         Size            =   72
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   1200
      TabIndex        =   7
      Top             =   480
      Width           =   12135
   End
   Begin VB.Label lblAboutDining 
      Alignment       =   2  'Center
      BackColor       =   &H000040C0&
      Caption         =   $"ParadiseCruisesDining.frx":8990
      BeginProperty Font 
         Name            =   "Broadway"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5655
      Left            =   600
      TabIndex        =   6
      Top             =   3840
      Width           =   3615
   End
   Begin VB.Label lblJade 
      Alignment       =   2  'Center
      BackColor       =   &H000040C0&
      Caption         =   "In Jade Garden you'll enjoy the innovative Asian cuisine prepared with a contemporary flair."
      BeginProperty Font 
         Name            =   "Broadway"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5520
      TabIndex        =   2
      Top             =   9480
      Width           =   4695
   End
   Begin VB.Label lblKyoto 
      Alignment       =   2  'Center
      BackColor       =   &H000040C0&
      Caption         =   $"ParadiseCruisesDining.frx":8C6F
      BeginProperty Font 
         Name            =   "Broadway"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   5760
      TabIndex        =   1
      Top             =   3600
      Width           =   4335
   End
   Begin VB.Label lblCrystal 
      Alignment       =   2  'Center
      BackColor       =   &H000040C0&
      Caption         =   $"ParadiseCruisesDining.frx":8CFB
      BeginProperty Font 
         Name            =   "Broadway"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   5880
      TabIndex        =   0
      Top             =   5640
      Width           =   4215
   End
End
Attribute VB_Name = "frmDining"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : ProjParadiseCruises (Meghan Horrell's VB Project.vbp)
'Form Name : frmDining (ParadiseCruisesDining.frm)
'Author: Meghan Horrell
'Date Written For: October 29, 2003
'Purpose of Form: To Display pictures of the various restaurants offered and to give information
                'about each restaurant
'Option Explicit is a command to force
'the user to explicitly declare all variables
'before they can be used.
Option Explicit
Private Sub cmdHome_Click()
    'Hides the Dining form and shows the Home form
    frmDining.Hide
    frmHome.Show
End Sub
Private Sub cmdQuit_Click()
    'Ends the program
    End
End Sub
