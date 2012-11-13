VERSION 5.00
Begin VB.Form frmAccent1 
   Caption         =   "Accent1"
   ClientHeight    =   9645
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12570
   LinkTopic       =   "Form1"
   ScaleHeight     =   9645
   ScaleWidth      =   12570
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdred 
      BackColor       =   &H000000FF&
      Caption         =   "RED"
      Height          =   1095
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4560
      UseMaskColor    =   -1  'True
      Width           =   2415
   End
   Begin VB.CommandButton cmdorange 
      BackColor       =   &H000080FF&
      Caption         =   "ORANGE"
      Height          =   1095
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1200
      UseMaskColor    =   -1  'True
      Width           =   2415
   End
   Begin VB.CommandButton cmdblue 
      BackColor       =   &H00FF0000&
      Caption         =   "BLUE"
      Height          =   1095
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2880
      UseMaskColor    =   -1  'True
      Width           =   2415
   End
   Begin VB.CommandButton cmdyellow 
      BackColor       =   &H0000FFFF&
      Caption         =   "YELLOW"
      Height          =   1095
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1200
      UseMaskColor    =   -1  'True
      Width           =   2415
   End
   Begin VB.CommandButton cmdgreen 
      BackColor       =   &H0000C000&
      Caption         =   "GREEN"
      Height          =   1095
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4560
      UseMaskColor    =   -1  'True
      Width           =   2415
   End
   Begin VB.CommandButton cmdwhite 
      BackColor       =   &H00FFFFFF&
      Caption         =   "WHITE"
      Height          =   1095
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1200
      UseMaskColor    =   -1  'True
      Width           =   2415
   End
   Begin VB.CommandButton cmdpurple 
      BackColor       =   &H00800080&
      Caption         =   "PURLE"
      Height          =   1095
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2880
      UseMaskColor    =   -1  'True
      Width           =   2415
   End
   Begin VB.CommandButton cmdblack 
      BackColor       =   &H00000000&
      Caption         =   "Command1"
      Height          =   1095
      Left            =   6720
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2880
      UseMaskColor    =   -1  'True
      Width           =   2415
   End
   Begin VB.CommandButton cmdsilver 
      BackColor       =   &H00C0C0C0&
      Caption         =   "SILVER"
      Height          =   1095
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4560
      UseMaskColor    =   -1  'True
      Width           =   2415
   End
   Begin VB.PictureBox piccolors 
      Height          =   2415
      Left            =   1440
      ScaleHeight     =   2355
      ScaleWidth      =   7515
      TabIndex        =   2
      Top             =   6000
      Width           =   7575
   End
   Begin VB.CommandButton cmdstop 
      Caption         =   "OK, Thats my Accent Color!! Now back to the Main Menu!! "
      Height          =   1335
      Left            =   9360
      TabIndex        =   1
      Top             =   5880
      Width           =   2055
   End
   Begin VB.CommandButton cmdquit 
      Caption         =   "QUIT"
      Height          =   975
      Left            =   9960
      TabIndex        =   0
      Top             =   7440
      Width           =   1095
   End
   Begin VB.Label lblcolors 
      Caption         =   "Please Select the Accent color You Would Like!!!"
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   960
      TabIndex        =   14
      Top             =   0
      Width           =   11055
   End
   Begin VB.Label lblblack 
      Caption         =   "<--- BLACK"
      Height          =   255
      Left            =   9240
      TabIndex        =   13
      Top             =   3240
      Width           =   855
   End
   Begin VB.Label lblcolorcost 
      Caption         =   "The Color cost is included in the Colst of the Shells!!!"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   12
         Charset         =   1
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   12
      Top             =   600
      Width           =   6375
   End
End
Attribute VB_Name = "frmAccent1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : Cheerleading (Cheerleading.vbp)
'Form Name : Color (Colors1.frm)
'Author: Heather Johnson
'Date Written: October 28, 2003
'Purpose of Form:'this form will ask you what colors you would like
                 'your uniforms to be
                 
Option Explicit
Private Sub cmdblue_Click()
    AccColor = "Blue"
    piccolors.Print AccColor
    'when you click on the blue button "BLUE" appears in the output box
End Sub
Private Sub cmdgreen_Click()
    AccColor = "Green"
    piccolors.Print AccColor
    'when you click on the green button "GREEN" appears in the output box
End Sub
Private Sub cmdorange_Click()
    AccColor = "Orange"
    piccolors.Print AccColor
    'when you click on the orange button "ORANGE" appears in the output box
End Sub

Private Sub cmdpurple_Click()
    AccColor = "Purple"
    piccolors.Print AccColor
    'when you click on the purple button "PURPLE" appears in the output box
End Sub

Private Sub cmdquit_Click()
End 'ends the form
End Sub

Private Sub cmdred_Click()
    AccColor = "Red"
    piccolors.Print AccColor
    'when you click on the red button "RED" appears in the output box
End Sub

Private Sub cmdsilver_Click()
    AccColor = "Silver"
    piccolors.Print AccColor
        'when you click on the silver button "SILVER" appears in the output box
End Sub

Private Sub cmdstop_Click()
frmOrder1.Show 'goes to the order form
frmAccent1.Hide 'hides the colors form
End Sub

Private Sub cmdwhite_Click()
    AccColor = "White"
    piccolors.Print AccColor
        'when you click on the white button "WHITE" appears in the output box
End Sub

Private Sub cmdyellow_Click()
    AccColor = "Yellow"
    piccolors.Print AccColor
    'when you click on the yellow button "YELLOW" appears in the output box
End Sub

Private Sub cmdblack_Click()
    AccColor = "Black"
    piccolors.Print AccColor
    'when you click on the black button "BLACK" appears in the output box
End Sub


