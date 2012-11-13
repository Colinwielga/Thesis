VERSION 5.00
Begin VB.Form frmColors1 
   Caption         =   "Color"
   ClientHeight    =   9780
   ClientLeft      =   3030
   ClientTop       =   2610
   ClientWidth     =   12375
   ForeColor       =   &H8000000B&
   LinkTopic       =   "Form1"
   ScaleHeight     =   9780
   ScaleWidth      =   12375
   Begin VB.CommandButton cmdquit 
      Caption         =   "QUIT"
      Height          =   975
      Left            =   9720
      TabIndex        =   13
      Top             =   7560
      Width           =   1095
   End
   Begin VB.CommandButton cmdstop 
      Caption         =   "OK, Thats my Body Color!!    Now on to the Accent Color!! "
      Height          =   1335
      Left            =   9120
      TabIndex        =   12
      Top             =   6000
      Width           =   2055
   End
   Begin VB.PictureBox piccolors 
      Height          =   2415
      Left            =   1200
      ScaleHeight     =   2355
      ScaleWidth      =   7515
      TabIndex        =   11
      Top             =   6120
      Width           =   7575
   End
   Begin VB.CommandButton cmdsilver 
      BackColor       =   &H00C0C0C0&
      Caption         =   "SILVER"
      Height          =   1095
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4680
      UseMaskColor    =   -1  'True
      Width           =   2415
   End
   Begin VB.CommandButton cmdblack 
      BackColor       =   &H00000000&
      Caption         =   "Command1"
      Height          =   1095
      Left            =   6480
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3000
      UseMaskColor    =   -1  'True
      Width           =   2415
   End
   Begin VB.CommandButton cmdpurple 
      BackColor       =   &H00800080&
      Caption         =   "PURLE"
      Height          =   1095
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3000
      UseMaskColor    =   -1  'True
      Width           =   2415
   End
   Begin VB.CommandButton cmdwhite 
      BackColor       =   &H00FFFFFF&
      Caption         =   "WHITE"
      Height          =   1095
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1320
      UseMaskColor    =   -1  'True
      Width           =   2415
   End
   Begin VB.CommandButton cmdgreen 
      BackColor       =   &H0000C000&
      Caption         =   "GREEN"
      Height          =   1095
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4680
      UseMaskColor    =   -1  'True
      Width           =   2415
   End
   Begin VB.CommandButton cmdyellow 
      BackColor       =   &H0000FFFF&
      Caption         =   "YELLOW"
      Height          =   1095
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1320
      UseMaskColor    =   -1  'True
      Width           =   2415
   End
   Begin VB.CommandButton cmdblue 
      BackColor       =   &H00FF0000&
      Caption         =   "BLUE"
      Height          =   1095
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3000
      UseMaskColor    =   -1  'True
      Width           =   2415
   End
   Begin VB.CommandButton cmdorange 
      BackColor       =   &H000080FF&
      Caption         =   "ORANGE"
      Height          =   1095
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1320
      UseMaskColor    =   -1  'True
      Width           =   2415
   End
   Begin VB.CommandButton cmdred 
      BackColor       =   &H000000FF&
      Caption         =   "RED"
      Height          =   1095
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4680
      UseMaskColor    =   -1  'True
      Width           =   2415
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
      Left            =   2640
      TabIndex        =   14
      Top             =   720
      Width           =   6375
   End
   Begin VB.Label lblblack 
      Caption         =   "<--- BLACK"
      Height          =   255
      Left            =   9000
      TabIndex        =   10
      Top             =   3360
      Width           =   855
   End
   Begin VB.Label lblcolors 
      Caption         =   "Please Select the Unifrom Body Color You Would Like!!!"
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
      Left            =   240
      TabIndex        =   9
      Top             =   120
      Width           =   12135
   End
End
Attribute VB_Name = "frmColors1"
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
    UniColor = "Blue"
    piccolors.Print UniColor
    'when you click on the blue button "BLUE" appears in the output box
End Sub
Private Sub cmdgreen_Click()
    UniColor = "Green"
    piccolors.Print UniColor
    'when you click on the green button "GREEN" appears in the output box
End Sub
Private Sub cmdorange_Click()
    UniColor = "Orange"
    piccolors.Print UniColor
    'when you click on the orange button "ORANGE" appears in the output box
End Sub

Private Sub cmdpurple_Click()
    UniColor = "Purple"
    piccolors.Print UniColor
    'when you click on the purple button "PURPLE" appears in the output box
End Sub

Private Sub cmdquit_Click()
End 'ends the form
End Sub

Private Sub cmdred_Click()
    UniColor = "Red"
    piccolors.Print UniColor
    'when you click on the red button "RED" appears in the output box
End Sub

Private Sub cmdsilver_Click()
    UniColor = "Silver"
    piccolors.Print UniColor
        'when you click on the silver button "SILVER" appears in the output box
End Sub

Private Sub cmdstop_Click()
frmAccent1.Show 'goes to the order form
frmColors1.Hide 'hides the colors form
End Sub

Private Sub cmdwhite_Click()
    UniColor = "White"
    piccolors.Print UniColor
        'when you click on the white button "WHITE" appears in the output box
End Sub

Private Sub cmdyellow_Click()
    UniColor = "Yellow"
    piccolors.Print UniColor
    'when you click on the yellow button "YELLOW" appears in the output box
End Sub

Private Sub cmdblack_Click()
    UniColor = "Black"
    piccolors.Print UniColor
    'when you click on the black button "BLACK" appears in the output box
End Sub

