VERSION 5.00
Begin VB.Form wwww
   BackColor       =   &H000000C0&
   Caption         =   "St Johns Fire Department Start Page"
   ClientHeight    =   6630
   ClientLeft      =   1830
   ClientTop       =   2205
   ClientWidth     =   10065
   FillColor       =   &H000000FF&
   FillStyle       =   0  'Solid
   ForeColor       =   &H8000000D&
   LinkTopic       =   "Form1"
   ScaleHeight     =   6630
   ScaleWidth      =   10065
   Begin VB.Frame Frame1
      BackColor       =   &H00000000&
      Caption         =   "Change Background Color"
      ForeColor       =   &H8000000E&
      Height          =   1095
      Left            =   7560
      TabIndex        =   7
      Top             =   5400
      Width           =   2415
      Begin VB.OptionButton optGreen
         BackColor       =   &H80000012&
         Caption         =   "Green"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   720
         Width           =   1095
      End
      Begin VB.OptionButton optBlue
         BackColor       =   &H80000012&
         Caption         =   "Blue"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   1095
      End
      Begin VB.OptionButton optRed
         BackColor       =   &H80000007&
         Caption         =   "Red"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdQuit
      Caption         =   "Quit"
      BeginProperty Font
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3360
      TabIndex        =   6
      Top             =   5640
      Width           =   3015
   End
   Begin VB.CommandButton cmdHistory
      Caption         =   "History of St. Johns Fire Department "
      Height          =   735
      Left            =   7680
      TabIndex        =   5
      Top             =   360
      Width           =   2175
   End
   Begin VB.CommandButton cmdTrucks
      Caption         =   "SJU FD Vehicles "
      Height          =   735
      Left            =   2640
      TabIndex        =   4
      Top             =   360
      Width           =   2175
   End
   Begin VB.CommandButton cmdRoster
      Caption         =   "Fire Department Members"
      Height          =   735
      Left            =   5160
      TabIndex        =   3
      Top             =   360
      Width           =   2175
   End
   Begin VB.CommandButton cmdPPE
      Caption         =   "Protective Equipment Quiz"
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   2175
   End
   Begin VB.Image Image1
      BorderStyle     =   1  'Fixed Single
      Height          =   2055
      Left            =   3720
      Picture         =   "sju_firedepartment.frx":0000
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   2415
   End
   Begin VB.Label Label2
      Alignment       =   2  'Center
      Caption         =   "St Johns University, Collegeville MN 56321"
      BeginProperty Font
         Name            =   "Lucida Sans Typewriter"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2400
      TabIndex        =   1
      Top             =   2280
      Width           =   5295
   End
   Begin VB.Label Label1
      Alignment       =   2  'Center
      Caption         =   "St. Johns Fire Department"
      BeginProperty Font
         Name            =   "Bernard MT Condensed"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2400
      TabIndex        =   0
      Top             =   1440
      Width           =   5295
   End
End
Attribute VB_Name = "wwww"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Project Name: Saint John's Fire Department
'Form Name: Main Page
'Authors: JT Trujillo and Matt Mollet
'Date Written: 2/20/2010
'Objective: To provide a main page which the user can choose different
            'activities to look at via clicking on buttons that
            'lead to different forms.
        'The overall purpose of the project is to inform the user
        'of some of the experiences of being on the Saint John's
        'Fire Department, and to provide information to the user
        'about the Saint John's Fire Department.

'This button will lead the user to a form which will display the
'historical story of the Saint John's Fire Department, mainly
'how it came to be, and how it changed over the years.
Private Sub aaaa_Click()
pppp.Visible = True
wwww.Visible = False
End Sub

'This button will lead the user to an activity in which they can learn
'through trial and error which order to correctly don (put on) a
'firefighter's protective gear.
Private Sub dddd_Click()
xxxx.Visible = True
wwww.Visible = False
End Sub

Private Sub cccc_Click()
'Quit Program
End
End Sub

'This button will lead the user to a form where they can view the
'fire department's entire roster, initially in order of rank, then
'the user can either put it into alphabetical order by last name, or
'search by last name to see if somebody is on our roster.
Private Sub bbbb_Click()
frmFirefighters.Visible = True
wwww.Visible = False
End Sub

'The user can click on this button to go to a form which holds an
'activity in which the user can view the fire department's trucks and
'other vehicles, and learn about what each vehicle is used for.
Private Sub eeee_Click()
qqqq.Visible = True
wwww.Visible = False
End Sub
'Changes background color
Private Sub ffff_Click()
wwww.BackColor = &HC00000
End Sub

Private Sub gggg_Click()
wwww.BackColor = &H4000&
End Sub

Private Sub hhhh_Click()
wwww.BackColor = &HC0&
End Sub

