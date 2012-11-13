VERSION 5.00
Begin VB.Form frmSchedule 
   BackColor       =   &H00000080&
   Caption         =   "Schedule"
   ClientHeight    =   8025
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10695
   LinkTopic       =   "Form1"
   ScaleHeight     =   8025
   ScaleWidth      =   10695
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next"
      Height          =   735
      Left            =   2640
      TabIndex        =   17
      Top             =   6840
      Width           =   1935
   End
   Begin VB.TextBox txt15 
      BackColor       =   &H00000080&
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Left            =   720
      TabIndex        =   16
      Text            =   "Fri, Dec 01   Minnesota State *  -  -  Mariucci Arena    7:07 p.m.  "
      Top             =   6240
      Width           =   9615
   End
   Begin VB.TextBox txt14 
      BackColor       =   &H00000080&
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Left            =   720
      TabIndex        =   15
      Text            =   "Sun, Nov 19   Wisconsin *  -  -  Mariucci Arena    5:07 p.m.  "
      Top             =   5880
      Width           =   9615
   End
   Begin VB.TextBox txt13 
      BackColor       =   &H00000080&
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Left            =   720
      TabIndex        =   14
      Text            =   "Sat, Nov 18   Wisconsin *  -  -  Mariucci Arena    6:07 p.m.  "
      Top             =   5520
      Width           =   9615
   End
   Begin VB.TextBox txt12 
      BackColor       =   &H00000080&
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Left            =   720
      TabIndex        =   13
      Text            =   "Sat, Nov 11   St. Cloud State *  -  -  at St. Cloud, Minn.    7:07 p.m.  "
      Top             =   5160
      Width           =   9615
   End
   Begin VB.TextBox txt11 
      BackColor       =   &H00000080&
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Left            =   720
      TabIndex        =   12
      Text            =   "Fri, Nov 10   St. Cloud State *  -  -  Mariucci Arena    7:07 p.m.  "
      Top             =   4800
      Width           =   9615
   End
   Begin VB.TextBox txt10 
      BackColor       =   &H00000080&
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Left            =   720
      TabIndex        =   11
      Text            =   "Sat, Nov 04   Minnesota Duluth *  -  -  at Duluth, Minn.    7:07 p.m.  "
      Top             =   4440
      Width           =   9615
   End
   Begin VB.TextBox txt9 
      BackColor       =   &H00000080&
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Left            =   720
      TabIndex        =   10
      Text            =   "Fri, Nov 03   Minnesota Duluth *  -  -  at Duluth, Minn.    7:07 p.m.  "
      Top             =   4080
      Width           =   9615
   End
   Begin VB.TextBox txt8 
      BackColor       =   &H00000080&
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Left            =   720
      TabIndex        =   9
      Text            =   "Sat, Oct 28   Colorado College * *  3  -  Mariucci Arena    7:07 p.m.  "
      Top             =   3720
      Width           =   9615
   End
   Begin VB.TextBox txt7 
      BackColor       =   &H00000080&
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Left            =   720
      TabIndex        =   8
      Text            =   " Fri, Oct 27   Colorado College * *  3  -  Mariucci Arena    7:07 p.m.  "
      Top             =   3360
      Width           =   9615
   End
   Begin VB.TextBox txt6 
      BackColor       =   &H00000080&
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Left            =   720
      TabIndex        =   7
      Text            =   "Sat, Oct 21   Ohio State  7  -  at Columbus, Ohio    7:05 p.m.  "
      Top             =   3000
      Width           =   9615
   End
   Begin VB.TextBox txt5 
      BackColor       =   &H00000080&
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Left            =   720
      TabIndex        =   6
      Text            =   "Fri, Oct 20   Ohio State  7  -  at Columbus, Ohio    6:05 p.m.  "
      Top             =   2640
      Width           =   9615
   End
   Begin VB.TextBox txt4 
      BackColor       =   &H00000080&
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Index           =   0
      Left            =   720
      TabIndex        =   5
      Text            =   "Sat, Oct 14   Wayne State  8  -  Mariucci Arena    7:07 p.m. "
      Top             =   2280
      Width           =   9615
   End
   Begin VB.TextBox txt3 
      BackColor       =   &H00000080&
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Left            =   720
      TabIndex        =   4
      Text            =   " Fri, Oct 13   Wayne State  8  -  Mariucci Arena    7:07 p.m.  "
      Top             =   1920
      Width           =   9615
   End
   Begin VB.TextBox txt2 
      BackColor       =   &H00000080&
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Left            =   720
      TabIndex        =   3
      Text            =   "Sun, Oct 08   Lethbridge  3  -  Mariucci Arena    7:07 p.m. "
      Top             =   1560
      Width           =   9615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00000080&
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Left            =   720
      TabIndex        =   2
      Text            =   "Fri, Oct 06   Maine  3  11  at St. Paul, Minn.    7:07 p.m.  "
      Top             =   1200
      Width           =   9615
   End
   Begin VB.CommandButton cmdHome 
      Caption         =   "Home"
      Height          =   735
      Left            =   720
      TabIndex        =   1
      Top             =   6840
      Width           =   1815
   End
   Begin VB.Label lblSchedule 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FFFF&
      Caption         =   "2006-2007 Schedule"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   3240
      TabIndex        =   0
      Top             =   360
      Width           =   4425
   End
End
Attribute VB_Name = "frmSchedule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Gopher Hockey
'frmSchedule
'Cole and John
'10/30/06
'Objective: The objective of this form is to present the user with the schedule
'for the 2006-2007 hockey season.

Option Explicit

Private Sub cmdHome_Click()
    frmSchedule.Visible = False     'allows user to go to main page
    frmMain.Visible = True
End Sub

Private Sub cmdNext_Click()
    frmSchedule2.Visible = True     'allows user to go to next page
    frmSchedule.Visible = False
End Sub
