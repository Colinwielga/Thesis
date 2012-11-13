VERSION 5.00
Begin VB.Form frmSchedule2 
   BackColor       =   &H00000080&
   Caption         =   "Form1"
   ClientHeight    =   7530
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9885
   LinkTopic       =   "Form1"
   ScaleHeight     =   7530
   ScaleWidth      =   9885
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdback 
      Caption         =   "Back"
      Height          =   735
      Index           =   1
      Left            =   2400
      TabIndex        =   13
      Top             =   6240
      Width           =   1935
   End
   Begin VB.CommandButton cmdHome 
      Caption         =   "Home"
      Height          =   735
      Index           =   0
      Left            =   240
      TabIndex        =   12
      Top             =   6240
      Width           =   1935
   End
   Begin VB.TextBox txtUND2 
      BackColor       =   &H00000080&
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Index           =   10
      Left            =   120
      TabIndex        =   11
      Text            =   "Sat, Jan 27   North Dakota*  - -  Mariucci Arena  7:07 p.m."
      Top             =   4560
      Width           =   9615
   End
   Begin VB.TextBox txtUND 
      BackColor       =   &H00000080&
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Index           =   9
      Left            =   120
      TabIndex        =   10
      Text            =   "Fri, Jan 26   North Dakota*  - -  Mariucci Arena  7:07 p.m."
      Top             =   4200
      Width           =   9615
   End
   Begin VB.TextBox txtDenver1 
      BackColor       =   &H00000080&
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Index           =   8
      Left            =   120
      TabIndex        =   9
      Text            =   "Sat, Jan 20  Denver*  - -  Mariucci Arena  7:07 p.m."
      Top             =   3840
      Width           =   9615
   End
   Begin VB.TextBox txtDenver 
      BackColor       =   &H00000080&
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Index           =   7
      Left            =   120
      TabIndex        =   8
      Text            =   "Fri, Jan 19  Denver*  - -  Mariucci Arena  7:07 p.m."
      Top             =   3480
      Width           =   9615
   End
   Begin VB.TextBox txtWisc2 
      BackColor       =   &H00000080&
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Index           =   6
      Left            =   120
      TabIndex        =   7
      Text            =   "Sat, Jan 13   Wisconsin *  -  -  at Madison, Wis.    7:07 p.m."
      Top             =   3120
      Width           =   9615
   End
   Begin VB.TextBox txtWisc 
      BackColor       =   &H00000080&
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Index           =   5
      Left            =   120
      TabIndex        =   6
      Text            =   "Fri, Jan 12   Wisconsin *  -  -  at Madison, Wis.    7:07 p.m.  "
      Top             =   2760
      Width           =   9615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00000080&
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Index           =   4
      Left            =   120
      TabIndex        =   5
      Text            =   "Sun, Jan 07   Minnesota State*  - -  Mariucci Arena  7:07 p.m."
      Top             =   2400
      Width           =   9615
   End
   Begin VB.TextBox txtMTU 
      BackColor       =   &H00000080&
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Index           =   3
      Left            =   120
      TabIndex        =   4
      Text            =   "Fri, Dec 08   Michigan Tech*  - -  at Houghton, Mich. 6:07 p.m."
      Top             =   1320
      Width           =   9615
   End
   Begin VB.TextBox txtMTU2 
      BackColor       =   &H00000080&
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Index           =   2
      Left            =   120
      TabIndex        =   3
      Text            =   "Sat, Dec 09   Michigan Tech*  - -  at Houghton, Mich. 6:07 p.m."
      Top             =   1680
      Width           =   9615
   End
   Begin VB.TextBox txtMSU2 
      BackColor       =   &H00000080&
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Text            =   "Fri, Jan 05   Minnesota State*  - -  at Mankato, Minn.  7:37 p.m."
      Top             =   2040
      Width           =   9615
   End
   Begin VB.TextBox txtMSU 
      BackColor       =   &H00000080&
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Text            =   "Sat, Dec 02   Minnesota State*  - -  at Mankato, Minn.  7:07 p.m."
      Top             =   960
      Width           =   9615
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
      Left            =   2640
      TabIndex        =   0
      Top             =   120
      Width           =   4425
   End
End
Attribute VB_Name = "frmSchedule2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Gopher Hockey
'frmSchedule2
'Cole and John
'10/30/06
'Objective: The objective of this form is to present the user with the schedule
'for the 2006-2007 hockey season.

Option Explicit


Private Sub cmdBack_Click(Index As Integer)
    frmSchedule.Visible = True      'see frmSchedule
    frmSchedule2.Visible = False
End Sub

Private Sub cmdHome_Click(Index As Integer)
    frmMain.Visible = True
    frmSchedule2.Visible = False
End Sub
