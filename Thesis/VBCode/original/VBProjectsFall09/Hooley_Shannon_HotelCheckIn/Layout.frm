VERSION 5.00
Begin VB.Form frmLayout 
   Caption         =   "Form1"
   ClientHeight    =   8535
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10380
   LinkTopic       =   "Form1"
   Picture         =   "Layout.frx":0000
   ScaleHeight     =   8535
   ScaleWidth      =   10380
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdList 
      BackColor       =   &H00FFFFFF&
      Caption         =   "See the avaliable rooms  from lowest to highest price"
      BeginProperty Font 
         Name            =   "Freestyle Script"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   5760
      Width           =   1935
   End
   Begin VB.CommandButton cmdReturn 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Return to Check In"
      BeginProperty Font 
         Name            =   "Freestyle Script"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   360
      Width           =   2775
   End
   Begin VB.TextBox txtRoom118 
      BackColor       =   &H00800080&
      BeginProperty Font 
         Name            =   "Freestyle Script"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0FF&
      Height          =   480
      Index           =   2
      Left            =   3720
      TabIndex        =   27
      Text            =   "Queen"
      Top             =   1920
      Width           =   855
   End
   Begin VB.TextBox txtRoom105 
      BackColor       =   &H00404000&
      BeginProperty Font 
         Name            =   "Freestyle Script"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   510
      Left            =   9000
      TabIndex        =   26
      Text            =   "Double"
      Top             =   4680
      Width           =   975
   End
   Begin VB.TextBox txtRoom118 
      BackColor       =   &H00800080&
      BeginProperty Font 
         Name            =   "Freestyle Script"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0FF&
      Height          =   480
      Index           =   1
      Left            =   5640
      TabIndex        =   25
      Text            =   "Queen"
      Top             =   1920
      Width           =   855
   End
   Begin VB.TextBox txtRoom101 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Freestyle Script"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Index           =   1
      Left            =   9480
      TabIndex        =   24
      Text            =   "King"
      Top             =   1560
      Width           =   735
   End
   Begin VB.TextBox txtRoom107 
      BackColor       =   &H00800000&
      BeginProperty Font 
         Name            =   "Freestyle Script"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   420
      Index           =   1
      Left            =   7920
      TabIndex        =   23
      Text            =   "Suite"
      Top             =   5280
      Width           =   735
   End
   Begin VB.TextBox txtRoom122 
      BackColor       =   &H00800000&
      BeginProperty Font 
         Name            =   "Freestyle Script"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   435
      Index           =   1
      Left            =   3240
      TabIndex        =   22
      Text            =   "Suite"
      Top             =   1200
      Width           =   735
   End
   Begin VB.TextBox txtroom113 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Freestyle Script"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Index           =   0
      Left            =   3480
      TabIndex        =   21
      Text            =   "King"
      Top             =   8040
      Width           =   735
   End
   Begin VB.TextBox txtRoom117 
      BackColor       =   &H00800000&
      BeginProperty Font 
         Name            =   "Freestyle Script"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   420
      Index           =   0
      Left            =   2520
      TabIndex        =   20
      Text            =   "Suite"
      Top             =   5640
      Width           =   735
   End
   Begin VB.TextBox txtRoom118 
      BackColor       =   &H00800080&
      BeginProperty Font 
         Name            =   "Freestyle Script"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0FF&
      Height          =   450
      Index           =   0
      Left            =   1080
      TabIndex        =   19
      Text            =   "Queen"
      Top             =   5595
      Width           =   855
   End
   Begin VB.TextBox txtRoom111 
      BackColor       =   &H0000FFFF&
      BeginProperty Font 
         Name            =   "Freestyle Script"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   495
      Left            =   7320
      TabIndex        =   18
      Text            =   "Presidential Suite"
      Top             =   7440
      Width           =   2055
   End
   Begin VB.TextBox txtRoom121 
      BackColor       =   &H0000FFFF&
      BeginProperty Font 
         Name            =   "Freestyle Script"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   0
      TabIndex        =   17
      Text            =   "Presidential Suite"
      Top             =   5040
      Width           =   2175
   End
   Begin VB.CommandButton cmdRoom101 
      BackColor       =   &H0000FFFF&
      Caption         =   "X"
      Height          =   375
      Left            =   9240
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   3120
      Width           =   375
   End
   Begin VB.CommandButton cmdRoom102 
      BackColor       =   &H0000FFFF&
      Caption         =   "X"
      Height          =   375
      Index           =   1
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   1800
      Width           =   495
   End
   Begin VB.CommandButton cmdRoom103 
      Caption         =   "X"
      Enabled         =   0   'False
      Height          =   375
      Left            =   8760
      TabIndex        =   14
      Top             =   3600
      Width           =   375
   End
   Begin VB.CommandButton cmdRoom104 
      Caption         =   "X"
      Enabled         =   0   'False
      Height          =   375
      Left            =   7080
      TabIndex        =   13
      Top             =   2400
      Width           =   375
   End
   Begin VB.CommandButton cmdRoom106 
      BackColor       =   &H0000FFFF&
      Caption         =   "X"
      Height          =   375
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3600
      Width           =   375
   End
   Begin VB.CommandButton cmdRoom105 
      BackColor       =   &H0000FFFF&
      Caption         =   "X"
      Height          =   375
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4080
      Width           =   375
   End
   Begin VB.CommandButton cmdRoom107 
      Caption         =   "X"
      Enabled         =   0   'False
      Height          =   375
      Index           =   0
      Left            =   7920
      TabIndex        =   10
      Top             =   4440
      Width           =   375
   End
   Begin VB.CommandButton cmdRoom111 
      BackColor       =   &H0000FFFF&
      Caption         =   "X"
      Height          =   375
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   7560
      Width           =   375
   End
   Begin VB.CommandButton cmdRoom118 
      BackColor       =   &H0000FFFF&
      Caption         =   "X"
      Height          =   375
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3840
      Width           =   375
   End
   Begin VB.CommandButton cmdRoom113 
      Caption         =   "X"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4080
      TabIndex        =   7
      Top             =   7560
      Width           =   375
   End
   Begin VB.CommandButton cmdRoom115 
      Caption         =   "X"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4080
      TabIndex        =   6
      Top             =   6240
      Width           =   375
   End
   Begin VB.CommandButton cmdRoom117 
      Caption         =   "X"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3480
      TabIndex        =   5
      Top             =   4920
      Width           =   375
   End
   Begin VB.CommandButton cmdRoom119 
      Caption         =   "X"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3000
      TabIndex        =   4
      Top             =   4080
      Width           =   375
   End
   Begin VB.CommandButton cmdRoom120 
      Caption         =   "X"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3240
      TabIndex        =   3
      Top             =   3120
      Width           =   375
   End
   Begin VB.CommandButton cmdRoom122 
      BackColor       =   &H0000FFFF&
      Caption         =   "X"
      Height          =   375
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1560
      Width           =   495
   End
   Begin VB.CommandButton cmdRoom121 
      BackColor       =   &H00FFFFFF&
      Caption         =   "X"
      Enabled         =   0   'False
      Height          =   375
      Left            =   600
      MaskColor       =   &H00800080&
      TabIndex        =   1
      Top             =   2400
      UseMaskColor    =   -1  'True
      Width           =   495
   End
   Begin VB.TextBox txtLakefront 
      BeginProperty Font 
         Name            =   "Freestyle Script"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Text            =   "Lake Front Inn Gift Shop"
      Top             =   0
      Width           =   3615
   End
   Begin VB.Line Line22 
      BorderWidth     =   3
      X1              =   4200
      X2              =   4200
      Y1              =   2280
      Y2              =   3240
   End
   Begin VB.Line Line21 
      BorderWidth     =   3
      X1              =   3840
      X2              =   3840
      Y1              =   2400
      Y2              =   2760
   End
   Begin VB.Line Line20 
      BorderWidth     =   3
      X1              =   9000
      X2              =   8640
      Y1              =   4680
      Y2              =   4560
   End
   Begin VB.Line Line19 
      BorderWidth     =   3
      X1              =   9120
      X2              =   9120
      Y1              =   4680
      Y2              =   3960
   End
   Begin VB.Line Line18 
      BorderWidth     =   3
      X1              =   6000
      X2              =   6360
      Y1              =   2400
      Y2              =   3240
   End
   Begin VB.Line Line17 
      BorderWidth     =   3
      X1              =   6360
      X2              =   6960
      Y1              =   2280
      Y2              =   2760
   End
   Begin VB.Line Line16 
      BorderWidth     =   3
      X1              =   9480
      X2              =   8040
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Line Line15 
      BorderWidth     =   3
      X1              =   9600
      X2              =   9600
      Y1              =   1920
      Y2              =   2640
   End
   Begin VB.Line Line14 
      BorderWidth     =   3
      X1              =   8160
      X2              =   8160
      Y1              =   5280
      Y2              =   4920
   End
   Begin VB.Line Line13 
      BorderWidth     =   3
      X1              =   3480
      X2              =   2880
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Line Line12 
      BorderWidth     =   3
      X1              =   3480
      X2              =   3480
      Y1              =   1560
      Y2              =   1920
   End
   Begin VB.Line Line11 
      BorderWidth     =   3
      X1              =   3600
      X2              =   4080
      Y1              =   7440
      Y2              =   7440
   End
   Begin VB.Line Line10 
      BorderWidth     =   3
      X1              =   3600
      X2              =   3600
      Y1              =   8040
      Y2              =   7440
   End
   Begin VB.Line Line9 
      BorderWidth     =   3
      X1              =   7920
      X2              =   7920
      Y1              =   6240
      Y2              =   7440
   End
   Begin VB.Line Line8 
      BorderWidth     =   3
      X1              =   1560
      X2              =   1560
      Y1              =   6480
      Y2              =   6000
   End
   Begin VB.Line Line7 
      BorderWidth     =   3
      X1              =   3240
      X2              =   3840
      Y1              =   5760
      Y2              =   5400
   End
   Begin VB.Line Line6 
      BorderWidth     =   3
      X1              =   1560
      X2              =   3960
      Y1              =   6480
      Y2              =   6480
   End
   Begin VB.Line Line5 
      BorderWidth     =   3
      X1              =   1920
      X2              =   2880
      Y1              =   5760
      Y2              =   4920
   End
   Begin VB.Line Line4 
      BorderWidth     =   3
      X1              =   7920
      X2              =   6360
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line Line3 
      BorderWidth     =   3
      X1              =   7920
      X2              =   6480
      Y1              =   6960
      Y2              =   6960
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00800000&
      BorderWidth     =   3
      DrawMode        =   1  'Blackness
      X1              =   960
      X2              =   1080
      Y1              =   5040
      Y2              =   3720
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      DrawMode        =   1  'Blackness
      X1              =   1080
      X2              =   2040
      Y1              =   5040
      Y2              =   4200
   End
   Begin VB.Image Image1 
      Height          =   2220
      Left            =   120
      Picture         =   "Layout.frx":17089
      Top             =   6600
      Width           =   3330
   End
End
Attribute VB_Name = "frmLayout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Hotel Check In
'frmLayout
'Shannon Hooley
'10/16/09
'This form allows the guest to see what rooms are avaliable, click on the room to get more details
'as well as get an organized list of avaliable rooms, low - high price

Private Sub cmdList_Click()
'brings the guest to low - high pricing list
frmLayout.Hide
frmPricing.Show
End Sub

Private Sub cmdReturn_Click()
'brings the guest back to check in
frmLayout.Hide
frmCheckIn.Show
End Sub

Private Sub cmdRoom101_Click()
'info on room 101
MsgBox ("Room 101 is a King Room, priced at $206 a night.")
End Sub

Private Sub cmdRoom102_Click(Index As Integer)
'info on room 102
MsgBox ("Room 102 is a King Room, priced at $206 a night.")
End Sub

Private Sub cmdRoom105_Click()
'info on room 105
MsgBox ("Room 105 is a Double Room, priced at $109 a night.")
End Sub

Private Sub cmdRoom106_Click()
'info on room 106
MsgBox ("Room 106 is a Queen Room, priced at $150 a night.")
End Sub

Private Sub cmdRoom111_Click()
'info on room 111
MsgBox ("Room 111 is a Presidential Suite, priced at $453 a night.")
End Sub

Private Sub cmdRoom118_Click()
'info on room 118
MsgBox ("Room 118 is a Queen Room, priced at $150 a night.")
End Sub

Private Sub cmdRoom122_Click()
'info on room 122
MsgBox ("Room 122 is a Suite, priced at $379 a night.")
End Sub
