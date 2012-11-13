VERSION 5.00
Begin VB.Form frmNavigate 
   Caption         =   "Navigate"
   ClientHeight    =   7875
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5850
   LinkTopic       =   "Form1"
   Picture         =   "frmNavigate.frx":0000
   ScaleHeight     =   7875
   ScaleWidth      =   5850
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAdmin 
      Caption         =   "Admin"
      Height          =   615
      Left            =   120
      TabIndex        =   8
      Top             =   6960
      Width           =   855
   End
   Begin VB.CommandButton cmdJumpSkateStore 
      Height          =   975
      Index           =   1
      Left            =   2040
      Picture         =   "frmNavigate.frx":9623
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4200
      Width           =   1575
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   1920
      Picture         =   "frmNavigate.frx":9F7B
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   240
      Width           =   1815
   End
   Begin VB.CommandButton cmdExit 
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2280
      Picture         =   "frmNavigate.frx":C7E8
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6360
      Width           =   1095
   End
   Begin VB.CommandButton cmdSkiStore 
      Caption         =   "!"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2040
      Picture         =   "frmNavigate.frx":D12C
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label lblSkateStore 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Click to go to our Skate Store"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   1920
      TabIndex        =   7
      Top             =   5400
      Width           =   1815
   End
   Begin VB.Label lblExit 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Click to Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   2040
      TabIndex        =   5
      Top             =   7200
      Width           =   1575
   End
   Begin VB.Label lblSkiStore 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Click to go to our Ski Store"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   1680
      TabIndex        =   4
      Top             =   3360
      Width           =   2415
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Click To go to our Store Front"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   1920
      TabIndex        =   3
      Top             =   1560
      Width           =   1935
   End
End
Attribute VB_Name = "frmNavigate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Dave and JB's ski and skate store
'frmFront
'Dave Lambert and John Brine
'Thursday March 23
'This form is used to navigate the site.  it goes to all of the areas of the site


Private Sub cmdAdmin_Click()
    Dim password As String
    Dim search As Single
    password = InputBox("Please input password", "Password")
    search = InStr(password, pass)
    If search > 0 Then
        'moves to the admin page from the navigate page
        frmAdmin.Visible = True
        frmNavigate.Visible = False
    Else
        frmNavigate.Visible = False
        frmFront.Visible = True
    End If
End Sub

Private Sub cmdBack_Click()
    'goes to the store front page
    frmFront.Visible = True
    frmNavigate.Visible = False
End Sub

Private Sub cmdExit_Click()
    'ends program and displays thank you message for comming
    MsgBox "Thanks For stopping! Come back soon!", , "Thanks!"
    End
End Sub

Private Sub cmdJumpSkateStore_Click(Index As Integer)
    'goes to the skate store page
    frmNavigate.Visible = False
    frmSkateStore.Visible = True
End Sub

Private Sub cmdSkiStore_Click()
    'goes to the ski store page
    frmSkiStore.Visible = True
    frmNavigate.Visible = False
End Sub


