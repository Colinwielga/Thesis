VERSION 5.00
Begin VB.Form frmsals 
   BackColor       =   &H007B5E02&
   Caption         =   "Sals"
   ClientHeight    =   8655
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11805
   LinkTopic       =   "Form1"
   ScaleHeight     =   8655
   ScaleWidth      =   11805
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdquit 
      BackColor       =   &H00008000&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro B"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6720
      Width           =   2655
   End
   Begin VB.CommandButton cmdleave 
      BackColor       =   &H00008000&
      Caption         =   "Continue on your tour de st. joe"
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro B"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5160
      Width           =   2655
   End
   Begin VB.CommandButton cmd2 
      BackColor       =   &H0003CCE9&
      Caption         =   "Pick us"
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro B"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7800
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.CommandButton cmd1 
      BackColor       =   &H0003CCE9&
      Caption         =   "Pick us"
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro B"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7800
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.CommandButton cmdhit 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Take home some hotties"
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro B"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   9600
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3600
      Width           =   1815
   End
   Begin VB.CommandButton cmdtalkto 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Talk to"
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro B"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3600
      Width           =   1815
   End
   Begin VB.CommandButton cmddance 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Dancing"
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro B"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3600
      Width           =   1815
   End
   Begin VB.CommandButton cmddrink 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Get Bombed"
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro B"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3600
      Width           =   1815
   End
   Begin VB.Label lbldoing 
      BackColor       =   &H007B5E02&
      Caption         =   "Click on what you would like to be doing..."
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro B"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4200
      TabIndex        =   8
      Top             =   3720
      Width           =   3255
   End
   Begin VB.Image Image6 
      Height          =   2520
      Left            =   8040
      Picture         =   "frmsals.frx":0000
      Top             =   480
      Visible         =   0   'False
      Width           =   3480
   End
   Begin VB.Image Image5 
      Height          =   3075
      Left            =   4200
      Picture         =   "frmsals.frx":1C902
      Top             =   120
      Visible         =   0   'False
      Width           =   4080
   End
   Begin VB.Image Image4 
      Height          =   2835
      Left            =   7560
      Picture         =   "frmsals.frx":456B4
      Top             =   4800
      Visible         =   0   'False
      Width           =   3765
   End
   Begin VB.Image Image3 
      Height          =   2865
      Left            =   3240
      Picture         =   "frmsals.frx":6851A
      Top             =   4800
      Visible         =   0   'False
      Width           =   3795
   End
   Begin VB.Image Image2 
      Height          =   2640
      Left            =   600
      Picture         =   "frmsals.frx":8BC64
      Top             =   240
      Visible         =   0   'False
      Width           =   4185
   End
   Begin VB.Image Image1 
      Height          =   2535
      Left            =   360
      Top             =   480
      Width           =   4095
   End
End
Attribute VB_Name = "frmsals"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

 
    'Project name:  Tour De St. Joe
    'Form:  frmsals, "Sals"
    'Author:  Brooke and Josh
    'Date:  3/28/08
    'Objective: To use dynamic picture loading to show what sorts of activities a user could "do" at Sal's.

Private Sub cmd1_Click()

    MsgBox ("Welcome to the Big Show!  We're going to have a good time.")

End Sub

Private Sub cmd2_Click()

    MsgBox ("We're going streaking!!  Up through the Quad and to the gymnasium.   Bring your big green hat.")

End Sub

Private Sub cmddance_Click()

    Image2.Visible = True
    Image3.Visible = False
    Image4.Visible = False
    Image5.Visible = False
    Image6.Visible = False
    
    cmd1.Visible = False
    cmd2.Visible = False

End Sub


Private Sub frmsals_Click()

    Image2.Visible = False
    Image3.Visible = False
    Image4.Visible = False
    Image5.Visible = False
    Image6.Visible = False
    
    cmd1.Visible = False
    cmd2.Visible = False
    cmddance.Visible = True
    cmddrink.Visible = True
    cmdhit.Visible = True
    cmdtalkto.Visible = True

End Sub

Private Sub cmddrink_Click()

    Image2.Visible = False
    Image3.Visible = False
    Image4.Visible = False
    Image5.Visible = True
    Image6.Visible = False

    cmd1.Visible = False
    cmd2.Visible = False

End Sub

Private Sub cmdhit_Click()

    Image2.Visible = False
    Image3.Visible = False
    Image4.Visible = False
    Image5.Visible = False
    Image6.Visible = True

    cmd1.Visible = False
    cmd2.Visible = False
    
End Sub

Private Sub cmdleave_Click()

    frmjoetown.Show
    frmsals.Hide

End Sub

Private Sub cmdquit_Click()
    End
End Sub

Private Sub cmdtalkto_Click()

    Image2.Visible = False
    Image3.Visible = True
    Image4.Visible = True
    Image5.Visible = False
    Image6.Visible = False
    
    cmd1.Visible = True
    cmd2.Visible = True

End Sub
