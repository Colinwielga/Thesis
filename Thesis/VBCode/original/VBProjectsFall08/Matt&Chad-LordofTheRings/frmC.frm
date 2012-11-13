VERSION 5.00
Begin VB.Form frmC 
   Caption         =   "Form1"
   ClientHeight    =   6405
   ClientLeft      =   3015
   ClientTop       =   1710
   ClientWidth     =   7695
   LinkTopic       =   "Form1"
   Picture         =   "frmC.frx":0000
   ScaleHeight     =   6405
   ScaleWidth      =   7695
   Begin VB.CommandButton Command4 
      Caption         =   "End Your Journey"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   5880
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Continue"
      Height          =   375
      Left            =   6120
      TabIndex        =   3
      Top             =   5880
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Summon the Army of the Dead"
      Height          =   495
      Left            =   5880
      TabIndex        =   2
      Top             =   2040
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ignite the Beacons"
      Height          =   495
      Left            =   5880
      TabIndex        =   1
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmC.frx":9753
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
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7455
   End
End
Attribute VB_Name = "frmC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Beacons As Boolean, Army As Boolean
Private Sub Command1_Click()
    MsgBox ("Great! More help is on the way!")
    Beacons = True
End Sub

Private Sub Command2_Click()
    MsgBox ("It's working! The army is coming to fulfill their oath!")
    Army = True
End Sub

Private Sub Command3_Click()
    If Beacons And Army Then
        frmC.Hide
        frmWIN.Show
    Else
        MsgBox ("We need more help! Please send for aide!")
    End If
End Sub
