VERSION 5.00
Begin VB.Form frmTunnel 
   Caption         =   "Into the tunnel..."
   ClientHeight    =   11625
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14685
   LinkTopic       =   "Form1"
   Picture         =   "frmTunnel.frx":0000
   ScaleHeight     =   11625
   ScaleWidth      =   14685
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cdmStats 
      Caption         =   "View your stats!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   11160
      TabIndex        =   2
      Top             =   3240
      Width           =   2055
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H80000012&
      BeginProperty Font 
         Name            =   "Myriad Condensed Web"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   2295
      Left            =   10320
      ScaleHeight     =   2235
      ScaleWidth      =   3795
      TabIndex        =   1
      Top             =   720
      Width           =   3855
   End
   Begin VB.CommandButton cmdContinue 
      Caption         =   "Continue to the tunnel's end..."
      BeginProperty Font 
         Name            =   "Myriad Condensed Web"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   10800
      TabIndex        =   0
      Top             =   5880
      Width           =   2895
   End
End
Attribute VB_Name = "frmtunnel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cdmStats_Click()    'show stats
    picResults.Cls
    picResults.Print "You've done well surviving the attack."
    picResults.Print "Your stats are as follows:"
    picResults.Print HP; " H.P."
    picResults.Print FormatCurrency(Money)
    picResults.Print Attack; " attack points."
    picResults.Print "Now click continue and see what lies ahead..."
End Sub

Private Sub cmdContinue_Click() 'continue to end of tunnel
    frmtunnel.Hide
    frmFinish.Show
    MsgBox ("You've done it!  You have escaped the aliens and found a safe place."), , ("Yay!")
    MsgBox ("Other people are here too and they say the army is almost done securing what's left of earth!"), , ("You are safe!")
    
    
End Sub

