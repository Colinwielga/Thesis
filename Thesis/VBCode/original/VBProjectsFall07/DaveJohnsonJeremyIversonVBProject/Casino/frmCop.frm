VERSION 5.00
Begin VB.Form frmCop 
   Caption         =   "Security"
   ClientHeight    =   5970
   ClientLeft      =   3000
   ClientTop       =   2580
   ClientWidth     =   8925
   LinkTopic       =   "Form1"
   ScaleHeight     =   5970
   ScaleWidth      =   8925
   Begin VB.PictureBox Picture1 
      Height          =   6735
      Left            =   -120
      Picture         =   "frmCop.frx":0000
      ScaleHeight     =   6675
      ScaleWidth      =   10275
      TabIndex        =   0
      Top             =   0
      Width           =   10335
      Begin VB.CommandButton cmdID 
         BackColor       =   &H000000C0&
         Caption         =   "Show ID"
         Height          =   735
         Left            =   2400
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   5040
         Width           =   2175
      End
      Begin VB.CommandButton cmdQuit 
         BackColor       =   &H00404080&
         Caption         =   "Leave Casino"
         Height          =   735
         Left            =   4680
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   5040
         Width           =   2175
      End
      Begin VB.CommandButton cmdPass 
         BackColor       =   &H00004080&
         Caption         =   "Sneak Past Security"
         Height          =   735
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   5040
         Width           =   2055
      End
   End
End
Attribute VB_Name = "frmCop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Project: Mystake Lake Casino
'Authors: David Johnson And Jeremy Iverson
'Date: Monday, November 5, 2007

'This form simulates the security at a Casino to check ID and make sure
'the user is of legal age to gamble

Private Sub cmdID_Click()
    'Show ID
    Dim d As Single
    d = MsgBox("Alright, " & nameglobal & ".  If this is your real name then this looks legit. We better not catch you doing any illegal business, ya hear! Amscray!", , "Amscray")
    frmCop.Hide
    frmLobby.Show
End Sub

Private Sub cmdPass_Click()
    Dim a As Single, b As Single, c As Single, age As Integer
    'Check age, you must be 18 to enter a casino, if not you are kicked out
    a = InputBox("Where do you think you're going? How old are you?", "Age")
    age = 18 - a
    If age <= 0 Then
        If a > 90 Then
            MsgBox "Dang that's ancient. Someone at your age will have a heart attack if you win the jackpot! I still need to see your ID"
        Else
            b = MsgBox("Did I hear a stutter? I'm going to need to see some identification, punk.", , "ID")
        End If
    Else
        c = MsgBox("Get out of this casino and come back in " & age & " years.", , "Age Verification")
        frmCop.Hide
        frmCasino.Show
    End If
End Sub

Private Sub cmdQuit_Click()
    'Gives the honest user who is not of age an opportunity to leave
    frmCop.Hide
    frmCasino.Show
End Sub
