VERSION 5.00
Begin VB.Form frmCalloway 
   Caption         =   "Long Drive Competition"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   Picture         =   "frmCalloway.frx":0000
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CommandButton cmdPro 
      Caption         =   "Return to Pro Shop"
      Height          =   615
      Left            =   480
      TabIndex        =   6
      Top             =   5400
      Width           =   1695
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return to Main Menu"
      Height          =   615
      Left            =   360
      TabIndex        =   5
      Top             =   4320
      Width           =   1935
   End
   Begin VB.TextBox txtClub 
      Height          =   975
      Left            =   1680
      TabIndex        =   4
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton cmdRipit 
      Caption         =   "Tee it up and RIP IT!!!"
      Height          =   1455
      Left            =   240
      TabIndex        =   2
      Top             =   2280
      Width           =   2055
   End
   Begin VB.TextBox txtPower 
      Height          =   735
      Left            =   1680
      TabIndex        =   0
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000013&
      Caption         =   "Select the power of club you wish to use on a scale from 1 to 10? 10 being the most powerful."
      Height          =   975
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label lblPower 
      BackColor       =   &H80000013&
      Caption         =   "How hard would you like to swing on a scale from 1 to 10?"
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   1455
   End
End
Attribute VB_Name = "frmCalloway"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdPro_Click()
    frmCalloway.Hide
    frmGolf.Show
End Sub

Private Sub cmdReturn_Click()
    frmHome.Show
    frmCalloway.Hide
End Sub

Private Sub cmdRipit_Click()
    'We are asking the user to input a club preference and power preference, using case select we have come up with different situations which will result in varying drive distances
    Dim power As Single, Calloway As Integer, club As Single, distance As Single
    club = txtClub.Text
    power = txtPower.Text
    Calloway = 1.1
    Select Case power
        Case Is = 10
            distance = ((Calloway - 0.4) * club * power) * 4
            MsgBox "Over swung on that one a bit and only put it out there: " & FormatNumber(distance, 2) & " yards"
        Case Is = 9
            distance = ((Calloway - 0.2) * club * power) * 4
            MsgBox "Forcing the issue a little too much resulted in: " & FormatNumber(distance, 2) & " yards"
        Case Is = 8
            distance = (Calloway * club * power) * 4
            MsgBox "That looked like a better swing as you managed: " & FormatNumber(distance, 2) & " yards"
        Case Is = 7
            distance = ((Calloway + 0.7) * club * power) * 4
            MsgBox "Beautiful swing, that one should be out there about: " & FormatNumber(distance, 2) & " yards"
        Case Is = 6
        distance = ((Calloway + 0.5) * club * power) * 4
            MsgBox "You are flirting with some good play here, that one was: " & FormatNumber(distance, 2) & " yards"
        Case Is = 5
        distance = ((Calloway + 0.2) * club * power) * 4
            MsgBox "That was a smooth sming, only good for: " & FormatNumber(distance, 2) & " yards"
        Case Is = 4
        distance = ((Calloway + 0.1) * club * power) * 4
            MsgBox "Come on show off those muscles next time, this one went: " & FormatNumber(distance, 2) & " yards"
        Case Is = 3
        distance = (Calloway * club * power) * 4
            MsgBox "Really? you cant do better than that, thats only: " & FormatNumber(distance, 2) & " yards"
        Case Is = 2
        distance = ((Calloway - 0.4) * club * power) * 4
            MsgBox "Come on that was a meager: " & FormatNumber(distance, 2) & " yards"
        Case Is = 1
        distance = ((Calloway - 1) * club * power) * 4
            MsgBox "Didn't go past the ladies tees with a crushing distance of: " & FormatNumber(distance, 2) & " yards... Dick Out!"
    End Select
End Sub
