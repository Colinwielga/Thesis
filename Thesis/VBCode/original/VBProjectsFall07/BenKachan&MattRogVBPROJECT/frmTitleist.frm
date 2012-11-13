VERSION 5.00
Begin VB.Form frmTitleist 
   Caption         =   "Long Drive Competition"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   Picture         =   "frmTitleist.frx":0000
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CommandButton cmdPro 
      Caption         =   "Return to Pro Shop"
      Height          =   735
      Left            =   240
      TabIndex        =   6
      Top             =   7320
      Width           =   1695
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return to Main Menu"
      Height          =   615
      Left            =   240
      TabIndex        =   5
      Top             =   6480
      Width           =   1935
   End
   Begin VB.CommandButton cmdPhil 
      Caption         =   "Step up like ""Lefty"" Phil Mickelson and unleash the Titleist"
      Height          =   2175
      Left            =   240
      TabIndex        =   4
      Top             =   3600
      Width           =   2055
   End
   Begin VB.TextBox txtPower 
      Height          =   1095
      Left            =   1440
      TabIndex        =   3
      Top             =   2040
      Width           =   615
   End
   Begin VB.TextBox txtClub 
      Height          =   1575
      Left            =   1440
      TabIndex        =   2
      Top             =   240
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "How hard would you like to swing on a scale from 1 to 10?"
      Height          =   1095
      Left            =   240
      TabIndex        =   1
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Select the power of club you wish to use on a scale from 1 to 10? 10 being the most powerful."
      Height          =   1455
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "frmTitleist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdPhil_Click()
Dim power As Single, Calloway As Integer, club As Single, distance As Single
    'We are asking the user to input a club preference and power preference, using case select we have come up with different situations which will result in varying drive distances
    club = txtClub.Text
    power = txtPower.Text
    Titleist = 1.05
    Select Case power
        Case Is = 10
            distance = ((Titleist - 0.5) * club * power) * 4
            MsgBox "Over swung on that one a bit and only put it out there: " & FormatNumber(distance, 2) & " yards"
        Case Is = 9
            distance = ((Titleist - 0.3) * club * power) * 4
            MsgBox "Forcing the issue a little too much resulted in: " & FormatNumber(distance, 2) & " yards"
        Case Is = 8
            distance = ((Titleist + 0.2) * club * power) * 4
            MsgBox "That looked like a better swing as you managed: " & FormatNumber(distance, 2) & " yards"
        Case Is = 7
            distance = ((Titleist + 0.5) * club * power) * 4
            MsgBox "Beautiful swing, that one should be out there about: " & FormatNumber(distance, 2) & " yards"
        Case Is = 6
        distance = ((Titleist + 0.4) * club * power) * 4
            MsgBox "You are flirting with some good play here, that one was: " & FormatNumber(distance, 2) & " yards"
        Case Is = 5
        distance = ((Titleist + 0.3) * club * power) * 4
            MsgBox "That was a smooth sming, only good for: " & FormatNumber(distance, 2) & " yards"
        Case Is = 4
        distance = ((Titleist + 0.3) * club * power) * 4
            MsgBox "Come on show off those muscles next time, this one went: " & FormatNumber(distance, 2) & " yards"
        Case Is = 3
        distance = ((Titleist + 0.2) * club * power) * 4
            MsgBox "Really? you cant do better than that, thats only: " & FormatNumber(distance, 2) & " yards"
        Case Is = 2
        distance = ((Titleist - 0.2) * club * power) * 4
            MsgBox "Come on that was a meager: " & FormatNumber(distance, 2) & " yards"
        Case Is = 1
        distance = ((Titleist - 1) * club * power) * 4
            MsgBox "Didn't go past the ladies tees with a crushing distance of: " & FormatNumber(distance, 2) & " yards... Dick Out!"
    End Select
End Sub

Private Sub cmdPro_Click()
    frmTitleist.Hide
    frmGolf.Show
End Sub

Private Sub cmdReturn_Click()
    frmHome.Show
    frmTitleist.Hide
End Sub
