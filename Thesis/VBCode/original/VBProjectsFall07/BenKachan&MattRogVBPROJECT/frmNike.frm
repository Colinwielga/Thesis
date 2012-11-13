VERSION 5.00
Begin VB.Form frmNike 
   Caption         =   "Long Drive Competition"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   Picture         =   "frmNike.frx":0000
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CommandButton cmdPro 
      Caption         =   "Return to Pro Shop"
      Height          =   735
      Left            =   480
      TabIndex        =   6
      Top             =   6840
      Width           =   1695
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return to Main Menu"
      Height          =   615
      Left            =   480
      TabIndex        =   5
      Top             =   5880
      Width           =   2055
   End
   Begin VB.CommandButton cmdRip 
      Caption         =   "So you think you're Tiger Woods? Step up and rip it with his gear."
      Height          =   1575
      Left            =   600
      TabIndex        =   4
      Top             =   3600
      Width           =   2055
   End
   Begin VB.TextBox txtPower 
      Height          =   975
      Left            =   1560
      TabIndex        =   3
      Top             =   1920
      Width           =   615
   End
   Begin VB.TextBox txtClub 
      Height          =   1335
      Left            =   1560
      TabIndex        =   2
      Top             =   120
      Width           =   615
   End
   Begin VB.Label lblPower 
      Caption         =   "How hard would you like to swing on a scale from 1 to 10?"
      Height          =   975
      Left            =   240
      TabIndex        =   1
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label lbClub 
      Caption         =   "Select the power of club you wish to use on a scale from 1 to 10? 10 being the most powerful."
      Height          =   1335
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "frmNike"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdPro_Click()
    frmNike.Hide
    frmGolf.Show
End Sub

Private Sub cmdReturn_Click()
    frmHome.Show
    frmNike.Hide
End Sub

Private Sub cmdRip_Click()
    'We are asking the user to input a club preference and power preference, using case select we have come up with different situations which will result in varying drive distances
    Dim power As Single, Calloway As Integer, club As Single, distance As Single
    club = txtClub.Text
    power = txtPower.Text
    Nike = 0.75
    Select Case power
        Case Is = 10
            distance = ((Nike + 0.9) * club * power) * 4
            MsgBox "Looked like Tiger with the power of that swing that one is out there: " & FormatNumber(distance, 2) & " yards"
        Case Is = 9
            distance = ((Nike - 0.2) * club * power) * 4
            MsgBox "Forcing the issue a little too much resulted in: " & FormatNumber(distance, 2) & " yards"
        Case Is = 8
            distance = (Nike * club * power) * 4
            MsgBox "That looked like a better swing as you managed: " & FormatNumber(distance, 2) & " yards"
        Case Is = 7
            distance = ((Nike + 0.6) * club * power) * 4
            MsgBox "Beautiful swing, that one should be out there about: " & FormatNumber(distance, 2) & " yards"
        Case Is = 6
        distance = ((Nike + 0.5) * club * power) * 4
            MsgBox "You are flirting with some good play here, that one was: " & FormatNumber(distance, 2) & " yards"
        Case Is = 5
        distance = ((Nike + 0.2) * club * power) * 4
            MsgBox "That was a smooth sming, only good for: " & FormatNumber(distance, 2) & " yards"
        Case Is = 4
        distance = ((Nike + 0.1) * club * power) * 4
            MsgBox "Come on show off those muscles next time, this one went: " & FormatNumber(distance, 2) & " yards"
        Case Is = 3
        distance = (Nike * club * power) * 4
            MsgBox "Really? you cant do better than that, thats only: " & FormatNumber(distance, 2) & " yards"
        Case Is = 2
        distance = ((Nike - 0.4) * club * power) * 4
            MsgBox "Come on that was a meager: " & FormatNumber(distance, 2) & " yards"
        Case Is = 1
        distance = ((Nike - 0.5) * club * power) * 4
            MsgBox "Didn't go past the ladies tees with a crushing distance of: " & FormatNumber(distance, 2) & " yards... Dick Out!"
    End Select
End Sub
