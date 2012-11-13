VERSION 5.00
Begin VB.Form frmPing 
   Caption         =   "Long Drive Competition"
   ClientHeight    =   8220
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10410
   LinkTopic       =   "Form1"
   Picture         =   "frmPing.frx":0000
   ScaleHeight     =   8220
   ScaleWidth      =   10410
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CommandButton cmdPro 
      Caption         =   "Return to Pro Shop"
      Height          =   735
      Left            =   360
      TabIndex        =   6
      Top             =   6840
      Width           =   1695
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return to Main Menu"
      Height          =   615
      Left            =   360
      TabIndex        =   5
      Top             =   6000
      Width           =   1935
   End
   Begin VB.CommandButton cmdHit 
      Caption         =   "Lets get that Ping zinging! See how far you can hit it."
      Height          =   1935
      Left            =   480
      TabIndex        =   4
      Top             =   3600
      Width           =   1935
   End
   Begin VB.TextBox txtPower 
      Height          =   1095
      Left            =   1440
      TabIndex        =   3
      Top             =   1800
      Width           =   615
   End
   Begin VB.TextBox txtClub 
      Height          =   1455
      Left            =   1440
      TabIndex        =   2
      Top             =   120
      Width           =   615
   End
   Begin VB.Label lblPower 
      Caption         =   "How hard would you like to swing on a scale from 1 to 10?"
      Height          =   1095
      Left            =   120
      TabIndex        =   1
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label lblClub 
      Caption         =   "Select the power of club you wish to use on a scale from 1 to 10? 10 being the most powerful."
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "frmPing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdHit_Click()
    'We are asking the user to input a club preference and power preference, using case select we have come up with different situations which will result in varying drive distances
    Dim power As Single, Calloway As Integer, club As Single, distance As Single
    club = txtClub.Text
    power = txtPower.Text
    Ping = 0.95
    Select Case power
        Case Is = 10
            distance = ((Ping - 0.3) * club * power) * 4
            MsgBox "Over swung on that one a bit and only put it out there: " & FormatNumber(distance, 2) & " yards"
        Case Is = 9
            distance = ((Ping - 0.1) * club * power) * 4
            MsgBox "Forcing the issue a little too much resulted in: " & FormatNumber(distance, 2) & " yards"
        Case Is = 8
            distance = (Ping * club * power) * 4
            MsgBox "That looked like a better swing as you managed: " & FormatNumber(distance, 2) & " yards"
        Case Is = 7
            distance = ((Ping + 1) * club * power) * 4
            MsgBox "That Ping definitely sent a lightning bolt down the fairway going: " & FormatNumber(distance, 2) & " yards"
        Case Is = 6
        distance = ((Ping + 0.8) * club * power) * 4
            MsgBox "You are flirting with some good play here, that one was: " & FormatNumber(distance, 2) & " yards"
        Case Is = 5
        distance = ((Ping + 0.4) * club * power) * 4
            MsgBox "That was a smooth sming, only good for: " & FormatNumber(distance, 2) & " yards"
        Case Is = 4
        distance = ((Ping + 0.2) * club * power) * 4
            MsgBox "Come on show off those muscles next time, this one went: " & FormatNumber(distance, 2) & " yards"
        Case Is = 3
        distance = (Ping * club * power) * 4
            MsgBox "Really? you cant do better than that, thats only: " & FormatNumber(distance, 2) & " yards"
        Case Is = 2
        distance = ((Ping - 0.4) * club * power) * 4
            MsgBox "Come on that was a meager: " & FormatNumber(distance, 2) & " yards"
        Case Is = 1
        distance = ((Ping - 1) * club * power) * 4
            MsgBox "Didn't go past the ladies tees with a crushing distance of: " & FormatNumber(distance, 2) & " yards... Dick Out!"
    End Select
End Sub

Private Sub cmdPro_Click()
    frmPing.Hide
    frmGolf.Show
End Sub

Private Sub cmdReturn_Click()
    frmHome.Show
    frmPing.Hide
End Sub
