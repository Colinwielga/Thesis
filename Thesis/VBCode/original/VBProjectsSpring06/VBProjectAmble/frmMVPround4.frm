VERSION 5.00
Begin VB.Form frmMVPround4 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form1"
   ClientHeight    =   1305
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6585
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   1305
   ScaleWidth      =   6585
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdback 
      Caption         =   "Back"
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton cmdcompute 
      Caption         =   "Enter"
      Height          =   495
      Left            =   5040
      TabIndex        =   2
      Top             =   240
      Width           =   1455
   End
   Begin VB.TextBox txtteam 
      Height          =   285
      Left            =   2520
      TabIndex        =   1
      Top             =   360
      Width           =   2415
   End
   Begin VB.Label lbljeff 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Created By: Jeff Amble"
      Height          =   255
      Left            =   2880
      TabIndex        =   4
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label lblexplain 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Enter team of the MVP you are looking for"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   2775
   End
End
Attribute VB_Name = "frmMVPround4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This form enables the user to enter a team name and'
'receive the MVP of that team, for that round'
Option Explicit
Dim pos As Integer
'This button enables the user to go back to the main page'
Private Sub cmdback_Click()
    frmMVPround4.Visible = False
    frmMVP.Visible = True
End Sub
'This button computes the users team and retrieves the MVP'
Private Sub cmdcompute_Click()
    Dim txt As String
    Dim pos As Single
    Dim N As String
    N = txtteam.Text
    Open App.Path & "\4throundMVP.txt" For Input As #1
        Do Until (EOF(1) Or pos <> 0)
            Input #1, txt
            pos = InStr(LCase(txt), LCase(N))
        Loop
    
        If pos <> 0 Then
            MsgBox txt, , "MVP"
        Else
            MsgBox "Team not found, Please enter new team name", , "Error"
        End If
    Close #1
End Sub


'This textbox displays the MVP'
Private Sub txtteam_Change()

End Sub
