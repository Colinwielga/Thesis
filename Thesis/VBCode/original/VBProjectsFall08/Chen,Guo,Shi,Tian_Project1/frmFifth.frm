VERSION 5.00
Begin VB.Form frmFifth 
   Caption         =   "Form1"
   ClientHeight    =   2820
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3555
   LinkTopic       =   "Form1"
   Picture         =   "frmFifth.frx":0000
   ScaleHeight     =   2820
   ScaleWidth      =   3555
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdEnter 
      Caption         =   "Enter"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      MaskColor       =   &H00FF0000&
      TabIndex        =   2
      Top             =   1440
      Width           =   2535
   End
   Begin VB.TextBox txtPassword 
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Text            =   "Password"
      Top             =   840
      Width           =   2535
   End
   Begin VB.TextBox txtUsername 
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Text            =   "Username"
      Top             =   240
      Width           =   2535
   End
End
Attribute VB_Name = "frmFifth"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdEnter_Click()
'the passport and username can be edited in the file "\Username and Passport.txt"
'if the username and passport match then the frmFourth (winning strategy of NIM) will be shown
Dim UserName(1 To 4) As String
Dim Password(1 To 4) As String
Dim M As Integer
Dim found As Boolean
Dim Z As Integer
Dim User As String
Dim PasswordInput As String
User = txtUsername.Text
PasswordInput = txtPassword.Text
M = 0
Z = 0
Open App.Path & "\Username and Passport.txt" For Input As #1
Do Until EOF(1)
    M = M + 1
    Input #1, UserName(M), Password(M)
Loop
Close #1
Do Until (found = True Or Z > 3)
    Z = Z + 1
    If UserName(Z) = User Then
        If PasswordInput = Password(Z) Then
            found = True
        End If
    End If
Loop

If found = True Then
    frmFourth.Visible = True
    frmFifth.Visible = False
Else: MsgBox "Sorry, the username doesn't exist or the username and password don't match.", , "Error"
    frmFifth.Visible = False
    frmThird.Visible = True
End If
End Sub
