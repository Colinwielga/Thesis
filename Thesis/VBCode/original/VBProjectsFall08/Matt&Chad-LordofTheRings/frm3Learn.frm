VERSION 5.00
Begin VB.Form frm3Learn 
   Caption         =   "Form1"
   ClientHeight    =   6450
   ClientLeft      =   2025
   ClientTop       =   1905
   ClientWidth     =   9420
   LinkTopic       =   "Form1"
   Picture         =   "frm3Learn.frx":0000
   ScaleHeight     =   6450
   ScaleWidth      =   9420
   Begin VB.CommandButton cmdGoBack 
      Caption         =   "Go Back"
      Height          =   255
      Left            =   7920
      TabIndex        =   6
      Top             =   6120
      Width           =   1095
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "End Your Journey"
      Height          =   255
      Left            =   7680
      TabIndex        =   5
      Top             =   5760
      Width           =   1575
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "Learn"
      Height          =   255
      Left            =   4320
      TabIndex        =   4
      Top             =   2160
      Width           =   855
   End
   Begin VB.TextBox txtName 
      Height          =   375
      Left            =   3960
      TabIndex        =   3
      Top             =   1680
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      Height          =   1215
      Left            =   120
      ScaleHeight     =   1155
      ScaleWidth      =   9195
      TabIndex        =   0
      Top             =   -240
      Width           =   9255
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Please type the name of the character you wish to learn more about"
      BeginProperty Font 
         Name            =   "Pristina"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   360
      TabIndex        =   2
      Top             =   1200
      Width           =   8655
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Frodo  Sam Merry Pippin Aragorn Gandalf Boromir Legolas Gimli"
      BeginProperty Font 
         Name            =   "Pristina"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   4215
      Left            =   240
      TabIndex        =   1
      Top             =   2040
      Width           =   1455
   End
End
Attribute VB_Name = "frm3Learn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdGo_Click()
 Dim Found As Boolean
    Dim CharacterName As String
    Dim Pos As Integer, CTR As Integer
    Dim Character(1 To 9) As String
    Dim Info(1 To 9) As String
    Dim NumLines As Integer
    Dim NewLine As String
    
    Open App.Path & "\Bio.txt" For Input As #1
    CTR = 0
    CharacterName = txtName.Text
    Do Until EOF(1)
        CTR = CTR + 1
        Input #1, Character(CTR), NumLines
        For Pos = 1 To NumLines
            Input #1, NewLine
            Info(CTR) = Info(CTR) & vbCrLf & NewLine
        Next Pos
    Loop
    Close #1
    Pos = 0
    Do While (Found = False And Pos < CTR)
        Pos = Pos + 1
        If LCase(Character(Pos)) = LCase(CharacterName) Then
            Found = True
        End If
    Loop
    picture1.Cls
    If Found = True Then
        picture1.Print Info(Pos)
    Else
        MsgBox "Please try again.  Make sure name is properly spelled.", , "Error" 'if there was no match, a messagebox will be displayed to the user informing them of the error
    End If
End Sub

Private Sub cmdGoBack_Click()
    frm3Learn.Hide
    frm2Characters.Show
End Sub

Private Sub cmdQuit_Click()
    End
End Sub
