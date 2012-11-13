VERSION 5.00
Begin VB.Form frmLevel5 
   Caption         =   "Level 5"
   ClientHeight    =   8505
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10245
   LinkTopic       =   "Form1"
   ScaleHeight     =   8505
   ScaleWidth      =   10245
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGet3 
      Caption         =   "Get Question 3"
      Height          =   375
      Left            =   4800
      TabIndex        =   11
      Top             =   6480
      Width           =   1215
   End
   Begin VB.CommandButton cmdGet2 
      Caption         =   "Get Question 2"
      Height          =   495
      Left            =   4680
      TabIndex        =   10
      Top             =   4320
      Width           =   1335
   End
   Begin VB.CommandButton cmdGet1 
      Caption         =   "Get Question 1"
      Height          =   375
      Left            =   4680
      TabIndex        =   9
      Top             =   2040
      Width           =   1455
   End
   Begin VB.CommandButton cmdRight3 
      Caption         =   "Are you right? "
      Height          =   615
      Left            =   5280
      TabIndex        =   8
      Top             =   6960
      Width           =   1695
   End
   Begin VB.CommandButton cmdRight2 
      Caption         =   "Are you right? "
      Height          =   615
      Left            =   5160
      TabIndex        =   7
      Top             =   5040
      Width           =   1815
   End
   Begin VB.CommandButton cmdRight1 
      Caption         =   "Are you right?"
      Height          =   615
      Left            =   5040
      TabIndex        =   6
      Top             =   2640
      Width           =   1695
   End
   Begin VB.TextBox txtAns3 
      Height          =   495
      Left            =   360
      TabIndex        =   5
      Top             =   7200
      Width           =   4095
   End
   Begin VB.TextBox txtAns2 
      Height          =   615
      Left            =   360
      TabIndex        =   4
      Top             =   5040
      Width           =   3975
   End
   Begin VB.TextBox txtAns1 
      Height          =   615
      Left            =   480
      TabIndex        =   3
      Top             =   2640
      Width           =   3735
   End
   Begin VB.PictureBox picResults3 
      BackColor       =   &H00C0FFC0&
      Height          =   495
      Left            =   240
      ScaleHeight     =   435
      ScaleWidth      =   4275
      TabIndex        =   2
      Top             =   6360
      Width           =   4335
   End
   Begin VB.PictureBox picResults2 
      BackColor       =   &H00C0FFC0&
      Height          =   495
      Left            =   240
      ScaleHeight     =   435
      ScaleWidth      =   4155
      TabIndex        =   1
      Top             =   4080
      Width           =   4215
   End
   Begin VB.PictureBox picResults1 
      BackColor       =   &H00C0FFC0&
      Height          =   495
      Left            =   240
      ScaleHeight     =   435
      ScaleWidth      =   4155
      TabIndex        =   0
      Top             =   1800
      Width           =   4215
   End
   Begin VB.Label lblDirections 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Caption         =   $"frmLevel5.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   360
      TabIndex        =   12
      Top             =   360
      Width           =   8775
   End
   Begin VB.Image imgQuit 
      Height          =   705
      Left            =   7800
      Picture         =   "frmLevel5.frx":00AF
      Top             =   7200
      Width           =   615
   End
   Begin VB.Image Image1 
      Height          =   8595
      Left            =   0
      Picture         =   "frmLevel5.frx":05B4
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10290
   End
End
Attribute VB_Name = "frmLevel5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim Grinch As String, Sam As String, Fish As String

Private Sub cmdGet1_Click()
    Grinch = "How the Grinch Stole Christmas"
    picResults1.Print Left(Grinch, 23)
End Sub

Private Sub cmdGet2_Click()
    Sam = "Green Eggs and Ham"
    picResults2.Print Right(Sam, 8)
End Sub

Private Sub cmdGet3_Click()
     Fish = "One Fish Two Fish Red Fish Blue Fish"
     picResults3.Print Mid(Fish, 4, 16)

End Sub

Private Sub cmdRight1_Click()
    txtAns1.Text = "How the Grinch Stole Christmas"
    If Grinch = "How the Grinch Stole Christmas" Then
        MsgBox YourName & " You are correct!!", , "Hooray!!"
    Else
        MsgBox YourName & " Try Again", , "Oops"
    End If
End Sub

Private Sub cmdRight2_Click()
     txtAns2.Text = "Green Eggs and Ham"
    If Sam = "Green Eggs and Ham" Then
        MsgBox YourName & " You are correct!!", , "Hooray!!"
    Else
        MsgBox YourName & " Try Again", , "Oops"
    End If

End Sub

Private Sub cmdRight3_Click()
      txtAns3.Text = "One Fish Two Fish Red Fish Blue Fish"
        If Fish = "One Fish Two Fish Red Fish Blue Fish" Then
            MsgBox YourName & " You are correct!!", , "Hooray!!"
            frmLevel5.Visible = False
            frmFinal.Visible = True
        Else
          MsgBox YourName & " Try Again", , "Oops"
        End If

End Sub

Private Sub imgQuit_Click()
    End
End Sub

