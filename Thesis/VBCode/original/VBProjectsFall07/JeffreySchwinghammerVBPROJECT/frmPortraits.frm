VERSION 5.00
Begin VB.Form frmPortraits 
   BackColor       =   &H80000007&
   Caption         =   "Portraits"
   ClientHeight    =   9315
   ClientLeft      =   870
   ClientTop       =   1575
   ClientWidth     =   10620
   LinkTopic       =   "Form1"
   Picture         =   "frmPortraits.frx":0000
   ScaleHeight     =   9315
   ScaleWidth      =   10620
   Begin VB.PictureBox picTure 
      BorderStyle     =   0  'None
      Height          =   6615
      Left            =   1560
      Picture         =   "frmPortraits.frx":9FA2A
      ScaleHeight     =   6615
      ScaleWidth      =   8295
      TabIndex        =   3
      Top             =   360
      Width           =   8295
   End
   Begin VB.CommandButton cmdEnterCombo 
      Caption         =   "Enter Combo"
      Height          =   495
      Left            =   4800
      TabIndex        =   2
      Top             =   7200
      Width           =   1695
   End
   Begin VB.PictureBox picPortraitstxt 
      BeginProperty Font 
         Name            =   "Century"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   2040
      ScaleHeight     =   1035
      ScaleWidth      =   8235
      TabIndex        =   1
      Top             =   7920
      Width           =   8295
   End
   Begin VB.CommandButton cmdExitPortraits 
      Caption         =   "Finished Looking"
      Height          =   735
      Left            =   240
      TabIndex        =   0
      Top             =   8160
      Width           =   1455
   End
End
Attribute VB_Name = "frmPortraits"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdEnterCombo_Click()
Dim Letter1 As String
Dim Letter2 As String
Dim Letter3 As String

Letter1 = InputBox("What is the FIRST letter?")
Letter2 = InputBox("What is the SECOND letter?")
Letter3 = InputBox("What is the THIRD letter?")


    If UCase(Letter1) = "G" And UCase(Letter2) = "U" And UCase(Letter3) = "N" Then
        picPortraitstxt.Cls
        picPortraitstxt.Print "The chest unlocks!"
        picPortraitstxt.Print " "
        picPortraitstxt.Print "You look inside and find a PISTOL. You hold on to it for safety."
        picPortraitstxt.Print "You feel more confident carrying a weapon. Who knows what's out there..."
        Gun = True
        cmdEnterCombo.Enabled = False
    Else
    
        picPortraitstxt.Cls
        picPortraitstxt.Print "That combo is incorrect. The chest did not open."
    
    End If


End Sub

Private Sub cmdExitPortraits_Click()
   Dim answer As Integer

answer = MsgBox("Are you sure you are done looking at the portraits?", vbYesNo)
    If answer = vbYes Then
        frmLibrary.Show
        frmPortraits.Hide
    End If
End Sub

Private Sub Form_activate()
    picPortraitstxt.Cls
    picPortraitstxt.Print "There are three portraits hanging on the wall. Beneath it is a small"
    picPortraitstxt.Print "chest with a combination lock on it. The combo is three letters long."
    picPortraitstxt.Print "You recognize the Japanese Symbol for color on the chest."
    picPortraitstxt.Print "You wonder how you knew that. 'Who am I,' you think to yourself..."
End Sub

