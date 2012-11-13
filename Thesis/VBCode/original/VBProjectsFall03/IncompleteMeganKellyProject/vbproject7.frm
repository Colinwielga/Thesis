VERSION 5.00
Begin VB.Form suspense7 
   BackColor       =   &H000000C0&
   Caption         =   "The suspense is building... can you feel it?"
   ClientHeight    =   5205
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5685
   LinkTopic       =   "Form7"
   ScaleHeight     =   5205
   ScaleWidth      =   5685
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Continue7 
      BackColor       =   &H000000C0&
      Caption         =   "Click to find out who won the battle..."
      Height          =   735
      Left            =   1200
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4440
      Width           =   3255
   End
   Begin VB.Image Image1 
      Height          =   4305
      Left            =   480
      Picture         =   "vbproject7.frx":0000
      Top             =   120
      Width           =   4725
   End
End
Attribute VB_Name = "suspense7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Continue7_Click()
Results.Visible = True
suspense7.Visible = False

Open "M:/cs130/MeganKellyProject/" & "messages.txt" For Input As #1
    For k = 1 To 5
        Input #1, opponentwinmessage(k), opponentlosemessage(k)
    Next k
Close #1

Open "M:/cs130/MeganKellyProject/" & "namefactor.txt" For Input As #1
    For k = 1 To 5
        Input #1, opponentname(i), opponentfactor(i)
    Next k
Close #1

If rival = opponentname(i) Then
    If sum > opponentfactor(i) Then
        FinalPicture.Picture = "M:\cs130\MeganKellyProject\" & opponentlosepic(i)
        GoAgain.Caption = opponentlosemessage(i)
    ElseIf sum <= opponentfactor(i) Then
        FinalPicture.Picture = "M:\cs130\MeganKellyProject\" & opponentwinpic(i)
        GoAgain.Caption = opponentwinmessage(i)
    End If
End If

        
End Sub
