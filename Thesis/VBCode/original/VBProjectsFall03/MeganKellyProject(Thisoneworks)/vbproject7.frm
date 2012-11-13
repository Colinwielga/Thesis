VERSION 5.00
Begin VB.Form suspense7 
   BackColor       =   &H000000C0&
   Caption         =   "The suspense is building... can you feel it?"
   ClientHeight    =   5505
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5820
   LinkTopic       =   "Form7"
   ScaleHeight     =   5505
   ScaleWidth      =   5820
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
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      Caption         =   "Programmed by Megan Kelly"
      Height          =   255
      Left            =   3120
      TabIndex        =   1
      Top             =   5280
      Width           =   2655
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
' Which not so nice person would you like to beat up today? "Megan'sVBProject.vbp"
'                       Intro1 (VBProject4.frm)
'                       Megan Kelly 11/03/03
' Purpose:  The purpose of this form is to provide a break between the questions and the results of the user.
suspense7.Visible = False

Open "M:/cs130/MeganKellyProject/" & "namefactor.txt" For Input As #2
    For k = 1 To 5
        Input #2, opponentname(k), opponentfactor(k)
    Next k
Close #2

'Tells computer which screen to display next, based on whether the user has won or lost.

If rival = opponentname(1) Then
    If sum > opponentfactor(1) Then
        beatbush.Visible = True
        suspense7.Visible = False
    ElseIf sum <= opponentfactor(1) Then
        George_W_Bush.Visible = True
        suspense7.Visible = False
    End If
ElseIf rival = opponentname(2) Then
    If sum > opponentfactor(2) Then
        stalindead.Visible = True
        suspense7.Visible = False
    ElseIf sum <= opponentfactor(2) Then
        stalinwins.Visible = True
        suspense7.Visible = False
    End If
ElseIf rival = opponentname(3) Then
    If sum > opponentfactor(3) Then
        hitlerloses.Visible = True
        suspense7.Visible = False
    ElseIf sum <= opponentfactor(3) Then
        HItlerwins.Visible = True
        suspense7.Visible = False
    End If
ElseIf rival = opponentname(4) Then
    If sum > opponentfactor(4) Then
        Mansonloses.Visible = True
        suspense7.Visible = False
    ElseIf sum <= opponentfactor(4) Then
        mansonwins.Visible = True
        suspense7.Visible = False
    End If
ElseIf rival = opponentname(5) Then
    If sum > opponentfactor(5) Then
        Asscroftloses.Visible = True
        suspense7.Visible = False
    ElseIf sum <= opponentfactor(5) Then
        Ashcroftwins.Visible = True
        suspense7.Visible = False
    End If
End If

End Sub

