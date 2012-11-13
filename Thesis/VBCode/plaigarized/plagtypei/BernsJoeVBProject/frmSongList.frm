VERSION 5.00
Begin VB.Form frmSongList
   Caption         =   "Song List"
   ClientHeight    =   4245
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5790
   LinkTopic       =   "Form1"
   ScaleHeight     =   4245
   ScaleWidth      =   5790
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdStop
      Caption         =   "Stop"
      Height          =   495
      Left            =   2160
      TabIndex        =   3
      Top             =   3360
      Width           =   1335
   End
   Begin VB.FileListBox File1
      Height          =   2430
      Left            =   720
      TabIndex        =   2
      Top             =   480
      Width           =   4335
   End
   Begin VB.CommandButton cmdPlay
      Caption         =   "Play"
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   3360
      Width           =   1335
   End
   Begin VB.CommandButton cmdClose
      Caption         =   "Close"
      Height          =   495
      Left            =   3840
      TabIndex        =   0
      Top             =   3360
      Width           =   1335
   End
End
Attribute VB_Name = "frmSongList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim playing   As Boolean
Dim songFile As String

Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    frmKeyboard.Enabled = True
End Sub

Private Sub Form_Load()
    File1.Path =   App.Path & "\songs\"
    File1.Refresh

End Sub

Private Sub cmdPlay_Click()

    Dim pos As   Integer
    playing = True
    songFile =   File1.FileName
    ctr = 0

    If songFile <> "" Then
        cmdPlay.Enabled = False

        Open App.Path & "\songs\" & songFile For Input As #5

        Do While Not EOF(5)
            ctr = ctr + 1
            Input #5,   notePitch(ctr), noteDuration(ctr)
        Loop

        If ctr > 1 Then
                For pos = 1 To (ctr - 1)
                    If playing = True Then

                        keyNote =   PlaySound(App.Path & "\sounds\" & notePitch(pos) & ".wav", 0, 1)
                        Pause noteDuration(pos + 1) / 60
                    End If
                Next pos
            Pause 0.3

            keyNote = PlaySound(vbNullString, 0, 0)
        End If

        Close
        playing =   False
        cmdPlay.Enabled = True
    End If

End Sub

Private Sub cmdStop_Click()
    keyNote   = PlaySound(vbNullString, 0, 0)
    playing = False

    Close
    cmdPlay.Enabled = True
End Sub

Private Sub cmdClose_Click()
    frmSongList.Visible   = False
    frmKeyboard.Enabled = True
End Sub

