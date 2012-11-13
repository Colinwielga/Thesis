VERSION 5.00
Begin VB.Form Piano 
   BackColor       =   &H00004080&
   Caption         =   "Play some Piano!"
   ClientHeight    =   9090
   ClientLeft      =   1650
   ClientTop       =   1500
   ClientWidth     =   12165
   LinkTopic       =   "Form1"
   ScaleHeight     =   9090
   ScaleWidth      =   12165
   Visible         =   0   'False
   Begin VB.CommandButton KeyShow 
      Caption         =   "Songs played in this key."
      Height          =   1455
      Left            =   9600
      TabIndex        =   23
      Top             =   7440
      Width           =   2295
   End
   Begin VB.CommandButton Lid 
      Caption         =   "Shut the Lid"
      Height          =   1455
      Left            =   240
      TabIndex        =   22
      Top             =   7440
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Leave the piano."
      Height          =   1455
      Left            =   240
      TabIndex        =   21
      Top             =   7440
      Width           =   2295
   End
   Begin VB.CommandButton Bten 
      BackColor       =   &H80000007&
      Caption         =   "&P"
      Height          =   3015
      Left            =   10560
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   2160
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton Bnine 
      BackColor       =   &H80000007&
      Caption         =   "&O"
      Height          =   3015
      Left            =   9480
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   2160
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton Begin 
      Caption         =   "Chose a beginning note!"
      Height          =   1455
      Left            =   360
      TabIndex        =   18
      Top             =   360
      Width           =   11415
   End
   Begin VB.CommandButton Setup 
      Caption         =   "Show me the keyboard!"
      Height          =   1455
      Left            =   360
      TabIndex        =   17
      Top             =   360
      Width           =   11415
   End
   Begin VB.CommandButton Wnine 
      BackColor       =   &H80000009&
      Caption         =   "&L"
      Height          =   4935
      Left            =   9840
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   2160
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Beight 
      BackColor       =   &H80000007&
      Caption         =   "&I"
      Height          =   3015
      Left            =   8400
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   2160
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton Bseven 
      BackColor       =   &H80000007&
      Caption         =   "&U"
      Height          =   3015
      Left            =   7320
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   2160
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton Bsix 
      BackColor       =   &H80000007&
      Caption         =   "&Y"
      Height          =   3015
      Left            =   6240
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   2160
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton Bfive 
      BackColor       =   &H80000007&
      Caption         =   "&T"
      Height          =   3015
      Left            =   5160
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2160
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton Bfour 
      BackColor       =   &H80000007&
      Caption         =   "&R"
      Height          =   3015
      Left            =   4080
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2160
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton Bthree 
      BackColor       =   &H80000007&
      Caption         =   "&E"
      Height          =   3015
      Left            =   3000
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2160
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton Bone 
      BackColor       =   &H80000007&
      Caption         =   "&Q"
      Height          =   3015
      Left            =   840
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2160
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton Btwo 
      BackColor       =   &H80000007&
      Caption         =   "&W"
      Height          =   3015
      Left            =   1920
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2160
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton Wseven 
      BackColor       =   &H80000009&
      Caption         =   "&J"
      Height          =   4935
      Left            =   7680
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2160
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Wsix 
      BackColor       =   &H80000009&
      Caption         =   "&H"
      Height          =   4935
      Left            =   6600
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2160
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Wfive 
      BackColor       =   &H80000009&
      Caption         =   "&G"
      Height          =   4935
      Left            =   5520
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2160
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Wfour 
      BackColor       =   &H80000009&
      Caption         =   "&F"
      Height          =   4935
      Left            =   4440
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2160
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Wthree 
      BackColor       =   &H80000009&
      Caption         =   "&D"
      Height          =   4935
      Left            =   3360
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2160
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Wtwo 
      BackColor       =   &H80000009&
      Caption         =   "&S"
      Height          =   4935
      Left            =   2280
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2160
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Wone 
      BackColor       =   &H80000009&
      Caption         =   "&A"
      Height          =   4935
      Left            =   1200
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2160
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Weight 
      BackColor       =   &H80000009&
      Caption         =   "&K"
      Height          =   4935
      Left            =   8760
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2160
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   1095
   End
End
Attribute VB_Name = "Piano"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Palonzison Piano
'This is the Piano Form
'Matthew Peterson and Nicholas Alonzi are the authors of this Form
'This form was written in 2009 in the month of March
'The object of this form is to provide an interface which the user can play the chose the first note
    'That their keyboard will begin with.  It also allows them to play the keys and it produces the sound.
'This form is the form that plays music and is the central form which everything works around

Private Sub Begin_Click()
 Key = InputBox("What note do you want to start on?  Indicate sharp using 's' and flat 'b' right after the note name")
 Begin.Visible = False
 Setup.Visible = True
End Sub

Private Sub Beight_Click()
Dim j As Integer
    If Key = "Cs" Or Key = "Db" Then
        j = 30
    ElseIf Key = "D" Then
        j = 30
    ElseIf Key = "Ds" Or Key = "Eb" Then
        j = 31
    ElseIf Key = "E" Or Key = "Fb" Then
        j = 31
    ElseIf Key = "Fs" Or Key = "Gb" Then
        j = 27
    ElseIf Key = "G" Then
        j = 27
    ElseIf Key = "Gs" Or Key = "Ab" Then
        j = 28
    ElseIf Key = "A" Then
        j = 28
    ElseIf Key = "Bb" Or Key = "As" Then
        j = 29
    ElseIf Key = "B" Or Key = "Cb" Then
        j = 29
    End If
    
    PlaySound (Notes(j)), SND_ASYNC 'Plays the note and allows it to be played exactly when the button is pressed
    
End Sub

Private Sub Bfive_Click()
Dim j As Integer
    If Key = "C" Then
        j = 27
    ElseIf Key = "Cs" Or Key = "Db" Then
        j = 28
    ElseIf Key = "D" Then
        j = 28
    ElseIf Key = "Ds" Or Key = "Eb" Then
        j = 29
    ElseIf Key = "E" Or Key = "Fb" Then
        j = 29
    ElseIf Key = "Fs" Or Key = "Gb" Then
        j = 25
    ElseIf Key = "G" Then
        j = 25
    ElseIf Key = "Gs" Or Key = "Ab" Then
        j = 26
    ElseIf Key = "A" Then
        j = 26
    End If
    PlaySound (Notes(j)), SND_ASYNC
    
End Sub

Private Sub Bfour_Click()
Dim j As Integer

     If Key = "Cs" Or Key = "Db" Then
        j = 27
    ElseIf Key = "D" Then
        j = 27
    ElseIf Key = "Ds" Or Key = "Eb" Then
        j = 28
    ElseIf Key = "E" Or Key = "Fb" Then
        j = 28
    ElseIf Key = "F" Or Key = "Es" Then
        j = 29
    ElseIf Key = "Gs" Or Key = "Ab" Then
        j = 25
    ElseIf Key = "A" Then
        j = 25
    ElseIf Key = "Bb" Or Key = "As" Then
        j = 26
    ElseIf Key = "B" Or Key = "Cb" Then
        j = 26
    End If
    PlaySound (Notes(j)), SND_ASYNC
End Sub

Private Sub Bnine_Click()
Dim j As Integer
    If Key = "C" Then
        j = 30
    ElseIf Key = "Cs" Or Key = "Db" Then
        j = 31
    ElseIf Key = "D" Then
        j = 31
    ElseIf Key = "Ds" Or Key = "Eb" Then
        j = 32
    ElseIf Key = "E" Or Key = "Fb" Then
        j = 32
    ElseIf Key = "F" Or Key = "Es" Then
        j = 32
    ElseIf Key = "Fs" Or Key = "Gb" Then
        j = 28
    ElseIf Key = "G" Then
        j = 28
    ElseIf Key = "Gs" Or Key = "Ab" Then
        j = 29
    ElseIf Key = "A" Then
        j = 29
    ElseIf Key = "Bb" Or Key = "As" Then
        j = 30
    ElseIf Key = "B" Or Key = "Cb" Then
        j = 30
    End If
    
    PlaySound (Notes(j)), SND_ASYNC
End Sub

Private Sub Bten_Click()
Dim j As Integer
    If Key = "C" Then
        j = 31
    ElseIf Key = "Ds" Or Key = "Eb" Then
        j = 32
    ElseIf Key = "E" Or Key = "Fb" Then
        j = 32
    ElseIf Key = "F" Or Key = "Es" Then
        j = 14
    ElseIf Key = "Fs" Or Key = "Gb" Then
        j = 29
    ElseIf Key = "G" Then
        j = 29
    ElseIf Key = "Bb" Or Key = "As" Then
        j = 30
    ElseIf Key = "B" Or Key = "Cb" Then
        j = 30
    End If
    
    PlaySound (Notes(j)), SND_ASYNC
End Sub

Private Sub Wnine_Click()
Dim j As Integer
    If Key = "C" Then
        j = 16
    ElseIf Key = "Cs" Or Key = "Db" Then
        j = 17
    ElseIf Key = "D" Then
        j = 17
    ElseIf Key = "Ds" Or Key = "Eb" Then
        j = 18
    ElseIf Key = "E" Or Key = "Fb" Then
        j = 18
    ElseIf Key = "F" Or Key = "Es" Then
        j = 19
    ElseIf Key = "Fs" Or Key = "Gb" Then
        j = 13
    ElseIf Key = "G" Then
        j = 13
    ElseIf Key = "Gs" Or Key = "Ab" Then
        j = 14
    ElseIf Key = "A" Then
        j = 14
    ElseIf Key = "Bb" Or Key = "As" Then
        j = 15
    ElseIf Key = "B" Or Key = "Cb" Then
        j = 15
    End If
    
    PlaySound (Notes(j)), SND_ASYNC
End Sub



Private Sub Bone_Click()
Dim j As Integer
    If Key = "Cs" Or Key = "Db" Then
        j = 25
    ElseIf Key = "Ds" Or Key = "Eb" Then
        j = 26
    ElseIf Key = "Fs" Or Key = "Gb" Then
        j = 22
    ElseIf Key = "Gs" Or Key = "Ab" Then
        j = 23
    ElseIf Key = "Bb" Or Key = "As" Then
        j = 24
    End If
    
    PlaySound (Notes(j)), SND_ASYNC
End Sub

Private Sub Bseven_Click()
Dim j As Integer
    If Key = "C" Then
        j = 29
    ElseIf Key = "Ds" Or Key = "Eb" Then
        j = 30
    ElseIf Key = "E" Or Key = "Fb" Then
        j = 30
    ElseIf Key = "F" Or Key = "Es" Then
        j = 31
    ElseIf Key = "Gs" Or Key = "Ab" Then
        j = 27
    ElseIf Key = "A" Then
        j = 27
    ElseIf Key = "Bb" Or Key = "As" Then
        j = 28
    ElseIf Key = "B" Or Key = "Cb" Then
        j = 28
    End If
    
    PlaySound (Notes(j)), SND_ASYNC
End Sub

Private Sub Bsix_Click()
Dim j As Integer
    If Key = "C" Then
        j = 28
    ElseIf Key = "Cs" Or Key = "Db" Then
        j = 29
    ElseIf Key = "D" Then
        j = 29
    ElseIf Key = "F" Or Key = "Es" Then
        j = 30
    ElseIf Key = "Fs" Or Key = "Gb" Then
        j = 26
    ElseIf Key = "G" Then
        j = 26
    ElseIf Key = "Bb" Or Key = "As" Then
        j = 27
    ElseIf Key = "B" Or Key = "Cb" Then
        j = 27
    End If
    
    PlaySound (Notes(j)), SND_ASYNC
End Sub

Private Sub Bthree_Click()
Dim j As Integer

If Key = "C" Then
        j = 26
    ElseIf Key = "Ds" Or Key = "Eb" Then
        j = 27
    ElseIf Key = "E" Or Key = "Fb" Then
        j = 27
    ElseIf Key = "F" Or Key = "Es" Then
        j = 28
    ElseIf Key = "Fs" Or Key = "Gb" Then
        j = 24
    ElseIf Key = "G" Then
        j = 24
    ElseIf Key = "Bb" Or Key = "As" Then
        j = 25
    ElseIf Key = "B" Or Key = "Cb" Then
        j = 25
    End If
    
    PlaySound (Notes(j)), SND_ASYNC
End Sub

Private Sub Btwo_Click()
Dim j As Integer

If Key = "C" Then
        j = 25
    ElseIf Key = "Cs" Or Key = "Db" Then
        j = 26
    ElseIf Key = "D" Then
        j = 26
    ElseIf Key = "F" Or Key = "Es" Then
        j = 27
    ElseIf Key = "Fs" Or Key = "Gb" Then
        j = 23
    ElseIf Key = "G" Then
        j = 23
    ElseIf Key = "Gs" Or Key = "Ab" Then
        j = 24
    ElseIf Key = "A" Then
        j = 24
    End If
    
    PlaySound (Notes(j)), SND_ASYNC
End Sub

Private Sub Command1_Click()
    End
End Sub

Private Sub KeyShow_Click()
    Ctr = 0
    Open App.Path & ("\songs.txt") For Input As #1
    Do While Not EOF(1)
        Ctr = Ctr + 1
        Input #1, KeySig(Ctr), Piece(Ctr)
    Loop
    Close #1
    OtherPiece.Show
End Sub

Private Sub Lid_Click()
        Bone.Visible = False
        Btwo.Visible = False
        Bthree.Visible = False
        Bfour.Visible = False
        Bfive.Visible = False
        Bsix.Visible = False
        Bseven.Visible = False
        Beight.Visible = False
        Bnine.Visible = False
        Bten.Visible = False
        Wone.Visible = False
        Wtwo.Visible = False
        Wthree.Visible = False
        Wfour.Visible = False
        Wfive.Visible = False
        Wsix.Visible = False
        Wseven.Visible = False
        Weight.Visible = False
        Wnine.Visible = False
    Lid.Visible = False
    MsgBox ("Thank you for using Matt and Nick's Piano.")
End Sub

Private Sub Setup_Click()
    Wone.Visible = True
    Wtwo.Visible = True
    Wthree.Visible = True
    Wfour.Visible = True
    Wfive.Visible = True
    Wsix.Visible = True
    Wseven.Visible = True
    Weight.Visible = True
    Wnine.Visible = True
If Key = "C" Then
        Bone.Visible = False
        Bone.Enabled = False
        Btwo.Visible = True
        Bthree.Visible = True
        Bfour.Visible = False
        Bfour.Enabled = False
        Bfive.Visible = True
        Bsix.Visible = True
        Bseven.Visible = True
        Beight.Visible = False
        Beight.Enabled = False
        Bnine.Visible = True
        Bten.Visible = True
    ElseIf Key = "Cs" Or Key = "Db" Then
        Bone.Visible = True
        Btwo.Visible = True
        Bthree.Visible = False
        Bthree.Enabled = False
        Bfour.Visible = True
        Bfive.Visible = True
        Bsix.Visible = True
        Bseven.Visible = False
        Bseven.Enabled = False
        Beight.Visible = True
        Bnine.Visible = True
        Bten.Visible = False
        Bten.Enabled = False
    ElseIf Key = "D" Then
        Bone.Visible = False
        Bone.Enabled = False
        Btwo.Visible = True
        Bthree.Visible = False
        Bthree.Enabled = False
        Bfour.Visible = True
        Bfive.Visible = True
        Bsix.Visible = True
        Bseven.Visible = False
        Bseven.Enabled = False
        Beight.Visible = True
        Bnine.Visible = True
        Bten.Visible = False
        Bten.Enabled = False
    ElseIf Key = "Ds" Or Key = "Eb" Then
        Bone.Visible = True
        Btwo.Visible = False
        Btwo.Enabled = False
        Bthree.Visible = True
        Bfour.Visible = True
        Bfive.Visible = True
        Bsix.Visible = False
        Bsix.Enabled = False
        Bseven.Visible = True
        Beight.Visible = True
        Bnine.Visible = False
        Bnine.Enabled = False
        Bten.Visible = True
    ElseIf Key = "E" Or Key = "Fb" Then
        Bone.Visible = False
        Bone.Enabled = False
        Btwo.Visible = False
        Btwo.Enabled = False
        Bthree.Visible = True
        Bfour.Visible = True
        Bfive.Visible = True
        Bsix.Visible = False
        Bsix.Visible = False
        Bseven.Visible = True
        Beight.Visible = True
        Bnine.Visible = False
        Bnine.Enabled = False
        Bten.Visible = True
    ElseIf Key = "F" Or Key = "Es" Then
        Bone.Visible = False
        Bone.Enabled = False
        Btwo.Visible = True
        Bthree.Visible = True
        Bfour.Visible = True
        Bfive.Visible = False
        Bfive.Enabled = False
        Bsix.Visible = True
        Bseven.Visible = True
        Beight.Visible = False
        Beight.Enabled = False
        Bnine.Visible = True
        Bten.Visible = True
    ElseIf Key = "Fs" Or Key = "Gb" Then
        Bone.Visible = True
        Btwo.Visible = True
        Bthree.Visible = True
        Bfour.Visible = False
        Bfour.Enabled = False
        Bfive.Visible = True
        Bsix.Visible = True
        Bseven.Visible = False
        Bseven.Enabled = False
        Beight.Visible = True
        Bnine.Visible = True
        Bten.Visible = True
    ElseIf Key = "G" Then
        Bone.Visible = False
        Bone.Enabled = False
        Btwo.Visible = True
        Bthree.Visible = True
        Bfour.Visible = False
        Bfour.Enabled = False
        Bfive.Visible = True
        Bsix.Visible = True
        Bseven.Visible = False
        Bseven.Enabled = False
        Beight.Visible = True
        Bnine.Visible = True
    ElseIf Key = "Gs" Or Key = "Ab" Then
        Bone.Visible = True
        Btwo.Visible = True
        Bthree.Visible = False
        Bthree.Enabled = False
        Bfour.Visible = True
        Bfive.Visible = True
        Bsix.Visible = False
        Bsix.Enabled = False
        Bseven.Visible = True
        Beight.Visible = True
        Bnine.Visible = True
        Bten.Visible = False
        Bten.Enabled = False
    ElseIf Key = "A" Then
        Bone.Visible = False
        Bone.Enabled = False
        Btwo.Visible = True
        Bthree.Visible = False
        Bthree.Enabled = False
        Bfour.Visible = True
        Bfive.Visible = True
        Bsix.Visible = False
        Bsix.Enabled = False
        Bseven.Visible = True
        Beight.Visible = True
        Bnine.Visible = True
        Bten.Visible = False
        Bten.Enabled = False
    ElseIf Key = "Bb" Or Key = "As" Then
        Bone.Visible = True
        Btwo.Visible = False
        Bthree.Visible = True
        Bfour.Visible = True
        Bfive.Visible = False
        Bsix.Visible = True
        Bseven.Visible = True
        Beight.Visible = True
        Bnine.Visible = False
        Bten.Visible = True
    ElseIf Key = "B" Or Key = "Cb" Then
        Bone.Visible = False
        Btwo.Visible = False
        Bthree.Visible = True
        Bfour.Visible = True
        Bfive.Visible = False
        Bsix.Visible = True
        Bseven.Visible = True
        Beight.Visible = True
        Bnine.Visible = False
        Bten.Visible = True
    Else
        MsgBox ("The starting not you gave.  Please try again.")
    End If
    Setup.Visible = False
    Begin.Visible = True
End Sub

Private Sub Weight_Click()
Dim j As Integer
    If Key = "C" Then
        j = 15
    ElseIf Key = "Cs" Or Key = "Db" Then
        j = 16
    ElseIf Key = "D" Then
        j = 16
    ElseIf Key = "Ds" Or Key = "Eb" Then
        j = 17
    ElseIf Key = "E" Or Key = "Fb" Then
        j = 17
    ElseIf Key = "F" Or Key = "Es" Then
        j = 18
    ElseIf Key = "Fs" Or Key = "Gb" Then
        j = 12
    ElseIf Key = "G" Then
        j = 12
    ElseIf Key = "Gs" Or Key = "Ab" Then
        j = 13
    ElseIf Key = "A" Then
        j = 13
    ElseIf Key = "Bb" Or Key = "As" Then
        j = 14
    ElseIf Key = "B" Or Key = "Cb" Then
        j = 14
    End If
    
    PlaySound (Notes(j)), SND_ASYNC
End Sub

Private Sub Wfive_Click()
Dim j As Integer
    If Key = "C" Then
        j = 12
    ElseIf Key = "Cs" Or Key = "Db" Then
        j = 13
    ElseIf Key = "D" Then
        j = 13
    ElseIf Key = "Ds" Or Key = "Eb" Then
        j = 14
    ElseIf Key = "E" Or Key = "Fb" Then
        j = 14
    ElseIf Key = "F" Or Key = "Es" Then
        j = 15
    ElseIf Key = "Fs" Or Key = "Gb" Then
        j = 9
    ElseIf Key = "G" Then
        j = 9
    ElseIf Key = "Gs" Or Key = "Ab" Then
        j = 10
    ElseIf Key = "A" Then
        j = 10
    ElseIf Key = "Bb" Or Key = "As" Then
        j = 11
    ElseIf Key = "B" Or Key = "Cb" Then
        j = 11
    End If
    
    PlaySound (Notes(j)), SND_ASYNC
End Sub

Private Sub Wfour_Click()
Dim j As Integer
    If Key = "C" Then
        j = 11
    ElseIf Key = "Cs" Or Key = "Db" Then
        j = 12
    ElseIf Key = "D" Then
        j = 12
    ElseIf Key = "Ds" Or Key = "Eb" Then
        j = 13
    ElseIf Key = "E" Or Key = "Fb" Then
        j = 13
    ElseIf Key = "F" Or Key = "Es" Then
        j = 14
    ElseIf Key = "Fs" Or Key = "Gb" Then
        j = 8
    ElseIf Key = "G" Then
        j = 8
    ElseIf Key = "Gs" Or Key = "Ab" Then
        j = 9
    ElseIf Key = "A" Then
        j = 9
    ElseIf Key = "Bb" Or Key = "As" Then
        j = 10
    ElseIf Key = "B" Or Key = "Cb" Then
        j = 10
    End If
    
    PlaySound (Notes(j)), SND_ASYNC
End Sub



Private Sub Wone_Click()
Dim j As Integer

    Select Case Key
        Case Is = "C"
            j = 8
            PlaySound (Notes(8)), SND_ASYNC
        Case Is = "Cs", "Db"
            j = 9
            PlaySound (Notes(9)), SND_ASYNC
        Case Is = "D"
            j = 9
            PlaySound (Notes(9)), SND_ASYNC
        Case Is = "Ds", "Eb"
            j = 10
            PlaySound (Notes(10)), SND_ASYNC
        Case Is = "E", "Fb"
            j = 10
            PlaySound (Notes(10)), SND_ASYNC
        Case Is = "F", "Es"
            j = 11
            PlaySound (Notes(11)), SND_ASYNC
        Case Is = "Fs", "Gb"
            j = 5
            PlaySound (Notes(5)), SND_ASYNC
        Case Is = "G"
            j = 5
            PlaySound (Notes(5)), SND_ASYNC
        Case Is = "Gs", "Ab"
            j = 6
            PlaySound (Notes(6)), SND_ASYNC
        Case Is = "A"
            j = 6
            PlaySound (Notes(6)), SND_ASYNC
        Case Is = "Bb", "As"
            j = 7
            PlaySound (Notes(7)), SND_ASYNC
        Case Is = "B", "Cb"
            j = 7
            PlaySound (Notes(7)), SND_ASYNC
    End Select

End Sub

Private Sub Wseven_Click()
Dim j As Integer
    If Key = "C" Then
        j = 14
    ElseIf Key = "Cs" Or Key = "Db" Then
        j = 15
    ElseIf Key = "D" Then
        j = 15
    ElseIf Key = "Ds" Or Key = "Eb" Then
        j = 16
    ElseIf Key = "E" Or Key = "Fb" Then
        j = 16
    ElseIf Key = "F" Or Key = "Es" Then
        j = 17
    ElseIf Key = "Fs" Or Key = "Gb" Then
        j = 11
    ElseIf Key = "G" Then
        j = 11
    ElseIf Key = "Gs" Or Key = "Ab" Then
        j = 12
    ElseIf Key = "A" Then
        j = 12
    ElseIf Key = "Bb" Or Key = "As" Then
        j = 13
    ElseIf Key = "B" Or Key = "Cb" Then
        j = 13
    End If
    
    PlaySound (Notes(j)), SND_ASYNC
End Sub

Private Sub Wsix_Click()
Dim j As Integer
    If Key = "C" Then
        j = 13
    ElseIf Key = "Cs" Or Key = "Db" Then
        j = 14
    ElseIf Key = "D" Then
        j = 14
    ElseIf Key = "Ds" Or Key = "Eb" Then
        j = 15
    ElseIf Key = "E" Or Key = "Fb" Then
        j = 15
    ElseIf Key = "F" Or Key = "Es" Then
        j = 16
    ElseIf Key = "Fs" Or Key = "Gb" Then
        j = 10
    ElseIf Key = "G" Then
        j = 10
    ElseIf Key = "Gs" Or Key = "Ab" Then
        j = 11
    ElseIf Key = "A" Then
        j = 11
    ElseIf Key = "Bb" Or Key = "As" Then
        j = 12
    ElseIf Key = "B" Or Key = "Cb" Then
        j = 12
    End If
    
    PlaySound (Notes(j)), SND_ASYNC
End Sub

Private Sub Wthree_Click()
Dim j As Integer
    If Key = "C" Then
        j = 10
    ElseIf Key = "Cs" Or Key = "Db" Then
        j = 11
    ElseIf Key = "D" Then
        j = 11
    ElseIf Key = "Ds" Or Key = "Eb" Then
        j = 12
    ElseIf Key = "E" Or Key = "Fb" Then
        j = 12
    ElseIf Key = "F" Or Key = "Es" Then
        j = 13
    ElseIf Key = "Fs" Or Key = "Gb" Then
        j = 7
    ElseIf Key = "G" Then
        j = 7
    ElseIf Key = "Gs" Or Key = "Ab" Then
        j = 8
    ElseIf Key = "A" Then
        j = 8
    ElseIf Key = "Bb" Or Key = "As" Then
        j = 9
    ElseIf Key = "B" Or Key = "Cb" Then
        j = 9
    End If
    
    PlaySound (Notes(j)), SND_ASYNC
End Sub

Private Sub Wtwo_Click()
Dim j As Integer
    If Key = "C" Then
        j = 9
    ElseIf Key = "Cs" Or Key = "Db" Then
        j = 10
    ElseIf Key = "D" Then
        j = 10
    ElseIf Key = "Ds" Or Key = "Eb" Then
        j = 11
    ElseIf Key = "E" Or Key = "Fb" Then
        j = 11
    ElseIf Key = "F" Or Key = "Es" Then
        j = 12
    ElseIf Key = "Fs" Or Key = "Gb" Then
        j = 6
    ElseIf Key = "G" Then
        j = 6
    ElseIf Key = "Gs" Or Key = "Ab" Then
        j = 7
    ElseIf Key = "A" Then
        j = 7
    ElseIf Key = "Bb" Or Key = "As" Then
        j = 8
    ElseIf Key = "B" Or Key = "Cb" Then
        j = 8
    End If
    
    PlaySound (Notes(j)), SND_ASYNC
End Sub
