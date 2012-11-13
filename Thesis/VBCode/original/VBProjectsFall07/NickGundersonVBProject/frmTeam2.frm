VERSION 5.00
Begin VB.Form frmTeam2 
   Caption         =   "Form1"
   ClientHeight    =   8505
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11235
   LinkTopic       =   "Form1"
   Picture         =   "frmTeam2.frx":0000
   ScaleHeight     =   8505
   ScaleWidth      =   11235
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Bodoni MT Black"
         Size            =   20.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   5760
      TabIndex        =   16
      Top             =   5760
      Width           =   3615
   End
   Begin VB.CommandButton cmdEnter 
      Caption         =   "Enter Stats"
      BeginProperty Font 
         Name            =   "Bodoni MT Black"
         Size            =   20.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   1440
      TabIndex        =   15
      Top             =   5760
      Width           =   3615
   End
   Begin VB.CommandButton cmd14 
      Caption         =   "Command14"
      BeginProperty Font 
         Name            =   "Bodoni MT Black"
         Size            =   8.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5760
      TabIndex        =   13
      Top             =   4920
      Width           =   2415
   End
   Begin VB.CommandButton cmd13 
      Caption         =   "Command13"
      BeginProperty Font 
         Name            =   "Bodoni MT Black"
         Size            =   8.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2640
      TabIndex        =   12
      Top             =   4920
      Width           =   2415
   End
   Begin VB.CommandButton cmd12 
      Caption         =   "Command12"
      BeginProperty Font 
         Name            =   "Bodoni MT Black"
         Size            =   8.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7320
      TabIndex        =   11
      Top             =   4080
      Width           =   2415
   End
   Begin VB.CommandButton cmd11 
      Caption         =   "Command11"
      BeginProperty Font 
         Name            =   "Bodoni MT Black"
         Size            =   8.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4200
      TabIndex        =   10
      Top             =   4080
      Width           =   2415
   End
   Begin VB.CommandButton cmd10 
      Caption         =   "Command10"
      BeginProperty Font 
         Name            =   "Bodoni MT Black"
         Size            =   8.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   9
      Top             =   4080
      Width           =   2415
   End
   Begin VB.CommandButton cmd9 
      Caption         =   "Command9"
      BeginProperty Font 
         Name            =   "Bodoni MT Black"
         Size            =   8.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7320
      TabIndex        =   8
      Top             =   3240
      Width           =   2415
   End
   Begin VB.CommandButton cmd8 
      Caption         =   "Command8"
      BeginProperty Font 
         Name            =   "Bodoni MT Black"
         Size            =   8.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4200
      TabIndex        =   7
      Top             =   3240
      Width           =   2415
   End
   Begin VB.CommandButton cmd7 
      Caption         =   "Command7"
      BeginProperty Font 
         Name            =   "Bodoni MT Black"
         Size            =   8.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   6
      Top             =   3240
      Width           =   2415
   End
   Begin VB.CommandButton cmd6 
      Caption         =   "Command6"
      BeginProperty Font 
         Name            =   "Bodoni MT Black"
         Size            =   8.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7320
      TabIndex        =   5
      Top             =   2400
      Width           =   2415
   End
   Begin VB.CommandButton cmd5 
      Caption         =   "Command5"
      BeginProperty Font 
         Name            =   "Bodoni MT Black"
         Size            =   8.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4200
      TabIndex        =   4
      Top             =   2400
      Width           =   2415
   End
   Begin VB.CommandButton cmd4 
      Caption         =   "Command4"
      BeginProperty Font 
         Name            =   "Bodoni MT Black"
         Size            =   8.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   3
      Top             =   2400
      Width           =   2415
   End
   Begin VB.CommandButton cmd3 
      Caption         =   "Command3"
      BeginProperty Font 
         Name            =   "Bodoni MT Black"
         Size            =   8.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7320
      TabIndex        =   2
      Top             =   1560
      Width           =   2415
   End
   Begin VB.CommandButton cmd2 
      Caption         =   "Command2"
      BeginProperty Font 
         Name            =   "Bodoni MT Black"
         Size            =   8.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4200
      TabIndex        =   1
      Top             =   1560
      Width           =   2415
   End
   Begin VB.CommandButton cmd1 
      Caption         =   "Command1"
      BeginProperty Font 
         Name            =   "Bodoni MT Black"
         Size            =   8.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   0
      Top             =   1560
      Width           =   2415
   End
   Begin VB.Label lblstarters 
      Alignment       =   2  'Center
      Caption         =   "Please Choose your starters"
      BeginProperty Font 
         Name            =   "Bodoni MT Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      TabIndex        =   17
      Top             =   960
      Width           =   8175
   End
   Begin VB.Label lblteam2 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Bodoni MT Black"
         Size            =   26.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1080
      TabIndex        =   14
      Top             =   0
      Width           =   8655
   End
End
Attribute VB_Name = "frmTeam2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'for notes on this screen please veiw frmstarters

Private Sub cmd1_Click()
cmd1.Enabled = False
If Pos2(1) = "QB" Then
    If QB = 0 Then
        QB = 1
    Else
        MsgBox "Sorry you have already Choosen a QB"
    End If
Else
    If Pos2(1) = "RB" Then
        If RB1 = 0 Then
            RB1 = 1
        Else
            If RB2 = 0 Then
                RB2 = 1
            Else
                If WRRB = 0 Then
                    WRRB = 1
                Else
                    MsgBox ("Sorry you have already started 2 RB's and Your WR/RB")
                End If
            End If
        End If
    Else
        If Pos2(1) = "WR" Then
          If WR1 = 0 Then
                WR1 = 1
            Else
                If WR2 = 0 Then
                    WR2 = 1
                Else
                    If WRRB = 0 Then
                        WRRB = 1
                    Else
                        MsgBox ("Sorry you have already started 2 WR's and Your WR/RB")
                    End If
                End If
            End If
        Else
            If Pos2(1) = "TE" Then
                If TE = 0 Then
                    TE = 1
                Else
                    MsgBox "You have already started a TE"
                End If
            Else
                If Pos2(1) = "K" Then
                    If K = 0 Then
                        K = 1
                    Else
                        MsgBox "Sorry You have already started a Kicker"
                    End If
                Else
                    If Pos2(1) = "DEF" Then
                        If Def = 0 Then
                            Def = 1
                        Else
                            MsgBox "Sorry You have already started a Defense"
                        End If
                    End If
                End If
            End If
        End If
    End If
End If

        
End Sub

Private Sub cmd10_Click()
cmd10.Enabled = False
If Pos2(10) = "QB" Then
    If QB = 0 Then
        QB = 10
    Else
        MsgBox "Sorry you have already Choosen a QB"
    End If
Else
    If Pos2(10) = "RB" Then
        If RB1 = 0 Then
            RB1 = 10
        Else
            If RB2 = 0 Then
                RB2 = 10
            Else
                If WRRB = 0 Then
                    WRRB = 10
                Else
                    MsgBox ("Sorry you have already started 2 RB's and Your WR/RB")
                End If
            End If
        End If
    Else
        If Pos2(10) = "WR" Then
          If WR1 = 0 Then
                WR1 = 10
            Else
                If WR2 = 0 Then
                    WR2 = 10
                Else
                    If WRRB = 0 Then
                        WRRB = 10
                    Else
                        MsgBox ("Sorry you have already started 2 WR's and Your WR/RB")
                    End If
                End If
            End If
        Else
            If Pos2(10) = "TE" Then
                If TE = 0 Then
                    TE = 10
                Else
                MsgBox "You have already started a TE"
                End If
            Else
                If Pos2(10) = "K" Then
                    If K = 0 Then
                        K = 10
                    Else
                        MsgBox "Sorry You have already started a Kicker"
                    End If
                Else
                    If Pos2(10) = "DEF" Then
                        If Def = 0 Then
                            Def = 10
                        Else
                            MsgBox "Sorry You have already started a Defense"
                        End If
                    End If
                End If
            End If
        End If
    End If
End If

End Sub

Private Sub cmd11_Click()
cmd11.Enabled = False
If Pos2(11) = "QB" Then
    If QB = 0 Then
        QB = 11
    Else
        MsgBox "Sorry you have already Choosen a QB"
    End If
Else
    If Pos2(11) = "RB" Then
        If RB1 = 0 Then
            RB1 = 11
        Else
            If RB2 = 0 Then
                RB2 = 11
            Else
                If WRRB = 0 Then
                    WRRB = 11
                Else
                    MsgBox ("Sorry you have already started 2 RB's and Your WR/RB")
                End If
            End If
        End If
    Else
        If Pos2(11) = "WR" Then
          If WR1 = 0 Then
                WR1 = 11
            Else
                If WR2 = 0 Then
                    WR2 = 11
                Else
                    If WRRB = 0 Then
                        WRRB = 11
                    Else
                        MsgBox ("Sorry you have already started 2 WR's and Your WR/RB")
                    End If
                End If
            End If
        Else
            If Pos2(11) = "TE" Then
                If TE = 0 Then
                    TE = 11
                Else
                MsgBox "You have already started a TE"
                End If
            Else
                If Pos2(11) = "K" Then
                    If K = 0 Then
                        K = 11
                    Else
                        MsgBox "Sorry You have already started a Kicker"
                    End If
                Else
                    If Pos2(11) = "DEF" Then
                        If Def = 0 Then
                            Def = 11
                        Else
                            MsgBox "Sorry You have already started a Defense"
                        End If
                    End If
                End If
            End If
        End If
    End If
End If

End Sub

Private Sub cmd12_Click()
cmd12.Enabled = False
If Pos2(12) = "QB" Then
    If QB = 0 Then
        QB = 12
    Else
        MsgBox "Sorry you have already Choosen a QB"
    End If
Else
    If Pos2(12) = "RB" Then
        If RB1 = 0 Then
            RB1 = 12
        Else
            If RB2 = 0 Then
                RB2 = 12
            Else
                If WRRB = 0 Then
                    WRRB = 12
                Else
                    MsgBox ("Sorry you have already started 2 RB's and Your WR/RB")
                End If
            End If
        End If
    Else
        If Pos2(12) = "WR" Then
          If WR1 = 0 Then
                WR1 = 12
            Else
                If WR2 = 0 Then
                    WR2 = 12
                Else
                    If WRRB = 0 Then
                        WRRB = 12
                    Else
                        MsgBox ("Sorry you have already started 2 WR's and Your WR/RB")
                    End If
                End If
            End If
        Else
            If Pos2(12) = "TE" Then
                If TE = 0 Then
                    TE = 12
                Else
                MsgBox "You have already started a TE"
                End If
            Else
                If Pos2(12) = "K" Then
                    If K = 0 Then
                        K = 12
                    Else
                        MsgBox "Sorry You have already started a Kicker"
                    End If
                Else
                    If Pos2(12) = "DEF" Then
                        If Def = 0 Then
                            Def = 12
                        Else
                            MsgBox "Sorry You have already started a Defense"
                        End If
                    End If
                End If
            End If
        End If
    End If
End If

End Sub

Private Sub cmd13_Click()
cmd13.Enabled = False
If Pos2(13) = "QB" Then
    If QB = 0 Then
        QB = 13
    Else
        MsgBox "Sorry you have already Choosen a QB"
    End If
Else
    If Pos2(13) = "RB" Then
        If RB1 = 0 Then
            RB1 = 13
        Else
            If RB2 = 0 Then
                RB2 = 13
            Else
                If WRRB = 0 Then
                    WRRB = 13
                Else
                    MsgBox ("Sorry you have already started 2 RB's and Your WR/RB")
                End If
            End If
        End If
    Else
        If Pos2(13) = "WR" Then
          If WR1 = 0 Then
                WR1 = 13
            Else
                If WR2 = 0 Then
                    WR2 = 13
                Else
                    If WRRB = 0 Then
                        WRRB = 13
                    Else
                        MsgBox ("Sorry you have already started 2 WR's and Your WR/RB")
                    End If
                End If
            End If
        Else
            If Pos2(13) = "TE" Then
                If TE = 0 Then
                    TE = 13
                Else
                MsgBox "You have already started a TE"
                End If
            Else
                If Pos2(13) = "K" Then
                    If K = 0 Then
                        K = 13
                    Else
                        MsgBox "Sorry You have already started a Kicker"
                    End If
                Else
                    If Pos2(13) = "DEF" Then
                        If Def = 0 Then
                            Def = 13
                        Else
                            MsgBox "Sorry You have already started a Defense"
                        End If
                    End If
                End If
            End If
        End If
    End If
End If

End Sub

Private Sub cmd14_Click()
cmd14.Enabled = False
If Pos2(14) = "QB" Then
    If QB = 0 Then
        QB = 14
    Else
        MsgBox "Sorry you have already Choosen a QB"
    End If
Else
    If Pos2(14) = "RB" Then
        If RB1 = 0 Then
            RB1 = 14
        Else
            If RB2 = 0 Then
                RB2 = 14
            Else
                If WRRB = 0 Then
                    WRRB = 14
                Else
                    MsgBox ("Sorry you have already started 2 RB's and Your WR/RB")
                End If
            End If
        End If
    Else
        If Pos2(14) = "WR" Then
          If WR1 = 0 Then
                WR1 = 14
            Else
                If WR2 = 0 Then
                    WR2 = 14
                Else
                    If WRRB = 0 Then
                        WRRB = 14
                    Else
                        MsgBox ("Sorry you have already started 2 WR's and Your WR/RB")
                    End If
                End If
            End If
        Else
            If Pos2(14) = "TE" Then
                If TE = 0 Then
                    TE = 14
                Else
                MsgBox "You have already started a TE"
                End If
            Else
                If Pos2(14) = "K" Then
                    If K = 0 Then
                        K = 14
                    Else
                        MsgBox "Sorry You have already started a Kicker"
                    End If
                Else
                    If Pos2(14) = "DEF" Then
                        If Def = 0 Then
                            Def = 14
                        Else
                            MsgBox "Sorry You have already started a Defense"
                        End If
                    End If
                End If
            End If
        End If
    End If
End If

End Sub

Private Sub cmd2_Click()
cmd2.Enabled = False
If Pos2(2) = "QB" Then
    If QB = 0 Then
        QB = 2
    Else
        MsgBox "Sorry you have already Choosen a QB"
    End If
Else
    If Pos2(2) = "RB" Then
        If RB1 = 0 Then
            RB1 = 2
        Else
            If RB2 = 0 Then
                RB2 = 2
            Else
                If WRRB = 0 Then
                    WRRB = 2
                Else
                    MsgBox ("Sorry you have already started 2 RB's and Your WR/RB")
                End If
            End If
        End If
    Else
        If Pos2(2) = "WR" Then
          If WR1 = 0 Then
                WR1 = 2
            Else
                If WR2 = 0 Then
                    WR2 = 2
                Else
                    If WRRB = 0 Then
                        WRRB = 2
                    Else
                        MsgBox ("Sorry you have already started 2 WR's and Your WR/RB")
                    End If
                End If
            End If
        Else
            If Pos2(2) = "TE" Then
                If TE = 0 Then
                    TE = 2
                Else
                MsgBox "You have already started a TE"
                End If
            Else
                If Pos2(2) = "K" Then
                    If K = 0 Then
                        K = 2
                    Else
                        MsgBox "Sorry You have already started a Kicker"
                    End If
                Else
                    If Pos2(2) = "DEF" Then
                        If Def = 0 Then
                            Def = 2
                        Else
                            MsgBox "Sorry You have already started a Defense"
                        End If
                    End If
                End If
            End If
        End If
    End If
End If

End Sub

Private Sub cmd3_Click()
cmd3.Enabled = False
If Pos2(3) = "QB" Then
    If QB = 0 Then
        QB = 3
    Else
        MsgBox "Sorry you have already Choosen a QB"
    End If
Else
    If Pos2(3) = "RB" Then
        If RB1 = 0 Then
            RB1 = 3
        Else
            If RB2 = 0 Then
                RB2 = 3
            Else
                If WRRB = 0 Then
                    WRRB = 3
                Else
                    MsgBox ("Sorry you have already started 2 RB's and Your WR/RB")
                End If
            End If
        End If
    Else
        If Pos2(3) = "WR" Then
          If WR1 = 0 Then
                WR1 = 3
            Else
                If WR2 = 0 Then
                    WR2 = 3
                Else
                    If WRRB = 0 Then
                        WRRB = 3
                    Else
                        MsgBox ("Sorry you have already started 2 WR's and Your WR/RB")
                    End If
                End If
            End If
        Else
            If Pos2(3) = "TE" Then
                If TE = 0 Then
                    TE = 3
                Else
                MsgBox "You have already started a TE"
                End If
            Else
                If Pos2(3) = "K" Then
                    If K = 0 Then
                        K = 3
                    Else
                        MsgBox "Sorry You have already started a Kicker"
                    End If
                Else
                    If Pos2(3) = "DEF" Then
                        If Def = 0 Then
                            Def = 3
                        Else
                            MsgBox "Sorry You have already started a Defense"
                        End If
                    End If
                End If
            End If
        End If
    End If
End If

End Sub

Private Sub cmd4_Click()
cmd4.Enabled = False
If Pos2(4) = "QB" Then
    If QB = 0 Then
        QB = 4
    Else
        MsgBox "Sorry you have already Choosen a QB"
    End If
Else
    If Pos2(4) = "RB" Then
        If RB1 = 0 Then
            RB1 = 4
        Else
            If RB2 = 0 Then
                RB2 = 4
            Else
                If WRRB = 0 Then
                    WRRB = 4
                Else
                    MsgBox ("Sorry you have already started 2 RB's and Your WR/RB")
                End If
            End If
        End If
    Else
        If Pos2(4) = "WR" Then
          If WR1 = 0 Then
                WR1 = 4
            Else
                If WR2 = 0 Then
                    WR2 = 4
                Else
                    If WRRB = 0 Then
                        WRRB = 4
                    Else
                        MsgBox ("Sorry you have already started 2 WR's and Your WR/RB")
                    End If
                End If
            End If
        Else
            If Pos2(4) = "TE" Then
                If TE = 0 Then
                    TE = 4
                Else
                MsgBox "You have already started a TE"
                End If
            Else
                If Pos2(4) = "K" Then
                    If K = 0 Then
                        K = 4
                    Else
                        MsgBox "Sorry You have already started a Kicker"
                    End If
                Else
                    If Pos2(4) = "DEF" Then
                        If Def = 0 Then
                            Def = 4
                        Else
                            MsgBox "Sorry You have already started a Defense"
                        End If
                    End If
                End If
            End If
        End If
    End If
End If

End Sub

Private Sub cmd5_Click()
cmd5.Enabled = False
If Pos2(5) = "QB" Then
    If QB = 0 Then
        QB = 5
    Else
        MsgBox "Sorry you have already Choosen a QB"
    End If
Else
    If Pos2(5) = "RB" Then
        If RB1 = 0 Then
            RB1 = 5
        Else
            If RB2 = 0 Then
                RB2 = 5
            Else
                If WRRB = 0 Then
                    WRRB = 5
                Else
                    MsgBox ("Sorry you have already started 2 RB's and Your WR/RB")
                End If
            End If
        End If
    Else
        If Pos2(5) = "WR" Then
          If WR1 = 0 Then
                WR1 = 5
            Else
                If WR2 = 0 Then
                    WR2 = 5
                Else
                    If WRRB = 0 Then
                        WRRB = 5
                    Else
                        MsgBox ("Sorry you have already started 2 WR's and Your WR/RB")
                    End If
                End If
            End If
        Else
            If Pos2(5) = "TE" Then
                If TE = 0 Then
                    TE = 5
                Else
                MsgBox "You have already started a TE"
                End If
            Else
                If Pos2(5) = "K" Then
                    If K = 0 Then
                        K = 5
                    Else
                        MsgBox "Sorry You have already started a Kicker"
                    End If
                Else
                    If Pos2(5) = "DEF" Then
                        If Def = 0 Then
                            Def = 5
                        Else
                            MsgBox "Sorry You have already started a Defense"
                        End If
                    End If
                End If
            End If
        End If
    End If
End If

End Sub

Private Sub cmd6_Click()
cmd6.Enabled = False
If Pos2(6) = "QB" Then
    If QB = 0 Then
        QB = 6
    Else
        MsgBox "Sorry you have already Choosen a QB"
    End If
Else
    If Pos2(6) = "RB" Then
        If RB1 = 0 Then
            RB1 = 6
        Else
            If RB2 = 0 Then
                RB2 = 6
            Else
                If WRRB = 0 Then
                    WRRB = 6
                Else
                    MsgBox ("Sorry you have already started 2 RB's and Your WR/RB")
                End If
            End If
        End If
    Else
        If Pos2(6) = "WR" Then
          If WR1 = 0 Then
                WR1 = 6
            Else
                If WR2 = 0 Then
                    WR2 = 6
                Else
                    If WRRB = 0 Then
                        WRRB = 6
                    Else
                        MsgBox ("Sorry you have already started 2 WR's and Your WR/RB")
                    End If
                End If
            End If
        Else
            If Pos2(6) = "TE" Then
                If TE = 0 Then
                    TE = 6
                Else
                    MsgBox "You have already started a TE"
                End If
            Else
                If Pos2(6) = "K" Then
                    If K = 0 Then
                        K = 6
                    Else
                        MsgBox "Sorry You have already started a Kicker"
                    End If
                Else
                    If Pos2(6) = "DEF" Then
                        If Def = 0 Then
                            Def = 6
                        Else
                            MsgBox "Sorry You have already started a Defense"
                        End If
                    End If
                End If
            End If
        End If
    End If
End If

End Sub

Private Sub cmd7_Click()
cmd7.Enabled = False
If Pos2(7) = "QB" Then
    If QB = 0 Then
        QB = 7
    Else
        MsgBox "Sorry you have already Choosen a QB"
    End If
Else
    If Pos2(7) = "RB" Then
        If RB1 = 0 Then
            RB1 = 7
        Else
            If RB2 = 0 Then
                RB2 = 7
            Else
                If WRRB = 0 Then
                    WRRB = 7
                Else
                    MsgBox ("Sorry you have already started 2 RB's and Your WR/RB")
                End If
            End If
        End If
    Else
        If Pos2(7) = "WR" Then
          If WR1 = 0 Then
                WR1 = 7
            Else
                If WR2 = 0 Then
                    WR2 = 7
                Else
                    If WRRB = 0 Then
                        WRRB = 7
                    Else
                        MsgBox ("Sorry you have already started 2 WR's and Your WR/RB")
                    End If
                End If
            End If
        Else
            If Pos2(7) = "TE" Then
                If TE = 0 Then
                    TE = 7
                Else
                MsgBox "You have already started a TE"
                End If
            Else
                If Pos2(7) = "K" Then
                    If K = 0 Then
                        K = 7
                    Else
                        MsgBox "Sorry You have already started a Kicker"
                    End If
                Else
                    If Pos2(7) = "DEF" Then
                        If Def = 0 Then
                            Def = 7
                        Else
                            MsgBox "Sorry You have already started a Defense"
                        End If
                    End If
                End If
            End If
        End If
    End If
End If

End Sub

Private Sub cmd8_Click()
cmd8.Enabled = False
If Pos2(8) = "QB" Then
    If QB = 0 Then
        QB = 8
    Else
        MsgBox "Sorry you have already Choosen a QB"
    End If
Else
    If Pos2(8) = "RB" Then
        If RB1 = 0 Then
            RB1 = 8
        Else
            If RB2 = 0 Then
                RB2 = 8
            Else
                If WRRB = 0 Then
                    WRRB = 8
                Else
                    MsgBox ("Sorry you have already started 2 RB's and Your WR/RB")
                End If
            End If
        End If
    Else
        If Pos2(8) = "WR" Then
          If WR1 = 0 Then
                WR1 = 8
            Else
                If WR2 = 0 Then
                    WR2 = 8
                Else
                    If WRRB = 0 Then
                        WRRB = 8
                    Else
                        MsgBox ("Sorry you have already started 2 WR's and Your WR/RB")
                    End If
                End If
            End If
        Else
            If Pos2(8) = "TE" Then
                If TE = 0 Then
                    TE = 8
                Else
                MsgBox "You have already started a TE"
                End If
            Else
                If Pos2(8) = "K" Then
                    If K = 0 Then
                        K = 8
                    Else
                        MsgBox "Sorry You have already started a Kicker"
                    End If
                Else
                    If Pos2(8) = "DEF" Then
                        If Def = 0 Then
                            Def = 8
                        Else
                            MsgBox "Sorry You have already started a Defense"
                        End If
                    End If
                End If
            End If
        End If
    End If
End If

End Sub

Private Sub cmd9_Click()
cmd9.Enabled = False
If Pos2(9) = "QB" Then
    If QB = 0 Then
        QB = 9
    Else
        MsgBox "Sorry you have already Choosen a QB"
    End If
Else
    If Pos2(9) = "RB" Then
        If RB1 = 0 Then
            RB1 = 9
        Else
            If RB2 = 0 Then
                RB2 = 9
            Else
                If WRRB = 0 Then
                    WRRB = 9
                Else
                    MsgBox ("Sorry you have already started 2 RB's and Your WR/RB")
                End If
            End If
        End If
    Else
        If Pos2(9) = "WR" Then
          If WR1 = 0 Then
                WR1 = 9
            Else
                If WR2 = 0 Then
                    WR2 = 9
                Else
                    If WRRB = 0 Then
                        WRRB = 9
                    Else
                        MsgBox ("Sorry you have already started 2 WR's and Your WR/RB")
                    End If
                End If
            End If
        Else
            If Pos2(9) = "TE" Then
                If TE = 0 Then
                    TE = 9
                Else
                MsgBox "You have already started a TE"
                End If
            Else
                If Pos2(9) = "K" Then
                    If K = 0 Then
                        K = 9
                    Else
                        MsgBox "Sorry You have already started a Kicker"
                    End If
                Else
                    If Pos2(9) = "DEF" Then
                        If Def = 0 Then
                            Def = 9
                        Else
                            MsgBox "Sorry You have already started a Defense"
                        End If
                    End If
                End If
            End If
        End If
    End If
End If

End Sub

Private Sub cmdEnter_Click()
If QB <> 0 And RB1 <> 0 And RB2 <> 0 And WRRB <> 0 And WR1 <> 0 And WR2 <> 0 And TE <> 0 And K <> 0 And Def <> 0 Then
    frmTeam2.Visible = False
    frmStats2.Visible = True
Else
    MsgBox "Sorry you have not entered the correct amount of players please reenter your players"
    QB = 0
    RB1 = 0
    RB2 = 0
    WRRB = 0
    WR1 = 0
    WR2 = 0
    TE = 0
    K = 0
    Def = 0
    cmd1.Enabled = True
    cmd2.Enabled = True
    cmd3.Enabled = True
    cmd4.Enabled = True
    cmd5.Enabled = True
    cmd6.Enabled = True
    cmd7.Enabled = True
    cmd8.Enabled = True
    cmd9.Enabled = True
    cmd10.Enabled = True
    cmd11.Enabled = True
    cmd12.Enabled = True
    cmd13.Enabled = True
    cmd14.Enabled = True
End If

End Sub

Private Sub cmdQuit_Click()
End
End Sub

Private Sub Form_Load()

lblTeam2.Caption = NTeam2(1)
cmd1.Caption = Player2(1)
cmd2.Caption = Player2(2)
cmd3.Caption = Player2(3)
cmd4.Caption = Player2(4)
cmd5.Caption = Player2(5)
cmd6.Caption = Player2(6)
cmd7.Caption = Player2(7)
cmd8.Caption = Player2(8)
cmd9.Caption = Player2(9)
cmd10.Caption = Player2(10)
cmd11.Caption = Player2(11)
cmd12.Caption = Player2(12)
cmd13.Caption = Player2(13)
cmd14.Caption = Player2(14)

End Sub

