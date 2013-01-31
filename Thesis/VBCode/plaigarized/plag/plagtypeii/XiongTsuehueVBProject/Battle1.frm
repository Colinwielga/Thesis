VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form ssswef
   Caption         =   "Battle 1"
   ClientHeight    =   8385
   ClientLeft      =   8100
   ClientTop       =   4035
   ClientWidth     =   9480
   LinkTopic       =   "Form1"
   ScaleHeight     =   8385
   ScaleWidth      =   9480
   Begin VB.CommandButton cmdReset
      Caption         =   "Reset"
      Height          =   375
      Left            =   4200
      TabIndex        =   12
      Top             =   7560
      Width           =   1215
   End
   Begin VB.CommandButton cmdQuit
      Caption         =   "Quit"
      Height          =   375
      Left            =   5520
      TabIndex        =   11
      Top             =   7560
      Width           =   1215
   End
   Begin VB.CommandButton cmdStartBattle
      Caption         =   "Start Battle"
      Height          =   375
      Left            =   2880
      TabIndex        =   10
      Top             =   7560
      Width           =   1215
   End
   Begin VB.PictureBox uuumbnbv
      Height          =   735
      Left            =   240
      ScaleHeight     =   675
      ScaleWidth      =   8955
      TabIndex        =   9
      Top             =   6720
      Width           =   9015
   End
   Begin VB.CommandButton zzzdf
      Caption         =   "Fight"
      Height          =   375
      Left            =   6120
      TabIndex        =   8
      Top             =   5640
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton oooisjf
      Caption         =   "Heal"
      Height          =   375
      Left            =   2520
      TabIndex        =   7
      Top             =   5880
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton nnoweiuf
      Caption         =   "rrxcvbhhhhhh"
      Height          =   375
      Left            =   2520
      TabIndex        =   6
      Top             =   5400
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton pppsldkf
      Caption         =   "Gain asdfjjj"
      Height          =   375
      Left            =   720
      TabIndex        =   5
      Top             =   5880
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton mmmiwef
      Caption         =   "Kunai Throw"
      Height          =   375
      Left            =   720
      TabIndex        =   4
      Top             =   5400
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.PictureBox iiie
      Height          =   615
      Left            =   4920
      ScaleHeight     =   555
      ScaleWidth      =   4275
      TabIndex        =   3
      Top             =   4680
      Width           =   4335
   End
   Begin VB.PictureBox llltr
      Height          =   615
      Left            =   240
      ScaleHeight     =   555
      ScaleWidth      =   4395
      TabIndex        =   2
      Top             =   4680
      Width           =   4455
   End
   Begin VB.PictureBox rrrewr
      Height          =   4455
      Left            =   4920
      ScaleHeight     =   4395
      ScaleWidth      =   4275
      TabIndex        =   1
      Top             =   120
      Width           =   4335
   End
   Begin VB.PictureBox qqqlijf
      Height          =   4455
      Left            =   240
      ScaleHeight     =   4395
      ScaleWidth      =   4395
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
   Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1
      Height          =   495
      Left            =   8520
      TabIndex        =   13
      Top             =   7800
      Visible         =   0   'False
      Width           =   615
      URL             =   "\\ad\homedir$\Students\T\t1xiong\Desktop\VBProject\VBMusic\NarutoTheme.mp3"
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   1085
      _cy             =   873
   End
End
Attribute VB_Name = "ssswef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim hhhh As Double, zxcv As Long, jjjh As Long, kkky As Long

Private Sub asas()

    Dim asdfasdf As Long
    asdfasdf = Int((Rnd * 9 - 0 + 1) + 0)

    If hhhh = 0 Then
        MsgBox "Click ---> Start Battle."
    ElseIf jjjh >= 5 And hhhh > 0 Then
        zxcv = zxcv - asdfasdf
        jjjh = jjjh - 5

        iiie.Cls
        iiie.Print "HP:" & zxcv; "/100"
        iiie.Print "asdfjjj:" & kkky; "/100"

        llltr.Cls
        llltr.Print "HP:" & hhhh; "/100"
        llltr.Print "asdfjjj:" & jjjh; "/100"

        uuumbnbv.Cls
        uuumbnbv.Print inputName & " attacks with Kunai Throw, dealing " & asdfasdf & " damages."

        mmmiwef.Visible = False
        nnoweiuf.Visible = False
        pppsldkf.Visible = False
        oooisjf.Visible = False
        zzzdf.Visible = True

            If hhhh <= 0 Then
                MsgBox "You have lost. Please try again."
                ssswef.Hide
                ttterre.Show
            ElseIf zxcv <= 0 Then
                MsgBox "Congrats. You are one battle closer towards becoming Hokage."
                sseefefefes.Show
                ssswef.Hide

            End If
    Else
        MsgBox "Inefficient chakara."
    End If

End Sub

Private Sub theef()

    Dim rrxcvbhhhhhh As Long
    rrxcvbhhhhhh = (Int((Rnd * 9 - 0 + 1) + 0)) * 3

    If hhhh = 0 Then
        MsgBox "Click ---> Start Battle."
    ElseIf jjjh >= 10 And hhhh > 0 Then
        zxcv = zxcv - rrxcvbhhhhhh
        jjjh = jjjh - 10

        iiie.Cls
        iiie.Print "HP:" & zxcv; "/100"
        iiie.Print "asdfjjj:" & kkky; "/100"

        llltr.Cls
        llltr.Print "HP:" & hhhh; "/100"
        llltr.Print "asdfjjj:" & jjjh; "/100"

        uuumbnbv.Cls
        uuumbnbv.Print inputName & " attack with rrxcvbhhhhhh, dealing " & rrxcvbhhhhhh & " damages."

        mmmiwef.Visible = False
        nnoweiuf.Visible = False
        pppsldkf.Visible = False
        oooisjf.Visible = False
        zzzdf.Visible = True

            If hhhh <= 0 Then
                MsgBox "You have lost. Please try again."
                ssswef.Hide
                ttterre.Show
            ElseIf zxcv <= 0 Then
                MsgBox "Congrats. You are one battle closer towards becoming Hokage."
                sseefefefes.Show
                ssswef.Hide

            End If
    Else
        MsgBox "Inefficient chakara."
    End If

End Sub

Private Sub qewq()

    Dim asdfjjj As Long
    asdfjjj = (Int((Rnd * 9 - 0 + 1) + 0)) * 4

    If hhhh = 0 Then
        MsgBox "Click ---> Start Battle."
    ElseIf jjjh >= 0 And hhhh > 0 Then
        jjjh = jjjh + asdfjjj

        iiie.Cls
        iiie.Print "HP:" & zxcv; "/100"
        iiie.Print "asdfjjj:" & kkky; "/100"

        llltr.Cls
        llltr.Print "HP:" & hhhh; "/100"
        llltr.Print "asdfjjj:" & jjjh; "/100"

        uuumbnbv.Cls
        uuumbnbv.Print inputName & " gained " & asdfjjj & " amount of asdfjjj."

        mmmiwef.Visible = False
        nnoweiuf.Visible = False
        pppsldkf.Visible = False
        oooisjf.Visible = False
        zzzdf.Visible = True

    End If

End Sub

Private Sub qrqr()

    Dim foodstuffs As Long
    foodstuffs = (Int((Rnd * 9 - 0 + 1) + 0)) * 4

    If hhhh = 0 Then
        MsgBox "Click ---> Start Battle."
    ElseIf hhhh > 0 And hhhh > 0 Then
        hhhh = hhhh + foodstuffs

        iiie.Cls
        iiie.Print "HP:" & zxcv; "/100"
        iiie.Print "asdfjjj:" & kkky; "/100"

        llltr.Cls
        llltr.Print "HP:" & hhhh; "/100"
        llltr.Print "asdfjjj:" & jjjh; "/100"

        uuumbnbv.Cls
        uuumbnbv.Print inputName & " gained " & foodstuffs & " Health."

        mmmiwef.Visible = False
        nnoweiuf.Visible = False
        pppsldkf.Visible = False
        oooisjf.Visible = False
        zzzdf.Visible = True

    End If

End Sub

Private Sub fqfq()

    Dim ctr As Long
    ctr = (Int((Rnd * 9 - 0 + 1) + 0))

    If hhhh = 0 Then
        MsgBox "Click ---> Start Battle."
    ElseIf hhhh > 0 Then
        If ctr >= 4 And ctr > 0 Then
            Dim asdfasdf As Long
            asdfasdf = Int((Rnd * 9 - 0 + 1) + 0)

            If kkky >= 5 And zxcv > 0 Then
                hhhh = hhhh - asdfasdf
                kkky = kkky - 5

                iiie.Cls
                iiie.Print "HP:" & zxcv; "/100"
                iiie.Print "asdfjjj:" & kkky; "/100"

                llltr.Cls
                llltr.Print "HP:" & hhhh; "/100"
                llltr.Print "asdfjjj:" & jjjh; "/100"

                uuumbnbv.Cls
                uuumbnbv.Print "Gaara attacks with Kunai Throw, dealing " & asdfasdf & " damages."

                mmmiwef.Visible = True
                nnoweiuf.Visible = True
                pppsldkf.Visible = True
                oooisjf.Visible = True
                zzzdf.Visible = False

            Else
                MsgBox "Inefficient chakara."
            End If
        ElseIf ctr <= 6 And ctr > 4 Then
            Dim mmmmoonman As Long
            mmmmoonman = (Int((Rnd * 9 - 0 + 1) + 0)) * 4

            If kkky >= 10 And zxcv > 0 Then
                hhhh = hhhh - mmmmoonman
                kkky = kkky - 30

                iiie.Cls
                iiie.Print "HP:" & zxcv; "/100"
                iiie.Print "asdfjjj:" & kkky; "/100"

                llltr.Cls
                llltr.Print "HP:" & hhhh; "/100"
                llltr.Print "asdfjjj:" & jjjh; "/100"

                uuumbnbv.Cls
                uuumbnbv.Print "Gaara attack with Sand Burial, dealing " & mmmmoonman & " damages."

                mmmiwef.Visible = True
                nnoweiuf.Visible = True
                pppsldkf.Visible = True
                oooisjf.Visible = True
                zzzdf.Visible = False

            Else
                MsgBox "Inefficient chakara."
            End If
        ElseIf ctr <= 8 And ctr > 6 Then
            Dim asdfjjj As Long
            asdfjjj = (Int((Rnd * 9 - 0 + 1) + 0)) * 4

            If kkky >= 0 And zxcv > 0 Then
                kkky = kkky + asdfjjj

                iiie.Cls
                iiie.Print "HP:" & zxcv; "/100"
                iiie.Print "asdfjjj:" & kkky; "/100"

                llltr.Cls
                llltr.Print "HP:" & hhhh; "/100"
                llltr.Print "asdfjjj:" & jjjh; "/100"

                uuumbnbv.Cls
                uuumbnbv.Print "Gaara gained " & asdfjjj & " amount of asdfjjj."

                mmmiwef.Visible = True
                nnoweiuf.Visible = True
                pppsldkf.Visible = True
                oooisjf.Visible = True
                zzzdf.Visible = False
            End If

        ElseIf ctr <= 9 And ctr > 8 Then
            Dim foodstuffs As Long
            foodstuffs = (Int((Rnd * 9 - 0 + 1) + 0)) * 4

            If zxcv > 0 And zxcv > 0 Then
                zxcv = zxcv + foodstuffs

                iiie.Cls
                iiie.Print "HP:" & zxcv; "/100"
                iiie.Print "asdfjjj:" & kkky; "/100"

                llltr.Cls
                llltr.Print "HP:" & hhhh; "/100"
                llltr.Print "asdfjjj:" & jjjh; "/100"

                uuumbnbv.Cls
                uuumbnbv.Print "Gaara gained " & foodstuffs & " Health."

                mmmiwef.Visible = True
                nnoweiuf.Visible = True
                pppsldkf.Visible = True
                oooisjf.Visible = True
                zzzdf.Visible = False
            End If
        End If

        If hhhh <= 0 Then
            MsgBox "You have lost. Please try again."
            ssswef.Hide
            ttterre.Show
        ElseIf zxcv <= 0 Then
            MsgBox "Congrats. You are one battle closer towards becoming Hokage."
            sseefefefes.Show
            ssswef.Hide
        End If
    End If

End Sub

Private Sub ffff()

    End

End Sub

Private Sub fefe()

    If hhhh = 0 Then
        MsgBox "Click ---> Start Battle."
    Else
        zxcv = 0
        kkky = 0

        hhhh = 0
        jjjh = 0

        iiie.Cls

        llltr.Cls

        uuumbnbv.Cls
    End If

End Sub

Private Sub fwfw()

    hhhh = 100
    jjjh = 100
    zxcv = 100
    kkky = 100

    iiie.Cls
    iiie.Print "HP:" & zxcv; "/100"
    iiie.Print "asdfjjj:" & kkky; "/100"

    llltr.Cls
    llltr.Print "HP:" & hhhh; "/100"
    llltr.Print "asdfjjj:" & jjjh; "/100"

    mmmiwef.Visible = True
    nnoweiuf.Visible = True
    pppsldkf.Visible = True
    oooisjf.Visible = True
    zzzdf.Visible = False

    qqqlijf.Picture = LoadPicture(App.Path & "\VBPicture\Naruto1.jpg")

    rrrewr.Picture = LoadPicture(App.Path & "\VBPicture\Gaara.jpg")

End Sub
