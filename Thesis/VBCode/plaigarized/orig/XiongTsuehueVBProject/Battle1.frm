VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form frmBattle1 
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
   Begin VB.PictureBox picResults 
      Height          =   735
      Left            =   240
      ScaleHeight     =   675
      ScaleWidth      =   8955
      TabIndex        =   9
      Top             =   6720
      Width           =   9015
   End
   Begin VB.CommandButton cmdFight 
      Caption         =   "Fight"
      Height          =   375
      Left            =   6120
      TabIndex        =   8
      Top             =   5640
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton cmdAttack4 
      Caption         =   "Heal"
      Height          =   375
      Left            =   2520
      TabIndex        =   7
      Top             =   5880
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton cmdAttack2 
      Caption         =   "Punch"
      Height          =   375
      Left            =   2520
      TabIndex        =   6
      Top             =   5400
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton cmdAttack3 
      Caption         =   "Gain Chakara"
      Height          =   375
      Left            =   720
      TabIndex        =   5
      Top             =   5880
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton cmdAttack1 
      Caption         =   "Kunai Throw"
      Height          =   375
      Left            =   720
      TabIndex        =   4
      Top             =   5400
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.PictureBox picEnemyInfo 
      Height          =   615
      Left            =   4920
      ScaleHeight     =   555
      ScaleWidth      =   4275
      TabIndex        =   3
      Top             =   4680
      Width           =   4335
   End
   Begin VB.PictureBox picPlayerInfo 
      Height          =   615
      Left            =   240
      ScaleHeight     =   555
      ScaleWidth      =   4395
      TabIndex        =   2
      Top             =   4680
      Width           =   4455
   End
   Begin VB.PictureBox picGaara 
      Height          =   4455
      Left            =   4920
      ScaleHeight     =   4395
      ScaleWidth      =   4275
      TabIndex        =   1
      Top             =   120
      Width           =   4335
   End
   Begin VB.PictureBox picNaruto1 
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
Attribute VB_Name = "frmBattle1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim PlayerHealth As Single, EnemyHealth As Integer, PlayerChakara As Integer, EnemyChakara As Integer

Private Sub cmdAttack1_Click()

    Dim KunaiThrow As Integer
    KunaiThrow = Int((Rnd * 9 - 0 + 1) + 0)
    
    If PlayerHealth = 0 Then
        MsgBox "Click ---> Start Battle."
    ElseIf PlayerChakara >= 5 And PlayerHealth > 0 Then
        EnemyHealth = EnemyHealth - KunaiThrow
        PlayerChakara = PlayerChakara - 5
        
        picEnemyInfo.Cls
        picEnemyInfo.Print "HP:" & EnemyHealth; "/100"
        picEnemyInfo.Print "Chakara:" & EnemyChakara; "/100"
        
        picPlayerInfo.Cls
        picPlayerInfo.Print "HP:" & PlayerHealth; "/100"
        picPlayerInfo.Print "Chakara:" & PlayerChakara; "/100"
        
        picResults.Cls
        picResults.Print inputName & " attacks with Kunai Throw, dealing " & KunaiThrow & " damages."
        
        cmdAttack1.Visible = False
        cmdAttack2.Visible = False
        cmdAttack3.Visible = False
        cmdAttack4.Visible = False
        cmdFight.Visible = True
        
            If PlayerHealth <= 0 Then
                MsgBox "You have lost. Please try again."
                frmBattle1.Hide
                frmTheEnd.Show
            ElseIf EnemyHealth <= 0 Then
                MsgBox "Congrats. You are one battle closer towards becoming Hokage."
                frmBattle2.Show
                frmBattle1.Hide
                
            End If
    Else
        MsgBox "Inefficient chakara."
    End If

End Sub

Private Sub cmdAttack2_Click()

    Dim Punch As Integer
    Punch = (Int((Rnd * 9 - 0 + 1) + 0)) * 3
    
    If PlayerHealth = 0 Then
        MsgBox "Click ---> Start Battle."
    ElseIf PlayerChakara >= 10 And PlayerHealth > 0 Then
        EnemyHealth = EnemyHealth - Punch
        PlayerChakara = PlayerChakara - 10
        
        picEnemyInfo.Cls
        picEnemyInfo.Print "HP:" & EnemyHealth; "/100"
        picEnemyInfo.Print "Chakara:" & EnemyChakara; "/100"
        
        picPlayerInfo.Cls
        picPlayerInfo.Print "HP:" & PlayerHealth; "/100"
        picPlayerInfo.Print "Chakara:" & PlayerChakara; "/100"
        
        picResults.Cls
        picResults.Print inputName & " attack with Punch, dealing " & Punch & " damages."
        
        cmdAttack1.Visible = False
        cmdAttack2.Visible = False
        cmdAttack3.Visible = False
        cmdAttack4.Visible = False
        cmdFight.Visible = True
        
            If PlayerHealth <= 0 Then
                MsgBox "You have lost. Please try again."
                frmBattle1.Hide
                frmTheEnd.Show
            ElseIf EnemyHealth <= 0 Then
                MsgBox "Congrats. You are one battle closer towards becoming Hokage."
                frmBattle2.Show
                frmBattle1.Hide
                
            End If
    Else
        MsgBox "Inefficient chakara."
    End If

End Sub

Private Sub cmdAttack3_Click()

    Dim Chakara As Integer
    Chakara = (Int((Rnd * 9 - 0 + 1) + 0)) * 4
    
    If PlayerHealth = 0 Then
        MsgBox "Click ---> Start Battle."
    ElseIf PlayerChakara >= 0 And PlayerHealth > 0 Then
        PlayerChakara = PlayerChakara + Chakara
        
        picEnemyInfo.Cls
        picEnemyInfo.Print "HP:" & EnemyHealth; "/100"
        picEnemyInfo.Print "Chakara:" & EnemyChakara; "/100"
        
        picPlayerInfo.Cls
        picPlayerInfo.Print "HP:" & PlayerHealth; "/100"
        picPlayerInfo.Print "Chakara:" & PlayerChakara; "/100"
        
        picResults.Cls
        picResults.Print inputName & " gained " & Chakara & " amount of Chakara."
        
        cmdAttack1.Visible = False
        cmdAttack2.Visible = False
        cmdAttack3.Visible = False
        cmdAttack4.Visible = False
        cmdFight.Visible = True
        
    End If
        
End Sub

Private Sub cmdAttack4_Click()

    Dim Eat As Integer
    Eat = (Int((Rnd * 9 - 0 + 1) + 0)) * 4
    
    If PlayerHealth = 0 Then
        MsgBox "Click ---> Start Battle."
    ElseIf PlayerHealth > 0 And PlayerHealth > 0 Then
        PlayerHealth = PlayerHealth + Eat
        
        picEnemyInfo.Cls
        picEnemyInfo.Print "HP:" & EnemyHealth; "/100"
        picEnemyInfo.Print "Chakara:" & EnemyChakara; "/100"
        
        picPlayerInfo.Cls
        picPlayerInfo.Print "HP:" & PlayerHealth; "/100"
        picPlayerInfo.Print "Chakara:" & PlayerChakara; "/100"
        
        picResults.Cls
        picResults.Print inputName & " gained " & Eat & " Health."
        
        cmdAttack1.Visible = False
        cmdAttack2.Visible = False
        cmdAttack3.Visible = False
        cmdAttack4.Visible = False
        cmdFight.Visible = True
        
    End If
    
End Sub

Private Sub cmdFight_Click()

    Dim ctr As Integer
    ctr = (Int((Rnd * 9 - 0 + 1) + 0))
    
    If PlayerHealth = 0 Then
        MsgBox "Click ---> Start Battle."
    ElseIf PlayerHealth > 0 Then
        If ctr >= 4 And ctr > 0 Then
            Dim KunaiThrow As Integer
            KunaiThrow = Int((Rnd * 9 - 0 + 1) + 0)
        
            If EnemyChakara >= 5 And EnemyHealth > 0 Then
                PlayerHealth = PlayerHealth - KunaiThrow
                EnemyChakara = EnemyChakara - 5
                
                picEnemyInfo.Cls
                picEnemyInfo.Print "HP:" & EnemyHealth; "/100"
                picEnemyInfo.Print "Chakara:" & EnemyChakara; "/100"
                
                picPlayerInfo.Cls
                picPlayerInfo.Print "HP:" & PlayerHealth; "/100"
                picPlayerInfo.Print "Chakara:" & PlayerChakara; "/100"
                
                picResults.Cls
                picResults.Print "Gaara attacks with Kunai Throw, dealing " & KunaiThrow & " damages."
                
                cmdAttack1.Visible = True
                cmdAttack2.Visible = True
                cmdAttack3.Visible = True
                cmdAttack4.Visible = True
                cmdFight.Visible = False
                
            Else
                MsgBox "Inefficient chakara."
            End If
        ElseIf ctr <= 6 And ctr > 4 Then
            Dim SandBurial As Integer
            SandBurial = (Int((Rnd * 9 - 0 + 1) + 0)) * 4
            
            If EnemyChakara >= 10 And EnemyHealth > 0 Then
                PlayerHealth = PlayerHealth - SandBurial
                EnemyChakara = EnemyChakara - 30
                
                picEnemyInfo.Cls
                picEnemyInfo.Print "HP:" & EnemyHealth; "/100"
                picEnemyInfo.Print "Chakara:" & EnemyChakara; "/100"
                
                picPlayerInfo.Cls
                picPlayerInfo.Print "HP:" & PlayerHealth; "/100"
                picPlayerInfo.Print "Chakara:" & PlayerChakara; "/100"
                
                picResults.Cls
                picResults.Print "Gaara attack with Sand Burial, dealing " & SandBurial & " damages."
                
                cmdAttack1.Visible = True
                cmdAttack2.Visible = True
                cmdAttack3.Visible = True
                cmdAttack4.Visible = True
                cmdFight.Visible = False
                
            Else
                MsgBox "Inefficient chakara."
            End If
        ElseIf ctr <= 8 And ctr > 6 Then
            Dim Chakara As Integer
            Chakara = (Int((Rnd * 9 - 0 + 1) + 0)) * 4
            
            If EnemyChakara >= 0 And EnemyHealth > 0 Then
                EnemyChakara = EnemyChakara + Chakara
                
                picEnemyInfo.Cls
                picEnemyInfo.Print "HP:" & EnemyHealth; "/100"
                picEnemyInfo.Print "Chakara:" & EnemyChakara; "/100"
                
                picPlayerInfo.Cls
                picPlayerInfo.Print "HP:" & PlayerHealth; "/100"
                picPlayerInfo.Print "Chakara:" & PlayerChakara; "/100"
                
                picResults.Cls
                picResults.Print "Gaara gained " & Chakara & " amount of Chakara."
                
                cmdAttack1.Visible = True
                cmdAttack2.Visible = True
                cmdAttack3.Visible = True
                cmdAttack4.Visible = True
                cmdFight.Visible = False
            End If
        
        ElseIf ctr <= 9 And ctr > 8 Then
            Dim Eat As Integer
            Eat = (Int((Rnd * 9 - 0 + 1) + 0)) * 4
            
            If EnemyHealth > 0 And EnemyHealth > 0 Then
                EnemyHealth = EnemyHealth + Eat
                
                picEnemyInfo.Cls
                picEnemyInfo.Print "HP:" & EnemyHealth; "/100"
                picEnemyInfo.Print "Chakara:" & EnemyChakara; "/100"
                
                picPlayerInfo.Cls
                picPlayerInfo.Print "HP:" & PlayerHealth; "/100"
                picPlayerInfo.Print "Chakara:" & PlayerChakara; "/100"
                
                picResults.Cls
                picResults.Print "Gaara gained " & Eat & " Health."
                
                cmdAttack1.Visible = True
                cmdAttack2.Visible = True
                cmdAttack3.Visible = True
                cmdAttack4.Visible = True
                cmdFight.Visible = False
            End If
        End If
        
        If PlayerHealth <= 0 Then
            MsgBox "You have lost. Please try again."
            frmBattle1.Hide
            frmTheEnd.Show
        ElseIf EnemyHealth <= 0 Then
            MsgBox "Congrats. You are one battle closer towards becoming Hokage."
            frmBattle2.Show
            frmBattle1.Hide
        End If
    End If

End Sub

Private Sub cmdQuit_Click()

    End
    
End Sub

Private Sub cmdReset_Click()

    If PlayerHealth = 0 Then
        MsgBox "Click ---> Start Battle."
    Else
        EnemyHealth = 0
        EnemyChakara = 0
    
        PlayerHealth = 0
        PlayerChakara = 0
        
        picEnemyInfo.Cls
            
        picPlayerInfo.Cls
        
        picResults.Cls
    End If
    
End Sub

Private Sub cmdStartBattle_Click()
    
    PlayerHealth = 100
    PlayerChakara = 100
    EnemyHealth = 100
    EnemyChakara = 100
    
    picEnemyInfo.Cls
    picEnemyInfo.Print "HP:" & EnemyHealth; "/100"
    picEnemyInfo.Print "Chakara:" & EnemyChakara; "/100"
        
    picPlayerInfo.Cls
    picPlayerInfo.Print "HP:" & PlayerHealth; "/100"
    picPlayerInfo.Print "Chakara:" & PlayerChakara; "/100"

    cmdAttack1.Visible = True
    cmdAttack2.Visible = True
    cmdAttack3.Visible = True
    cmdAttack4.Visible = True
    cmdFight.Visible = False
    
    picNaruto1.Picture = LoadPicture(App.Path & "\VBPicture\Naruto1.jpg")
    
    picGaara.Picture = LoadPicture(App.Path & "\VBPicture\Gaara.jpg")

End Sub
