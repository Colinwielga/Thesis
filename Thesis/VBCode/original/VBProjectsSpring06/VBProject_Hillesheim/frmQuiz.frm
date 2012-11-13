VERSION 5.00
Begin VB.Form frmQuiz 
   BackColor       =   &H80000001&
   Caption         =   "Guess Who Won!"
   ClientHeight    =   8220
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10680
   LinkTopic       =   "Form1"
   ScaleHeight     =   8220
   ScaleWidth      =   10680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdMidway 
      Caption         =   "Midway"
      Height          =   495
      Left            =   4560
      TabIndex        =   11
      Top             =   6600
      Width           =   1455
   End
   Begin VB.CommandButton cmdSavo 
      Caption         =   "Savo Island"
      Height          =   495
      Left            =   480
      TabIndex        =   10
      Top             =   7200
      Width           =   1455
   End
   Begin VB.CommandButton cmdVella 
      Caption         =   "Vella Gulf"
      Height          =   495
      Left            =   2520
      TabIndex        =   9
      Top             =   7200
      Width           =   1455
   End
   Begin VB.CommandButton cmdJava 
      Caption         =   "Java Sea"
      Height          =   495
      Left            =   480
      TabIndex        =   8
      Top             =   6600
      Width           =   1455
   End
   Begin VB.CommandButton cmdCoral 
      Caption         =   "Coral Sea"
      Height          =   495
      Left            =   2520
      TabIndex        =   7
      Top             =   6600
      Width           =   1455
   End
   Begin VB.CommandButton cmdMTS 
      Caption         =   "Philippine Sea"
      CausesValidation=   0   'False
      Height          =   495
      Left            =   4560
      TabIndex        =   6
      Top             =   7200
      Width           =   1455
   End
   Begin VB.CommandButton cmdLeyte 
      Caption         =   "Leyte Gulf"
      Height          =   495
      Left            =   6600
      TabIndex        =   5
      Top             =   7200
      Width           =   1455
   End
   Begin VB.CommandButton cmdGuad 
      Caption         =   "Guadalcanal"
      Height          =   495
      Left            =   6600
      TabIndex        =   4
      Top             =   6600
      Width           =   1455
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load Stats"
      Height          =   375
      Left            =   4200
      TabIndex        =   3
      Top             =   6000
      Width           =   2655
   End
   Begin VB.PictureBox picDisplay 
      Height          =   5535
      Left            =   240
      ScaleHeight     =   5475
      ScaleWidth      =   10155
      TabIndex        =   2
      Top             =   360
      Width           =   10215
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return to Previous Page"
      Height          =   615
      Left            =   9000
      TabIndex        =   1
      Top             =   6360
      Width           =   1215
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   615
      Left            =   9000
      TabIndex        =   0
      Top             =   7320
      Width           =   1215
   End
   Begin VB.Label lblDirections 
      BackColor       =   &H80000001&
      Caption         =   "Click on Load Stats to see statistics of the battles. Click on a battle below to give your answer."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   1080
      TabIndex        =   13
      Top             =   0
      Width           =   8895
   End
   Begin VB.Label lblName 
      BackColor       =   &H80000001&
      Caption         =   "By Jacob Hillesheim"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   7800
      Width           =   2655
   End
End
Attribute VB_Name = "frmQuiz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Naval History (Naval.vpb)
'Quiz (frmQuiz.frm)
'Jacob Hillesheim
'March 20,2006
'The purpose of this form is to make the user analyze the data,
'predict who won each battle and explain the ramifications of each battle on both sides.
Option Explicit
Dim a As String
Private Sub cmdCoral_Click()
    'Asks who won battle
    a = InputBox("Do you think Japan or America won the Battle of the Coral Sea?", "Battle of the Coral Sea")
    
    'provides answer
    If a = "America" Then
        MsgBox "You are right. The Battle of the Coral Sea was a strategic victory for the Americans. It came at a cost the US Navy could ill afford, but turning back Japan's invasion of Port Moresby kept the communications and supply lines between America and Australia open.", , "Correct!"
    ElseIf a = "Japan" Then
        MsgBox "You are correct. The Battle of the Coral Sea was a tactical victory for the Japanese. They lost a small carrier and a destroyer while sinking a destroyer and a large carrier, which America could not afford to lose.", , "Correct!"
    Else
        MsgBox "Please enter either Japan or America", , "Error"
    End If
End Sub
Private Sub cmdGuad_Click()
    'Asks who won battle
    a = InputBox("Do you think Japan or America won the Naval Battle of Guadalcanal?", "Naval Battle of Guadalcanal")
    
    'provides answer
    If a = "America" Then
        MsgBox "Although the Americans lost more ships than the Japanese, the Naval Battle of Guadalcanal was an American victory. The US Navy sunk two Japanese battleships in the battle, but more important were the strategic consequences. Because of the battle, the Japanese surface group was not allowed to bombard the US Marines on Guadalcanal. The Marines were able to fight back a vicious Japanese assault. Defeats in the Land and Naval Battles of Guadalcanal persuaded the Japanese to pull out of the island.", , "Correct"
    ElseIf a = "Japan" Then
        MsgBox "While Japan did sink more American ships, the Naval Battle of Guadalcanal was a resounding American victory. The Japanese ships were not allowed to bombard Marine emplacements on the island and Japanese transports trying to land troops on Guadalcanal were massacred.", , "Almost"
    Else
        MsgBox "Please enter either Japan or America", , "Error"
    End If
End Sub
Private Sub cmdJava_Click()
    'Asks who won battle
    a = InputBox("Do you think Japan or America won the Battle of the Java Sea?", "Battle of the Java Sea")
    
    'provides answer
    If a = "America" Then
        MsgBox "Sorry. The Battle of the Java Sea wiped out the combined Asiatic Fleets of the US, Britain, and the Netherlands. Japan was able to easily capture the Dutch East Indies as a result of this impressive victory over a group of antiquidated Allied vessels.", , "Incorrect"
    ElseIf a = "Japan" Then
        MsgBox "You are correct. The Japanese juggernaut continued to roll over the Allies in this battle. With this battle, the attack on Pearl Harbor, and the Indian Ocean Raid, Japan controlled the Indian Ocean and West, South, and Central Pacific.", , "Correct"
    Else
        MsgBox "Please enter either Japan or America", , "Error"
    End If
End Sub
Private Sub cmdLeyte_Click()
    'Asks who won battle
    a = InputBox("Do you think Japan or America won the Battle of Leyte Gulf?", "Battle of Leyte Gulf")
    
    'provides answer
    If a = "America" Then
        MsgBox "You are correct. Leyte Gulf remains the largest naval battle in history and an American victory. Leyte Gulf was an epic for many reasons, one of which being that the once invincible Japanese Fleet was no longer a viable threat.", , "You are correct"
    ElseIf a = "Japan" Then
        MsgBox "You are almost correct. This battle was almost a huge victory for the Japanese, but American heroism saved the day. The courageous charge of the US destroyers and destroyer-escorts against Japanese battleships was the only thing that saved a task force of escort carriers and troop transports landing on Leyte Island.", , "Sorry"
    Else
        MsgBox "Please enter either Japan or America", , "Error"
    End If
End Sub
Private Sub cmdLoad_Click()
    'Clears display
    picDisplay.Cls
    
    'inputs data from file into arrays
    Open App.Path & "\battles.txt" For Input As #1
        For pos = 1 To 8
            Input #1, battles(pos), ACV(pos), ABB(pos), ACA(pos), ADD(pos), JCV(pos), JBB(pos), JCA(pos), JDD(pos)
        Next pos
    Close #1
    
    'prints column headings
    picDisplay.Print "HERE ARE EIGHT FAMOUS NAVAL BATTLES OF WORLD WAR II AND CORRESPONDING LOSSES."
    picDisplay.Print
    picDisplay.Print "AMERICAN LOSSES"
    picDisplay.Print "Battle "; Tab(35); "US Carriers", "    US Battleships             ", "  US Cruisers", , "  US Destroyers"
    picDisplay.Print "----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
    
    'Displays information in arrays pertaining to American loses
    For pos = 1 To 8
        picDisplay.Print battles(pos); Tab(40); ACV(pos); Tab(65); ABB(pos); Tab(90); ACA(pos); Tab(120); ADD(pos)
    Next pos
    
    'Prints column headings
    picDisplay.Print
    picDisplay.Print "JAPANESE LOSSES"
    picDisplay.Print "Battle "; Tab(32); "Japanese Carriers", "Japanese Battleships", "Japanese Cruisers", "Japanese Destoryers"
    picDisplay.Print "----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
    
    'Displays information in arrays pertaining to Japanese loses
    For pos = 1 To 8
        picDisplay.Print battles(pos); Tab(40); JCV(pos); Tab(65); JBB(pos); Tab(90); JCA(pos); Tab(120); JDD(pos)
    Next pos
End Sub
Private Sub cmdMidway_Click()
    'Asks who won battle
    a = InputBox("Do you think Japan or America won the Battle of Midway?", "Battle of Midway")
    
    'provides answer
    If a = "America" Then
        MsgBox "You are correct. The Battle of Midway is considered to be the turning point of World War II in the Pacific.", , "Correct"
    ElseIf a = "Japan" Then
        MsgBox "Did you actually read the statistics at all? Midway was a crushing defeat for the Japanese, showing that they were not invincible. From here, it was all down hill for Japan.", , "Not quite"
    Else
        MsgBox "Please enter either Japan or America", , "Error"
    End If
End Sub
Private Sub cmdMTS_Click()
    'Asks who won battle
    a = InputBox("Do you think Japan or America won the Battle of the Philippine Sea?", "Battle of the Philippine Sea")
    
    'provides answer
    If a = "America" Then
        MsgBox "Yes. The Battle of the Philippine Sea was another victory for Admiral Spruance. US forces sunk three Japanese carriers and shot down enough novice pilots that surviving Japanese aircraft carriers no longer were able to carry more than a few planes.", , "Easy one"
    ElseIf a = "Japan" Then
        MsgBox "Does the phrase Marianas Turkey Shoot mean anything to you? Japan lost three carriers and most importantly lost so many pilots and planes that Japan was forced to see that their pilots could no longer hit American targets without flying directly into them.", , "Wow..."
    Else
        MsgBox "Please enter either Japan or America", , "Error"
    End If
End Sub
Private Sub cmdQuit_Click()
    'ends program
    End
End Sub
Private Sub cmdReturn_Click()
    'returns user to Battle Page
    frmBattle.Show
    frmQuiz.Hide
    
End Sub
Private Sub cmdSavo_Click()
    'Asks who won battle
    a = InputBox("Do you think Japan or America won the Battle of Savo Island?", "Battle of Savo Island")
    
    'provides answer
    If a = "America" Then
        MsgBox "Did you remember to load the statistics? The Battle of Savo Island was a complete disaster for the Allies. If the Japanese would have pressed forward, they would have been able to bombard US positions on Guadalcanal and destroy supply transports, perhaps resulting in Japanese victory on Guadalcanal.", , "Way off"
    ElseIf a = "Japan" Then
        MsgBox "You are correct. The Battle of Savo Island was the most lopsided victory for Japan since Pearl Harbor.", , "Absolutely"
    Else
        MsgBox "Please enter either Japan or America", , "Error"
    End If
End Sub
Private Sub cmdVella_Click()
    'Asks who won battle
    a = InputBox("Do you think Japan or America won the Battle of Vella Gulf?", "Battle of Vella Gulf")
    
    'provides answer
    If a = "America" Then
        MsgBox "Okay, that was an easy one. Although it was a small battle, it showed that American destroyer skippers were just as good as their famed Japanese counterparts.", , "Yup"
    ElseIf a = "Japan" Then
        MsgBox "Sorry. Japan suffered a defeat in the Battle of Vella Gulf. While only a few small ships were lost, Japan was not able to make good its losses at this stage in the war.", , "Nope"
    Else
        MsgBox "Please enter either Japan or America", , "Error"
    End If
End Sub
Private Sub Form_Load()
    'explains instructions for quiz to user
    MsgBox "Here is your chance to name the winner of some famous battles based on the ships lost. First, load and study the statistics. When you are finished, click on a battle and name your winner. You must spell your answer correctly. Have fun!", , "Instructions"
    MsgBox "For a special treat, enter both countries for both perspectives of the outcome. Not all of these battles had a clear-cut result.", , "Hint"
End Sub
