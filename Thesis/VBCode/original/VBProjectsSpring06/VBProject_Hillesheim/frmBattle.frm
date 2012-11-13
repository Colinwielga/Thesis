VERSION 5.00
Begin VB.Form frmBattle 
   BackColor       =   &H80000001&
   Caption         =   "Battle Statistics"
   ClientHeight    =   8550
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10635
   LinkTopic       =   "Form1"
   ScaleHeight     =   8550
   ScaleWidth      =   10635
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdTotal 
      Caption         =   "Find Total Losses"
      Height          =   615
      Left            =   3480
      TabIndex        =   7
      Top             =   6600
      Width           =   1695
   End
   Begin VB.CommandButton cmdData 
      Caption         =   "Sort"
      Height          =   615
      Left            =   8040
      TabIndex        =   6
      Top             =   6600
      Width           =   1695
   End
   Begin VB.CommandButton cmdQuiz 
      Caption         =   "Guess Victors"
      Height          =   615
      Left            =   5760
      TabIndex        =   5
      Top             =   6600
      Width           =   1695
   End
   Begin VB.CommandButton cmdDisplay 
      Caption         =   "Display Statistics"
      Height          =   615
      Left            =   840
      TabIndex        =   4
      Top             =   6600
      Width           =   2055
   End
   Begin VB.PictureBox picDisplay 
      Height          =   5775
      Left            =   120
      ScaleHeight     =   5715
      ScaleWidth      =   10275
      TabIndex        =   3
      Top             =   600
      Width           =   10335
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Return to Main Page"
      Height          =   735
      Left            =   5760
      TabIndex        =   2
      Top             =   7680
      Width           =   1695
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   735
      Left            =   8040
      TabIndex        =   0
      Top             =   7680
      Width           =   1695
   End
   Begin VB.Label lblBanner 
      BackColor       =   &H80000001&
      Caption         =   "CASUALTIES OF WAR: SHIPS SUNK IN EIGHT PACIFIC NAVAL ENGAGEMENTS"
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   1320
      TabIndex        =   8
      Top             =   120
      Width           =   8055
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
      Left            =   240
      TabIndex        =   1
      Top             =   8160
      Width           =   2175
   End
End
Attribute VB_Name = "frmBattle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Naval History (Naval.vpb)
'Battle Information (frmBattle.frm)
'Jacob Hillesheim
'March 20,2006
'The purpose of this form is to show the user statistics of the battles
'and provide command buttons to take the user to a page where they can
'guess who won each battle and another where they can manipulate the data

Private Sub cmdBack_Click()
    
    'Returns user to Main Page
    frmBattle.Hide
    frmMain.Show
    
    'Clears picture box and greets user
    frmMain.picWelcome.Cls
    frmMain.picWelcome.Print "Welcome, Admiral "; Left(x, 1); ". " & Left(y, 1); ". " & z

End Sub

Private Sub cmdData_Click()
    'Takes user to Stat sorting page
    frmBattle.Hide
    frmSort.Show
    
End Sub

Private Sub cmdDisplay_Click()
    'clears Display box
    picDisplay.Cls
    
    'inputs data from file into arrays
    Open App.Path & "\battles.txt" For Input As #1
        For pos = 1 To 8
            Input #1, battles(pos), ACV(pos), ABB(pos), ACA(pos), ADD(pos), JCV(pos), JBB(pos), JCA(pos), JDD(pos)
        Next pos
    Close #1
    
    'prints headings
    picDisplay.Print "HERE ARE EIGHT FAMOUS NAVAL BATTLES OF WORLD WAR II AND CORRESPONDING LOSSES."
    picDisplay.Print
    picDisplay.Print "AMERICAN LOSSES"
    picDisplay.Print "Battle "; Tab(35); "US Carriers", "    US Battleships             ", "  US Cruisers", , "  US Destroyers"
    picDisplay.Print "----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
    
    'prints data in arrays pertaining to American losses
    For pos = 1 To 8
        picDisplay.Print battles(pos); Tab(40); ACV(pos); Tab(65); ABB(pos); Tab(90); ACA(pos); Tab(120); ADD(pos)
    Next pos
    
    'prints headings
    picDisplay.Print
    picDisplay.Print "JAPANESE LOSSES"
    picDisplay.Print "Battle "; Tab(32); "Japanese Carriers", "Japanese Battleships", "Japanese Cruisers", "Japanese Destoryers"
    picDisplay.Print "----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
    
    'prints data in arrays pertaining to Japanese losses
    For pos = 1 To 8
        picDisplay.Print battles(pos); Tab(40); JCV(pos); Tab(65); JBB(pos); Tab(90); JCA(pos); Tab(120); JDD(pos)
    Next pos

End Sub

Private Sub cmdQuit_Click()
    'ends program
    End
End Sub

Private Sub cmdQuiz_Click()
    'Takes user to Quiz page
    frmBattle.Hide
    frmQuiz.Show
End Sub

Private Sub cmdTotal_Click()
    Dim TACV, TABB, TACA, TADD, TJCV, TJBB, TJCA, TJDD As Integer
    
    'inputs data from file into arrays
    Open App.Path & "\battles.txt" For Input As #1
        For pos = 1 To 8
            Input #1, battles(pos), ACV(pos), ABB(pos), ACA(pos), ADD(pos), JCV(pos), JBB(pos), JCA(pos), JDD(pos)
        Next pos
    Close #1
    
    'initializes variables
    TACV = 0
    TABB = 0
    TACA = 0
    TADD = 0
    TJCV = 0
    TJBB = 0
    TJCA = 0
    TJDD = 0
    
    'finds the sum of numbers in each array
    For pos = 1 To 8
        TACV = TACV + ACV(pos)
    Next pos
    For pos = 1 To 8
        TABB = TABB + ABB(pos)
    Next pos
    For pos = 1 To 8
        TACA = TACA + ACA(pos)
    Next pos
    For pos = 1 To 8
        TADD = TADD + ADD(pos)
    Next pos
    For pos = 1 To 8
        TJCV = TJCV + JCV(pos)
    Next pos
    For pos = 1 To 8
        TJBB = TJBB + JBB(pos)
    Next pos
    For pos = 1 To 8
        TJCA = TJCA + JCA(pos)
    Next pos
    For pos = 1 To 8
        TJDD = TJDD + JDD(pos)
    Next pos
    
    'gives totals to user
    MsgBox "America lost " & TACV & " carriers, " & TABB & " battleships, " & TACA & " cruisers, and " & TADD & " destroyers", , "American Total"
    MsgBox "Japan lost " & TJCV & " carriers, " & TJBB & " battleships, " & TJCA & " cruisers, and " & TJDD & " destroyers", , "Japanese Total"
End Sub
