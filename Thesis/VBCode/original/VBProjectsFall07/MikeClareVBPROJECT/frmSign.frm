VERSION 5.00
Begin VB.Form frmSign 
   BackColor       =   &H00004040&
   Caption         =   "the mysterious sign..."
   ClientHeight    =   7395
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10485
   LinkTopic       =   "Form1"
   ScaleHeight     =   7395
   ScaleWidth      =   10485
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdContinue 
      Caption         =   "Continue..."
      BeginProperty Font 
         Name            =   "Myriad Condensed Web"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   4680
      TabIndex        =   4
      Top             =   6000
      Width           =   2895
   End
   Begin VB.CommandButton cmdAge 
      Caption         =   "Sort ages..."
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Myriad Condensed Web"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   480
      TabIndex        =   3
      Top             =   3960
      Width           =   2055
   End
   Begin VB.CommandButton cmdAlpha 
      Caption         =   "Sort names..."
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Myriad Condensed Web"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   480
      TabIndex        =   2
      Top             =   2280
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Wipe off the dust and blood so you can read it better"
      BeginProperty Font 
         Name            =   "Myriad Condensed Web"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   480
      TabIndex        =   1
      Top             =   360
      Width           =   2055
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H0000C0C0&
      BeginProperty Font 
         Name            =   "Myriad Condensed Web"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5415
      Left            =   2880
      ScaleHeight     =   5355
      ScaleWidth      =   7155
      TabIndex        =   0
      Top             =   240
      Width           =   7215
   End
End
Attribute VB_Name = "frmSign"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 Dim Pass As Integer
    Dim TempName As String
    Dim TempAge As Single

Private Sub cmdAge_Click()  'sort the sign by ages from youngest to oldest
    picResults.Cls
    For Pass = 1 To (CTR - 1)
        For Pos = 1 To (CTR - Pass)
            If agesArray(Pos) > agesArray(Pos + 1) Then
                TempAge = agesArray(Pos)
                agesArray(Pos) = agesArray(Pos + 1)
                agesArray(Pos + 1) = TempAge
                
                TempName = namesArray(Pos)
                namesArray(Pos) = namesArray(Pos + 1)
                namesArray(Pos + 1) = TempName
                               
            End If
        Next Pos
    Next Pass
    
    picResults.Print Right("So, now These are the names of those who died in the alien attack", 57)
    picResults.Print "------------------------------------------------------------------------------"
    For Pos = 1 To CTR
    picResults.Print namesArray(Pos), ; Tab(32); "Age "; agesArray(Pos)
    Next Pos
End Sub

Private Sub cmdAlpha_Click()        'sort the names alphabetically
   
    
    picResults.Cls
    
    For Pass = 1 To (CTR - 1)
        For Pos = 1 To (CTR - Pass)
            If namesArray(Pos) > namesArray(Pos + 1) Then
                TempName = namesArray(Pos)
                namesArray(Pos) = namesArray(Pos + 1)
                namesArray(Pos + 1) = TempName
                
               
                TempAge = agesArray(Pos)
                agesArray(Pos) = agesArray(Pos + 1)
                agesArray(Pos + 1) = TempAge
            End If
        Next Pos
    Next Pass
    
    picResults.Print Right("So, now These are the names of those who died in the alien attack", 57)
    picResults.Print "------------------------------------------------------------------------------"
    For Pos = 1 To CTR
    picResults.Print namesArray(Pos), ; Tab(32); "Age "; agesArray(Pos)
    Next Pos
End Sub

Private Sub cmdContinue_Click() 'continue to a fight!
    MsgBox ("Okay, enough with this sign.  I'm still alive and I don't need to know who has died!"), , ("Moving on")
    MsgBox ("You turn around to leave, but find yourself face to face with another Alien!!"), , ("Ahhhhh!")
    frmSign.Hide
    frmFight.Show
End Sub

Private Sub Command1_Click()        'load an array (file sign.txt)  and print it out
    Open App.Path & "\sign.txt" For Input As #3
    CTR = 0
    Do Until EOF(3)
        CTR = CTR + 1
        Input #3, namesArray(CTR), agesArray(CTR)
    Loop
    Close #3
    picResults.Cls
    picResults.Print Right("So, now These are the names of those who died in the alien attack", 57)
    picResults.Print "------------------------------------------------------------------------------"
    For Pos = 1 To CTR
        picResults.Print namesArray(Pos), ; Tab(32); "Age "; agesArray(Pos)
    Next Pos
    MsgBox ("You also notice there are a couple buttons next to the sign.  Try clicking them..."), , ("Other buttons")
    cmdAlpha.Enabled = True
    cmdAge.Enabled = True
End Sub
