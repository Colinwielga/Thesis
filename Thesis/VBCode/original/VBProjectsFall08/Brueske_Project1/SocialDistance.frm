VERSION 5.00
Begin VB.Form SocialDistance 
   BackColor       =   &H8000000D&
   Caption         =   "Form2"
   ClientHeight    =   4455
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9105
   LinkTopic       =   "Form2"
   ScaleHeight     =   4455
   ScaleWidth      =   9105
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Prompt"
      Height          =   1095
      Left            =   120
      TabIndex        =   16
      Top             =   0
      Width           =   7575
      Begin VB.Label Label5 
         Caption         =   "(Enter a check a mark for associations you would permit)"
         BeginProperty Font 
            Name            =   "Gill Sans Ultra Bold Condensed"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   600
         Width           =   7335
      End
      Begin VB.Label Label4 
         Caption         =   "If it were up to you, would you permit members of this group to:"
         BeginProperty Font 
            Name            =   "Gill Sans Ultra Bold Condensed"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   7335
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Associations"
      Height          =   2055
      Left            =   2400
      TabIndex        =   5
      Top             =   1200
      Width           =   6615
      Begin VB.CheckBox Check3 
         Height          =   375
         Left            =   3120
         TabIndex        =   15
         Top             =   840
         Width           =   255
      End
      Begin VB.CheckBox Check2 
         Height          =   375
         Left            =   1800
         TabIndex        =   14
         Top             =   840
         Width           =   255
      End
      Begin VB.CheckBox Check4 
         Height          =   375
         Left            =   4440
         TabIndex        =   13
         Top             =   840
         Width           =   255
      End
      Begin VB.CheckBox Check5 
         Height          =   375
         Left            =   5760
         TabIndex        =   12
         Top             =   840
         Width           =   255
      End
      Begin VB.CheckBox Check1 
         Height          =   375
         Left            =   480
         MaskColor       =   &H80000003&
         TabIndex        =   11
         Top             =   840
         Width           =   255
      End
      Begin VB.Label lbl1 
         Alignment       =   2  'Center
         Caption         =   "Live In Your Country"
         Height          =   495
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label lbl2 
         Alignment       =   2  'Center
         Caption         =   "Live In Your Community"
         Height          =   495
         Left            =   1440
         TabIndex        =   9
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Live In Your Neighborhood"
         Height          =   495
         Left            =   2760
         TabIndex        =   8
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Live Next Door"
         Height          =   495
         Left            =   4080
         TabIndex        =   7
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Marry Your Children"
         Height          =   495
         Left            =   5400
         TabIndex        =   6
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdHome 
      Caption         =   "Home"
      Height          =   495
      Left            =   2160
      TabIndex        =   4
      Top             =   3480
      Width           =   1455
   End
   Begin VB.CommandButton cmdscore 
      Caption         =   "Score"
      Enabled         =   0   'False
      Height          =   495
      Left            =   480
      TabIndex        =   3
      Top             =   3480
      Width           =   1455
   End
   Begin VB.CommandButton cmddata 
      Caption         =   "Load Data"
      Height          =   495
      Left            =   480
      TabIndex        =   2
      Top             =   2040
      Width           =   1455
   End
   Begin VB.PictureBox picoutput 
      BackColor       =   &H0080FFFF&
      Height          =   735
      Left            =   120
      ScaleHeight     =   675
      ScaleWidth      =   2115
      TabIndex        =   1
      Top             =   1200
      Width           =   2175
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next Group"
      Enabled         =   0   'False
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   2760
      Width           =   1455
   End
End
Attribute VB_Name = "SocialDistance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ctr As Single
Dim group(1 To 50) As String
Dim score As Single
Dim tally As Single
'Alienation and Social Distance Project
'Results Form
'Kevin Brueske
'Created Oct 25, 2008
'Objective
    'Loads preset ethnic groups, obtains a value from the user for each ethnic group, tabulates a scores,
    'updates user profile



Private Sub cmddata_Click()
 'load ethnic groups into arrays
    ctr = 0
    Open App.Path & "\socialdistance.txt" For Input As #1
       Do Until EOF(1)
        ctr = ctr + 1
        Input #1, group(ctr)
        Loop
        Close #1
   cmdNext.Enabled = True
End Sub

Private Sub cmdhome_Click()
        
        'Reset form values
            Check1.Value = Unchecked
            Check2.Value = Unchecked
            Check3.Value = Unchecked
            Check4.Value = Unchecked
            Check5.Value = Unchecked
            cmdscore.Enabled = False
            ctr = 0
            score = 0
            tally = 0
         'change forms
         SocialDistance.Hide
         Home.Show
End Sub

Private Sub cmdNext_Click()
'If/Then scoring procedure to keep track of social distance score
'Highest number inputed will be the number added to the score
        If Check1.Value = Checked Then
            score = score + 1
         If Check2.Value = Checked Then
            score = score + 1
                If Check3.Value = Checked Then
                score = score + 1
                  If Check4.Value = Checked Then
                score = score + 1
                    If Check5.Value = Checked Then
                    score = score + 1
                    End If
                End If
                End If
            End If
         End If
        
         
     
       
         
         
    'Reset option buttons
         Check1.Value = Unchecked
         Check2.Value = Unchecked
         Check3.Value = Unchecked
         Check4.Value = Unchecked
         Check5.Value = Unchecked
  'IF/THEN statement determining whether to continue the questions or if the questions have been completed
   If tally < ctr Then
    
        tally = tally + 1
        picoutput.Cls
        picoutput.Print group(tally)
        
     
   Else: picoutput.Cls
        picoutput.Print "Complete."
        picoutput.Print "Press score."
        cmdscore.Enabled = True
        cmdNext.Enabled = False
    End If
         

End Sub


Private Sub cmdscore_Click()
    'Outputs user score
    picoutput.Cls
    'Score is determined by the largest number given for each question divided by the number of questions
    score = FormatNumber((score / 13), 1)
    picoutput.Print score
    picoutput.Print "Most tolerant is 5"
    picoutput.Print "Most intolerant is 0"
    'Updates user profile with the social distance score
    distanceScore(usrnum) = score
End Sub

