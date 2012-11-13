VERSION 5.00
Begin VB.Form frmEffects 
   BackColor       =   &H000000C0&
   Caption         =   "Physical Effects"
   ClientHeight    =   5415
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   19080
   FillColor       =   &H000000C0&
   BeginProperty Font 
      Name            =   "MV Boli"
      Size            =   18
      Charset         =   1
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H000000C0&
   LinkTopic       =   "Form1"
   ScaleHeight     =   5415
   ScaleWidth      =   19080
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdNext 
      Caption         =   "More Info:"
      Height          =   615
      Left            =   10320
      TabIndex        =   2
      Top             =   3720
      Width           =   1815
   End
   Begin VB.PictureBox picEffects 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Myriad Condensed Web"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   120
      ScaleHeight     =   1695
      ScaleWidth      =   20655
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1680
      Width           =   20655
      Begin VB.Label lblClick 
         Caption         =   "click the white box:"
         BeginProperty Font 
            Name            =   "MV Boli"
            Size            =   14.25
            Charset         =   1
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   3
         Top             =   1320
         Width           =   4335
      End
   End
   Begin VB.Label lblEffects 
      BackColor       =   &H000000C0&
      Caption         =   "Physical Effects"
      BeginProperty Font 
         Name            =   "Rosewood Std Regular"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   6360
      TabIndex        =   0
      Top             =   240
      Width           =   9495
   End
End
Attribute VB_Name = "frmEffects"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdNext_Click()
'hides current form and shows conclusion form
frmEffects.Hide
frmConclusion.Show
End Sub

Private Sub picEffects_Click()
'we find the range which the user's bal fits and display the corresponding effects of this range
   Select Case Bal
        Case Is <= 0#
            picEffects.Print "You have not consumed any alcohol.  Therefore you have no physical affects from consumption"
        Case Is <= 0.03
             picEffects.Print "No loss of coordination, slight euphoria and loss of shyness. Depressant effects are not apparent.  Mildly relaxed and maybe a little lightheaded"
        Case Is <= 0.06
             picEffects.Print "Feeling of well-being, relaxation, lower inhibitions, sesnation of   warmth. Euphoria. Some minor impairment of reasoning an memory, lowering of caution. "
        Case Is <= 0.09
             picEffects.Print "Slight impairment of balance, speech, vision, reaction time, and hearing.  Euphoria. Judgement and self-control are reduced "
        Case Is <= 0.125
             picEffects.Print "Significant impairment of motor coordination and loss of goodjudgment. Speech may be slurred; balance, vision, reaction time and hearing will be impaired. Euphoria."
        Case Is <= 0.15
             picEffects.Print "Gross motor impairment and lack of physical control. Blurred vision and major loss of balance. Judgment and perception are severely  impaired."
        Case Is <= 0.19
             picEffects.Print "Dysphoria predominates, nausea may appear. The drinker has the appearance of a sloppy drunk."
        Case Is <= 0.2
             picEffects.Print "Feeling dazed/confused. May need help to stand/walk. If you injure yourself you may not feel the pain. The gag reflex is impaired and you can choke if you do vomit."
        Case Is <= 0.25
             picEffects.Print "All mental, physical and sensory functions are severely impaired. Increased risk of asphyxiation from choking on vomit."
        Case Is <= 0.3
             picEffects.Print "STUPOR. You have little comprehension of where you are. You  may pass out suddenly and be difficult to awaken."
        Case Is <= 0.35
             picEffects.Print " A Coma is possible. This is the level of surgical anesthesia."
        Case Is <= 0.4
             picEffects.Print " Onset of a coma, and possible death due to respiratory arrest"
        Case Else
            picEffects.Print "You are most likely legally dead.  Please call 911."
    End Select
End Sub

