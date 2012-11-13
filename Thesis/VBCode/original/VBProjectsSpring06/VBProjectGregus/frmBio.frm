VERSION 5.00
Begin VB.Form frmBio 
   BackColor       =   &H80000007&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Team Bio Page"
   ClientHeight    =   9240
   ClientLeft      =   105
   ClientTop       =   885
   ClientWidth     =   11370
   BeginProperty Font 
      Name            =   "Gill Sans Ultra Bold Condensed"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   FontTransparent =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9240
   ScaleWidth      =   11370
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CommandButton cmdInstructions 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Instructions"
      Height          =   495
      Left            =   4320
      MaskColor       =   &H000000C0&
      TabIndex        =   19
      Top             =   7440
      Width           =   2175
   End
   Begin VB.Timer TimerJ 
      Interval        =   3000
      Left            =   4920
      Top             =   4200
   End
   Begin VB.PictureBox picBoxJ 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1095
      Left            =   5280
      Picture         =   "frmBio.frx":0000
      ScaleHeight     =   1095
      ScaleWidth      =   615
      TabIndex        =   17
      Top             =   5400
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Timer TimerU 
      Interval        =   4000
      Left            =   6240
      Top             =   4320
   End
   Begin VB.Timer TimerS 
      Interval        =   2000
      Left            =   4200
      Top             =   4200
   End
   Begin VB.PictureBox picBoxU 
      BackColor       =   &H80000012&
      BorderStyle     =   0  'None
      Height          =   1095
      Left            =   6000
      Picture         =   "frmBio.frx":022E
      ScaleHeight     =   1095
      ScaleWidth      =   975
      TabIndex        =   16
      Top             =   5400
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.PictureBox picBoxS 
      BackColor       =   &H80000012&
      BorderStyle     =   0  'None
      Height          =   1095
      Left            =   4440
      Picture         =   "frmBio.frx":050B
      ScaleHeight     =   1095
      ScaleWidth      =   735
      TabIndex        =   15
      Top             =   5400
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox picboxGoalie 
      BorderStyle     =   0  'None
      Height          =   1815
      Left            =   9360
      Picture         =   "frmBio.frx":086B
      ScaleHeight     =   1815
      ScaleWidth      =   1455
      TabIndex        =   14
      Top             =   6960
      Width           =   1455
   End
   Begin VB.ComboBox Goalies 
      BeginProperty Font 
         Name            =   "Gulim"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "frmBio.frx":20CC
      Left            =   7440
      List            =   "frmBio.frx":20D6
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   6480
      Width           =   1815
   End
   Begin VB.PictureBox picBoxDefense 
      BorderStyle     =   0  'None
      Height          =   1695
      Left            =   240
      Picture         =   "frmBio.frx":20F4
      ScaleHeight     =   1695
      ScaleWidth      =   2055
      TabIndex        =   11
      Top             =   6960
      Width           =   2055
   End
   Begin VB.ComboBox Defense 
      BeginProperty Font 
         Name            =   "Gulim"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "frmBio.frx":359C
      Left            =   2400
      List            =   "frmBio.frx":35AC
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   6480
      Width           =   1695
   End
   Begin VB.PictureBox picBoxMiddie 
      BorderStyle     =   0  'None
      Height          =   1455
      Left            =   9360
      Picture         =   "frmBio.frx":35E0
      ScaleHeight     =   1455
      ScaleWidth      =   1455
      TabIndex        =   8
      Top             =   4920
      Width           =   1455
   End
   Begin VB.ComboBox Midfield 
      BeginProperty Font 
         Name            =   "Gulim"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "frmBio.frx":44EB
      Left            =   7440
      List            =   "frmBio.frx":44FB
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   4440
      Width           =   1815
   End
   Begin VB.PictureBox picBoxAttack 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   240
      Picture         =   "frmBio.frx":4534
      ScaleHeight     =   1455
      ScaleWidth      =   1935
      TabIndex        =   5
      Top             =   4920
      Width           =   1935
   End
   Begin VB.ComboBox Attack 
      BeginProperty Font 
         Name            =   "Gulim"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "frmBio.frx":5A94
      Left            =   2280
      List            =   "frmBio.frx":5AA4
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   4440
      Width           =   1815
   End
   Begin VB.PictureBox picBoxSJU 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   0
      Picture         =   "frmBio.frx":5AE0
      ScaleHeight     =   855
      ScaleWidth      =   4455
      TabIndex        =   2
      Top             =   3600
      Width           =   4455
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back to the Main Page"
      Height          =   735
      Left            =   4200
      TabIndex        =   1
      Top             =   8280
      Width           =   2535
   End
   Begin VB.PictureBox picboxHeader 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   0
      Picture         =   "frmBio.frx":6C0B
      ScaleHeight     =   3495
      ScaleWidth      =   10935
      TabIndex        =   0
      Top             =   0
      Width           =   10935
   End
   Begin VB.Label lblCredit 
      BackColor       =   &H00000000&
      Caption         =   "Project by: Dan Gregus"
      BeginProperty Font 
         Name            =   "Gill Sans Ultra Bold Condensed"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   7320
      TabIndex        =   18
      Top             =   8880
      Width           =   2535
   End
   Begin VB.Label lblGoalies 
      Caption         =   "   Goalies!"
      BeginProperty Font 
         Name            =   "Gill Sans Ultra Bold Condensed"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9360
      TabIndex        =   12
      Top             =   6480
      Width           =   1455
   End
   Begin VB.Label lblDefense 
      Caption         =   "      Defense!"
      BeginProperty Font 
         Name            =   "Gill Sans Ultra Bold Condensed"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Top             =   6480
      Width           =   2055
   End
   Begin VB.Label Middies 
      Caption         =   "Midfielders!"
      BeginProperty Font 
         Name            =   "Gill Sans Ultra Bold Condensed"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9360
      TabIndex        =   6
      Top             =   4440
      Width           =   1455
   End
   Begin VB.Label lblAttack 
      BackColor       =   &H8000000E&
      Caption         =   "On the Attack!"
      BeginProperty Font 
         Name            =   "Gill Sans Ultra Bold Condensed"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   4440
      Width           =   1935
   End
End
Attribute VB_Name = "frmBio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim player As String
'SJU Lacrosse Guide (Final Project 1.VBP)
'frmBio (frmBio.frm)
'Dan Gregus
'3/20/06
'Objective: To create a secondary page that is linked th the main page ->
    '(frmSJULacrosse) that displays an SJU graphic on three separate timers as ->
    'well as provide access to the player biography pages

'Adds attackmen to the box titeled Attack
Private Sub Attack_Change()
    Attack.AddItem ("John Carlson")
    Attack.AddItem ("Erick Peterson")
    Attack.AddItem ("Adam Rietz")
    Attack.AddItem ("Kyle Hinners")
    
    
    
End Sub

Private Sub Attack_Click()
    'The statements made here apply to all of the comboboxes on the page
    
    'This next line is IMPORTANT as it brings the list into the equation
    player = Attack.List(Attack.ListIndex)
    'If we get a new recruit we will need a new If statement
    'Brings selected entry to the desired page
    If player = "John Carlson" Then
        frmJC.Visible = True
    
    ElseIf player = "Erick Peterson" Then
        frmEP.Visible = True
    
    ElseIf player = "Adam Rietz" Then
        frmAR.Visible = True
    
    ElseIf player = "Kyle Hinners" Then
        frmKH.Visible = True
    End If
    
    frmBio.Visible = False
    
End Sub

'Brings the user back to the main page
Private Sub cmdBack_Click()
    frmSJULacrosse.Visible = True
    frmBio.Visible = False
End Sub

Private Sub cmdInstructions_Click()
MsgBox "In order to view a player profile, simply click on a combo box and select his name", , "Page instructions"
End Sub

Private Sub Defense_Change()
    Defense.AddItem ("Adam Benny")
    Defense.AddItem ("Alex Kady")
    Defense.AddItem ("Brian Jensen")
    Defense.AddItem ("Tim Herby")
End Sub

Private Sub Defense_Click()
    
    'This next line is IMPORTANT
    player = Defense.List(Defense.ListIndex)
    'If we get a new recruit we will need a new If statement
    If player = "Adam Benny" Then
        frmAB.Visible = True
    
    ElseIf player = "Brian Jensen" Then
        frmBJ.Visible = True
    
    ElseIf player = "Tim Herby" Then
        frmTH.Visible = True
    
    ElseIf player = "Alex Kady" Then
        frmAK.Visible = True
    End If
    
    frmBio.Visible = False
    
End Sub

Private Sub Goalies_Change()
    Goalies.AddItem ("Will Durbin")
    Goalies.AddItem ("John Broich")
End Sub
Private Sub Goalies_Click()
    
    'This next line is IMPORTANT
    player = Goalies.List(Goalies.ListIndex)
    'If we get a new recruit we will need a new If statement
    If player = "Will Durbin" Then
        frmWD.Visible = True
    
    ElseIf player = "John Broich" Then
        frmJBr.Visible = True
    
    End If
    
    frmBio.Visible = False
  
End Sub

Private Sub Midfield_Change()
    Midfield.AddItem ("Dan Gregus")
    Midfield.AddItem ("Mark Bachand")
    Midfield.AddItem ("Joe Boone")
    Midfield.AddItem ("Justin Gervais")

End Sub


Private Sub Midfield_Click()
    
    'This next line is IMPORTANT
    player = Midfield.List(Midfield.ListIndex)
    'If we get a new recruit we will need a new If statement
    If player = "Dan Gregus" Then
        frmDG.Visible = True
    
    ElseIf player = "Mark Bachand" Then
        frmMB.Visible = True
    
    ElseIf player = "Joe Boone" Then
        frmJB.Visible = True
    
    ElseIf player = "Justin Gervais" Then
        frmJG.Visible = True
    End If
    
    frmBio.Visible = False
    
End Sub

'Timer here displays the J picture at a predetermined interval
Private Sub TimerJ_Timer()
    picBoxJ.Visible = True
    TimerJ = False

End Sub

'Timer here displays the S picture at a predetermined interval
Private Sub TimerS_Timer()
    picBoxS.Visible = True
    TimerS = False

End Sub

'Timer here displays the U picture at a predetermined interval
Private Sub TimerU_Timer()
    picBoxU.Visible = True
    TimerU = False

End Sub
