VERSION 5.00
Begin VB.Form frmQuarter 
   BackColor       =   &H00008000&
   Caption         =   "Select a Quarter"
   ClientHeight    =   6240
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8100
   LinkTopic       =   "Form1"
   ScaleHeight     =   6240
   ScaleWidth      =   8100
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtQuarter 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000004&
      Height          =   735
      Left            =   5280
      TabIndex        =   3
      Text            =   "0"
      Top             =   1440
      Width           =   2175
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H8000000E&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5400
      Width           =   1095
   End
   Begin VB.CommandButton cmdQuery 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Enter"
      DisabledPicture =   "frmQuarter.frx":0000
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5280
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2520
      Width           =   2175
   End
   Begin VB.Label lblSelect 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "What Quarter Are You in?"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   1035
      Left            =   4680
      TabIndex        =   1
      Top             =   240
      Width           =   3285
   End
   Begin VB.Image imgFootball 
      Height          =   7200
      Left            =   0
      Picture         =   "frmQuarter.frx":A254
      Top             =   0
      Width           =   9600
   End
End
Attribute VB_Name = "frmQuarter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    'Football Playcalling sheet
    'frmQuarter
    'Ben Dimond
    '10/16/09
        'This program takes user input to gather information about a football game and produces
        'an output with plays that would work best in each situation
        'The background picture was found at this web address: http://www.cnycentral.com/uploadedImages/Shared/Sports/National_stories/Football.jpg

Dim notFirst As Boolean

    
Private Sub cmdQuery_Click()
    'Clicking this button tells the program what quarter it is
       'the position of the ball, the down and the distance to a first down based on user input
    
    'Define Quarter
    Quarter = txtQuarter.Text
    
    'This loop will make sure the numbers input by user are valid
    If Quarter <= 4 And Quarter >= 1 Then
        Position = InputBox("Please enter the position of the Ball (-49(1-yardline) through -1(49-yardline) for own side and 50 through 1 for opposing side of the field):", "What is the Position of the Ball?")
            If Position > 50 Or Position < -49 Or Position = 0 Then
                MsgBox "Please Enter a valid Ball Position", , "Uh oh!"
            Else
                Down = InputBox("Please enter the Down:", "What Down is it?")
                If Down > 4 Or Down < 1 Then
                     MsgBox "Please Enter a valid Down", , "Uh oh!"
                Else
                   Distance = InputBox("Please Enter the Distance:", "How much to go?")
                   If Distance > 100 Or Distance <= 0 Then
                     MsgBox "Please enter a valid Distance", , "Uh oh!"
                   Else
                     MsgBox "You are in the " & Quarter & " quarter. The ball is on the " & Position & " yard line. It is " & Down & " down with " & Distance & " yard(s) to go.", , "Just Double-Checking:"
                    frmPlays.Show
                    frmQuarter.Visible = False
                   End If
                End If
            End If
    Else
        MsgBox "Please Enter a valid Quarter", , "Uh oh!"
    End If
    
    
    
End Sub

Private Sub Form_Load()
'This ensures that the welcome screen shows up first
  If Not notFirst Then
    frmWelcome.Show
    frmQuarter.Hide
    notFirst = True
  End If
End Sub



Private Sub cmdQuit_Click()
    MsgBox "Go get em'!", , "Good Luck!"
    End
End Sub
