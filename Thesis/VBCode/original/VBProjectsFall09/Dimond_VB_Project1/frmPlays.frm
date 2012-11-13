VERSION 5.00
Begin VB.Form frmPlays 
   BackColor       =   &H00004000&
   Caption         =   "Suggested Plays"
   ClientHeight    =   12975
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14970
   LinkTopic       =   "Form1"
   ScaleHeight     =   12975
   ScaleWidth      =   14970
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CommandButton cmdReturn 
      BackColor       =   &H0000FFFF&
      Caption         =   "Return to Home Screen"
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
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   3000
      Width           =   3015
   End
   Begin VB.CommandButton cmdFieldGoal 
      BackColor       =   &H80000009&
      Caption         =   "Field Goal"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11640
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   10560
      Width           =   3135
   End
   Begin VB.CommandButton cmdDoubleOption 
      BackColor       =   &H80000009&
      Caption         =   "Double Option"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   360
      Width           =   3135
   End
   Begin VB.CommandButton cmdPunt 
      Caption         =   "Punt"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   11
      Top             =   10560
      Width           =   3135
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H0000FFFF&
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
      Left            =   6000
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   9240
      Width           =   3135
   End
   Begin VB.PictureBox picResults1 
      Height          =   5175
      Left            =   3600
      ScaleHeight     =   5115
      ScaleWidth      =   7755
      TabIndex        =   9
      Top             =   3840
      Width           =   7815
   End
   Begin VB.CommandButton cmdHBPass 
      BackColor       =   &H80000009&
      Caption         =   "Halfback Pass"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   11640
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   8520
      Width           =   3135
   End
   Begin VB.CommandButton cmdPAPass 
      BackColor       =   &H80000009&
      Caption         =   "Play-action Pass"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11640
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6600
      Width           =   3135
   End
   Begin VB.CommandButton cmdSlantFades 
      BackColor       =   &H80000009&
      Caption         =   "Slant-Fades"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11640
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4560
      Width           =   3135
   End
   Begin VB.CommandButton cmdStops 
      BackColor       =   &H80000009&
      Caption         =   "Quick Stops"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11640
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2400
      Width           =   3135
   End
   Begin VB.CommandButton cmdHailMary 
      BackColor       =   &H80000009&
      Caption         =   "Hail Mary"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11640
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   360
      Width           =   3135
   End
   Begin VB.CommandButton cmdSweep 
      BackColor       =   &H80000009&
      Caption         =   "Sweep"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   8520
      Width           =   3135
   End
   Begin VB.CommandButton cmdQBDraw 
      BackColor       =   &H80000009&
      Caption         =   "Quarterback Draw"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6600
      Width           =   3135
   End
   Begin VB.CommandButton cmdStraightShots 
      BackColor       =   &H80000009&
      Caption         =   "Straight Shots"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4560
      Width           =   3135
   End
   Begin VB.CommandButton cmdInsideTrap 
      BackColor       =   &H80000009&
      Caption         =   "Inside Trap"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2400
      Width           =   3135
   End
   Begin VB.Image Image2 
      Height          =   2700
      Left            =   6840
      Picture         =   "frmPlays.frx":0000
      Top             =   10080
      Width           =   1350
   End
   Begin VB.Image Image1 
      Height          =   2700
      Left            =   6840
      Picture         =   "frmPlays.frx":08A6
      Top             =   120
      Width           =   1350
   End
   Begin VB.Image ImgGrass9 
      Height          =   4755
      Left            =   10080
      Picture         =   "frmPlays.frx":1180
      Top             =   9360
      Width           =   5145
   End
   Begin VB.Image ImgGrass8 
      Height          =   4755
      Left            =   5040
      Picture         =   "frmPlays.frx":EC25
      Top             =   9240
      Width           =   5145
   End
   Begin VB.Image ImgGrass7 
      Height          =   4755
      Left            =   0
      Picture         =   "frmPlays.frx":1C6CA
      Top             =   9240
      Width           =   5145
   End
   Begin VB.Image ImgGrass6 
      Height          =   4755
      Left            =   9960
      Picture         =   "frmPlays.frx":2A16F
      Top             =   4680
      Width           =   5145
   End
   Begin VB.Image ImgGrass5 
      Height          =   4755
      Left            =   5040
      Picture         =   "frmPlays.frx":37C14
      Top             =   4680
      Width           =   5145
   End
   Begin VB.Image ImgGrass4 
      Height          =   4755
      Left            =   0
      Picture         =   "frmPlays.frx":456B9
      Top             =   4680
      Width           =   5145
   End
   Begin VB.Image ImgGrass3 
      Height          =   4755
      Left            =   10080
      Picture         =   "frmPlays.frx":5315E
      Top             =   0
      Width           =   5145
   End
   Begin VB.Image ImgGrass2 
      Height          =   4755
      Left            =   5040
      Picture         =   "frmPlays.frx":60C03
      Top             =   0
      Width           =   5145
   End
   Begin VB.Image ImgGrass1 
      Height          =   4755
      Left            =   0
      Picture         =   "frmPlays.frx":6E6A8
      Top             =   0
      Width           =   5145
   End
End
Attribute VB_Name = "frmPlays"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Football Playcalling sheet
'frmPlays
'Ben Dimond
'10/16/09
    'This form has cmdButtons for each play that are visible/invisible depending on the Quarter,
    'Down,Posistion and Distance given by the user
    'The field goal picture was found at this URL: http://www.ripten.com/wp-content/uploads/2008/07/field-goal-post1.jpg
    'The grass picture was found at this URL: http://www.flooringwebsite.com/Carpet/Indoor_Outdoor/Grass-Tex/Board_1/Front/All_Sports_Turf/All_Sports_Turf.jpg

'This will Print the desired play picture in the picture box (adapted from LordOfTheRings)
Private Sub cmdFieldGoal_Click()
    picResults1.Picture = LoadPicture(App.Path & "\FieldGoal.jpg") 'Loads the picture of the Field Goal Play
End Sub

Private Sub cmdHailMary_Click()
    picResults1.Picture = LoadPicture(App.Path & "\HailMary.jpg") 'Loads the picture of the Hail Mary Play
End Sub

Private Sub cmdHBPass_Click()
    picResults1.Picture = LoadPicture(App.Path & "\HBPass.jpg") 'Loads the picture of the Halfback Pass Play
End Sub

Private Sub cmdInsideTrap_Click()
    picResults1.Picture = LoadPicture(App.Path & "\InsideTrap.jpg") 'Loads the picture of the Inside Trap play
End Sub

Private Sub cmdPAPass_Click()
    picResults1.Picture = LoadPicture(App.Path & "\PAPass.jpg") 'Loads the picture of the Play-action Pass Play
End Sub

Private Sub cmdPunt_Click()
    picResults1.Picture = LoadPicture(App.Path & "\Punt.jpg") 'Loads the picture of the Punt Play
End Sub

Private Sub cmdQBDraw_Click()
    picResults1.Picture = LoadPicture(App.Path & "\QuarterbackDraw.jpg") 'Loads the picture of the Quarterback Draw Play
End Sub

Private Sub cmdReturn_Click() 'This returns the user to frmQuarter
    frmPlays.Visible = False
    frmQuarter.Show
        
End Sub

Private Sub cmdSlantFades_Click()
    picResults1.Picture = LoadPicture(App.Path & "\SlantFades.jpg") 'Loads the picture of the Slant-Fades Play
End Sub

Private Sub cmdStops_Click()
    picResults1.Picture = LoadPicture(App.Path & "\Stops.jpg") 'Loads the picture of the Stops Play
End Sub

Private Sub cmdStraightShots_Click()
    picResults1.Picture = LoadPicture(App.Path & "\StraightShots.jpg") 'Loads the picture of the Straight Shots Play
End Sub


Private Sub cmdSweep_Click()
    picResults1.Picture = LoadPicture(App.Path & "\Sweep.jpg") 'Loads the picture of the Sweep Play
End Sub


'This will clear the Picture box everytime the form is loaded (adapted from LordOfTheRings)
Private Sub Form_Activate()
    picResults1.Picture = LoadPicture("") 'This clears the picture box each time the form is opened

'This if-statement will analyze the user input and make the "Play" buttons visible or invisible based on that info
    If Quarter >= 1 And Quarter <= 3 Then
        Select Case Down
        Case Is = 1
            If Distance <= 10 And Distance >= 1 Then
                cmdDoubleOption.Enabled = True
                cmdInsideTrap.Enabled = True
                cmdStraightShots.Enabled = True
                cmdQBDraw.Enabled = False
                cmdSweep.Enabled = True
                cmdPunt.Enabled = False
                cmdHailMary.Enabled = False
                cmdStops.Enabled = True
                cmdSlantFades.Enabled = True
                cmdPAPass.Enabled = True
                cmdHBPass.Enabled = True
                cmdFieldGoal.Enabled = False
            Else
                cmdDoubleOption.Enabled = True
                cmdInsideTrap.Enabled = False
                cmdStraightShots.Enabled = False
                cmdQBDraw.Enabled = True
                cmdSweep.Enabled = True
                cmdPunt.Enabled = False
                cmdHailMary.Enabled = False
                cmdStops.Enabled = True
                cmdSlantFades.Enabled = True
                cmdPAPass.Enabled = True
                cmdHBPass.Enabled = True
                cmdFieldGoal.Enabled = False
            End If
        Case Is = 2
            If Distance <= 10 And Distance >= 1 Then
                    cmdDoubleOption.Enabled = True
                    cmdInsideTrap.Enabled = True
                    cmdStraightShots.Enabled = True
                    cmdQBDraw.Enabled = True
                    cmdSweep.Enabled = True
                    cmdPunt.Enabled = False
                    cmdHailMary.Enabled = False
                    cmdStops.Enabled = True
                    cmdSlantFades.Enabled = True
                    cmdPAPass.Enabled = True
                    cmdHBPass.Enabled = True
                    cmdFieldGoal.Enabled = False
                Else
                    cmdDoubleOption.Enabled = True
                    cmdInsideTrap.Enabled = False
                    cmdStraightShots.Enabled = False
                    cmdQBDraw.Enabled = True
                    cmdSweep.Enabled = True
                    cmdPunt.Enabled = False
                    cmdHailMary.Enabled = False
                    cmdStops.Enabled = True
                    cmdSlantFades.Enabled = True
                    cmdPAPass.Enabled = True
                    cmdHBPass.Enabled = True
                    cmdFieldGoal.Enabled = False
                End If
        Case Is = 3
                    If Distance <= 10 And Distance >= 1 Then
                        cmdDoubleOption.Enabled = True
                        cmdInsideTrap.Enabled = True
                        cmdStraightShots.Enabled = True
                        cmdQBDraw.Enabled = True
                        cmdSweep.Enabled = True
                        cmdPunt.Enabled = False
                        cmdHailMary.Enabled = False
                        cmdStops.Enabled = True
                        cmdSlantFades.Enabled = True
                        cmdPAPass.Enabled = True
                        cmdHBPass.Enabled = True
                        cmdFieldGoal.Enabled = False
                    Else
                        cmdDoubleOption.Enabled = True
                        cmdInsideTrap.Enabled = False
                        cmdStraightShots.Enabled = False
                        cmdQBDraw.Enabled = True
                        cmdSweep.Enabled = True
                        cmdPunt.Enabled = False
                        cmdHailMary.Enabled = False
                        cmdStops.Enabled = True
                        cmdSlantFades.Enabled = True
                        cmdPAPass.Enabled = True
                        cmdHBPass.Enabled = True
                        cmdFieldGoal.Enabled = False
                    End If
         Case Is = 4
                        If Distance < 5 Then
                            cmdDoubleOption.Enabled = True
                            cmdInsideTrap.Enabled = True
                            cmdStraightShots.Enabled = True
                            cmdQBDraw.Enabled = True
                            cmdSweep.Enabled = True
                            cmdPunt.Enabled = False
                            cmdHailMary.Enabled = False
                            cmdStops.Enabled = True
                            cmdSlantFades.Enabled = True
                            cmdPAPass.Enabled = True
                            cmdHBPass.Enabled = True
                            cmdFieldGoal.Enabled = False
                                           
                        Else
                            If Position > 20 Or Position < 0 Then
                                cmdDoubleOption.Enabled = False
                                cmdInsideTrap.Enabled = False
                                cmdStraightShots.Enabled = False
                                cmdQBDraw.Enabled = False
                                cmdSweep.Enabled = False
                                cmdPunt.Enabled = True
                                cmdHailMary.Enabled = False
                                cmdStops.Enabled = False
                                cmdSlantFades.Enabled = False
                                cmdPAPass.Enabled = False
                                cmdHBPass.Enabled = False
                                cmdFieldGoal.Enabled = False
                            Else
                                cmdDoubleOption.Enabled = False
                                cmdInsideTrap.Enabled = False
                                cmdStraightShots.Enabled = False
                                cmdQBDraw.Enabled = False
                                cmdSweep.Enabled = False
                                cmdPunt.Enabled = False
                                cmdHailMary.Enabled = False
                                cmdStops.Enabled = False
                                cmdSlantFades.Enabled = False
                                cmdPAPass.Enabled = False
                                cmdHBPass.Enabled = False
                                cmdFieldGoal.Enabled = True
                            End If
                        End If
        End Select
    Else
        Select Case Down
        Case Is = 1
            If Distance <= 10 And Distance >= 1 Then
                cmdDoubleOption.Enabled = True
                cmdInsideTrap.Enabled = True
                cmdStraightShots.Enabled = True
                cmdQBDraw.Enabled = False
                cmdSweep.Enabled = True
                cmdPunt.Enabled = False
                cmdHailMary.Enabled = False
                cmdStops.Enabled = True
                cmdSlantFades.Enabled = True
                cmdPAPass.Enabled = True
                cmdHBPass.Enabled = True
                cmdFieldGoal.Enabled = False
            Else
                cmdDoubleOption.Enabled = True
                cmdInsideTrap.Enabled = False
                cmdStraightShots.Enabled = False
                cmdQBDraw.Enabled = True
                cmdSweep.Enabled = True
                cmdPunt.Enabled = False
                cmdHailMary.Enabled = False
                cmdStops.Enabled = True
                cmdSlantFades.Enabled = True
                cmdPAPass.Enabled = True
                cmdHBPass.Enabled = True
                cmdFieldGoal.Enabled = False
            End If
        Case Is = 2
                If Distance <= 10 And Distance >= 1 Then
                    cmdDoubleOption.Enabled = True
                    cmdInsideTrap.Enabled = True
                    cmdStraightShots.Enabled = True
                    cmdQBDraw.Enabled = True
                    cmdSweep.Enabled = True
                    cmdPunt.Enabled = False
                    cmdHailMary.Enabled = False
                    cmdStops.Enabled = True
                    cmdSlantFades.Enabled = True
                    cmdPAPass.Enabled = True
                    cmdHBPass.Enabled = True
                    cmdFieldGoal.Enabled = False
                Else
                    cmdDoubleOption.Enabled = True
                    cmdInsideTrap.Enabled = False
                    cmdStraightShots.Enabled = False
                    cmdQBDraw.Enabled = True
                    cmdSweep.Enabled = True
                    cmdPunt.Enabled = False
                    cmdHailMary.Enabled = False
                    cmdStops.Enabled = True
                    cmdSlantFades.Enabled = True
                    cmdPAPass.Enabled = True
                    cmdHBPass.Enabled = True
                    cmdFieldGoal.Enabled = False
                End If
       Case Is = 3
                    If Distance <= 10 And Distance >= 1 Then
                        cmdDoubleOption.Enabled = True
                        cmdInsideTrap.Enabled = True
                        cmdStraightShots.Enabled = True
                        cmdQBDraw.Enabled = True
                        cmdSweep.Enabled = True
                        cmdPunt.Enabled = False
                        cmdHailMary.Enabled = False
                        cmdStops.Enabled = True
                        cmdSlantFades.Enabled = True
                        cmdPAPass.Enabled = True
                        cmdHBPass.Enabled = True
                        cmdFieldGoal.Enabled = False
                    Else
                        cmdDoubleOption.Enabled = True
                        cmdInsideTrap.Enabled = False
                        cmdStraightShots.Enabled = False
                        cmdQBDraw.Enabled = True
                        cmdSweep.Enabled = True
                        cmdPunt.Enabled = False
                        cmdHailMary.Enabled = False
                        cmdStops.Enabled = True
                        cmdSlantFades.Enabled = True
                        cmdPAPass.Enabled = True
                        cmdHBPass.Enabled = True
                        cmdFieldGoal.Enabled = False
                    End If
       Case Is = 4
                        If Distance < 5 Then
                            cmdDoubleOption.Enabled = True
                            cmdInsideTrap.Enabled = True
                            cmdStraightShots.Enabled = True
                            cmdQBDraw.Enabled = True
                            cmdSweep.Enabled = True
                            cmdPunt.Enabled = False
                            cmdHailMary.Enabled = False
                            cmdStops.Enabled = True
                            cmdSlantFades.Enabled = True
                            cmdPAPass.Enabled = True
                            cmdHBPass.Enabled = True
                            cmdFieldGoal.Enabled = True
                        Else
                            If Position > 20 Or Position < 0 Then
                                cmdDoubleOption.Enabled = False
                                cmdInsideTrap.Enabled = False
                                cmdStraightShots.Enabled = False
                                cmdQBDraw.Enabled = False
                                cmdSweep.Enabled = False
                                cmdPunt.Enabled = True
                                cmdHailMary.Enabled = True
                                cmdStops.Enabled = True
                                cmdSlantFades.Enabled = True
                                cmdPAPass.Enabled = True
                                cmdHBPass.Enabled = True
                                cmdFieldGoal.Enabled = False
                            Else
                                cmdDoubleOption.Enabled = False
                                cmdInsideTrap.Enabled = False
                                cmdStraightShots.Enabled = False
                                cmdQBDraw.Enabled = False
                                cmdSweep.Enabled = False
                                cmdPunt.Enabled = False
                                cmdHailMary.Enabled = True
                                cmdStops.Enabled = True
                                cmdSlantFades.Enabled = True
                                cmdPAPass.Enabled = True
                                cmdHBPass.Enabled = True
                                cmdFieldGoal.Enabled = True
                            End If
                        End If
        End Select
    End If
End Sub

Private Sub cmdQuit_Click()
    'quit button with a friendly message
    MsgBox "Go get em'!", , "Good Luck!"
    End
End Sub


Private Sub cmdDoubleOption_Click()
    picResults1.Picture = LoadPicture(App.Path & "\DoubleOption.jpg") 'This prints the picture for the Double Option Play
End Sub
