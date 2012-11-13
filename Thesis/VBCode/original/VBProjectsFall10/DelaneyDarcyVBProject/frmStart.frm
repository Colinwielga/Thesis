VERSION 5.00
Begin VB.Form frmStart 
   BackColor       =   &H8000000D&
   Caption         =   "Cash Flow Fantasy"
   ClientHeight    =   11760
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   16530
   LinkTopic       =   "Form1"
   PaletteMode     =   2  'Custom
   ScaleHeight     =   11760
   ScaleWidth      =   16530
   WindowState     =   2  'Maximized
   Begin VB.PictureBox PicHiring 
      Height          =   4215
      Left            =   5280
      Picture         =   "frmStart.frx":0000
      ScaleHeight     =   4155
      ScaleWidth      =   5115
      TabIndex        =   23
      Top             =   5880
      Width           =   5175
   End
   Begin VB.PictureBox PicJobs 
      Height          =   5175
      Left            =   5280
      Picture         =   "frmStart.frx":19E59
      ScaleHeight     =   5115
      ScaleWidth      =   5115
      TabIndex        =   22
      Top             =   720
      Width           =   5175
   End
   Begin VB.CommandButton cmdSelectProfession 
      BackColor       =   &H00FF8080&
      Caption         =   "Select A Profession"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   9840
      Width           =   10455
   End
   Begin VB.PictureBox picGrad 
      Height          =   7815
      Left            =   3000
      Picture         =   "frmStart.frx":34583
      ScaleHeight     =   7755
      ScaleWidth      =   10275
      TabIndex        =   20
      Top             =   1800
      Width           =   10335
   End
   Begin VB.CommandButton cmdMortician 
      BackColor       =   &H00000080&
      Caption         =   "Mortician"
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   9840
      Width           =   2295
   End
   Begin VB.CommandButton cmdArchitect 
      BackColor       =   &H00000080&
      Caption         =   "Architect"
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   13680
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   6240
      Width           =   2295
   End
   Begin VB.CommandButton cmdBanker 
      BackColor       =   &H00000080&
      Caption         =   "Banker"
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   8640
      Width           =   2295
   End
   Begin VB.CommandButton cmdTeacher 
      BackColor       =   &H00000080&
      Caption         =   "Teacher"
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   13680
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   5040
      Width           =   2295
   End
   Begin VB.CommandButton cmdPolice 
      BackColor       =   &H00000080&
      Caption         =   "Police Officer"
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   13680
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   3840
      Width           =   2295
   End
   Begin VB.CommandButton cmdPolitician 
      BackColor       =   &H00000080&
      Caption         =   "Politician"
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   13680
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   2640
      Width           =   2295
   End
   Begin VB.CommandButton cmdProAthlete 
      BackColor       =   &H00000080&
      Caption         =   "Professional Athlete"
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   13680
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   1440
      Width           =   2295
   End
   Begin VB.CommandButton cmdActor 
      BackColor       =   &H00000080&
      Caption         =   "Actor"
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   7440
      Width           =   2295
   End
   Begin VB.CommandButton cmdPilot 
      BackColor       =   &H00000080&
      Caption         =   "Pilot"
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   13680
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   7440
      Width           =   2295
   End
   Begin VB.CommandButton cmdSecretary 
      BackColor       =   &H00000080&
      Caption         =   "Secretary"
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   13680
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   240
      Width           =   2295
   End
   Begin VB.CommandButton cmdCeo 
      BackColor       =   &H00000080&
      Caption         =   "CEO"
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   13680
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   8640
      Width           =   2295
   End
   Begin VB.CommandButton cmdNurse 
      BackColor       =   &H00000080&
      Caption         =   "Nurse"
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6240
      Width           =   2295
   End
   Begin VB.CommandButton cmdMechanic 
      BackColor       =   &H00000080&
      Caption         =   "Mechanic"
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   13680
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   9840
      Width           =   2295
   End
   Begin VB.CommandButton cmdEngineer 
      BackColor       =   &H00000080&
      Caption         =   "Engineer"
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5040
      Width           =   2295
   End
   Begin VB.CommandButton cmdProfessor 
      BackColor       =   &H00000080&
      Caption         =   "Professor"
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3840
      Width           =   2295
   End
   Begin VB.CommandButton cmdAccountant 
      BackColor       =   &H00000080&
      Caption         =   "Accountant"
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2640
      Width           =   2295
   End
   Begin VB.CommandButton cmdLawyer 
      BackColor       =   &H00000080&
      Caption         =   "Lawyer"
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1440
      Width           =   2295
   End
   Begin VB.CommandButton cmdDoctor 
      BackColor       =   &H00000080&
      Caption         =   "Doctor"
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   240
      Width           =   2295
   End
   Begin VB.CommandButton cmdStart 
      BackColor       =   &H0000FFFF&
      Caption         =   "Start"
      BeginProperty Font 
         Name            =   "Snap ITC"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2115
      Left            =   5160
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3600
      Width           =   3015
   End
   Begin VB.PictureBox picResults1 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   975
      Left            =   3000
      ScaleHeight     =   915
      ScaleWidth      =   10275
      TabIndex        =   0
      Top             =   360
      Width           =   10335
   End
   Begin VB.Label lblStart 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      Caption         =   $"frmStart.frx":478AA
      BeginProperty Font 
         Name            =   "High Tower Text"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1815
      Left            =   2160
      TabIndex        =   24
      Top             =   1440
      Width           =   9615
   End
End
Attribute VB_Name = "frmStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This form is the starting point of the program. It houses the professions there are to choose from and leads the user to
'Transition into the invividual profession forms. All the profession select buttons assign the relevat profession salary to variable
Option Explicit

'Retrieves salary and assign it to the specified professions' variable
Private Function getSalary(ByVal profession As String) As Single
    Dim Pos As Integer
    Dim Found As Boolean
    Do Until Found = True Or Pos >= ctr
        Pos = Pos + 1
        If LCase(Professions(Pos)) = LCase(profession) Then
            Found = True
        End If
    Loop
    If Found = True Then
        getSalary = Salaries(Pos)
    Else
         getSalary = -1
    End If
End Function
'assign salary to variable
Private Sub cmdAccountant_Click()
    yourSalary = getSalary("Accountant")
    frmStart.Visible = False
    frmAccountant.Visible = True

End Sub
'assign salary to variable
Private Sub cmdActor_Click()
    yourSalary = getSalary("Actor")
    frmStart.Visible = False
    frmActor.Visible = True

End Sub
'assign salary to variable
Private Sub cmdArchitect_Click()
    frmStart.Visible = False
    frmArchitect.Visible = True
    yourSalary = getSalary("Architect")
End Sub

Private Sub cmdBanker_Click()
    yourSalary = getSalary("Banker")
    frmStart.Visible = False
    frmBanker.Visible = True

End Sub

Private Sub cmdCeo_Click()
    yourSalary = getSalary("CEO")
    frmStart.Visible = False
    frmCEO.Visible = True

End Sub

Private Sub cmdDoctor_Click()
    yourSalary = getSalary("Doctor")
    frmStart.Visible = False
    frmDoctor.Visible = True

End Sub

Private Sub cmdEngineer_Click()
    yourSalary = getSalary("Engineer")
    frmStart.Visible = False
    frmEngineer.Visible = True

End Sub

Private Sub cmdLawyer_Click()
    yourSalary = getSalary("Lawyer")
    frmStart.Visible = False
    frmLawyer.Visible = True

End Sub

Private Sub cmdMechanic_Click()
    yourSalary = getSalary("Mechanic")
    frmStart.Visible = False
    frmMechanic.Visible = True

End Sub

Private Sub cmdMortician_Click()
    yourSalary = getSalary("Mortician")
    frmStart.Visible = False
    frmMortician.Visible = True

End Sub

Private Sub cmdNurse_Click()
    yourSalary = getSalary("Nurse")
    frmStart.Visible = False
    frmNurse.Visible = True


End Sub

Private Sub cmdPilot_Click()
    yourSalary = getSalary("Pilot")
    frmStart.Visible = False
    frmPilot.Visible = True

End Sub

Private Sub cmdPolice_Click()
    yourSalary = getSalary("Police")
    frmStart.Visible = False
frmPolice.Visible = True

End Sub

Private Sub cmdPolitician_Click()
    yourSalary = getSalary("Politician")
    frmStart.Visible = False
    frmPolitician.Visible = True

End Sub

Private Sub cmdProAthlete_Click()
    yourSalary = getSalary("Athlete")
    frmStart.Visible = False
    frmAthlete.Visible = True

End Sub

Private Sub cmdProfessor_Click()
    yourSalary = getSalary("Professor")
    frmStart.Visible = False
    frmProfessor.Visible = True

End Sub

Private Sub cmdSecretary_Click()
    yourSalary = getSalary("Secretary")
    frmStart.Visible = False
    frmSecretary.Visible = True

End Sub
'makes selecy profession buttons visible
Private Sub cmdSelectProfession_Click()
    cmdSelectProfession.Visible = False
    picResults1.Visible = False
    picGrad.Visible = False
    PicJobs.Visible = True
    PicHiring.Visible = True
    cmdDoctor.Visible = True
    cmdLawyer.Visible = True
    cmdAccountant.Visible = True
    cmdProfessor.Visible = True
    cmdEngineer.Visible = True
    cmdCeo.Visible = True
    cmdMortician.Visible = True
    cmdSecretary.Visible = True
    cmdTeacher.Visible = True
    cmdPilot.Visible = True
    cmdNurse.Visible = True
    cmdActor.Visible = True
    cmdBanker.Visible = True
    cmdMechanic.Visible = True
    cmdPolice.Visible = True
    cmdProAthlete.Visible = True
    cmdArchitect.Visible = True
    cmdPolitician.Visible = True

'Reads salary information into arrays
Open App.Path & "\Salary.txt" For Input As #1
    ctr = 0
    Do Until EOF(1)
        ctr = ctr + 1
        Input #1, Professions(ctr), Salaries(ctr)
    Loop
Close #1

End Sub
'Starts program and leads to the profession selections screen
Private Sub cmdStart_Click()
    cmdStart.Visible = False
    lblStart.Visible = False
    cmdSelectProfession.Visible = True
    picResults1.Visible = True
    picResults1.Cls
    picGrad.Visible = True
    picGrad.Cls
    picResults1.Print "Congratulations " & UserName & "! You have just graduated from college. "
    picResults1.Print "Now, it is time to get a job! Please select a profession!"

End Sub
'Assign teacher's salary to unique variable
Private Sub cmdTeacher_Click()
    yourSalary = getSalary("Teacher")
    frmStart.Visible = False
    frmTeacher.Visible = True

End Sub
' Load form
Private Sub Form_Load()
    cmdSelectProfession.Visible = False
    picResults1.Visible = False
    picGrad.Visible = False
    PicJobs.Visible = False
    PicHiring.Visible = False
    cmdDoctor.Visible = False
    cmdLawyer.Visible = False
    cmdAccountant.Visible = False
    cmdProfessor.Visible = False
    cmdEngineer.Visible = False
    cmdCeo.Visible = False
    cmdMortician.Visible = False
    cmdSecretary.Visible = False
    cmdTeacher.Visible = False
    cmdPilot.Visible = False
    cmdNurse.Visible = False
    cmdActor.Visible = False
    cmdBanker.Visible = False
    cmdMechanic.Visible = False
    cmdPolice.Visible = False
    cmdProAthlete.Visible = False
    cmdArchitect.Visible = False
    cmdPolitician.Visible = False
End Sub

