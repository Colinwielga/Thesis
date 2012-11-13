VERSION 5.00
Begin VB.Form frmIntroCC 
   BackColor       =   &H00008000&
   Caption         =   "Introduction Page"
   ClientHeight    =   8220
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11550
   LinkTopic       =   "Form1"
   ScaleHeight     =   8220
   ScaleWidth      =   11550
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDate 
      BackColor       =   &H000000FF&
      Caption         =   "Show Date"
      Height          =   855
      Left            =   8040
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4200
      Width           =   1695
   End
   Begin VB.PictureBox picResults 
      Height          =   1695
      Left            =   7200
      ScaleHeight     =   1635
      ScaleWidth      =   3555
      TabIndex        =   8
      Top             =   2280
      Width           =   3615
   End
   Begin VB.CommandButton cmdSources 
      BackColor       =   &H000000FF&
      Caption         =   "Sources Used"
      Height          =   855
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   7080
      Width           =   1695
   End
   Begin VB.CommandButton cmdBiografies 
      BackColor       =   &H000000FF&
      Caption         =   "Info on Famous Runners"
      BeginProperty Font 
         Name            =   "MS UI Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5880
      Width           =   1695
   End
   Begin VB.CommandButton cmdVO2Max 
      BackColor       =   &H000000FF&
      Caption         =   "VO2 Max Calculator"
      BeginProperty Font 
         Name            =   "MS UI Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1560
      MaskColor       =   &H00C00000&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5880
      Width           =   1695
   End
   Begin VB.CommandButton cmdBMI 
      BackColor       =   &H000000FF&
      Caption         =   "BMI Calculator"
      BeginProperty Font 
         Name            =   "MS UI Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5400
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5880
      Width           =   1695
   End
   Begin VB.CommandButton cmdCalculator 
      BackColor       =   &H000000FF&
      Caption         =   "Pace Calculator"
      BeginProperty Font 
         Name            =   "MS UI Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3480
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5880
      Width           =   1695
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00FF0000&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "MS UI Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7080
      Width           =   1695
   End
   Begin VB.CommandButton cmdRaceResults 
      BackColor       =   &H000000FF&
      Caption         =   "Find Past Race Results"
      BeginProperty Font 
         Name            =   "MS UI Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7080
      Width           =   1695
   End
   Begin VB.Image Image2 
      Height          =   4350
      Left            =   360
      Picture         =   "frmIntroCC.frx":0000
      Top             =   1200
      Width           =   6450
   End
   Begin VB.Label lblRunningInfo 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      Caption         =   "Welcome to the Running Information Page!"
      BeginProperty Font 
         Name            =   "MS UI Gothic"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   1215
      Left            =   2160
      TabIndex        =   5
      Top             =   120
      Width           =   6135
   End
End
Attribute VB_Name = "frmIntroCC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'changes from homepage to runner's information page'
Private Sub cmdBiografies_Click()
    frmIntroCC.Hide
    frmRunners.Show
End Sub
'changes from homepage to body mass index page'
Private Sub cmdBMI_Click()
    frmIntroCC.Hide
    frmBMI.Show
End Sub
'Changes from Homepage to Pace Calculator page'
Private Sub cmdCalculator_Click()
    frmIntroCC.Hide
    frmPace.Show
End Sub
'Ends program'
Private Sub cmdQuit_Click()
    End
End Sub

Private Sub cmdRaceResults_Click() 'Changes from homepage to race results page'
    frmIntroCC.Hide
    frmRaceResults.Show
End Sub

'goes from homepage to sources used page'
Private Sub cmdSources_Click()
    frmSources.Show
    frmIntroCC.Hide
End Sub

Private Sub cmdDate_Click()
picResults.Cls
If cmdDate.BackColor <> vbRed Then
    cmdDate.BackColor = vbRed
Else
    cmdDate.BackColor = vbButtonFace
End If
picResults.Print ("The current date is: ")
picResults.Print (FormatDateTime(Date, 1))
'this step shows the correct time and date'
End Sub

'changes from homepage to VO2 max page'
Private Sub cmdVO2Max_Click()
    frmIntroCC.Hide
    frmVO2Max.Show
End Sub
