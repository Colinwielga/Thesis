VERSION 5.00
Begin VB.Form frmStormBball 
   BackColor       =   &H0000C000&
   Caption         =   "Storm 10th Grade Basketball"
   ClientHeight    =   7110
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9810
   LinkTopic       =   "Form1"
   ScaleHeight     =   7110
   ScaleWidth      =   9810
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSources 
      BackColor       =   &H0000FFFF&
      Caption         =   "View Bibliography"
      Height          =   1335
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5400
      Width           =   2295
   End
   Begin VB.CommandButton cmdLink 
      BackColor       =   &H0000FFFF&
      Caption         =   "Follow the link to the Sauk Rapids High School Website"
      Height          =   1335
      Left            =   7440
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3120
      Width           =   2055
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H0000FFFF&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5280
      Width           =   2175
   End
   Begin VB.CommandButton cmdFindfact 
      Caption         =   "Find a fact!"
      Height          =   615
      Left            =   4320
      TabIndex        =   5
      Top             =   6120
      Width           =   1455
   End
   Begin VB.TextBox txtQuestion 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3960
      TabIndex        =   3
      Top             =   5280
      Width           =   2295
   End
   Begin VB.CommandButton cmdShootingpercent 
      BackColor       =   &H0000FFFF&
      Caption         =   "Find a Players Shooting Percentage"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2880
      Width           =   2535
   End
   Begin VB.CommandButton cmdSchedule 
      BackColor       =   &H0000FFFF&
      Caption         =   "See Schedule"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   7440
      MaskColor       =   &H0000FFFF&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   360
      Width           =   2235
   End
   Begin VB.CommandButton CmdRoster 
      BackColor       =   &H0000FFFF&
      Caption         =   "Display Team Roster"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      MaskColor       =   &H8000000F&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   360
      Width           =   2175
   End
   Begin VB.Label lblFront 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      Caption         =   "2006-2007 Storm 10th Grade Boys Basketball"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3000
      TabIndex        =   7
      Top             =   2040
      Width           =   4215
   End
   Begin VB.Label lblQuestion 
      Alignment       =   2  'Center
      Caption         =   "Enter your favorite college basketball team for a fun fact about them!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3480
      TabIndex        =   4
      Top             =   4080
      Width           =   3375
   End
   Begin VB.Image Image1 
      Height          =   1635
      Left            =   3720
      Picture         =   "frmStormBball.frx":0000
      Stretch         =   -1  'True
      Top             =   240
      Width           =   2580
   End
End
Attribute VB_Name = "frmStormBball"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdFindfact_Click()
If txtQuestion.Text = "Duke" Then
    MsgBox "Duke is the GREATEST!", , "FACT"
Else: MsgBox "Your favorite team sucks compared to Duke", , "FACT"
End If
End Sub

Private Sub cmdLink_Click()
frmLink.Show
frmStormBball.Hide
End Sub

Private Sub cmdQuit_Click()
End
End Sub


Private Sub CmdRoster_Click()
frmStormBball.Hide
frmRoster.Show
End Sub

Private Sub cmdSchedule_Click()
Dim MonthNumber As Single
MonthNumber = InputBox("Enter the number of the month (1-12) you wish to see", "Schedule")
Select Case MonthNumber
    Case Is = 11
        frmStormBball.Hide
        frmNovember.Show
    Case 12
        frmStormBball.Hide
        frmDecember.Show
    Case 1
        frmStormBball.Hide
        frmJanuary.Show
    Case 2
        frmStormBball.Hide
        frmFebruary.Show
    Case 3
        frmStormBball.Hide
        frmMarch.Show
    Case 4 To 10
        MsgBox "Basketball season runs from November (11) through March (3)", , "Try Again"
    End Select

End Sub






Private Sub cmdShootingpercent_Click()
frmStormBball.Hide 'this will hide the main form
frmPercentage.Show 'this brings up the shooting percentage form
End Sub

Private Sub Command1_Click()
frmLink.Show
frmStormBball.Hide
End Sub

Private Sub cmdSources_Click()
frmStormBball.Hide
frmSources.Show
End Sub
