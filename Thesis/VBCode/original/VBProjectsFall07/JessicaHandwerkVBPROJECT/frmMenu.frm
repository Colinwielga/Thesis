VERSION 5.00
Begin VB.Form frmMenu 
   BackColor       =   &H80000006&
   Caption         =   "Form1"
   ClientHeight    =   8190
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9885
   FillColor       =   &H00E0E0E0&
   LinkTopic       =   "Form1"
   ScaleHeight     =   8190
   ScaleWidth      =   9885
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdFunds 
      Caption         =   "Funds, Donations, and Sponsors"
      BeginProperty Font 
         Name            =   "@Gungsuh"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   6
      Left            =   480
      TabIndex        =   8
      Top             =   4320
      Width           =   2895
   End
   Begin VB.CommandButton cmdVolunteer 
      Caption         =   "Volunteer"
      BeginProperty Font 
         Name            =   "@Gungsuh"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   4
      Left            =   480
      TabIndex        =   7
      Top             =   5160
      Width           =   2895
   End
   Begin VB.CommandButton cmdHabitat 
      Caption         =   "Habitat For Humanity"
      BeginProperty Font 
         Name            =   "@Gungsuh"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   3
      Left            =   480
      TabIndex        =   6
      Top             =   2640
      Width           =   2895
   End
   Begin VB.CommandButton cmdRequirements 
      Caption         =   "Requirements for Homeowners"
      BeginProperty Font 
         Name            =   "@Gungsuh"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   1
      Left            =   480
      TabIndex        =   5
      Top             =   3480
      Width           =   2895
   End
   Begin VB.CommandButton cmdComm 
      Caption         =   "Communication 387"
      BeginProperty Font 
         Name            =   "@Gungsuh"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   0
      Left            =   480
      TabIndex        =   4
      Top             =   1800
      Width           =   2895
   End
   Begin VB.CommandButton cmdTryIt 
      Caption         =   "See the Habitat Website!"
      BeginProperty Font 
         Name            =   "@Gungsuh"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   5040
      TabIndex        =   2
      Top             =   6840
      Width           =   3375
   End
   Begin VB.CommandButton cmdWorksCited 
      Caption         =   "View Works Cited"
      BeginProperty Font 
         Name            =   "@Gungsuh"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   480
      TabIndex        =   1
      Top             =   6000
      Width           =   2895
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Adobe Garamond Pro"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   8400
      TabIndex        =   0
      Top             =   6840
      Width           =   495
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "Habitat for Humanity"
      BeginProperty Font 
         Name            =   "Tekton Pro Ext"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   480
      TabIndex        =   10
      Top             =   240
      Width           =   2895
   End
   Begin VB.Line Line7 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   12
      X1              =   480
      X2              =   3360
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line6 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   10
      X1              =   480
      X2              =   3360
      Y1              =   8160
      Y2              =   8160
   End
   Begin VB.Line Line5 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   10
      Index           =   1
      X1              =   360
      X2              =   360
      Y1              =   0
      Y2              =   8160
   End
   Begin VB.Line Line5 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   10
      Index           =   0
      X1              =   3480
      X2              =   3480
      Y1              =   0
      Y2              =   8160
   End
   Begin VB.Line Line4 
      BorderColor     =   &H8000000D&
      BorderStyle     =   5  'Dash-Dot-Dot
      BorderWidth     =   2
      X1              =   4560
      X2              =   9000
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line Line3 
      BorderColor     =   &H8000000D&
      BorderStyle     =   5  'Dash-Dot-Dot
      BorderWidth     =   2
      X1              =   9000
      X2              =   9000
      Y1              =   120
      Y2              =   6120
   End
   Begin VB.Line Line2 
      BorderColor     =   &H8000000D&
      BorderStyle     =   5  'Dash-Dot-Dot
      BorderWidth     =   2
      X1              =   4560
      X2              =   4560
      Y1              =   120
      Y2              =   6120
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000D&
      BorderStyle     =   5  'Dash-Dot-Dot
      BorderWidth     =   2
      X1              =   9000
      X2              =   4560
      Y1              =   6120
      Y2              =   6120
   End
   Begin VB.Label lblMainMenu 
      BackColor       =   &H80000012&
      Caption         =   "Main Menu"
      BeginProperty Font 
         Name            =   "Tekton Pro Ext"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   855
      Left            =   720
      TabIndex        =   9
      Top             =   600
      Width           =   2535
   End
   Begin VB.Label lblBuildYours 
      BackColor       =   &H80000012&
      Caption         =   "Build Your Own Home!"
      BeginProperty Font 
         Name            =   "@Gungsuh"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   615
      Left            =   4920
      TabIndex        =   3
      Top             =   6360
      Width           =   3855
   End
   Begin VB.Image Image1 
      Height          =   5685
      Left            =   4920
      Picture         =   "frmMenu.frx":0000
      Top             =   240
      Width           =   3720
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdComm_Click(Index As Integer)
   'move to student form
    frmStudents.Show
    frmMenu.Hide

End Sub

Private Sub cmdExit_Click()
    End

End Sub

Private Sub cmdFunds_Click(Index As Integer)
    'move to funds form
    frmFunds.Show
    frmMenu.Hide

End Sub

Private Sub cmdHabitat_Click(Index As Integer)
    'move to habitat form
    frmHabitat.Show
    frmMenu.Hide
End Sub

Private Sub cmdRequirements_Click(Index As Integer)
    'move to requirement form
    frmReqi.Show
    frmMenu.Hide

End Sub

Private Sub cmdTryIt_Click()
    'allows the user to get to webpage
    Dim Web As Long
    Web = ShellExecute(frmMenu.hwnd, "Open", "www.centralminnesotahabitat.org", vbNullString, vbNullString, SW_ShowNormal)
End Sub

Private Sub cmdVolunteer_Click(Index As Integer)
    'move to volunteer form
    frmVolunteer.Show
    frmMenu.Hide
End Sub

Private Sub cmdWorksCited_Click()
frmWorksCited.Show
frmMenu.Hide

End Sub



