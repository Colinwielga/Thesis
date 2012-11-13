VERSION 5.00
Begin VB.Form frmEmployment 
   BackColor       =   &H000000C0&
   Caption         =   "Employment Information"
   ClientHeight    =   5775
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8520
   LinkTopic       =   "Form1"
   ScaleHeight     =   5775
   ScaleWidth      =   8520
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Back"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7440
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5160
      Width           =   975
   End
   Begin VB.CommandButton cmdJobFair 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Minnesota's Private Colleges Job and Internship Fair "
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3240
      Width           =   2535
   End
   Begin VB.CommandButton cmdSchools 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Current Employment Opportunities at MPCC Schools"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3240
      Width           =   2535
   End
   Begin VB.CommandButton cmdCurrent 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Current Employment Opportunities at MPCC"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1320
      Width           =   2535
   End
   Begin VB.CommandButton cmdCareerServices 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Click Here to Find Career Services contact information"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1320
      Width           =   2535
   End
   Begin VB.Label lblLooking 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Looking for a Job?"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4920
      TabIndex        =   2
      Top             =   240
      Width           =   2655
   End
   Begin VB.Label lblEmployers 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Information for Employers:"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   960
      TabIndex        =   1
      Top             =   240
      Width           =   2655
   End
End
Attribute VB_Name = "frmEmployment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ie As Object

Private Sub cmdCareerServices_Click()

'Enables text/label to be clicked to access webpage on Internet Explorer
'Source: http://www.mrexcel.com/forum/showthread.php?t=28421

Const url As String = "http://www.mnprivatecolleges.org/college_links/career_services.php"

    Set ie = CreateObject("internetexplorer.application")
    With ie
        .Visible = True
        .navigate url
    End With
    Set ie = Nothing

End Sub

Private Sub cmdCurrent_Click()

frmEmployment.Hide
frmAdmin.Show

End Sub

Private Sub cmdJobFair_Click()

'Enables text/label to be clicked to access webpage on Internet Explorer
'Source: http://www.mrexcel.com/forum/showthread.php?t=28421

Const url As String = "http://www.mnpcfair.org/"

    Set ie = CreateObject("internetexplorer.application")
    With ie
        .Visible = True
        .navigate url
    End With
    Set ie = Nothing

End Sub

Private Sub cmdSchools_Click()

'Enables text/label to be clicked to access webpage on Internet Explorer
'Source: http://www.mrexcel.com/forum/showthread.php?t=28421

Const url As String = "http://www.mnprivatecolleges.org/employment/index.php"

    Set ie = CreateObject("internetexplorer.application")
    With ie
        .Visible = True
        .navigate url
    End With
    Set ie = Nothing

End Sub

Private Sub Command1_Click()
frmEmployment.Hide
frmAboutMPCC.Show
End Sub

