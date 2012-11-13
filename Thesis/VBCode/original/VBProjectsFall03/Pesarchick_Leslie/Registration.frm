VERSION 5.00
Begin VB.Form frmRegistration 
   BackColor       =   &H00404000&
   Caption         =   "Registration"
   ClientHeight    =   8640
   ClientLeft      =   1050
   ClientTop       =   1125
   ClientWidth     =   10860
   LinkTopic       =   "Form1"
   ScaleHeight     =   8640
   ScaleWidth      =   10860
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00808000&
      Caption         =   "Back"
      Height          =   735
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7680
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      Height          =   4335
      Left            =   3360
      Picture         =   "Registration.frx":0000
      ScaleHeight     =   4275
      ScaleWidth      =   5955
      TabIndex        =   5
      Top             =   3720
      Width           =   6015
   End
   Begin VB.CommandButton cmdJazz 
      BackColor       =   &H00808000&
      Caption         =   "Jazz Classes"
      BeginProperty Font 
         Name            =   "MS Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1920
      Width           =   2055
   End
   Begin VB.CommandButton cmdTap 
      BackColor       =   &H00808000&
      Caption         =   "Tap Classes"
      BeginProperty Font 
         Name            =   "MS Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   360
      Width           =   1935
   End
   Begin VB.CommandButton cmdHipHop 
      BackColor       =   &H00808000&
      Caption         =   "Hip Hop Classes"
      BeginProperty Font 
         Name            =   "MS Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3480
      Width           =   1935
   End
   Begin VB.CommandButton cmdLyric 
      BackColor       =   &H00808000&
      Caption         =   "Lyrical or Modern Dance Classes"
      BeginProperty Font 
         Name            =   "MS Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1920
      Width           =   1935
   End
   Begin VB.CommandButton cmdBallet 
      BackColor       =   &H00808000&
      Caption         =   "Ballet Classes"
      BeginProperty Font 
         Name            =   "MS Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   360
      Width           =   1935
   End
   Begin VB.Label lblName 
      BackColor       =   &H00808000&
      Caption         =   "Created by Leslie Pesarchick"
      Height          =   375
      Left            =   8160
      TabIndex        =   8
      Top             =   8160
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "Choose the type of Dance Class you would like to take."
      BeginProperty Font 
         Name            =   "MS Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5400
      TabIndex        =   7
      Top             =   480
      Width           =   4935
   End
End
Attribute VB_Name = "frmRegistration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : ProjectDanceInfo (DanceProject.prj.vbp)
'Form Name : frmRegistration (Registration.frm)
'Author: Leslie Pesarchick
'Date Written: October 27, 2003
'Purpose of Form: to have the user register for classes
                    'they can choose between ballet, tap, jazz
                    'lyrical, modern, and hip hop.

Option Explicit
'Option Explicit is a command to force the user to explicitly declare all
'variables before they can be used.
Private Sub cmdBack_Click()
    frmMain.Show
    frmRegistration.Hide
End Sub

Private Sub cmdBallet_Click()
    frmBalletClasses.Show
    frmRegistration.Hide
    frmBalletClasses.picResults.Print "Age Group"; Tab(20); "Level"; Tab(41); "Day and Time of Practice"
    frmBalletClasses.picResults.Print "*******************************************************************************************************************"

End Sub

Private Sub cmdHipHop_Click()
    frmHipHopClasses.Show
    frmRegistration.Hide
    frmHipHopClasses.picResults.Print "Age Group"; Tab(20); "Level"; Tab(41); "Day and Time of Practice"
    frmHipHopClasses.picResults.Print "*******************************************************************************************************************"

End Sub

Private Sub cmdJazz_Click()
    frmJazzClasses.Show
    frmRegistration.Hide
    frmJazzClasses.picResults.Print "Age Group"; Tab(20); "Level"; Tab(41); "Day and Time of Practice"
    frmJazzClasses.picResults.Print "*******************************************************************************************************************"

End Sub

Private Sub cmdLyric_Click()
    frmLyricClasses.Show
    frmRegistration.Hide
    frmLyricClasses.picResults.Print "Age Group"; Tab(20); "Level"; Tab(41); "Day and Time of Practice"
    frmLyricClasses.picResults.Print "*******************************************************************************************************************"

End Sub

Private Sub cmdTap_Click()
    frmTapClasses.Show
    frmRegistration.Hide
    frmTapClasses.picResults.Print "Age Group"; Tab(20); "Level"; Tab(41); "Day and Time of Practice"
    frmTapClasses.picResults.Print "*******************************************************************************************************************"

End Sub
