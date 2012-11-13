VERSION 5.00
Begin VB.Form frmCourseInfo 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Course Information"
   ClientHeight    =   6450
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9720
   LinkTopic       =   "Form1"
   ScaleHeight     =   6450
   ScaleWidth      =   9720
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Main Menu"
      Height          =   495
      Left            =   7200
      TabIndex        =   7
      Top             =   5880
      Width           =   2175
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Select Another Course"
      Height          =   855
      Left            =   0
      TabIndex        =   4
      Top             =   4920
      Width           =   2175
   End
   Begin VB.CommandButton cmdScores 
      Caption         =   "Enter Scores for a Round of Golf"
      Height          =   1095
      Left            =   0
      TabIndex        =   3
      Top             =   3480
      Width           =   2175
   End
   Begin VB.CommandButton cmdContact 
      Caption         =   "Contact Information"
      Height          =   975
      Left            =   0
      TabIndex        =   2
      Top             =   2160
      Width           =   2175
   End
   Begin VB.CommandButton cmdRates 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Rates"
      Height          =   855
      Left            =   0
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   1
      Top             =   840
      Width           =   2175
   End
   Begin VB.Label lblTitle3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Blackberry Ridge Golf Course"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   360
      TabIndex        =   6
      Top             =   120
      Visible         =   0   'False
      Width           =   9255
   End
   Begin VB.Label lblTitle2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Albany Golf Club"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   5
      Top             =   120
      Visible         =   0   'False
      Width           =   9135
   End
   Begin VB.Image imgRates2 
      Height          =   4920
      Left            =   2520
      Picture         =   "frmRichSpring.frx":0000
      Stretch         =   -1  'True
      Top             =   720
      Visible         =   0   'False
      Width           =   6840
   End
   Begin VB.Image imgContact3 
      Height          =   4935
      Left            =   2520
      Picture         =   "frmRichSpring.frx":3B45E
      Stretch         =   -1  'True
      Top             =   720
      Visible         =   0   'False
      Width           =   6855
   End
   Begin VB.Image imgRates3 
      Height          =   4935
      Left            =   2520
      Picture         =   "frmRichSpring.frx":6A9A2
      Stretch         =   -1  'True
      Top             =   720
      Visible         =   0   'False
      Width           =   6855
   End
   Begin VB.Image imgContact2 
      Height          =   4920
      Left            =   2520
      Picture         =   "frmRichSpring.frx":A95F8
      Stretch         =   -1  'True
      Top             =   720
      Visible         =   0   'False
      Width           =   6840
   End
   Begin VB.Image imgRates1 
      Height          =   4800
      Left            =   2520
      Picture         =   "frmRichSpring.frx":D29D4
      Stretch         =   -1  'True
      Top             =   840
      Visible         =   0   'False
      Width           =   6840
   End
   Begin VB.Image imgContact1 
      Height          =   4785
      Left            =   2520
      Picture         =   "frmRichSpring.frx":F4BCA
      Stretch         =   -1  'True
      Top             =   840
      Visible         =   0   'False
      Width           =   6855
   End
   Begin VB.Label lblTitle1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Rich-Spring Golf Club"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   8895
   End
End
Attribute VB_Name = "frmCourseInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: GolfGuide
':Form Name:  frmCourseInfo
':Author:   Tyler Cash
':Date written:  March 19, 2009


'This form displays specific course information about the course the user chose.
'The form is designed to load images and labels specific to the course chosen.
'The form also allows the user to display course rates and contact information.

Option Explicit
Private Sub cmdContact_Click()
'When the user clicks the contact information button, the program displays an image
'with the contact information of the course that was selected by the user.

'Checking which course was chosen.
'Displays contact info and hides rates information
    If Course = 1 Then
        imgContact1.Visible = True
        imgRates1.Visible = False
    ElseIf Course = 2 Then
        imgContact2.Visible = True
        imgRates2.Visible = False
    ElseIf Course = 3 Then
        imgContact3.Visible = True
        imgRates3.Visible = False
    End If
        
        
End Sub

Private Sub cmdExit_Click()
'This button returns the user to the form where they can choose a new course.

'Changing forms
    frmCourses.Show

'The form needs to be unloaded because it performs tasks when it loads.
'If the user selects a new course, this form needs to do those tasks again.
    Unload Me
            
End Sub

Private Sub cmdQuit_Click()
'This button returns the user to the title menu.

'Changing forms
    frmTitle.Show
    
'The form needs to be unloaded because it performs tasks when it loads.
'If the user selects a new course, this form needs to do those tasks again.
    Unload Me
    
End Sub

Private Sub cmdRates_Click()
'When the user clicks the Rates button, the program displays an image
'with the rates of the course that was selected by the user.

'Checking which course was chosen.
'Displays rates and hides contact information
    If Course = 1 Then
        imgContact1.Visible = False
        imgRates1.Visible = True
    ElseIf Course = 2 Then
        imgContact2.Visible = False
        imgRates2.Visible = True
    ElseIf Course = 3 Then
        imgContact3.Visible = False
        imgRates3.Visible = True
    End If
End Sub

Private Sub cmdScores_Click()
'This button changes forms to Stat Tracking portion of the program
    
'Changing forms
    frmStats.Show
    
'The form needs to be unloaded because it performs tasks when it loads.
'If the user selects a new course, this form needs to do those tasks again.
    Unload Me
        
End Sub

Private Sub Form_Load()
'When the form loads it checks which course has been selected.
'The background picture for the form is set depending on the course selected.

'Checking which course was chosen and setting background image
    If Course = 1 Then
        frmCourseInfo.Picture = LoadPicture(App.Path & "\Images" & "\Rich-Spring Golf Course Big.JPG")
        lblTitle1.Visible = True
    ElseIf Course = 2 Then
        frmCourseInfo.Picture = LoadPicture(App.Path & "\Images" & "\Albany Big.bmp")
        lblTitle2.Visible = True
    ElseIf Course = 3 Then
        frmCourseInfo.Picture = LoadPicture(App.Path & "\Images" & "\Blackberry Ridge Big.bmp")
        lblTitle3.Visible = True
    End If
    
End Sub
