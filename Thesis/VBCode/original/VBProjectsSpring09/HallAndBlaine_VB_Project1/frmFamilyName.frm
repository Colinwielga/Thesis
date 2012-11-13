VERSION 5.00
Begin VB.Form frmFamilyName 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Family Name"
   ClientHeight    =   7485
   ClientLeft      =   5715
   ClientTop       =   4665
   ClientWidth     =   15015
   LinkTopic       =   "Form1"
   ScaleHeight     =   7485
   ScaleWidth      =   15015
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   4215
      Left            =   5040
      Picture         =   "frmFamilyName.frx":0000
      ScaleHeight     =   4215
      ScaleWidth      =   6015
      TabIndex        =   5
      Top             =   2760
      Width           =   6015
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   2775
      Left            =   480
      Picture         =   "frmFamilyName.frx":98C6
      ScaleHeight     =   2775
      ScaleWidth      =   4215
      TabIndex        =   4
      Top             =   3480
      Width           =   4215
   End
   Begin VB.CommandButton cmdFormSurvey1 
      BackColor       =   &H000080FF&
      Caption         =   "Proceed to Survey Question 1"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   14.25
         Charset         =   1
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   11280
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2760
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.CommandButton CreateFamilyName 
      BackColor       =   &H80000003&
      Caption         =   "Create Family Name"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   14.25
         Charset         =   1
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   840
      Width           =   2775
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H000000FF&
      Caption         =   "Quit "
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   18
         Charset         =   1
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   10200
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   960
      Width           =   2295
   End
   Begin VB.CommandButton cmdReturn 
      BackColor       =   &H0000FFFF&
      Caption         =   "Return To Main Page"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   14.25
         Charset         =   1
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   840
      Width           =   2295
   End
End
Attribute VB_Name = "frmFamilyName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Family Feud
'frmFamilyName
'Colin Hall and Andre Blaine
'March 15
'This form will take a name from the user and display it as the family name,
'will return to the Main Page form,
'and will quit.


Private Sub CreateFamilyName_Click()

    'This button will ask the user for their family's last name.
    FamilyName = InputBox("Enter a last name for your family.", "Last Name")
    
    'This will display the family's last name in a message box.
    MsgBox "Your family's name is The " & FamilyName & "'s.", , "Family's Name"
    
    'This will make the button to proceed to Survey Question 1 visible and able to be clicked.
    cmdFormSurvey1.Visible = True
    
End Sub

Private Sub cmdFormSurvey1_Click()

    'This will hide Family Name form and show Survey 1 form.
    frmFamilyName.Hide
    frmSurvey1.Show


End Sub
Private Sub cmdReturn_Click()

    'This button will hide the Creators form and will open the Main Page form.
    frmMainPage.Show
    frmCreators.Hide
    
End Sub

Private Sub cmdQuit_Click()

    'This button will end the Visual Basic Program.
    End
    
End Sub
