VERSION 5.00
Begin VB.Form frmHipHopClasses 
   BackColor       =   &H00C0C000&
   Caption         =   "Hip Hop Classes"
   ClientHeight    =   9225
   ClientLeft      =   450
   ClientTop       =   930
   ClientWidth     =   13125
   LinkTopic       =   "Form1"
   ScaleHeight     =   9225
   ScaleWidth      =   13125
   Visible         =   0   'False
   Begin VB.PictureBox picResults 
      BackColor       =   &H00FFFFC0&
      Height          =   6975
      Left            =   7440
      ScaleHeight     =   6915
      ScaleWidth      =   5355
      TabIndex        =   6
      Top             =   1680
      Width           =   5415
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Back"
      Height          =   735
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   8280
      Width           =   1215
   End
   Begin VB.CommandButton cmdAge 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Find Your Hip Hop Class"
      BeginProperty Font 
         Name            =   "MS Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2400
      Width           =   2655
   End
   Begin VB.TextBox txtAge 
      Height          =   735
      Left            =   3000
      TabIndex        =   1
      Top             =   1440
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
      Height          =   4575
      Left            =   120
      Picture         =   "HipHopClasses.frx":0000
      ScaleHeight     =   4515
      ScaleWidth      =   7035
      TabIndex        =   0
      Top             =   3360
      Width           =   7095
   End
   Begin VB.Label lblName 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Created by Leslie Pesarchick"
      Height          =   375
      Left            =   10560
      TabIndex        =   7
      Top             =   8760
      Width           =   2295
   End
   Begin VB.Label lblHipHop 
      BackColor       =   &H00FFFFC0&
      Caption         =   "  Hip Hop Classes"
      BeginProperty Font 
         Name            =   "MS Gothic"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   975
      Left            =   3120
      TabIndex        =   4
      Top             =   240
      Width           =   4455
   End
   Begin VB.Label lblAge 
      BackColor       =   &H00FFFFC0&
      Caption         =   " Enter your Age"
      BeginProperty Font 
         Name            =   "MS Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   735
      Left            =   240
      TabIndex        =   2
      Top             =   1440
      Width           =   2535
   End
End
Attribute VB_Name = "frmHipHopClasses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : ProjectDanceInfo (DanceProject.prj.vbp)
'Form Name : frmHipHopClasses (HipHopClasses.frm)
'Author: Leslie Pesarchick
'Date Written: October 27, 2003
'Purpose of Form: to have the user register for Hip Hop Classes
                    'user inputs their age
                    'user finds out what level they are at
                    'and when their class is
                    'gives the user the number to call to register for that class

Option Explicit
'Option Explicit is a command to force the user to explicitly declare all
'variables before they can be used.

Private Sub cmdAge_Click()
Dim Ages(1 To 6) As Single
Dim Agegroup(1 To 6) As String
Dim Level(1 To 6) As String
Dim Time(1 To 6) As String
Dim Age As Single
Dim NotFound As Boolean
Dim I As Integer
Dim CTR As Integer
picResults.Cls
picResults.Print "Age Group"; Tab(20); "Level"; Tab(41); "Day and Time of Practice"
picResults.Print "*******************************************************************************************************************"

Open Path & "Notepads\Hip Hop, Ages.txt" For Input As #1 'opens the notepad to use as inputs

For CTR = 1 To 6
    Input #1, Ages(CTR), Agegroup(CTR), Level(CTR), Time(CTR) 'makes CTR into an array
Next CTR
Close #1 'closes the file
Age = txtAge.Text
I = 0
NotFound = True
Do While NotFound And I < 6
    I = I + 1
    If Ages(I) <= Age Then
        NotFound = False
    End If
Loop
If NotFound Then
        MsgBox "No dance classes offered for this age", , "No Classes"
    Else
        picResults.Print Agegroup(I), Level(I), Time(I)
        picResults.Print
        picResults.Print
        picResults.Print
        picResults.Print "*To Register for Classes, please call 1-800-FUN-DANZ"
        picResults.Print
        picResults.Print
End If
End Sub

Private Sub cmdBack_Click()
    frmRegistration.Show
    frmHipHopClasses.Hide
End Sub

Private Sub Form_Load()
Path = "N:\CS130\handin\Pesarchick_Leslie\"
End Sub
