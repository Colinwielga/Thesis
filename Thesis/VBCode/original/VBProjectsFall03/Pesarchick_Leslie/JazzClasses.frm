VERSION 5.00
Begin VB.Form frmJazzClasses 
   BackColor       =   &H00FF8080&
   Caption         =   "Jazz Classes"
   ClientHeight    =   9450
   ClientLeft      =   450
   ClientTop       =   540
   ClientWidth     =   13065
   LinkTopic       =   "Form1"
   ScaleHeight     =   9450
   ScaleWidth      =   13065
   Visible         =   0   'False
   Begin VB.PictureBox picResults 
      BackColor       =   &H00FFC0C0&
      Height          =   6255
      Left            =   7440
      ScaleHeight     =   6195
      ScaleWidth      =   5235
      TabIndex        =   6
      Top             =   2160
      Width           =   5295
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Back"
      Height          =   735
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   8040
      Width           =   1215
   End
   Begin VB.CommandButton cmdAge 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Find Your Jazz Class"
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
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2640
      Width           =   2655
   End
   Begin VB.TextBox txtAge 
      Height          =   735
      Left            =   2880
      TabIndex        =   1
      Top             =   1560
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
      Height          =   3975
      Left            =   120
      Picture         =   "JazzClasses.frx":0000
      ScaleHeight     =   3915
      ScaleWidth      =   7035
      TabIndex        =   0
      Top             =   3720
      Width           =   7095
   End
   Begin VB.Label lblName 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Created by Leslie Pesarchick"
      Height          =   375
      Left            =   10440
      TabIndex        =   7
      Top             =   8880
      Width           =   2295
   End
   Begin VB.Label lblJazz 
      BackColor       =   &H00FFC0C0&
      Caption         =   " Jazz Classes"
      BeginProperty Font 
         Name            =   "MS Gothic"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3840
      TabIndex        =   4
      Top             =   360
      Width           =   3615
   End
   Begin VB.Label lblAge 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Enter Your Age"
      BeginProperty Font 
         Name            =   "MS Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   2655
   End
End
Attribute VB_Name = "frmJazzClasses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : ProjectDanceInfo (DanceProject.prj.vbp)
'Form Name : frmJazzClasses(JazzClasses.frm)
'Author: Leslie Pesarchick
'Date Written: October 27, 2003
'Purpose of Form: to have the user register for Jazz Classes
                    'the users input their age
                    'the user finds out what level they are in
                    'the user finds out when their dance classes are

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

Open Path & "Notepads\Jazz, Ages.txt" For Input As #1 'opens the notepad to use as inputs

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
    frmJazzClasses.Hide
End Sub

Private Sub Form_Load()
Path = "N:\CS130\handin\Pesarchick_Leslie\"
End Sub
