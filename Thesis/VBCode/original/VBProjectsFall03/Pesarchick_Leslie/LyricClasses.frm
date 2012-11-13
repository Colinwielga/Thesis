VERSION 5.00
Begin VB.Form frmLyricClasses 
   BackColor       =   &H00404080&
   Caption         =   "Lyrical Classes"
   ClientHeight    =   9705
   ClientLeft      =   255
   ClientTop       =   540
   ClientWidth     =   13230
   LinkTopic       =   "Form1"
   ScaleHeight     =   9705
   ScaleWidth      =   13230
   Visible         =   0   'False
   Begin VB.PictureBox picResults 
      BackColor       =   &H00C0C0FF&
      Height          =   6375
      Left            =   7080
      ScaleHeight     =   6315
      ScaleWidth      =   5595
      TabIndex        =   7
      Top             =   2400
      Width           =   5655
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Back"
      Height          =   735
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   8760
      Width           =   1215
   End
   Begin VB.CommandButton cmdLyrical 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Find Your Lyrical Dance Class"
      BeginProperty Font 
         Name            =   "MS Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2400
      Width           =   2775
   End
   Begin VB.CommandButton cmdModern 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Find Your Modern Dance Class"
      BeginProperty Font 
         Name            =   "MS Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2400
      Width           =   2775
   End
   Begin VB.TextBox txtAge 
      Height          =   735
      Left            =   3120
      TabIndex        =   1
      Top             =   1440
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
      Height          =   4935
      Left            =   120
      Picture         =   "LyricClasses.frx":0000
      ScaleHeight     =   4875
      ScaleWidth      =   6555
      TabIndex        =   0
      Top             =   3600
      Width           =   6615
   End
   Begin VB.Label lblName 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Created by Leslie Pesarchick"
      Height          =   375
      Left            =   10440
      TabIndex        =   8
      Top             =   9000
      Width           =   2295
   End
   Begin VB.Label lblLyrical 
      BackColor       =   &H00C0C0FF&
      Caption         =   " Modern Dance and   Lyrical Classes"
      BeginProperty Font 
         Name            =   "MS Gothic"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   1215
      Left            =   3600
      TabIndex        =   5
      Top             =   120
      Width           =   4455
   End
   Begin VB.Label lblAge 
      BackColor       =   &H00C0C0FF&
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
      ForeColor       =   &H00000040&
      Height          =   735
      Left            =   240
      TabIndex        =   2
      Top             =   1440
      Width           =   2775
   End
End
Attribute VB_Name = "frmLyricClasses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : ProjectDanceInfo (DanceProject.prj.vbp)
'Form Name : frmLyricClasses (LyricClasses.frm)
'Author: Leslie Pesarchick
'Date Written: October 27, 2003
'Purpose of Form: to have the user register for Lyrical or Modern Dance Classes
                    'the user inputs his/her age
                    'the user finds out what level he/she is in
                    'for Lyrical and/or Modern Dance
                    'the user finds out when his/her class is

Option Explicit
'Option Explicit is a command to force the user to explicitly declare all
'variables before they can be used.
Private Sub cmdBack_Click()
    frmRegistration.Show
    frmLyricClasses.Hide
End Sub

Private Sub cmdLyrical_Click()
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

Open Path & "Notepads\Lyrical, Ages.txt" For Input As #1 'opens the notepad to use as inputs

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

Private Sub cmdModern_Click()
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

Open Path & "Notepads\Modern, Ages.txt" For Input As #1 'opens the notepad to use as inputs

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

Private Sub Form_Load()
Path = "N:\CS130\handin\Pesarchick_Leslie\"
End Sub
