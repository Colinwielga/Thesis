VERSION 5.00
Begin VB.Form frmBalletClasses 
   BackColor       =   &H0080C0FF&
   Caption         =   "Ballet Classes"
   ClientHeight    =   9240
   ClientLeft      =   255
   ClientTop       =   930
   ClientWidth     =   13260
   LinkTopic       =   "Form1"
   ScaleHeight     =   9240
   ScaleWidth      =   13260
   Visible         =   0   'False
   Begin VB.PictureBox picResults 
      BackColor       =   &H00C0E0FF&
      Height          =   5535
      Left            =   7920
      ScaleHeight     =   5475
      ScaleWidth      =   5115
      TabIndex        =   6
      Top             =   2160
      Width           =   5175
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Back"
      Height          =   735
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   8160
      Width           =   1215
   End
   Begin VB.CommandButton cmdFindClass 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Find Your Ballet Class"
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
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2160
      Width           =   2655
   End
   Begin VB.PictureBox Picture1 
      Height          =   4815
      Left            =   120
      Picture         =   "BalletClasses.frx":0000
      ScaleHeight     =   4755
      ScaleWidth      =   7515
      TabIndex        =   2
      Top             =   3120
      Width           =   7575
   End
   Begin VB.TextBox txtAge 
      Height          =   615
      Left            =   2760
      TabIndex        =   0
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label lblName 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Created by Leslie Pesarchick"
      Height          =   375
      Left            =   10920
      TabIndex        =   7
      Top             =   8640
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0E0FF&
      Caption         =   " Ballet Classes"
      BeginProperty Font 
         Name            =   "MS Gothic"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   855
      Left            =   4320
      TabIndex        =   4
      Top             =   360
      Width           =   3735
   End
   Begin VB.Label lblage 
      BackColor       =   &H00C0E0FF&
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
      ForeColor       =   &H00004080&
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   2535
   End
End
Attribute VB_Name = "frmBalletClasses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : ProjectDanceInfo (DanceProject.prj.vbp)
'Form Name : frmBalletClasses (BalletClasses.frm)
'Author: Leslie Pesarchick
'Date Written: October 27, 2003
'Purpose of Form: to have the user register for a Ballet Class
                    'have the user input their age
                    'tells the user which level their class is
                    'tells the user when their class is, and what time
                    'tells the user to call the number given to register
                    'for that class
Option Explicit
'Option Explicit is a command to force the user to explicitly declare all
'variables before they can be used.

Private Sub cmdBack_Click()
    frmRegistration.Show
    frmBalletClasses.Hide
End Sub

Private Sub cmdFindClass_Click()
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

Open Path & "Notepads\Ballet, Ages.txt" For Input As #1 'opens the notepad to use as inputs

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
