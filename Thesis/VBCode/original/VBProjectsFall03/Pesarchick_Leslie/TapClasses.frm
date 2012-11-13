VERSION 5.00
Begin VB.Form frmTapClasses 
   BackColor       =   &H00C0E0FF&
   Caption         =   "Tap Classes"
   ClientHeight    =   10035
   ClientLeft      =   645
   ClientTop       =   345
   ClientWidth     =   13320
   LinkTopic       =   "Form1"
   ScaleHeight     =   10035
   ScaleWidth      =   13320
   Visible         =   0   'False
   Begin VB.PictureBox picResults 
      BackColor       =   &H00C0C0FF&
      Height          =   6735
      Left            =   7440
      ScaleHeight     =   6675
      ScaleWidth      =   5355
      TabIndex        =   6
      Top             =   2400
      Width           =   5415
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Back"
      Height          =   735
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   9000
      Width           =   1095
   End
   Begin VB.CommandButton cmdAge 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Find Your Tap Class"
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
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2520
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
      Left            =   480
      Picture         =   "TapClasses.frx":0000
      ScaleHeight     =   4515
      ScaleWidth      =   6555
      TabIndex        =   0
      Top             =   3720
      Width           =   6615
   End
   Begin VB.Label lblName 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Created by Leslie Pesarchick"
      Height          =   375
      Left            =   10440
      TabIndex        =   7
      Top             =   9480
      Width           =   2415
   End
   Begin VB.Label lblTap 
      BackColor       =   &H00C0C0FF&
      Caption         =   " Tap Classes"
      BeginProperty Font 
         Name            =   "MS Gothic"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   975
      Left            =   3840
      TabIndex        =   4
      Top             =   360
      Width           =   3375
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
      ForeColor       =   &H00004080&
      Height          =   735
      Left            =   240
      TabIndex        =   2
      Top             =   1440
      Width           =   2655
   End
End
Attribute VB_Name = "frmTapClasses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : ProjectDanceInfo (DanceProject.prj.vbp)
'Form Name : frmTapClasses (TapClasses.frm)
'Author: Leslie Pesarchick
'Date Written: October 27, 2003
'Purpose of Form: to have the user register for classes
                    'the user inputs his/her age
                    'the user finds out what level he/she is
                    'the user finds out when his/her class is

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

Open Path & "Notepads\Tap, Ages.txt" For Input As #1 'opens the notepad to use as inputs

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
    frmTapClasses.Hide
End Sub

Private Sub Form_Load()
Path = "N:\CS130\handin\Pesarchick_Leslie\"
End Sub
