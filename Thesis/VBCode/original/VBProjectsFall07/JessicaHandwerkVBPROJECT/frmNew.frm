VERSION 5.00
Begin VB.Form frmStudents 
   BackColor       =   &H80000007&
   Caption         =   "Form2"
   ClientHeight    =   5910
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8595
   LinkTopic       =   "Form2"
   ScaleHeight     =   5910
   ScaleWidth      =   8595
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picGroups 
      Height          =   3975
      Left            =   5640
      ScaleHeight     =   3915
      ScaleWidth      =   2235
      TabIndex        =   6
      Top             =   960
      Width           =   2295
   End
   Begin VB.CommandButton cmdGroup 
      Caption         =   "Create Volunteer Groups"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2400
      TabIndex        =   5
      Top             =   5040
      Width           =   3135
   End
   Begin VB.CommandButton cmdAlpha 
      Caption         =   "Alphabetize Names"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1080
      TabIndex        =   4
      Top             =   5040
      Width           =   1335
   End
   Begin VB.CommandButton cmdGetNames 
      BackColor       =   &H80000007&
      Caption         =   "Get Names"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   3
      Top             =   5040
      Width           =   975
   End
   Begin VB.PictureBox picNames 
      Height          =   3975
      Left            =   120
      ScaleHeight     =   3915
      ScaleWidth      =   2115
      TabIndex        =   2
      Top             =   960
      Width           =   2175
   End
   Begin VB.CommandButton cmdMenu 
      Caption         =   "Main Menu"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6000
      TabIndex        =   1
      Top             =   5040
      Width           =   855
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6840
      TabIndex        =   0
      Top             =   5040
      Width           =   855
   End
   Begin VB.Label lbltitle 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      Caption         =   "       Karyl Daughters'       Communication 387 Class"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   855
      Left            =   1560
      TabIndex        =   7
      Top             =   0
      Width           =   4935
   End
   Begin VB.Image Image1 
      Height          =   4035
      Left            =   2400
      Picture         =   "frmNew.frx":0000
      Stretch         =   -1  'True
      Top             =   960
      Width           =   3105
   End
End
Attribute VB_Name = "frmStudents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim J(1 To 12) As String
Dim ctr As Integer

Private Sub cmdAlpha_Click()
picNames.Cls
Dim Pass As Integer
Dim Pos As Integer
Dim Temp As String
    For Pass = 1 To ctr - 1
        For Pos = 1 To ctr - Pass
            If J(Pos) > J(Pos + 1) Then
                Temp = J(Pos)
                J(Pos) = J(Pos + 1)
                J(Pos + 1) = Temp
            End If
        Next Pos
    Next Pass
cmdGetNames.Enabled = False
cmdAlpha.Enabled = False
cmdGroup.Enabled = True
picNames.Print "Comm. 387 Students"
For Pos = 1 To ctr
    picNames.Print J(Pos)
Next Pos
    

End Sub

Private Sub cmdExit_Click()
End
End Sub

Private Sub cmdGetNames_Click()

Dim Pos As Integer


ctr = 0

For Pos = 1 To 12
ctr = ctr + 1
J(Pos) = InputBox("Please Enter A Name (You will need to enter 12 different names total)", "Enter Names")
Next Pos

picNames.Print "Comm. 387 Students"
'Prints all names entered
For Pos = 1 To ctr
    picNames.Print J(Pos)
Next Pos
cmdAlpha.Enabled = True
cmdGetNames.Enabled = False

End Sub


Private Sub cmdGroup_Click()
Dim Pos As Integer
'puts names into four groups
For Pos = 1 To ctr
    Select Case Pos
        Case Is = 1
            picGroups.Print "Group One"
        Case Is = 5
            picGroups.Print "Group Two"
        Case Is = 9
            picGroups.Print "Group Three"
    End Select

    picGroups.Print J(Pos)
Next Pos
cmdGroup.Enabled = False
End Sub

Private Sub cmdMenu_Click()
frmMenu.Show
frmStudents.Hide
End Sub
