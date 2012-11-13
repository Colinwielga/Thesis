VERSION 5.00
Begin VB.Form frmRecord 
   Caption         =   "Wolves Season So Far:"
   ClientHeight    =   8745
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10515
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8745
   ScaleWidth      =   10515
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdRecord 
      BackColor       =   &H00C00000&
      Caption         =   "Show Wolves Games and Record"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   7680
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5520
      Width           =   2655
   End
   Begin VB.PictureBox picOutput 
      BackColor       =   &H00FF8080&
      Height          =   11055
      Left            =   0
      ScaleHeight     =   10995
      ScaleWidth      =   10275
      TabIndex        =   0
      Top             =   0
      Width           =   10335
      Begin VB.CommandButton cmdBack 
         BackColor       =   &H00C00000&
         Caption         =   "Back"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   7680
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   7200
         Width           =   2655
      End
      Begin VB.Label lblName 
         Caption         =   "By: Chad Henfling"
         Height          =   255
         Left            =   6360
         TabIndex        =   3
         Top             =   8280
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmRecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Minnesota Timberwolves Center (MinnesotaTimberwovlesbyChadHenfling.vbp)
'Main Form (frmRecord.frm)
'Chad Henfling
'Created March 23, 2006
'This form displays all the games the Wolves have played this year and their record througout the year along with some other statistics.
Option Explicit
Dim Schedule As String
Dim counter As Integer

Private Sub cmdBack_Click()
    'go back to main form
    frmRecord.Visible = False
    frm1.Visible = True
End Sub

Private Sub cmdRecord_Click()
    'Opening file and reading information
    Open App.Path & "\Schedule.txt" For Input As #2
    counter = 0
    'reading the whole file and printing it
    Do Until EOF(2)
        counter = counter + 1
        Input #2, Schedule
        picOutput.Print Schedule
    Loop
    Close #2
End Sub
