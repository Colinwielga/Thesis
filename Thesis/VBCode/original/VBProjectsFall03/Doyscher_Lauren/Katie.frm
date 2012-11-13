VERSION 5.00
Begin VB.Form KatieForm 
   BackColor       =   &H000000FF&
   Caption         =   "Form1"
   ClientHeight    =   7485
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10395
   LinkTopic       =   "Form1"
   ScaleHeight     =   7485
   ScaleWidth      =   10395
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton optBlack 
      BackColor       =   &H00000000&
      Caption         =   "BLACK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   1920
      TabIndex        =   14
      Top             =   5760
      Width           =   1215
   End
   Begin VB.CommandButton cmdCalc 
      BackColor       =   &H00E0E0E0&
      Caption         =   "See If You Are Correct"
      Enabled         =   0   'False
      Height          =   735
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   6600
      Width           =   1935
   End
   Begin VB.OptionButton optBrown 
      BackColor       =   &H00004080&
      Caption         =   "BROWN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1920
      TabIndex        =   11
      Top             =   5160
      Width           =   1215
   End
   Begin VB.OptionButton optBlue 
      BackColor       =   &H00FF0000&
      Caption         =   "BLUE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1920
      TabIndex        =   10
      Top             =   3960
      Width           =   1215
   End
   Begin VB.OptionButton optPurple 
      BackColor       =   &H00C000C0&
      Caption         =   "PURPLE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1920
      TabIndex        =   9
      Top             =   4560
      Width           =   1215
   End
   Begin VB.OptionButton optGreen 
      BackColor       =   &H0000C000&
      Caption         =   "GREEN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   720
      TabIndex        =   8
      Top             =   5760
      Width           =   1215
   End
   Begin VB.OptionButton optYellow 
      BackColor       =   &H0000FFFF&
      Caption         =   "YELLOW"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   720
      TabIndex        =   7
      Top             =   5160
      Width           =   1215
   End
   Begin VB.OptionButton optOrange 
      BackColor       =   &H000080FF&
      Caption         =   "ORANGE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   720
      TabIndex        =   6
      Top             =   4560
      Width           =   1215
   End
   Begin VB.OptionButton optRed 
      BackColor       =   &H000000FF&
      Caption         =   "RED"
      Height          =   375
      Left            =   720
      TabIndex        =   5
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton cmdMainForm 
      BackColor       =   &H000000FF&
      Caption         =   "Main Page"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   6360
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6360
      UseMaskColor    =   -1  'True
      Width           =   1815
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H000000FF&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6360
      UseMaskColor    =   -1  'True
      Width           =   1815
   End
   Begin VB.PictureBox Picture1 
      Height          =   2655
      Left            =   360
      Picture         =   "Katie.frx":0000
      ScaleHeight     =   2595
      ScaleWidth      =   2115
      TabIndex        =   0
      Top             =   480
      Width           =   2175
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Lauren Doyscher"
      Height          =   255
      Left            =   8880
      TabIndex        =   15
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Guess Katie's Favorite Color"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   12
      Top             =   3480
      Width           =   3135
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Hometown: Duluth, MN   Team: National                Age: 19"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   2760
      TabIndex        =   4
      Top             =   1320
      Width           =   3135
   End
   Begin VB.Line Line1 
      BorderWidth     =   5
      X1              =   2760
      X2              =   5400
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Katie Ness"
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   2880
      TabIndex        =   2
      Top             =   480
      Width           =   3735
   End
End
Attribute VB_Name = "KatieForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: SophmoreDancers (VBProject.vbp)
'Form Name: KatieForm (Katie.frm)
'Author: Lauren Doyscher
'Date Written: 10/27/03
'This form shows Katie's information
Option Explicit
'According to color picked by user, a message box appears telling user if correct or incorrect
Private Sub cmdCalc_Click()
    If optBlue = True Then
        MsgBox "     You Are Correct!     ", , "Good Job!"
    Else: MsgBox "Incorrect!  Katie's Favorite Color is Blue", , "Sorry!"
    End If
End Sub
Private Sub cmdMainForm_Click()
'Brings you to Main Page
KatieForm.Hide
MainForm.Show
End Sub
Private Sub cmdQuit_Click()
End
End Sub
'Enables "see if you are correct" calculation button
Private Sub optBlack_Click()
cmdCalc.Enabled = True
End Sub
'Enables "see if you are correct" calculation button
Private Sub optBlue_Click()
cmdCalc.Enabled = True
End Sub
'Enables "see if you are correct" calculation button
Private Sub optBrown_Click()
cmdCalc.Enabled = True
End Sub
'Enables "see if you are correct" calculation button
Private Sub optGreen_Click()
cmdCalc.Enabled = True
End Sub
'Enables "see if you are correct" calculation button
Private Sub optOrange_Click()
cmdCalc.Enabled = True
End Sub
'Enables "see if you are correct" calculation button
Private Sub optPurple_Click()
cmdCalc.Enabled = True
End Sub
'Enables "see if you are correct" calculation button
Private Sub optRed_Click()
cmdCalc.Enabled = True
End Sub
'Enables "see if you are correct" calculation button
Private Sub optYellow_Click()
cmdCalc.Enabled = True
End Sub