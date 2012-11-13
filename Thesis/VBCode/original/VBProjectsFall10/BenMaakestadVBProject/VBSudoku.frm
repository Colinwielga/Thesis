VERSION 5.00
Begin VB.Form frmOption 
   BackColor       =   &H00404000&
   Caption         =   "Form1"
   ClientHeight    =   4365
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10770
   LinkTopic       =   "Form1"
   ScaleHeight     =   4365
   ScaleWidth      =   10770
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox ckAns 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Check Answer"
      Height          =   615
      Left            =   4080
      MaskColor       =   &H00C0E0FF&
      TabIndex        =   12
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton cmdInstruction 
      BackColor       =   &H00000080&
      Caption         =   "Instructions"
      Height          =   495
      Left            =   480
      TabIndex        =   11
      Top             =   2520
      Width           =   1335
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00FFFF00&
      Caption         =   "Back"
      Height          =   495
      Left            =   480
      MaskColor       =   &H00FFFF00&
      TabIndex        =   8
      Top             =   3360
      Width           =   1335
   End
   Begin VB.PictureBox picShow 
      Height          =   1815
      Left            =   5640
      ScaleHeight     =   1755
      ScaleWidth      =   4995
      TabIndex        =   7
      Top             =   960
      Width           =   5055
   End
   Begin VB.CommandButton cmdProcess 
      Caption         =   "Submit!"
      Height          =   615
      Left            =   9120
      TabIndex        =   6
      Top             =   3240
      Width           =   1335
   End
   Begin VB.CheckBox ckShow 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Answer Sheet"
      Height          =   615
      Left            =   4080
      TabIndex        =   5
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CheckBox ckMultiple 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Multiple Frebbie"
      Height          =   615
      Left            =   4080
      TabIndex        =   4
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CheckBox ckOne 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Freebie"
      Height          =   495
      Left            =   4080
      MaskColor       =   &H00FF0000&
      TabIndex        =   3
      Top             =   1200
      Width           =   1215
   End
   Begin VB.OptionButton optdifficult 
      BackColor       =   &H000000FF&
      Caption         =   "Difficult"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2160
      TabIndex        =   2
      Top             =   2640
      Width           =   1455
   End
   Begin VB.OptionButton optIntermediate 
      BackColor       =   &H0000FFFF&
      Caption         =   "Intermediate"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2160
      TabIndex        =   1
      Top             =   1920
      Width           =   1455
   End
   Begin VB.OptionButton optEasy 
      BackColor       =   &H0000FF00&
      Caption         =   "Easy"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2160
      TabIndex        =   0
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "Handicaps"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   4200
      TabIndex        =   10
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C000&
      Caption         =   "Difficulty"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   255
      Left            =   2400
      TabIndex        =   9
      Top             =   840
      Width           =   855
   End
End
Attribute VB_Name = "frmOption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'This form gives the user instructions on how to select their options and
'get to the puzzle... there is a back button if the user wishes to change their name






Private Sub ckAns_Click()
    If ckAns.Value = 1 Then     'check box function instructions
        picShow.Cls
        picShow.Print "This feature will go over every answer the user has input"
        picShow.Print "and will let the you know one by one whether each input"
        picShow.Print "was right or wrong. Be warned this may only be used once!"
    End If
End Sub

Private Sub ckMultiple_Click()
    If ckMultiple.Value = 1 Then        'check box function instructions
        picShow.Cls
        picShow.Print "This feature will randomly fill in one number you have not "
        picShow.Print "yet solved (regardless of correct/incorrect answers). "
        picShow.Print "This feature is available without limit!"
    End If
End Sub

Private Sub ckOne_Click()
    If ckOne.Value = 1 Then     'check box function instructions
        picShow.Cls
        picShow.Print "This feature will randomly fill in one number you have not "
        picShow.Print "yet solved (regardless of correct/incorrect answers). "
        picShow.Print "This feature is useable only once!"
    End If
    
    
End Sub

Private Sub ckShow_Click()
    If ckShow.Value = 1 Then        'check box function instructions
        picShow.Cls
        picShow.Print "This feature will reveal a copy of the solved the puzzle "
        picShow.Print "for a few seconds! This feature is available only once!"
    End If
End Sub

Private Sub cmdBack_Click() 'brings the user back to the name form
    frmName.Show
    frmOption.Hide
End Sub

Private Sub cmdInstruction_Click()  'Gives a full set of instructions to the user
    picShow.Cls
    picShow.Print User; ", you need to select the puzzle difficulty level in the left column."
    picShow.Print "If you would like some assistance, check some of the options "
    picShow.Print "in the right Column. By clicking on some of the Options you can see"
    picShow.Print "what each one does. When you are finished Click submit! "
    
End Sub

Private Sub cmdProcess_Click()      'creates the look of the puzzle form through check boxes and option buttons
    If optEasy.Value = False _
        And optIntermediate.Value = False _
        And optdifficult.Value = False Then
        MsgBox "You need to select a difficulty level!"
    ElseIf optEasy.Value = True Then
        Level = "Easy"
    ElseIf optIntermediate.Value = True Then
        Level = "Intermediate"
    Else
        Level = "Difficult"
    End If
    picShow.Print Level
    
    If ckOne.Value = 1 Then
        frmPuzzle.cmdFreebie.Visible = True
    End If
    If ckMultiple.Value = 1 Then
        frmPuzzle.cmdMultiple.Visible = True
    End If
    If ckShow.Value = 1 Then
        frmPuzzle.cmdCheat.Visible = True
    End If
    If ckAns.Value = 1 Then
        frmPuzzle.cmdCheck.Visible = True
    End If
    
    
    
    
    frmPuzzle.Show
    frmOption.Hide
    
    
        
    
End Sub




Private Sub Form_Load()     'declared check/options
    optEasy.Value = False
    optIntermediate.Value = False
    optdifficult.Value = False
    ckOne.Value = 0
    ckMultiple.Value = 0
    ckShow.Value = 0
    ckAns.Value = 0
    
End Sub

