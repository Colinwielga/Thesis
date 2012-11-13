VERSION 5.00
Begin VB.Form frmLoad 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Load Data"
   ClientHeight    =   7215
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9615
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   13.5
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   7215
   ScaleWidth      =   9615
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   6600
      TabIndex        =   4
      Top             =   2640
      Width           =   2535
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   6600
      TabIndex        =   3
      Top             =   480
      Width           =   2535
   End
   Begin VB.DirListBox Dir1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   360
      TabIndex        =   2
      Top             =   1920
      Width           =   5175
   End
   Begin VB.DriveListBox Drive1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   360
      TabIndex        =   1
      Top             =   960
      Width           =   5055
   End
   Begin VB.FileListBox filLoad 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1650
      Left            =   360
      Pattern         =   "*.txt"
      TabIndex        =   0
      Top             =   5280
      Width           =   5175
   End
   Begin VB.Label lblLoad 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Locate your previous data file:"
      Height          =   615
      Left            =   360
      TabIndex        =   5
      Top             =   120
      Width           =   5535
   End
End
Attribute VB_Name = "frmLoad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: GolfGuide
':Form Name:  frmLoad
':Author:   Tyler Cash
':Date written:  March 21, 2009


'This form allows the user to select a text file that contains their scoring data.
'The user indicates the drive and directory that their file is located in.
'The file listbox displays the text files located in the specified directory.
'The user selects a file and then clicks the load button.

Option Explicit

Private Sub cmdCancel_Click()
'This button returns the user to the Stat Tracking title form

'Changing forms
    frmLoad.Hide
    frmStats.Show
End Sub

Private Sub cmdLoad_Click()
'This button loads the user indicated text file when clicked.

'Checking to see if the user picked a file
'If not, display error and make them start over
    If filLoad.FileName = "" Then
        MsgBox "Please select a file", , "Error"
        Exit Sub
    End If
    
'Saving the directory the user selected
    FileSelected = filLoad.Path
    
'Checking to make sure the directory containing the file ends in "\"
'Combines the directory and the filename
    If Right(FileSelected, 1) = "\" Then
        FileName = FileSelected & filLoad.FileName
    Else
        FileName = FileSelected & "\" & filLoad.FileName
    End If
    
    
'Msgbox displaying the entire file location (for troubleschooting)
    'MsgBox FileName
            
'Changing forms to the form allowing the user to enter new scoring data
    frmLoad.Hide
    frmScoring.Show
End Sub

Private Sub Dir1_Change()
'This sub tells the file listbox which directory to get its files from.
    
    On Error GoTo Error
    filLoad.FileName = Dir1.Path
    Exit Sub
    
Error: MsgBox "That directory is unavailable.  Select a different directory.", , "Error"
    Exit Sub
End Sub

Private Sub Drive1_Change()
'This sub tells the drive listbox which drive to get its directories from
    On Error GoTo Error
    Dir1.Path = Drive1.Drive
    Exit Sub
    
Error: MsgBox "That drive is unavailable.  Select a different drive.", , "Error"
    Exit Sub
End Sub

