VERSION 5.00
Begin VB.Form frmCreate 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Create Data File"
   ClientHeight    =   7515
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10470
   LinkTopic       =   "Form1"
   ScaleHeight     =   7515
   ScaleWidth      =   10470
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   1095
      Left            =   7200
      TabIndex        =   4
      Top             =   2040
      Width           =   2655
   End
   Begin VB.DirListBox Dir1 
      Height          =   2340
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   3855
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   3855
   End
   Begin VB.CommandButton cmdCreate 
      Caption         =   "Create New Text File"
      Height          =   1095
      Left            =   7200
      TabIndex        =   0
      Top             =   360
      Width           =   2655
   End
   Begin VB.Label lblCreate 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Choose a location to save your data:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Width           =   5415
   End
End
Attribute VB_Name = "frmCreate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: GolfGuide
':Form Name:  frmCreate
':Author:   Tyler Cash
':Date written:  March 21, 2009

'This form allows the user to create a new text file to save their scoring data in.
'The user locates the location that they would like to save their text file.
'Then the user gives the text file a name.
'The program then creates the text file.

Option Explicit


Private Sub cmdCancel_Click()
'This button changes forms back the the stat tracking main menu.

'Changing forms
    frmStats.Show
    frmCreate.Hide
End Sub

Private Sub cmdCreate_Click()
'This button creates the text file in the location specified by the user.

 Dim fso As FileSystemObject
 
 'Getting the location where the text file will be saved from a Directory box
    FolderSelected = Dir1.Path
    
'Making sure the location name ends in "\"
    If Right(FolderSelected, 1) <> "\" Then
        FolderSelected = FolderSelected & "\"
    End If
    
'Inputbox asks user to give their file a name.  Extra spaces are removed from the name.
    FileName = Trim(InputBox("Give your data file a name." & vbNewLine & vbNewLine & vbNewLine & vbNewLine & vbNewLine & "*Please exclude file extension type*", "Save"))
    
'Combines the location to be saved and the file name
    FileName = FolderSelected & FileName & ".txt"
        
 
 'Displays save location in msgbox (Comment line for troubleshooting.)
    'MsgBox FileName

    Set fso = New FileSystemObject
   
'Creates a text file named whatever the user chose in the location the user chose
On Error GoTo Error
    fso.CreateTextFile FileName, True
   
'Changing forms to the form allowing the user to enter new scoring data.
    frmCreate.Hide
    frmScoring.Show
    
'Exit sub so we don't go through the error loop.
    Exit Sub

'In case the selected folder is read only
Error:  MsgBox "You can't save a file in that location.  Choose a new location.", , "Error"
        Exit Sub
        
End Sub

Private Sub Drive1_Change()
'This sub is for a Drive listbox that allows the user to select which drive they would
'like to save their text file onto.

'Displays the directories on the drive in the directory listbox
On Error GoTo Error
    Dir1.Path = Drive1.Drive
    Exit Sub
    
Error:  MsgBox "That Drive is unavailable.  Please select another drive", , "Error"
        Exit Sub
End Sub

