VERSION 5.00
Begin VB.Form frmavoid 
   BackColor       =   &H00000000&
   Caption         =   "Methods of Avoidance"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picFight 
      Height          =   615
      Left            =   5400
      Picture         =   "frmavoid.frx":0000
      ScaleHeight     =   555
      ScaleWidth      =   3075
      TabIndex        =   4
      Top             =   9360
      Width           =   3135
   End
   Begin VB.CommandButton cmdCasefiles 
      BackColor       =   &H0000FF00&
      Caption         =   "Return to Case Files"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   9720
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   9240
      Width           =   1935
   End
   Begin VB.CommandButton cmdtitle 
      BackColor       =   &H0000FF00&
      Caption         =   "Go back to Title Screen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   9240
      Width           =   1815
   End
   Begin VB.PictureBox picavoid 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   8295
      Left            =   120
      ScaleHeight     =   8295
      ScaleWidth      =   14775
      TabIndex        =   1
      Top             =   840
      Width           =   14775
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Methods of Avoidance"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   5040
      TabIndex        =   0
      Top             =   120
      Width           =   5775
   End
End
Attribute VB_Name = "frmavoid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Allows the user to go back to select another case file or attempt the hard case
Private Sub cmdCasefiles_Click()
    frmavoid.Hide
    frmCasefiles.Show
    
End Sub

'allows the user to go the title screen which is also the only quit button
'to the entire program.
Private Sub cmdtitle_Click()
    frmavoid.Hide
    frmTitleScreen.Show

End Sub

Private Sub Form_Activate()
'I declared pos and ctr throughout only under certain buttons so i can load
'more than one file during my projet.
Dim pos As Integer, ctr As Integer
'This will clear the picture box should someone go back to this form
'Otherwise it will print itself twice and that looks back.
picavoid.Cls
    'I made a file in notepad and am now opening it up for input as number 6
    'i did this just to keep track of how many files i am using total.
    Open App.Path & "\avoid.txt" For Input As #6
    Do Until EOF(6) 'reads the entire file until the end
        pos = pos + 1 'keeps track of each line so it gives it a certain value.
        Input #6, avoid(pos) 'I say what i am going to use this for.
    Loop 'Do this process until the entire file ahs been read.
    Close #6 'closes the file because now i already have it all read under a differnt
            'varible name. This will make sure i don't cross files or something.
    
For ctr = 1 To 26 'i a notepad file and i want it to dispaly all the valuse for the
                    'file i made. Since my file takes up only 26 lines i just let it
                    'Run until then.
    picavoid.Print avoid(ctr) 'This prints out every line of my file one by one
Next ctr 'This repeats until the end of my file is reached.

End Sub

