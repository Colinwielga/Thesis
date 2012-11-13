VERSION 5.00
Begin VB.Form frmCasefiles 
   BackColor       =   &H00000000&
   Caption         =   "Case files"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdLearnpedo 
      BackColor       =   &H000080FF&
      Caption         =   "Click to read a brief description of all the  pedophile typologies. "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   11880
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2640
      Width           =   1935
   End
   Begin VB.CommandButton cmdLearnRape 
      BackColor       =   &H000080FF&
      Caption         =   "Click to read a brief description of the four rapist typologies. "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   840
      MaskColor       =   &H000080FF&
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2640
      Width           =   2055
   End
   Begin VB.PictureBox picHeader 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   2175
      Left            =   5520
      Picture         =   "frmCasefiles.frx":0000
      ScaleHeight     =   2175
      ScaleWidth      =   5175
      TabIndex        =   10
      Top             =   5640
      Width           =   5175
   End
   Begin VB.PictureBox picInstruction 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1935
      Left            =   4080
      ScaleHeight     =   1935
      ScaleWidth      =   6495
      TabIndex        =   9
      Top             =   2880
      Width           =   6495
   End
   Begin VB.CommandButton cmdHARD 
      BackColor       =   &H00FF00FF&
      Caption         =   $"frmCasefiles.frx":2BFE
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   7800
      Width           =   3255
   End
   Begin VB.CommandButton cmdpedo2 
      BackColor       =   &H0000FF00&
      Caption         =   "Read Case File"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   11520
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   360
      Width           =   2655
   End
   Begin VB.CommandButton cmdRape2 
      BackColor       =   &H0000FF00&
      Caption         =   "Read Case File"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   360
      Width           =   2775
   End
   Begin VB.CommandButton cmdPedo 
      BackColor       =   &H0000FF00&
      Caption         =   "Read Case File"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   360
      Width           =   2655
   End
   Begin VB.CommandButton cmdRape1 
      BackColor       =   &H0000FF00&
      Caption         =   "Read Case File"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   360
      Width           =   2415
   End
   Begin VB.Label lblpedo2 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Pedophile Case File #2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   615
      Left            =   11520
      TabIndex        =   7
      Top             =   2160
      Width           =   2895
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Rapist Case File #2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   615
      Left            =   3960
      TabIndex        =   5
      Top             =   2160
      Width           =   2895
   End
   Begin VB.Label lblpedo1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Pedophile Case File #1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   615
      Left            =   7920
      TabIndex        =   4
      Top             =   2160
      Width           =   2775
   End
   Begin VB.Label lblrapeone 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Rapist Case File #1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   615
      Left            =   720
      TabIndex        =   1
      Top             =   2160
      Width           =   2295
   End
End
Attribute VB_Name = "frmCasefiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdHARD_Click()
'hides the case files form and shows the hard case form
    frmCasefiles.Hide
    frmHardCase.Show
    
End Sub

Private Sub cmdLearnpedo_Click()
'Hides the case files form and shows the rape description form
    frmCasefiles.Hide
    frmDescription2.Show
    
End Sub

Private Sub cmdLearnRape_Click()
'Hides the case files form and shows the pedophile description form
    frmCasefiles.Hide
    frmDescription1.Show
    
End Sub

Private Sub cmdPedo_Click()
'Hides the case files form and shows the first pedophile case form
    frmCasefiles.Hide
    frmCasePedo1.Show
End Sub

Private Sub cmdpedo2_Click()
'Hides the case files form and shows the second pedophile case form
    frmCasefiles.Hide
    frmCasePedo2.Show
End Sub

Private Sub cmdRape1_Click()
'Hides the case files form and shows the first rape case form
    frmCasefiles.Hide
    frmCaseRape1Read.Show
    
End Sub

Private Sub cmdRape2_Click()
'Hides the case files form and shows the second rape case form
    frmCasefiles.Hide
    frmCaseRape2Read.Show
    
End Sub

Private Sub Form_Activate()
'declare my pos variable for my file list i am going to open
Dim pos As Integer
picInstruction.Cls 'Make sure my picture box is clear of old data
    picInstruction.Print "Please click on a case file that you would" 'Prints a line of standard text
    picInstruction.Print "like to read "; nam 'uses the name the user inputed in the second form and displays is.
                                                'making it more personal.
    
    Open App.Path & "\CaseFile.txt" For Input As #1 'Open file i made with a listing of all case files
        
    Do Until EOF(1) 'I read the list until every line has been read
        pos = pos + 1 'this keeps track of how many times this is done
        Input #1, CaseFile(pos) 'i tell the program i am using casefiles array for the data
    Loop 'I loop it until it has reached the entire file
    
    Close #1 'Prevents me having a clash with another file. Since this file has already
            'been read i don't need to keep it open.
    
    
End Sub
