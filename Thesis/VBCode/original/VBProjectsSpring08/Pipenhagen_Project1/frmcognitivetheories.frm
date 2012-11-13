VERSION 5.00
Begin VB.Form frmcognitivetheories 
   BackColor       =   &H00008000&
   Caption         =   "Form1"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   3735
      Left            =   10680
      Picture         =   "frmcognitivetheories.frx":0000
      ScaleHeight     =   3675
      ScaleWidth      =   4155
      TabIndex        =   6
      Top             =   2520
      Width           =   4215
   End
   Begin VB.CommandButton cmdsystematicdesensitization 
      BackColor       =   &H00C0C000&
      Caption         =   "Systematic Desensitization"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   720
      Width           =   3375
   End
   Begin VB.CommandButton cmdbehavioralrehersal 
      BackColor       =   &H00C0C000&
      Caption         =   "Behavioral Rehersal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2160
      Width           =   3375
   End
   Begin VB.PictureBox picresults 
      BackColor       =   &H00C0C0C0&
      Height          =   7335
      Left            =   3840
      ScaleHeight     =   7275
      ScaleWidth      =   6555
      TabIndex        =   3
      Top             =   720
      Width           =   6615
   End
   Begin VB.CommandButton cmdcontingencymanagement 
      BackColor       =   &H00C0C000&
      Caption         =   "Contingency Management"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3600
      Width           =   3375
   End
   Begin VB.CommandButton cmdrationalreconstructuring 
      BackColor       =   &H00C0C000&
      Caption         =   "Rational Reconstructuring"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5040
      Width           =   3375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C000&
      Caption         =   "Back"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6480
      Width           =   3375
   End
End
Attribute VB_Name = "frmcognitivetheories"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project name: A Review of Theoretical Orientations in Clinical Psychology
'Form name: frmcognitivetheories
'Author: Calvin Pipenhagen
'Date Written: March 12, 2008
'Objective: To present information about important cognitive-behavioral theories
Option Explicit
Private Sub cmdbehavioralrehersal_Click() 'loads a data file with text relating to behavioral rehersal.
Dim behavioralr(1 To 100) As String
Dim ctr As Integer
Dim n As Integer
picresults.Cls
Open App.Path & "\behavioralr.txt" For Input As #1 'the file is comprised of lines of text describing behavioral rehersal
    Do Until EOF(1)
        ctr = ctr + 1
        Input #1, behavioralr(ctr)
    Loop
    Close #1
For n = 1 To ctr 'each line of the file is printed
    picresults.Print behavioralr(n)
Next n

End Sub

Private Sub cmdcontingencymanagement_Click() 'loads a data file with text relating to contingency management
Dim contingencym(1 To 100) As String
Dim ctr As Integer
Dim n As Integer
picresults.Cls
Open App.Path & "\contingencym.txt" For Input As #1
    Do Until EOF(1)
        ctr = ctr + 1
        Input #1, contingencym(ctr)
    Loop
    Close #1
For n = 1 To ctr
    picresults.Print contingencym(n)
Next n
End Sub

Private Sub cmdrationalreconstructuring_Click() 'loads a data file with text relating to cognitive restructuring
Dim rationalr(1 To 100) As String
Dim ctr As Integer
Dim n As Integer
picresults.Cls
Open App.Path & "\rationalr.txt" For Input As #1
    Do Until EOF(1)
        ctr = ctr + 1
        Input #1, rationalr(ctr)
    Loop
    Close #1
For n = 1 To ctr
    picresults.Print rationalr(n)
Next n
End Sub

Private Sub cmdsystematicdesensitization_Click() 'loads a data file with text relating to systematic desensitization
Dim systematicd(1 To 100) As String
Dim ctr As Integer
Dim n As Integer
picresults.Cls
Open App.Path & "\systematicd.txt" For Input As #1
    Do Until EOF(1)
        ctr = ctr + 1
        Input #1, systematicd(ctr)
    Loop
    Close #1
For n = 1 To ctr
    picresults.Print systematicd(n)
Next n
End Sub


Private Sub Command1_Click() 'returns to the main cognitive-behavioral page
frmcognitivetheories.Hide
frmcognitivebehavioral.Show
End Sub

