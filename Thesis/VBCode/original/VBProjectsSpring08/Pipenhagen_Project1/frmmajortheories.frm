VERSION 5.00
Begin VB.Form frmmajortheories 
   BackColor       =   &H00FF80FF&
   Caption         =   "Form1"
   ClientHeight    =   9495
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14880
   LinkTopic       =   "Form1"
   ScaleHeight     =   9495
   ScaleWidth      =   14880
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   3495
      Left            =   11040
      Picture         =   "frmmajortheories.frx":0000
      ScaleHeight     =   3435
      ScaleWidth      =   3435
      TabIndex        =   6
      Top             =   2040
      Width           =   3495
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFC0FF&
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
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6240
      Width           =   3375
   End
   Begin VB.CommandButton cmdfreeassosiation 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Free Association"
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
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4800
      Width           =   3375
   End
   Begin VB.CommandButton cmddefensemechanisms 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Defense Mechanisms"
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
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3360
      Width           =   3375
   End
   Begin VB.PictureBox picresults 
      Height          =   7335
      Left            =   3840
      ScaleHeight     =   7275
      ScaleWidth      =   6675
      TabIndex        =   2
      Top             =   360
      Width           =   6735
   End
   Begin VB.CommandButton cmdpsychosexualstages 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Psychosexual Stages"
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
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1920
      Width           =   3375
   End
   Begin VB.CommandButton cmdpersonality 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Personality Structures"
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
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   360
      Width           =   3375
   End
End
Attribute VB_Name = "frmmajortheories"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project name: A Review of Theoretical Orientations in Clinical Psychology
'Form name: frmmajortheories
'Author: Calvin Pipenhagen
'Date Written: March 9, 2008
'Objective: To present information about important Psychodynamic theories. All info is
           'loaded from data files.
Option Explicit
Private Sub cmddefensemechanisms_Click() 'loads the data file corresponding to defense mechanisms.
Dim egodefense(1 To 100) As String
Dim ctr As Integer
Dim n As Integer
picresults.Cls
Open App.Path & "\egodefenses.txt" For Input As #1 'each line of text in this file is enclosed in quotes
    Do Until EOF(1)
        ctr = ctr + 1
        Input #1, egodefense(ctr)
    Loop
    Close #1
For n = 1 To ctr 'each line of the file is printed.
    picresults.Print egodefense(n)
Next n
End Sub

Private Sub cmdfreeassosiation_Click() 'loads the data file corresponding to free association.
Dim freeassociation(1 To 100) As String
Dim ctr As Integer
Dim n As Integer
picresults.Cls
Open App.Path & "\freeassociation.txt" For Input As #1
    Do Until EOF(1)
        ctr = ctr + 1
        Input #1, freeassociation(ctr)
    Loop
    Close #1
For n = 1 To ctr
    picresults.Print freeassociation(n)
Next n
End Sub

Private Sub cmdpersonality_Click() 'loads the data file corresponding to Freud's theory of personality.
Dim personality(1 To 100) As String
Dim ctr As Integer
Dim n As Integer
picresults.Cls
Open App.Path & "\personalitystructure.txt" For Input As #1
    Do Until EOF(1)
        ctr = ctr + 1
        Input #1, personality(ctr)
    Loop
    Close #1
For n = 1 To ctr
    picresults.Print personality(n)
Next n
End Sub



Private Sub cmdpsychosexualstages_Click() 'loads the data file corresponding to the psychosexual stages.
Dim psychosexual(1 To 100) As String
Dim ctr As Integer
Dim n As Integer
picresults.Cls
Open App.Path & "\psychosexualstages.txt" For Input As #1
    Do Until EOF(1)
        ctr = ctr + 1
        Input #1, psychosexual(ctr)
    Loop
    Close #1
For n = 1 To ctr
    picresults.Print psychosexual(n)
Next n
End Sub

Private Sub Command1_Click() 'goes back to the main psychodynamic page
frmmajortheories.Hide
frmpsychodynamic.Show
End Sub
