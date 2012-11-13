VERSION 5.00
Begin VB.Form frmhumanistictheories 
   BackColor       =   &H00008080&
   Caption         =   "Form1"
   ClientHeight    =   9495
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12375
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9495
   ScaleWidth      =   12375
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   3495
      Left            =   1320
      Picture         =   "frmhumanistictheories.frx":0000
      ScaleHeight     =   3435
      ScaleWidth      =   8955
      TabIndex        =   6
      Top             =   -1200
      Width           =   9015
   End
   Begin VB.CommandButton cmdback 
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
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7920
      Width           =   3375
   End
   Begin VB.CommandButton cmdlogotherapy 
      BackColor       =   &H00C0C000&
      Caption         =   "Logotherapy"
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
      Top             =   6720
      Width           =   3375
   End
   Begin VB.CommandButton cmdgestalt 
      BackColor       =   &H00C0C000&
      Caption         =   "Gestalt"
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
      Top             =   5520
      Width           =   3375
   End
   Begin VB.PictureBox picresults 
      BackColor       =   &H00C0C0C0&
      Height          =   5895
      Left            =   3960
      ScaleHeight     =   5835
      ScaleWidth      =   6435
      TabIndex        =   2
      Top             =   3120
      Width           =   6495
   End
   Begin VB.CommandButton cmdgrowthpotential 
      BackColor       =   &H00C0C000&
      Caption         =   "Growth Potential"
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
      Top             =   4320
      Width           =   3375
   End
   Begin VB.CommandButton cmdcorefeatures 
      BackColor       =   &H00C0C000&
      Caption         =   "Core Theories"
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
      Top             =   3120
      Width           =   3375
   End
End
Attribute VB_Name = "frmhumanistictheories"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project name: A Review of Theoretical Orientations in Clinical Psychology
'Form name: frmhumanistictheories
'Author: Calvin Pipenhagen
'Date Written: March 25, 2008
'Objective: To present information about important humanistic theories.
Option Explicit
Private Sub cmdback_Click() 'returns to the main humanistic page
frmhumanistictheories.Hide
frmhumanistic.Show
End Sub

Private Sub cmdgrowthpotential_Click() 'loads data file containing information about growth potential
Dim growthpotential(1 To 100) As String
Dim ctr As Integer
Dim n As Integer
picresults.Cls
Open App.Path & "\growthpotential.txt" For Input As #1 'each line of the data file is a line of text related to growth potential
    Do Until EOF(1)
        ctr = ctr + 1
        Input #1, growthpotential(ctr)
    Loop
    Close #1
For n = 1 To ctr
    picresults.Print growthpotential(n) 'each line of the data file is printed in a picturebox
Next n
End Sub

Private Sub cmdgestalt_Click() 'loads data file containing information about gestalt therapy.
Dim gestalt(1 To 100) As String
Dim ctr As Integer
Dim n As Integer
picresults.Cls
Open App.Path & "\gestalt.txt" For Input As #1
    Do Until EOF(1)
        ctr = ctr + 1
        Input #1, gestalt(ctr)
    Loop
    Close #1
For n = 1 To ctr
    picresults.Print gestalt(n)
Next n
End Sub

Private Sub cmdlogotherapy_Click() 'loads a data file containing information about logotherapy
Dim logotherapy(1 To 100) As String
Dim ctr As Integer
Dim n As Integer
picresults.Cls
Open App.Path & "\logotherapy.txt" For Input As #1
    Do Until EOF(1)
        ctr = ctr + 1
        Input #1, logotherapy(ctr)
    Loop
    Close #1
For n = 1 To ctr
    picresults.Print logotherapy(n)
Next n
End Sub

Private Sub cmdcorefeatures_Click() 'loads a data file containing information about the core features of client-centered therapy.
Dim corefeatures(1 To 100) As String
Dim ctr As Integer
Dim n As Integer
picresults.Cls
Open App.Path & "\corefeatures.txt" For Input As #1
    Do Until EOF(1)
        ctr = ctr + 1
        Input #1, corefeatures(ctr)
    Loop
    Close #1
For n = 1 To ctr
    picresults.Print corefeatures(n)
Next n
End Sub




