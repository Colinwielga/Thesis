VERSION 5.00
Begin VB.Form psychentryform 
   BackColor       =   &H000000C0&
   Caption         =   "PSYCHOLOGY TERMS AND DEFINITIONS by SARAH AHLFS "
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox psychpic 
      Height          =   3975
      Left            =   2280
      Picture         =   "psychform2.frx":0000
      ScaleHeight     =   3915
      ScaleWidth      =   9675
      TabIndex        =   3
      Top             =   3120
      Width           =   9735
   End
   Begin VB.CommandButton cmdquit 
      BackColor       =   &H00C0C000&
      Caption         =   "QUIT"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   18
         Charset         =   1
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   9480
      Width           =   2775
   End
   Begin VB.CommandButton cmdenterproject 
      BackColor       =   &H00FF0000&
      Caption         =   "ENTER PROJECT"
      BeginProperty Font 
         Name            =   "Batang"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7560
      Width           =   3855
   End
   Begin VB.Label lblpsych2 
      Caption         =   "PSYCHOLOGY TERMS AND DEFINITIONS PROGRAM"
      BeginProperty Font 
         Name            =   "Informal Roman"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   1440
      TabIndex        =   2
      Top             =   240
      Width           =   11295
   End
End
Attribute VB_Name = "psychentryform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Psychology Terms and Definitions (psych_project.vbp)
'Psychinfoform (under psych_form.frm and psych_project.vbp)
'Psychentryform (under psychform2.frm and psych_project.vbp)
'Sarah Ahlfs
'October 20th, 2003
'Purpose: The overall purpose of this project is to have a database of psychology terms and definitions (since I'm a psych major) that can be easily seen, alphabetized, and searched through.  The pictures are reminders of some of the most significant people in psychology and 2 of my favorite psychology studies that have been done.
         'I will be able to use this throughout my career as a psychology major.  I can add to it and continually use it for quick reference and help on assignments if needed.
         'the purpose of this specific form (psychentryform) is just to serve as an entrance into the next form which contains all the information
Option Explicit
Private Sub cmdenterproject_Click() 'when this button is hit it does the following:
psychinfoform.Visible = True 'this says to show psychform1 - the form with all the info (other form)
psychentryform.Visible = False 'this says to hide psychform2 - the entry form(this form)
End Sub

Private Sub cmdquit_Click() 'this ends the running of the program
End
End Sub

'NOTE: psychology defitions were taken from:Psychology Applied to Modern Life study guide and pictures were taken from http://psych.wisc.edu/henriques/resources/Images.html
