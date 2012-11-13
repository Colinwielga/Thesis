VERSION 5.00
Begin VB.Form Form24 
   BackColor       =   &H00000000&
   Caption         =   "Form24"
   ClientHeight    =   8835
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7875
   LinkTopic       =   "Form24"
   Picture         =   "Music.frx":0000
   ScaleHeight     =   8835
   ScaleWidth      =   7875
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Main Menu"
      Height          =   435
      Left            =   0
      TabIndex        =   0
      Top             =   8400
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "*Because of limited disk space, only one song can be available"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   720
      TabIndex        =   4
      Top             =   8040
      Width           =   7455
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "(Click Box)"
      Height          =   255
      Left            =   3240
      TabIndex        =   3
      Top             =   5640
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Dream Theatre - Another Day"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   1800
      TabIndex        =   2
      Top             =   2760
      Width           =   4335
   End
   Begin VB.OLE OLE1 
      BackStyle       =   0  'Transparent
      Class           =   "SoundRec"
      Height          =   1335
      Left            =   2400
      OleObjectBlob   =   "Music.frx":BD70
      TabIndex        =   1
      Top             =   3840
      Width           =   2895
   End
End
Attribute VB_Name = "Form24"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'WesternArt Project
'Form24
'Bursh,Wrobel
'11-1-06
'This is our Music Form, which allows you to play Dream Theatres - Another Day
'while taking the Western Art History Tour.  Unfortunately there is no found way
'to stop the file from playing.
Option Explicit
Private Sub Command1_Click()
 Form1.Show
 Form24.Hide
End Sub

