VERSION 5.00
Begin VB.Form frmBass 
   Caption         =   "Learning the Bass Clef"
   ClientHeight    =   8250
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10815
   LinkTopic       =   "Form1"
   Picture         =   "frmBass.frx":0000
   ScaleHeight     =   8250
   ScaleWidth      =   10815
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdTest 
      BackColor       =   &H0080C0FF&
      Caption         =   "Test Yourself"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   9240
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6840
      Width           =   1455
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Back To Main Page"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6840
      Width           =   1335
   End
   Begin VB.PictureBox picBass 
      AutoSize        =   -1  'True
      Height          =   2985
      Left            =   1560
      Picture         =   "frmBass.frx":32685A
      ScaleHeight     =   2925
      ScaleWidth      =   7260
      TabIndex        =   1
      Top             =   3480
      Width           =   7320
   End
   Begin VB.Label lblBass2 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   $"frmBass.frx":36BAA0
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   1935
      Left            =   480
      TabIndex        =   4
      Top             =   1320
      Width           =   9735
   End
   Begin VB.Label lblBass 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Learning the Bass Clef"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   975
      Left            =   1440
      TabIndex        =   0
      Top             =   240
      Width           =   7815
   End
End
Attribute VB_Name = "frmBass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'The purpose of this page is soley to educate the user about the Bass Clef
'there is a picture box as the main focus with a label above it that help with learning the Bass Clef.
'it also has two buttons located on the page one brings the user back to frmLessonMainPage and the other leads to a quiz on the Bass Clef

Private Sub cmdBack_Click()     'This button changes forms
    frmBass.Hide                    'this hides frmBass
    frmLessonMainPage.Show          'this makes frmLessonMainPage visible
End Sub

Private Sub cmdTest_Click()         'This button changes forms
    frmBass.Hide                    'this hides frmBass
    frmBass2.Show                   'this makes frmBass2 visible
End Sub
