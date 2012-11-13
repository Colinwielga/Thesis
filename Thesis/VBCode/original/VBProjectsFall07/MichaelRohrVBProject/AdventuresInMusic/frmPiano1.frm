VERSION 5.00
Begin VB.Form frmPiano1 
   Caption         =   "Learning the Keyboard"
   ClientHeight    =   9120
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11220
   LinkTopic       =   "Form1"
   Picture         =   "frmPiano1.frx":0000
   ScaleHeight     =   9120
   ScaleWidth      =   11220
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdNext 
      BackColor       =   &H0080FF80&
      Caption         =   "Test Yourself!"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   9600
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7560
      Width           =   1455
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Back"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7560
      Width           =   1455
   End
   Begin VB.PictureBox picKeyboard_With_Notes 
      AutoSize        =   -1  'True
      Height          =   4815
      Left            =   2160
      Picture         =   "frmPiano1.frx":2DBBC2
      ScaleHeight     =   4755
      ScaleWidth      =   6930
      TabIndex        =   0
      Top             =   2760
      Width           =   6990
   End
   Begin VB.Label lblKeyboard 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      Caption         =   $"frmPiano1.frx":3472C0
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   1335
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   10935
   End
   Begin VB.Label lblPiano 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      Caption         =   "LEARNING THE KEYBOARD"
      BeginProperty Font 
         Name            =   "Tempus Sans ITC"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   1095
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   10695
   End
End
Attribute VB_Name = "frmPiano1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'This page informs the user about the piano keyboard using a picture to explain the keyboard, the notes and the correlation to the notes on the staff
'it then proceeds to a quiz about the keyboard

Private Sub cmdBack_Click()     'This button changes forms to frmLessonMainPage
    frmPiano1.Hide                  'this hides frmPiano1
    frmLessonMainPage.Show          'this makes frmLessonMainPage visible
End Sub

Private Sub cmdNext_Click()     'This button changes forms to frmPiano2
    frmPiano1.Hide                  'this hide frmPiano1
    frmPiano2.Show                  'this makes frmPiano2 visible
End Sub
