VERSION 5.00
Begin VB.Form FrmMain 
   BackColor       =   &H00000000&
   Caption         =   "Main"
   ClientHeight    =   9315
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11535
   LinkTopic       =   "Form1"
   ScaleHeight     =   9315
   ScaleWidth      =   11535
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00FF80FF&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Broadway"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   8520
      Width           =   5415
   End
   Begin VB.CommandButton cmdInBetweenMain 
      BackColor       =   &H00800080&
      Caption         =   "STEP #3: Everything Else in Between"
      BeginProperty Font 
         Name            =   "Broadway"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5640
      Width           =   3495
   End
   Begin VB.CommandButton cmdExerciseMain 
      BackColor       =   &H00FF0000&
      Caption         =   "STEP #2: Exercise"
      BeginProperty Font 
         Name            =   "Broadway"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5640
      Width           =   3495
   End
   Begin VB.CommandButton cmdNutritionMain 
      BackColor       =   &H0000FF00&
      Caption         =   "STEP #1: Nutrition"
      BeginProperty Font 
         Name            =   "Broadway"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5640
      Width           =   3615
   End
   Begin VB.Label lblMain3 
      BackColor       =   &H0000FFFF&
      Caption         =   "Just Follow These Three Easy Steps!!!"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1800
      TabIndex        =   5
      Top             =   4200
      Width           =   7815
   End
   Begin VB.Label lblMain2 
      BackColor       =   &H000080FF&
      Caption         =   "Find Out How Healthy You Are......."
      BeginProperty Font 
         Name            =   "Blackadder ITC"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   720
      TabIndex        =   1
      Top             =   2640
      Width           =   10215
   End
   Begin VB.Label lblMain1 
      BackColor       =   &H000000FF&
      Caption         =   "  Bennie Health 101:"
      BeginProperty Font 
         Name            =   "Onyx"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   2040
      TabIndex        =   0
      Top             =   600
      Width           =   7335
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    'Bennie Health Project
    'FrmMain
    'Heidi Donnelly
    'Written: 9/23
    'The purpose of this form is to display the three main areas of health that college-aged women should be aware of: Nutrition, Exercise, and Everything Else and allow navigation to these areas main forms.
    
Private Sub cmdExerciseMain_Click()
'leads to exercise form
    MsgBox ("I am sure you already checked out the nutrition portion but now we're all curious just how much ") & UserName & (" exercises...! he-he!")
    FrmMain.Hide
    FrmExerciseMain.Show
End Sub

Private Sub cmdInBetweenMain_Click()
'leads to everything else form
    FrmMain.Hide
    FrmEverythingElseInBetweenMain.Show
End Sub

Private Sub cmdNutritionMain_Click()
'leads to nutrition form
    FrmMain.Hide
    FrmNutritionMain.Show
End Sub

Private Sub cmdQuit_Click()
    End
End Sub
