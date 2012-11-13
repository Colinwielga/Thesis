VERSION 5.00
Begin VB.Form FrmMain 
   BackColor       =   &H00000000&
   Caption         =   "Main Menu"
   ClientHeight    =   7350
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9705
   FillColor       =   &H00FFFFFF&
   BeginProperty Font 
      Name            =   "Bradley Hand ITC"
      Size            =   27.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   7350
   ScaleWidth      =   9705
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H0080FFFF&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Bradley Hand ITC"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6480
      Width           =   1575
   End
   Begin VB.CommandButton cmdQuiz 
      BackColor       =   &H00FF00FF&
      Caption         =   "Test my Knowledge"
      BeginProperty Font 
         Name            =   "Bradley Hand ITC"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5160
      Width           =   3495
   End
   Begin VB.CommandButton cmdSleep 
      BackColor       =   &H00FFFF00&
      Caption         =   "Learn about Sleep"
      BeginProperty Font 
         Name            =   "Bradley Hand ITC"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3720
      Width           =   2175
   End
   Begin VB.CommandButton cmdExercise 
      BackColor       =   &H000080FF&
      Caption         =   "Learn about Exercise"
      BeginProperty Font 
         Name            =   "Bradley Hand ITC"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3720
      Width           =   2175
   End
   Begin VB.CommandButton cmdFood 
      BackColor       =   &H000000FF&
      Caption         =   "Learn about Food"
      BeginProperty Font 
         Name            =   "Bradley Hand ITC"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   600
      MaskColor       =   &H00FF8080&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3720
      Width           =   2175
   End
   Begin VB.Image Image3 
      Height          =   1140
      Left            =   6840
      Picture         =   "FrmMain.frx":0000
      Top             =   2040
      Width           =   1860
   End
   Begin VB.Image Image2 
      Height          =   1455
      Left            =   4200
      Picture         =   "FrmMain.frx":0BFD
      Top             =   1920
      Width           =   1200
   End
   Begin VB.Image Image1 
      Height          =   1200
      Left            =   960
      Picture         =   "FrmMain.frx":14B4
      Top             =   1920
      Width           =   1470
   End
   Begin VB.Label lblSubtitle 
      BackColor       =   &H80000012&
      Caption         =   "   Click on a button to learn more about staying healthy!"
      BeginProperty Font 
         Name            =   "Bradley Hand ITC"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   615
      Left            =   960
      TabIndex        =   1
      Top             =   1200
      Width           =   7575
   End
   Begin VB.Label lblTitle 
      BackColor       =   &H00000000&
      Caption         =   "                Be Healthy......."
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   840
      TabIndex        =   0
      Top             =   240
      Width           =   7455
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Be Healthy (VBFinalProject.vbp)
'Form Name: Main Menu (FrmMain.frm)
'Author: Ali Kluthe
'Date: 10/27/2005
'Purpose: The purpose of this project is to teach the user about healthy habits and evaluate their health.
'Objective: This form is main the main menu for the project. It allows the user to chose a button and move to a different form.

Private Sub cmdExercise_Click() 'This button allows the user to see the exercise form.
FrmMain.Hide 'Hides the main form
FrmExercise.Show 'Shows the exercise form

End Sub

Private Sub cmdFood_Click() 'This button allows the user to see the food form.
FrmMain.Hide 'Hides the main form
FrmFood.Show 'Shows the food form

End Sub

Private Sub cmdQuit_Click() 'This button ends the program.
End 'Ends the program
End Sub

Private Sub cmdQuiz_Click() 'This button allows the user to see the quiz form.
FrmMain.Hide 'Hides the main form
FrmQuiz.Show 'Shows the quiz form
End Sub

Private Sub cmdSleep_Click() 'This button allows the user to see the sleep form.
FrmMain.Hide 'Hides the main form
FrmSleep.Show 'Shows the sleep form
End Sub

