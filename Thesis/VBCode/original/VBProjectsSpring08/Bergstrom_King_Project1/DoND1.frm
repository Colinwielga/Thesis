VERSION 5.00
Begin VB.Form Start 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   7455
   ClientLeft      =   7125
   ClientTop       =   4695
   ClientWidth     =   9855
   FillColor       =   &H00008080&
   ForeColor       =   &H80000014&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   2  'Custom
   ScaleHeight     =   7455
   ScaleWidth      =   9855
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picTitle 
      Height          =   3495
      Left            =   1440
      Picture         =   "DoND1.frx":0000
      ScaleHeight     =   3435
      ScaleWidth      =   7035
      TabIndex        =   6
      Top             =   120
      Width           =   7095
   End
   Begin VB.PictureBox picResults 
      Height          =   1215
      Left            =   240
      ScaleHeight     =   1155
      ScaleWidth      =   9315
      TabIndex        =   5
      Top             =   4560
      Width           =   9375
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H80000016&
      Caption         =   "Goodbye!"
      Height          =   1335
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6000
      Width           =   1695
   End
   Begin VB.CommandButton cmdWelcome 
      BackColor       =   &H80000016&
      Caption         =   "Welcome to Deal or No Deal!"
      Height          =   1335
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6000
      Width           =   1695
   End
   Begin VB.CommandButton cmdCaseValue 
      BackColor       =   &H80000016&
      Caption         =   "How much money is in those cases?"
      Height          =   1335
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6000
      Width           =   1695
   End
   Begin VB.CommandButton cmdRules 
      BackColor       =   &H80000016&
      Caption         =   "What are the Rules?"
      Height          =   1335
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6000
      Width           =   1695
   End
   Begin VB.CommandButton cmdHistory 
      BackColor       =   &H80000016&
      Caption         =   "History of the Show!"
      Height          =   1335
      Left            =   6000
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6000
      Width           =   1695
   End
   Begin VB.Label lblDemonstration 
      Caption         =   "Interactive Demonstration"
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2640
      TabIndex        =   7
      Top             =   3840
      Width           =   4695
   End
End
Attribute VB_Name = "Start"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Deal or No Deal Introduction
'Form Name: Start
'Authors: Chris Bergstrom and Brady King
'Date Written: March 27th, 2008


'Objective of Project:
'The Overall objective of the project is to help the user learn more about Deal or No Deal
'and how it opperates on the television game show. The project offers the following features: an overview of the rules of the game,
'a look at the history of the show, information about past winners, and an understanding of the monetary values of the cases.

'Objective of the Form: The objective of this form is to greet the user.
'Clicking on the buttons will open another form which will take you into more indepths sections about the tv game show.
'Or if not interested, the user can exit.

Private Sub cmdCaseValue_Click() 'Takes the user to the case value form.
Start.Hide 'Hides the Greeting form.
CaseValue.Show 'Displays the Case Value form.
End Sub

Private Sub cmdHistory_Click() 'Takes the user to the history form
Start.Hide 'Hides the Greeting form.
History.Show 'Displays the History form.
End Sub


Private Sub cmdQuit_Click() 'Halts the program.
End 'This ends the project.
End Sub

Private Sub cmdRules_Click() 'Takes the user to the rules form.
Start.Hide 'Hides the Greeting form.
Rules.Show 'Displays the Rules form.
End Sub

Private Sub cmdWelcome_Click() 'Displays a welcoming greeting to the user.
'The picResults.Print, displays the greeting.
picResults.Print "Welcome to Deal Or No Deal!"
picResults.Print "The exhilarating hit game show where contestants play and deal in a high-energy contest of nerves, instincts and raw intuition."
picResults.Print "Click on the links below to find out more about the TV show that is taking the world by storm!"
End Sub
