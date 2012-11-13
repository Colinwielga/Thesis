VERSION 5.00
Begin VB.Form frmSummary 
   BackColor       =   &H00C0C000&
   Caption         =   "Summary"
   ClientHeight    =   11010
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Mathematica5"
         Size            =   12
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   1800
      TabIndex        =   2
      Top             =   8400
      Width           =   2295
   End
   Begin VB.CommandButton cmdSummarize 
      Caption         =   "Summarize Results"
      BeginProperty Font 
         Name            =   "Mathematica5"
         Size            =   36
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3600
      Left            =   960
      TabIndex        =   1
      Top             =   1560
      Width           =   3975
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7815
      Left            =   5640
      ScaleHeight     =   32.313
      ScaleMode       =   4  'Character
      ScaleWidth      =   65.625
      TabIndex        =   0
      Top             =   840
      Width           =   7935
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFC0&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00808000&
      FillStyle       =   3  'Vertical Line
      Height          =   9615
      Left            =   720
      Top             =   360
      Width           =   13455
   End
End
Attribute VB_Name = "frmSummary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: The Game Of Life
'Form Name: frmSummary
'Authors: Pam Arndt and Alisa Davis
'Date Written: 3/4/09
'Objective:This is our ending form for the project
'All entries the user has made have been saved in the program and are now displayed on one form to sum up the choices made.
'Note: If user entered more than one choice, or went back and redid a form, the most recent answer will have been the one saved,
'and will be displayed just as the user had entered them.
'There is also a plug for a great play, Noises Off, which runs 3/26/09-4/4/09

Option Explicit

Private Sub cmdQuit_Click()
End 'Quits program
End Sub

Private Sub cmdSummarize_Click()
picResults.Cls 'clears any information that may be in picturebox (if user clicks button twice for example)
picResults.Print Tab(7); "What an interesting life you have ahead of you, " & UserFirstName & " " & UserLastName & "!"
picResults.Print
picResults.Print "Spouse: " & Spouse
picResults.Print
picResults.Print "Name of first child: " & ChildName
picResults.Print
picResults.Print "Career: " & Career
picResults.Print
picResults.Print "House: " & House
picResults.Print
picResults.Print "Lottery winnings: " & Lottery
picResults.Print
picResults.Print
picResults.Print
picResults.Print Tab(25); "Thank you for playing!"
picResults.Print
picResults.Print
picResults.Print
picResults.Print
picResults.Print
picResults.Print
picResults.Print Tab(20); "Also, you should go see Noises Off!"
End Sub

