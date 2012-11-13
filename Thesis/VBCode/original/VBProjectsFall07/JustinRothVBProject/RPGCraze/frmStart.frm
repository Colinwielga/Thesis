VERSION 5.00
Begin VB.Form frmStart 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Start"
   ClientHeight    =   5745
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5880
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   5880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   615
      Left            =   2280
      TabIndex        =   3
      Top             =   4800
      Width           =   1455
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start!"
      Height          =   735
      Left            =   1920
      TabIndex        =   1
      Top             =   3960
      Width           =   2175
   End
   Begin VB.TextBox txtName 
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      TabIndex        =   0
      Top             =   3000
      Width           =   2175
   End
   Begin VB.Label lblSlogan 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   """A role-playing game"""
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   375
      Left            =   1200
      TabIndex        =   4
      Top             =   2040
      Width           =   3855
   End
   Begin VB.Image imgLogo 
      Height          =   1740
      Left            =   1320
      Picture         =   "frmStart.frx":0000
      Top             =   0
      Width           =   3915
   End
   Begin VB.Line ln1 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   5760
      Y1              =   3720
      Y2              =   3720
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Enter a name for your character:"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   1200
      TabIndex        =   2
      Top             =   2640
      Width           =   3735
   End
End
Attribute VB_Name = "frmStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project name: RPGCraze
'Form name: frmStart
'Author: Justin Roth
'Date Written: Sunday, November 4th, 2007
'Objective of project: The objective of my project is to introduce a new crowd to the RPG (Role-Playing Game) Genre of gameplay.
        'Although it is not full-fledged, it has many of the RPG characteristics.
        'I created this program so that people could have some fun playing around with an RPG-like game.
'Objective of form: This form is the start of the program.
        'This form asks the user to input a name for a character, as well as other customizable character information.
        
Option Explicit

Private Sub cmdStart_Click()
    N = txtName.Text    'defines where the name (N) of the user's character is entered
    
    MsgBox "Please answer the following questions about your character."    'prompts the user to prepare for the upcoming questions.
    
    Do Until CTR = 1
        CTR = CTR + 1   'Increases CTR so the loop can determine whether to stop or not.
        
        Hght = InputBox("What is your character's height in inches?", "Character Height")   'Asks the user to input a height for their character and stores it in the module as public.
        Weight = InputBox("What is your character's weight in lbs?", "Character Weight")    'Asks the user to input a weight for their character and stores it in the module as public.
        Gender = InputBox("What is your character's gender?", "Character Gender")   'Asks the user to input a gender for their character and stores it in the module as public.
        Age = InputBox("What is your character's age in years?", "Character Age")   'Asks the user to input an age for their character and stores it in the module as public.
        
        frmCharacter.Show   'Shows the Character Form (frmCharacter).
        frmStart.Hide   'Hides the Start Form (frmStart) to move on to the next form.
    Loop
    
End Sub

Private Sub cmdQuit_Click()
    End 'Quits the program.
End Sub
