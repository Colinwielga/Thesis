VERSION 5.00
Begin VB.Form frmScreen1 
   BackColor       =   &H80000007&
   Caption         =   "Wheel of Fortune!"
   ClientHeight    =   5265
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7200
   LinkTopic       =   "Form1"
   ScaleHeight     =   5265
   ScaleWidth      =   7200
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000009&
      Caption         =   "Play Game!"
      BeginProperty Font 
         Name            =   "JazzText"
         Size            =   14.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3360
      Width           =   1815
   End
   Begin VB.Image Image1 
      Height          =   2655
      Left            =   1800
      Picture         =   "WheelFortune2.frx":0000
      Stretch         =   -1  'True
      Top             =   360
      Width           =   3135
   End
End
Attribute VB_Name = "frmScreen1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Wheel of Fortune!(WheelofFortune.vbp)
'Form name: frmScreen1(WheelFortune2.frm); Form caption: Wheel of Fortune!
'Author: Maria Zipp
'Date written: 1st November, 2006
'Form Objective: this is the first form that allows the user to enter
'               his or her name. Then it reads a file of consonants
'               and one of vowels and stores them in seperate arrays.
'               This form is then hidden and the main form (screen2)
'               is visible.


Private Sub Command1_Click()
    nombre = InputBox("What is your name?", "To Start")
    size = 0
    'reads consonants and vowels into seperate arrays
    Open App.Path & "\alphabet.txt" For Input As #1
    Do Until EOF(1)
        Input #1, letter
        size = size + 1
        alphaArray(size) = letter
    Loop
    Close #1
    size = 0
    Open App.Path & "\vowels.txt" For Input As #2
    Do Until EOF(2)
        Input #2, vowel
        size = size + 1
        vowelArray(size) = vowel
    Loop
    Close #2
    'hides first screen, shows game screen
    frmScreen1.Visible = False
    frmScreen2.Visible = True
        
End Sub
