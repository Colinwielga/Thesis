VERSION 5.00
Begin VB.Form frmResults 
   BackColor       =   &H003D30AD&
   Caption         =   "Results"
   ClientHeight    =   9975
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   ScaleHeight     =   9975
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picResults 
      BackColor       =   &H0084C11E&
      Height          =   1335
      Left            =   3840
      ScaleHeight     =   1275
      ScaleWidth      =   6555
      TabIndex        =   6
      Top             =   4560
      Width           =   6615
   End
   Begin VB.CommandButton cmdhome 
      Caption         =   "Go Back to Homepage"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   8520
      TabIndex        =   5
      Top             =   6360
      Width           =   2055
   End
   Begin VB.PictureBox picLumiere 
      Height          =   4215
      Left            =   11040
      Picture         =   "frmResults.frx":0000
      ScaleHeight     =   4155
      ScaleWidth      =   3075
      TabIndex        =   4
      Top             =   120
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.PictureBox picGaston 
      Height          =   3495
      Left            =   7920
      Picture         =   "frmResults.frx":2244
      ScaleHeight     =   3435
      ScaleWidth      =   2955
      TabIndex        =   3
      Top             =   360
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.CommandButton cmdResults 
      Caption         =   "Click to find out which character you are"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   2520
      TabIndex        =   2
      Top             =   6240
      Width           =   3255
   End
   Begin VB.PictureBox picBeast 
      Height          =   3135
      Left            =   4320
      Picture         =   "frmResults.frx":A41A
      ScaleHeight     =   3075
      ScaleWidth      =   3435
      TabIndex        =   1
      Top             =   480
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.PictureBox picBelle 
      Height          =   3855
      Left            =   720
      Picture         =   "frmResults.frx":DE03
      ScaleHeight     =   3795
      ScaleWidth      =   3435
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   3495
   End
End
Attribute VB_Name = "frmResults"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdhome_Click()
'keeps the picResults from printing repeatedly
picResults.Cls
picBelle.Visible = False
picBeast.Visible = False
picGaston.Visible = False
picLumiere.Visible = False
'brings user back to main form
frmResults.Hide
frmPersonality.Show
'set counters to 0 so character does not repeat
CtrA = 0
CtrB = 0
CtrC = 0
CtrD = 0


End Sub



Private Sub cmdResults_Click()

    picBelle.Visible = False
    picBeast.Visible = False
    picGaston.Visible = False
    picLumiere.Visible = False
    picResults.Cls

'if statement that displays the a characters picture and a description if the user chooses 3 out of 5 options for that character
If CtrA >= 3 Then
    picBelle.Visible = True
    picBeast.Visible = False
    picGaston.Visible = False
    picLumiere.Visible = False
    picResults.Print "You are most like Belle!"
    picResults.Print "You must be beautiful, intelligent, and kind."
ElseIf CtrB >= 3 Then
    picBelle.Visible = False
    picBeast.Visible = True
    picGaston.Visible = False
    picLumiere.Visible = False
    picResults.Print "You are most like the Beast!"
    picResults.Print "You must be hairy, have a short temper,"
    picResults.Print "and potential to be a hot prince. "
ElseIf CtrC >= 3 Then
    picBelle.Visible = False
    picBeast.Visible = False
    picGaston.Visible = True
    picLumiere.Visible = False
    picResults.Print "You are most like Gaston!"
    picResults.Print " You must be strong, conceited and macho."
ElseIf CtrD >= 3 Then
    picBelle.Visible = False
    picBeast.Visible = False
    picGaston.Visible = False
    picLumiere.Visible = True
    picResults.Print "You are most like Lumiere!"
    picResults.Print "You must be romantic, determined and suave."
Else
    picBelle.Visible = True
    picBeast.Visible = True
    picGaston.Visible = True
    picLumiere.Visible = True
    picResults.Print "You have a part of each of these characters in you!"
    picResults.Print "You must be intelligent, hairy, conceited and romantic."
End If

    
    
    
    

End Sub


