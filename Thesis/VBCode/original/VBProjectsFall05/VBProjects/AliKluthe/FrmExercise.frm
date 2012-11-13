VERSION 5.00
Begin VB.Form FrmExercise 
   BackColor       =   &H00000000&
   Caption         =   "Exercise"
   ClientHeight    =   6945
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9510
   FillColor       =   &H0000C000&
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "NancyBlue"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   ScaleHeight     =   6945
   ScaleWidth      =   9510
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdJump 
      BackColor       =   &H000080FF&
      Caption         =   "Calculate    my    BMI"
      Height          =   975
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4440
      Width           =   1455
   End
   Begin VB.CommandButton cmdMenu 
      BackColor       =   &H0000FFFF&
      Caption         =   "Return to Main Menu"
      Height          =   1095
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5520
      Width           =   1575
   End
   Begin VB.PictureBox picCalories 
      BackColor       =   &H00000000&
      FillColor       =   &H0000FFFF&
      BeginProperty Font 
         Name            =   "Pristina"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   2295
      Left            =   3120
      ScaleHeight     =   2235
      ScaleWidth      =   3915
      TabIndex        =   3
      Top             =   4320
      Width           =   3975
   End
   Begin VB.CommandButton cmdExercise 
      BackColor       =   &H000000C0&
      Caption         =   "View Calories Burned from common exercises"
      BeginProperty Font 
         Name            =   "Goudy Old Style"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   360
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4920
      Width           =   2295
   End
   Begin VB.Label lblExercise 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   $"FrmExercise.frx":0000
      BeginProperty Font 
         Name            =   "Eras Light ITC"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   1815
      Left            =   480
      TabIndex        =   4
      Top             =   2640
      Width           =   8775
   End
   Begin VB.Label lblExercise1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   $"FrmExercise.frx":00BD
      BeginProperty Font 
         Name            =   "Maiandra GD"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1815
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   8775
   End
   Begin VB.Label lblTitle 
      BackColor       =   &H00000000&
      Caption         =   "Exercise"
      BeginProperty Font 
         Name            =   "GungsuhChe"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   1455
      Left            =   3240
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "FrmExercise"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Be Healthy (VBFinalProject.vbp)
'Form Name: Main Menu (FrmMain.frm)
'Author: Ali Kluthe
'Date: 10/27/2005
'Objective: The purpose of this form is to inform the user about the benefits of exercise. It also shows the user the number of calories burned from different exercises.

Private Sub cmdExercise_Click() 'This button allows the user to view an array showing the calories burned from different exercises.
Dim Name(1 To 5) As String 'Declares Name as a string variable
Dim Low(1 To 5) As Single 'Declares Low as a single variable
Dim High(1 To 5) As Single 'Declares High as a single variable
Dim I As Integer 'Declares I as an integer
Open App.Path & "\Exercise.txt" For Input As #1 'Opens the Exercise file and sets it as input #1
picCalories.Print "Activity"; "       "; "Range of Calories Burned" 'prints the word in quotations
picCalories.Print "*************************************************" 'prints the word in quotations
For I = 1 To 5 'While I is between 1 and 5 do steps below
    Input #1, Name(I), Low(I), High(I) 'Sets the variables in the array
    picCalories.Print Name(I), Low(I); " "; "to"; " "; High(I) 'prints the array values in a picture box
Next I 'Moves to the next I

End Sub

Private Sub cmdJump_Click() 'This button allows the user to see the BMI form.
FrmExercise.Hide 'Hides the exercise form
FrmBMI.Show 'Shows the BMI form

End Sub

Private Sub CmdMenu_Click() 'This button allows the user to return to the main menu.
FrmExercise.Hide 'Hides the exercise form
FrmMain.Show 'Shows the main form

End Sub

