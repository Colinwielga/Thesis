VERSION 5.00
Begin VB.Form frmCalculateCalories 
   BackColor       =   &H00FFFFFF&
   Caption         =   "frmCalculate"
   ClientHeight    =   9585
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14340
   LinkTopic       =   "Form1"
   Picture         =   "frmCalculateCalories.frx":0000
   ScaleHeight     =   9585
   ScaleWidth      =   14340
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   12480
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   8880
      Width           =   1335
   End
   Begin VB.TextBox txtMinutes 
      Alignment       =   1  'Right Justify
      Height          =   615
      Left            =   6360
      TabIndex        =   3
      Text            =   "0"
      Top             =   3960
      Width           =   1455
   End
   Begin VB.CommandButton cmdReturn 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Return to the Main Screen"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   12000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   8160
      Width           =   2295
   End
   Begin VB.CommandButton cmdCaloriesBurned 
      Appearance      =   0  'Flat
      BackColor       =   &H00000080&
      Caption         =   "Second: Push to Calculate the Number of Calories you Burn While Rowing!"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   1560
      MaskColor       =   &H00FFC0C0&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4080
      Width           =   2895
   End
   Begin VB.Label lblInput 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "First: Input Time Spent Rowing ( in Minutes)"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   735
      Left            =   5400
      TabIndex        =   4
      Top             =   3240
      Width           =   3255
   End
   Begin VB.Label lblCalculate 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "How Hard are You Working??"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   2535
      Left            =   2280
      TabIndex        =   0
      Top             =   120
      Width           =   10215
   End
End
Attribute VB_Name = "frmCalculateCalories"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project name: CSB/SJU Crew
'Form name: frmMeettheMembers
'Authors: Lauren Nephew and Rachel Stalley
'Date: October 16th, 2009
'Objective: Allow the user to calculate their calories burned by inputting minutes into a text box and then printing the results by using a message box.
Option Explicit

Private Sub cmdCaloriesBurned_Click()
Dim Minutes As Long
Dim Results As Long
Minutes = txtMinutes.Text 'This makes the text box input a variable to be calculated
Results = Minutes * 13 'This calculates the calories burned
Select Case Results 'this creates a message box that shows the calories burned and adds a statment depending on how many callories burned.
    Case 0 To 100
        MsgBox "You Burned " & Results & " calories. Good job!", , "Calories Burned"
    Case 101 To 200
         MsgBox "You Burned " & Results & " calories. Wow, way to go!", , "Calories Burned"
    Case 201 To 500
        MsgBox "You Burned " & Results & " calories. You are awesome!", , "Calories Burned"
    Case 501 To 800
        MsgBox "You Burned " & Results & " calories. Great work out!", , "Calories Burned"
    Case 801 To 1500
        MsgBox "You Burned " & Results & " calories. Great row!", , "Calories Burned"
    Case Is > 1500
        MsgBox "You Burned " & Results & " calories. You are an amazing rower!", , "Calories Burned"
    End Select

End Sub

Private Sub cmdQuit_Click()
End
End Sub

Private Sub cmdReturn_Click()
'this brings the user back to the main menu screen
frmCSBSJUCrewMain.Show
frmCalculateCalories.Hide
End Sub
