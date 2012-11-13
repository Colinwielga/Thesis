VERSION 5.00
Begin VB.Form FrmQuiz 
   BackColor       =   &H008080FF&
   Caption         =   "Quiz"
   ClientHeight    =   8220
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10425
   LinkTopic       =   "Form1"
   ScaleHeight     =   8220
   ScaleWidth      =   10425
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdMenu 
      BackColor       =   &H00C0C000&
      Caption         =   "Return to the Main Menu"
      Height          =   375
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   7800
      Width           =   3495
   End
   Begin VB.PictureBox PicResults 
      BackColor       =   &H008080FF&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   3135
      Left            =   5760
      ScaleHeight     =   3075
      ScaleWidth      =   3555
      TabIndex        =   14
      Top             =   4440
      Width           =   3615
   End
   Begin VB.CommandButton cmdTotal 
      BackColor       =   &H000080FF&
      Caption         =   "See my results!"
      Height          =   975
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3360
      Width           =   3615
   End
   Begin VB.TextBox txtQ5 
      BeginProperty Font 
         Name            =   "Tw Cen MT Condensed Extra Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7800
      TabIndex        =   12
      Top             =   2160
      Width           =   1695
   End
   Begin VB.TextBox txtQ4 
      BeginProperty Font 
         Name            =   "Tw Cen MT Condensed Extra Bold"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   2520
      TabIndex        =   10
      Top             =   6600
      Width           =   1455
   End
   Begin VB.TextBox txtQ3 
      BeginProperty Font 
         Name            =   "Tw Cen MT Condensed Extra Bold"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   2520
      TabIndex        =   8
      Top             =   5040
      Width           =   1455
   End
   Begin VB.TextBox txtQ2 
      BeginProperty Font 
         Name            =   "Tw Cen MT Condensed Extra Bold"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   2520
      TabIndex        =   6
      Top             =   3600
      Width           =   1455
   End
   Begin VB.TextBox txtQ1 
      BeginProperty Font 
         Name            =   "Tw Cen MT Condensed Extra Bold"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   2520
      TabIndex        =   4
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Image ImgQ5 
      Height          =   285
      Left            =   9720
      Picture         =   "FrmQuiz.frx":0000
      Top             =   2400
      Width           =   300
   End
   Begin VB.Image ImgQ4 
      Height          =   285
      Left            =   4320
      Picture         =   "FrmQuiz.frx":0342
      Top             =   7080
      Width           =   300
   End
   Begin VB.Image ImgQ3 
      Height          =   285
      Left            =   4320
      Picture         =   "FrmQuiz.frx":0684
      Top             =   5520
      Width           =   300
   End
   Begin VB.Image ImgQ2 
      Height          =   285
      Left            =   4320
      Picture         =   "FrmQuiz.frx":09C6
      Top             =   3960
      Width           =   300
   End
   Begin VB.Image ImgQ1 
      Height          =   285
      Left            =   4320
      Picture         =   "FrmQuiz.frx":0D08
      Top             =   2520
      Width           =   300
   End
   Begin VB.Label lblQ5 
      BackColor       =   &H008080FF&
      Caption         =   "Question 5:  Fill in the Blank, Sleep helps you avoid getting sick and reduces ________. "
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   975
      Left            =   5040
      TabIndex        =   11
      Top             =   2160
      Width           =   2415
   End
   Begin VB.Label lblQ4 
      BackColor       =   &H008080FF&
      Caption         =   "Question 4:                          If you are trying to prevent weight gain how many minutes should you exercise each day?"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1215
      Left            =   0
      TabIndex        =   9
      Top             =   6480
      Width           =   2175
   End
   Begin VB.Label lblQ3 
      BackColor       =   &H008080FF&
      Caption         =   "Question 3:                     How many minutes should you exercise each day?"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   1095
      Left            =   0
      TabIndex        =   7
      Top             =   5040
      Width           =   2175
   End
   Begin VB.Label lblQ2 
      BackColor       =   &H008080FF&
      Caption         =   "Question 2:                 What kind of fat should you try to avoid?"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   1095
      Left            =   120
      TabIndex        =   5
      Top             =   3600
      Width           =   2055
   End
   Begin VB.Label lblQ1 
      BackColor       =   &H008080FF&
      Caption         =   "Question 1:                  How many cups of Fruit do you need each day?"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   975
      Left            =   120
      TabIndex        =   3
      Top             =   2280
      Width           =   2055
   End
   Begin VB.Label lblInstructions 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      Caption         =   $"FrmQuiz.frx":104A
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   10095
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      Caption         =   "Quiz"
      BeginProperty Font 
         Name            =   "Bradley Hand ITC"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   615
      Left            =   3120
      TabIndex        =   1
      Top             =   0
      Width           =   4215
   End
   Begin VB.Label lblSubtitle 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      Caption         =   "Test your knowlege of healthy habits! "
      BeginProperty Font 
         Name            =   "Bradley Hand ITC"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1095
      Left            =   3360
      TabIndex        =   0
      Top             =   600
      Width           =   3855
   End
End
Attribute VB_Name = "FrmQuiz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Be Healthy (VBFinalProject.vbp)
'Form Name: Main Menu (FrmMain.frm)
'Author: Ali Kluthe
'Date: 10/27/2005
'Objective: The purpose of this form is to test the user's knowledge of healthy habits.
Option Explicit
Dim Sum As Integer 'Declares Sum as an integer


Private Sub CmdMenu_Click() 'This button allows the user to see the main menu form.
FrmQuiz.Hide 'Hides the quiz form
FrmMain.Show 'Shows the main form

End Sub

Private Sub cmdTotal_Click() 'This button allows the user to see their total score and percent correct.
Dim Percent As Single 'Declares percent as a single variable
Percent = Sum / 5 'Percent is calulated using that equation
PicResults.Print "Your Total Score is", Sum; " "; "out of 5" 'prints the words in quotations and the sum
PicResults.Print "Percent Correct", FormatPercent(Percent) 'prints the words in quotations and the percentage


End Sub

Private Sub ImgQ1_Click() 'This button allows the user to see if they answered question 1 correctly.
Dim AnsQ1 As Integer 'Declares AnsQ1 as an integer
AnsQ1 = txtQ1.Text 'AnsQ1 is found from the text box txtQ1
    If AnsQ1 = 2 Then 'If the answer is 2 then move onto the next step
        Sum = Sum + 1 'Adds one to the sum
        MsgBox "Correct! Move onto the next question. ", , "Correct Answer" 'prints the words in quotations
    Else: MsgBox "Incorrect! The correct answer is 2. Move onto the next question.", , "Wrong Answer" 'If the answer is inncorect then the words in quotations are printed
    End If 'Closes the If statement
    

End Sub

Private Sub ImgQ2_Click() 'This button allows the user to see if they answered question 2 correctly.
Dim AnsQ2 As String 'Declares AnsQ2 as a string variable
AnsQ2 = txtQ2.Text 'AnsQ2 is found from text box txtQ2
    If AnsQ2 = "trans" Then 'If the answer is trans then move onto the next step
        Sum = Sum + 1 'Adds one to the sum
        MsgBox "Correct! Move onto the next question.", , "Correct Answer" 'prints the words in quotations
    Else: MsgBox "Incorrect! The correct answer is Trans Fat. Move onto the next question.", , "Wrong Answer" 'If the answer is inncorect then the words in quotations are printed
    End If 'Closes the If statement
    
    

End Sub

Private Sub ImgQ3_Click() 'This button allows the user to see if they answered question 3 correctly.
Dim AnsQ3 As Integer 'Declares Ans Q3 as an integer
AnsQ3 = txtQ3.Text 'AnsQ3 is found from text box txtQ3
    If AnsQ3 = 30 Then 'If the answer is 30 then move onto the next step
        Sum = Sum + 1 'Adds one to the sum
        MsgBox "Correct! Move onto the next question.", , "Correct Answer" 'prints the words in quotations
    Else: MsgBox "Incorrect! The correct answer is 30. Move onto the next question.", , "Wrong Answer" 'If the answer is inncorect then the words in quotations are printed
    End If 'Closes the If statement
    
End Sub

Private Sub ImgQ4_Click() 'This button allows the user to see if they answered question 4 correctly.
Dim AnsQ4 As Integer 'Declares AnsQ4 as an integer
AnsQ4 = txtQ4.Text 'AnsQ4 is found from text box txtQ4
    If AnsQ4 = 60 Then 'If the answer is 60 then move onto the next step
        Sum = Sum + 1 'Adds one to the sum
        MsgBox "Correct! Move onto the next question.", , "Correct Answer" 'Prints the words in quotations
    Else: MsgBox "Incorrect! The correct answer is 60. Move onto the next quesion.", , "Wrong Answer" 'If the answer is inncorect then the words in quotations are printed
    End If 'Closes the If statement
End Sub

Private Sub ImgQ5_Click() 'This button allows the user to see if they answered question 5 correctly.
Dim AnsQ5 As String 'Declares AnsQ5 as a string variable
AnsQ5 = txtQ5.Text 'AnsQ5 is found from text box txtQ5
    If AnsQ5 = "stress" Then 'If the answer is stress then move onto the next step
        Sum = Sum + 1 'Adds one to the sum
        MsgBox "Correct! Click on the see me results button to see total score.", , "Correct Answer" 'Prints the words in quotations
    Else: MsgBox "Incorrect! The correct answer is stress. Click on the see me results button to see total score.", , "Wrong Answer" 'If the answer is inncorect then the words in quotations are printed
    End If 'Closes the If statement
    
End Sub

