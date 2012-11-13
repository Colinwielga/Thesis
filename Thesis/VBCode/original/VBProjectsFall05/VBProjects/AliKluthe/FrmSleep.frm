VERSION 5.00
Begin VB.Form FrmSleep 
   BackColor       =   &H00FFFF00&
   Caption         =   "Sleep"
   ClientHeight    =   6990
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9615
   LinkTopic       =   "Form1"
   ScaleHeight     =   6990
   ScaleWidth      =   9615
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdMenu 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Return to Main Menu"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6360
      Width           =   1575
   End
   Begin VB.PictureBox picSleep 
      BackColor       =   &H00FFFF80&
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   360
      ScaleHeight     =   1515
      ScaleWidth      =   8955
      TabIndex        =   8
      Top             =   4560
      Width           =   9015
   End
   Begin VB.CommandButton cmdSleep 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Am I getting enough sleep?"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3600
      Width           =   2655
   End
   Begin VB.TextBox txtHour 
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   7440
      TabIndex        =   6
      Top             =   2280
      Width           =   1695
   End
   Begin VB.TextBox txtAge 
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   2640
      TabIndex        =   4
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Label lblHours 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "Enter the average hours of sleep you get per night"
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
      Left            =   5040
      TabIndex        =   5
      Top             =   2400
      Width           =   1935
   End
   Begin VB.Label lblAge 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "Enter your age in years"
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
      Left            =   360
      TabIndex        =   3
      Top             =   2400
      Width           =   1935
   End
   Begin VB.Label lblQuestion 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "Are you getting enough sleep?  Input your info below to find out!"
      BeginProperty Font 
         Name            =   "Bradley Hand ITC"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   735
      Left            =   240
      TabIndex        =   2
      Top             =   1440
      Width           =   9135
   End
   Begin VB.Label lblSleep 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "Sleep is an important part of staying healthy. Sleep helps you avoid getting sick, and reduces stress."
      BeginProperty Font 
         Name            =   "Bradley Hand ITC"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   855
      Left            =   480
      TabIndex        =   1
      Top             =   720
      Width           =   8655
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "Sleep"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1215
      Left            =   2280
      TabIndex        =   0
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "FrmSleep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Be Healthy (VBFinalProject.vbp)
'Form Name: Main Menu (FrmMain.frm)
'Author: Ali Kluthe
'Date: 10/27/2005
'Objective: The purpose of this form is to inform the user about the benefits of sleep. It also evaluates whether or not the user is getting enough sleep.

Private Sub CmdMenu_Click() 'This button allows the user to return to the main form.
FrmSleep.Hide 'Hides the sleep form
FrmMain.Show 'Shows the main form

End Sub

Private Sub cmdSleep_Click() 'This button allows the user to see if they are getting enough sleep.
picSleep.Cls 'Clears the picture box
Dim Age As Integer 'Declares Age as an integer
Dim Hour As Single 'Declares Hour as a single variable
Age = txtAge.Text 'Age is found in the text box txtage.txt
Hour = txtHour.Text 'Hour is found in the txthour.txt
If Age <= 3 Then 'If the age is less than three move onto the next step
    If Hour >= 14 Then 'If hours are greater than or equal to 14 then move onto the next step
    picSleep.Print "Congratulations! You are getting enough sleep." 'prints the words in quotations
    Else: picSleep.Print "Sorry! You are not getting enough sleep. It is reccomended that you get at least 14 hours of sleep per night." 'If the conditions are not met then the words in quotations are printed
    End If 'Closes the inside if statement
End If 'Closes the outside if statement
If Age > 3 And Age <= 12 Then 'If the age is between 3 and 15 move to the next step
    If Hour >= 10 Then 'If hour is greater than or equal to 10 then move onto the next step
    picSleep.Print "Congratulations! You are getting enough sleep." 'prints words in qutations
    Else: picSleep.Print "Sorry! You are not getting enough sleep. It is reccomended that you get at least 10 hours of sleep per night." 'If the conditions are not met then the words in quotations are printed
    End If 'Closes the inside if statement
End If 'Closes the outside if statement
If Age > 12 And Age <= 19 Then 'If the age is between 12 and 19 then move onto the next step
    If Hour >= 9 Then 'If Hour is greater than or equal to 9 then move onto the next step
    picSleep.Print "Congratulations! You are getting enough sleep." 'prints the words in quotations
    Else: picSleep.Print "Sorry! You are not getting enough sleep. It is reccomended that you get at least 9 hours of sleep per night." 'If the conditions are not met then the words in quotations are printed
    End If 'Closes the inside if statement
End If 'Closes the outside if statement
If Age > 19 Then 'If the age is greater than 19 move onto the next step
    If Hour >= 8 Then 'If Hour is greater than or equal to 8 then move onto the next step
    picSleep.Print "Congratulations! You are getting enough sleep." 'prints the words in quotations
    Else: picSleep.Print "Sorry! You are not getting enough sleep. It is reccomended that you get at least 8 hours of sleep per night." 'If the conditions are not met then the words in quotations are printed
    End If 'Closes the inside if statement
End If 'Closes the outside if statement


End Sub

