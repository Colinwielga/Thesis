VERSION 5.00
Begin VB.Form FrmBMI 
   BackColor       =   &H00FF8080&
   Caption         =   "BMI"
   ClientHeight    =   7860
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10470
   LinkTopic       =   "Form1"
   ScaleHeight     =   7860
   ScaleWidth      =   10470
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdMain 
      BackColor       =   &H00FFFF00&
      Caption         =   "Return to Main Menu"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6960
      Width           =   2055
   End
   Begin VB.CommandButton cmdExercise 
      BackColor       =   &H0080FF80&
      Caption         =   "Return to Exercise Page"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6960
      Width           =   2055
   End
   Begin VB.CommandButton cmdInfo 
      BackColor       =   &H0080FFFF&
      Caption         =   "What does my BMI mean?"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1800
      Width           =   2055
   End
   Begin VB.CommandButton cmdBMI 
      BackColor       =   &H00FF80FF&
      Caption         =   "My BMI"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2280
      MaskColor       =   &H0000FF00&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1800
      Width           =   1815
   End
   Begin VB.PictureBox picBMI 
      BackColor       =   &H00FF8080&
      Height          =   3735
      Left            =   1080
      ScaleHeight     =   3675
      ScaleWidth      =   7515
      TabIndex        =   1
      Top             =   3000
      Width           =   7575
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      Caption         =   "Calculate Your BMI"
      BeginProperty Font 
         Name            =   "Myriad Condensed Web"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   1095
      Left            =   840
      TabIndex        =   0
      Top             =   240
      Width           =   8895
   End
End
Attribute VB_Name = "FrmBMI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Be Healthy (VBFinalProject.vbp)
'Form Name: BMI (FrmBMI.frm)
'Author: Ali Kluthe
'Date: 10/27/2005
'Objective: The purpose of this form is to calculate the user's BMI and inform them of what their BMI means.

Private Sub cmdBMI_Click() 'This button allows the user to calculate their BMI
picBMI.Cls 'clears the picture box
Dim Height As Single 'Declares height as a single variable
Dim Weight As Single 'Declares weight as a single variable
Dim BMI As Single 'Declares BMI as a single variable
Height = InputBox("Enter your height in inches", "Height") 'Height is entered by the user through an input box
Weight = InputBox("Enter your weight in pounds", "Weight") ' Weight is enetered by the user through an input box
BMI = ((Weight * 703) / (Height ^ 2)) 'The BMI is calulated using this equation
    If BMI < 18.5 Then 'If the BMI is less than 18.5 then move onto the next step
        picBMI.Print "Your BMI is"; " "; FormatNumber(BMI, 2); "."; " "; "You are underweight." 'Prints the BMI and a message that says "You are underweight"
    ElseIf 18.5 <= BMI And BMI <= 24.9 Then 'If the BMI is between 18.5 and 24.9 then move onto the next step
        picBMI.Print "Your BMI is"; " "; FormatNumber(BMI, 2); "."; " "; "You are normal." 'Prints the BMI and the message "You are normal"
    ElseIf 25 <= BMI And BMI <= 29.9 Then 'If the BMI is between 25 and 29.9 then move onto the next step
        picBMI.Print "Your BMI is"; " "; FormatNumber(BMI, 2); "."; " "; "You are overweight." 'Prints the BMI and the message "You are overweight"
    Else: picBMI.Print "Your BMI is"; " "; FormatNumber(BMI, 2); "."; " "; "You are obese." 'Otherwise it print the BMI and the message "You are obese"
    End If 'Ends the If statement

End Sub

Private Sub cmdExercise_Click() 'This button allows the user to go to the exercise form
FrmBMI.Hide 'Hides the BMI Form
FrmExercise.Show 'Shows the Exercise Form
End Sub

Private Sub cmdInfo_Click() 'This button allows the user to see information about their BMI.
picBMI.Print "The higher your BMI, the greater your risk for diseases such as diabetes, heart disease, arthritis," 'Prints the information in the quotes in a picture box.
picBMI.Print "and certain cancers.Because BMI does not show the difference between fat and muscle, it does not" 'Prints the information in the quotes in a picture box.
picBMI.Print "always accurately predict when weight could lead to health problems. For example, someone with a lot" 'Prints the information in the quotes in a picture box.
picBMI.Print "of muscle (such as a body builder) may have a BMI in the unhealthy range, but still be healthy and" 'Prints the information in the quotes in a picture box.
picBMI.Print "have little risk of developing diabetes or having a heart attack." 'Prints the information in the quotes in a picture box.

End Sub

Private Sub cmdMain_Click() 'This button allows the user to return to the main menu.
FrmBMI.Hide 'Hides the BMI form
FrmMain.Show 'Shows the Main form
End Sub
