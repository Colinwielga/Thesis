VERSION 5.00
Begin VB.Form frmInfo 
   BackColor       =   &H00FF00FF&
   Caption         =   "Info"
   ClientHeight    =   8655
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11685
   FillColor       =   &H00FF00FF&
   BeginProperty Font 
      Name            =   "Script MT Bold"
      Size            =   15.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FF00FF&
   LinkTopic       =   "Form2"
   ScaleHeight     =   8655
   ScaleWidth      =   11685
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtHours 
      Height          =   855
      Left            =   4440
      TabIndex        =   7
      Top             =   4200
      Width           =   2775
   End
   Begin VB.CommandButton cmdFemale 
      BackColor       =   &H0000FFFF&
      Caption         =   "FEMALE"
      BeginProperty Font 
         Name            =   "Script MT Bold"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   3120
      MaskColor       =   &H0000FFFF&
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6720
      Width           =   2055
   End
   Begin VB.CommandButton cmdMale 
      BackColor       =   &H000080FF&
      Caption         =   "MALE"
      BeginProperty Font 
         Name            =   "Mathematica6"
         Size            =   18
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   360
      MaskColor       =   &H000080FF&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6720
      Width           =   2175
   End
   Begin VB.TextBox txtWeight 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5880
      TabIndex        =   3
      Top             =   2400
      Width           =   2775
   End
   Begin VB.TextBox txtAge 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6000
      TabIndex        =   2
      Top             =   600
      Width           =   2535
   End
   Begin VB.Image barimage 
      Height          =   3735
      Left            =   7680
      Picture         =   "Form2.frx":0000
      Top             =   3840
      Width           =   3750
   End
   Begin VB.Label lblHours 
      BackColor       =   &H00FFFF80&
      Caption         =   "How Long Have you Been Drinking? (in hours)"
      Height          =   855
      Left            =   360
      TabIndex        =   8
      Top             =   4200
      Width           =   3495
   End
   Begin VB.Label lblGender 
      BackColor       =   &H00FFFF00&
      Caption         =   "Please Choose your Gender by Clicking on the options below:"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   600
      TabIndex        =   4
      Top             =   5760
      Width           =   4455
   End
   Begin VB.Label lblWeight 
      BackColor       =   &H00FFFF80&
      Caption         =   "And your Weight? (In Lbs)"
      BeginProperty Font 
         Name            =   "MS UI Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3960
      TabIndex        =   1
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label lblAge 
      BackColor       =   &H00FFFF80&
      Caption         =   "What is your Age?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3840
      TabIndex        =   0
      Top             =   480
      Width           =   1935
   End
   Begin VB.Image martiniimage 
      Height          =   3750
      Left            =   0
      Picture         =   "Form2.frx":73B6
      Top             =   120
      Width           =   3750
   End
End
Attribute VB_Name = "frmInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdFemale_Click()
'dim the age and hours as variables
Dim Age As Integer, Hours As Integer
'the information the user inputs in the textboxes for weight age and hours should be used for the complimentary variable names in our equations.
Weight = txtWeight.Text
Age = txtAge.Text
Hours = txtHours.Text
'we need to add all the sums of different types of consumptions in order to find the overall total sum of consumption
sum = (sum1 + sum2 + sum3 + sum4)
'this is the formula for finding blood alcohol level based upon the user's data entries
Bal = (sum * 5.14 / Weight * 0.66) - 0.015 * Hours
'If the user is under the age of 21, it is not legal for them to be consuming.  we inform the user via a message box and then stop the program
If Age < 21 Then
    MsgBox ("You are not of the legal age to be consuming alcohol.  Please discontinue use of Drinking buddy at this time and stop drinking IMMEDIATELY.")
End
End If
'we are switching from the current form to the next form which will display the BAL
frmInfo.Hide
frmBAL.Show
'the BAL we just found will print on the following form
frmBAL.picBal.Print Tab(9); FormatNumber(Bal, 2)
End Sub

Private Sub cmdMale_Click()
'dim the age and hours as variables
Dim Age As Integer, Hours As Integer
'the information the user inputs in the textboxes for weight age and hours should be used for the complimentary variable names in our equations.

Weight = txtWeight.Text
Age = txtAge.Text
Hours = txtHours.Text
'we need to add all the sums of different types of consumptions in order to find the overall total sum of consumption
sum = (sum1 + sum2 + sum3 + sum4)

Bal = (sum * 5.14 / Weight * 0.73) - 0.015 * Hours
'If the user is under the age of 21, it is not legal for them to be consuming.  we inform the user via a message box and then stop the program
If Age < 21 Then
    MsgBox ("You are not of the legal age to be consuming alcohol.  Please discontinue use of Drinking buddy at this time and stop drinking IMMEDIATELY.")
End
End If
'we are switching from the current form to the next form which will display the BAL
frmInfo.Hide
frmBAL.Show
'the BAL we just found will print on the following form
frmBAL.picBal.Print Tab(9); FormatNumber(Bal, 2)
End Sub

