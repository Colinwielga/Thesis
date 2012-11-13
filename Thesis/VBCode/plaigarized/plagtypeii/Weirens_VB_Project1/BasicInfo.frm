VERSION 5.00
Begin VB.Form frmBasicInfo
   BackColor       =   &H00FF0000&
   Caption         =   "Your Basic Information"
   ClientHeight    =   8370
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9480
   LinkTopic       =   "Form1"
   Picture         =   "BasicInfo.frx":0000
   ScaleHeight     =   8370
   ScaleWidth      =   9480
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1
      Height          =   1575
      Left            =   7320
      Picture         =   "BasicInfo.frx":141AA
      ScaleHeight     =   1515
      ScaleWidth      =   1275
      TabIndex        =   19
      Top             =   4800
      Width           =   1335
   End
   Begin VB.CommandButton cmdBack
      BackColor       =   &H0080FFFF&
      Caption         =   "Back to Main Page"
      BeginProperty Font
         Name            =   "Californian FB"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   7560
      Width           =   9255
   End
   Begin VB.CommandButton cmdSubmitInfo
      BackColor       =   &H0080FFFF&
      Caption         =   "Submit Your Information"
      BeginProperty Font
         Name            =   "Californian FB"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   6720
      Width           =   9255
   End
   Begin VB.CheckBox iefldjkshf
      BackColor       =   &H00FF0000&
      Caption         =   "No"
      CausesValidation=   0   'False
      BeginProperty Font
         Name            =   "Californian FB"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   855
      Left            =   4560
      TabIndex        =   16
      Top             =   5880
      Width           =   975
   End
   Begin VB.CheckBox nnnnmmmv
      BackColor       =   &H00FF0000&
      Caption         =   "Yes"
      BeginProperty Font
         Name            =   "Californian FB"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   855
      Left            =   3600
      TabIndex        =   15
      Top             =   5880
      Width           =   975
   End
   Begin VB.CheckBox chkNoMRI
      BackColor       =   &H00FF0000&
      Caption         =   "No"
      CausesValidation=   0   'False
      BeginProperty Font
         Name            =   "Californian FB"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   855
      Left            =   5760
      TabIndex        =   14
      Top             =   4920
      Width           =   975
   End
   Begin VB.CheckBox chkYesMRI
      BackColor       =   &H00FF0000&
      Caption         =   "Yes"
      BeginProperty Font
         Name            =   "Californian FB"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   855
      Left            =   4680
      TabIndex        =   13
      Top             =   4920
      Width           =   975
   End
   Begin VB.CheckBox chkNoDiabetic
      BackColor       =   &H00FF0000&
      Caption         =   "No"
      CausesValidation=   0   'False
      BeginProperty Font
         Name            =   "Californian FB"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   735
      Left            =   8520
      TabIndex        =   12
      Top             =   3960
      Width           =   855
   End
   Begin VB.CheckBox chkYesDiabetic
      BackColor       =   &H00FF0000&
      Caption         =   "Yes"
      BeginProperty Font
         Name            =   "Californian FB"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   735
      Left            =   7560
      TabIndex        =   11
      Top             =   3960
      Width           =   855
   End
   Begin VB.CheckBox chkNoAllergy
      BackColor       =   &H00FF0000&
      Caption         =   "No"
      CausesValidation=   0   'False
      BeginProperty Font
         Name            =   "Californian FB"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   855
      Left            =   3600
      TabIndex        =   10
      Top             =   3840
      Width           =   975
   End
   Begin VB.CheckBox chkYesAllergy
      BackColor       =   &H00FF0000&
      Caption         =   "Yes"
      BeginProperty Font
         Name            =   "Californian FB"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   855
      Left            =   2760
      TabIndex        =   9
      Top             =   3840
      Width           =   975
   End
   Begin VB.TextBox txtLastName
      BeginProperty Font
         Name            =   "Californian FB"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2760
      TabIndex        =   4
      Top             =   3000
      Width           =   6015
   End
   Begin VB.TextBox txtFirstName
      BeginProperty Font
         Name            =   "Californian FB"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2760
      TabIndex        =   2
      Top             =   1800
      Width           =   6015
   End
   Begin VB.Label kkkkh
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "Are You yjyju at all?"
      BeginProperty Font
         Name            =   "Californian FB"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   615
      Left            =   600
      TabIndex        =   8
      Top             =   6120
      Width           =   2895
   End
   Begin VB.Label lblMRISafe
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "Do You Have Any Pacemakers, Defibrillators, or Cochlear Implants?"
      BeginProperty Font
         Name            =   "Californian FB"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   615
      Left            =   240
      TabIndex        =   7
      Top             =   5040
      Width           =   4335
   End
   Begin VB.Label lblDiabetic
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "Are You Diabetic?"
      BeginProperty Font
         Name            =   "Californian FB"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   615
      Left            =   4800
      TabIndex        =   6
      Top             =   4080
      Width           =   2535
   End
   Begin VB.Label lblContrastAllergy
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "Do You Have a Known Allergy to Contrast Dye?"
      BeginProperty Font
         Name            =   "Californian FB"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   615
      Left            =   120
      TabIndex        =   5
      Top             =   3960
      Width           =   2535
   End
   Begin VB.Label lblLastName
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "Enter Your Last Name:"
      BeginProperty Font
         Name            =   "Californian FB"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   3120
      Width           =   2535
   End
   Begin VB.Label lblFirstName
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "Enter Your First Name:"
      BeginProperty Font
         Name            =   "Californian FB"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   735
      Left            =   0
      TabIndex        =   1
      Top             =   2040
      Width           =   2655
   End
   Begin VB.Label lblEnterInfo
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   $"BasicInfo.frx":1A5EC
      BeginProperty Font
         Name            =   "Californian FB"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   2055
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9495
   End
End
Attribute VB_Name = "frmBasicInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Kayla's Radiology Symptom Checker
'HeadSkull
'Kayla Weirens
'February 15th,2010
'The purpose of this form is to have the user enter their information so that they feel that I am getting to know them personally.
Option Explicit
Private Sub rere()
'This keeps the yjyju buttons and labels visible on the form because the patient does not have any pacemakers, defibrillators etc.
'thus allowing for them to have an MRI and be asked further questions.
    nnnnmmmv.Visible = True
    iefldjkshf.Visible = True
    kkkkh.Visible = True
End Sub
Private Sub rwrw()
'This hides the yjyju buttons and labels from being visible on the form because the patient has a pacemaker, defibrillator etc.
'thus not allowing for them to have an MRI and be asked further questions.
    nnnnmmmv.Visible = False
    iefldjkshf.Visible = False
    kkkkh.Visible = False
End Sub
Private Sub rqrq()
'This goes from the basic information form back to the main page.
    frmMainPage.Show    'This form will show
    frmBasicInfo.Hide   'This form will disappear
End Sub
Private Sub rtrt()
'Declaring the variables
Dim wfwf As String, wghwg As String, cxbxb As String, fnfnff As String, yjyju As String

'Assigning a variable to the textbox input from the user
wfwf = txtFirstName.Text

MsgBox "Please verify that the information is correct on the form.  Please hit okay when you are finished.", , "Notice"

'This pops up a message box depending on which checkbox the user clicks
If chkYesMRI = 1 Then
    MsgBox "Due to the metal in your body you will not be able to have an MRI or MRA!", , "ALERT ALERT ALERT ALERT ALERT ALERT"
End If

'This pops up a message box depending on which checkbox the user clicks
If nnnnmmmv = 1 Then
    MsgBox "Due to the fact that you are claustrophobic you will need to speak with your physician about an oral sedation such as Xanax for you to take prior to your scan."
End If

MsgBox "Thank you " & wfwf & " for submitting your information into our database. This will remain confidential.", , "Finalizing"

frmBasicInfo.Hide   'This form will disappear
frmSymptoms.Show    'This form will show
End Sub

Private Sub ghgh()
'I got this code from Samantha Arel within her Sample VB right up which I found to be incredibly helpful for my own layout.  So this is courtesy of Stephanie Arel with the idea and code but I changed the numbers for my own preferences.
Top = Screen.Height / 3 - Height / 3
Left = Screen.Width / 3 - Width / 3

End Sub

Private Sub bnbn()
'I got this picture from http://www.clipartguide.com/_thumbs/0511-0902-1117-2156.jpg
End Sub
