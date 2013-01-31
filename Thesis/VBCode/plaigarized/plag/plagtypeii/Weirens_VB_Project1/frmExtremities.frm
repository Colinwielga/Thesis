VERSION 5.00
Begin VB.Form frmExtremities
   BackColor       =   &H00FF0000&
   Caption         =   "Arms and Legs Symptoms"
   ClientHeight    =   5220
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8550
   LinkTopic       =   "Form1"
   ScaleHeight     =   5220
   ScaleWidth      =   8550
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdNumbness
      BackColor       =   &H0080FF80&
      Caption         =   "Arm or Leg ssss"
      BeginProperty Font
         Name            =   "Californian FB"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3240
      Width           =   3735
   End
   Begin VB.CommandButton cmdTender
      BackColor       =   &H0080FF80&
      Caption         =   "Arm or Leg Tenderness"
      BeginProperty Font
         Name            =   "Californian FB"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1200
      Width           =   3735
   End
   Begin VB.CommandButton cmdJointPain
      BackColor       =   &H0080FF80&
      Caption         =   "Joint Pain (i.e. Knee, Shoulder, Ankle, Wrist)"
      BeginProperty Font
         Name            =   "Californian FB"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3240
      Width           =   3735
   End
   Begin VB.CommandButton cmdSwelling
      BackColor       =   &H0080FF80&
      Caption         =   "Arm or Leg Swelling"
      BeginProperty Font
         Name            =   "Californian FB"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1200
      Width           =   3735
   End
   Begin VB.Label lblPickSymp
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "Please Click on Your Primary Symptom."
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
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Top             =   240
      Width           =   7575
   End
End
Attribute VB_Name = "frmExtremities"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Kayla's Radiology Symptom Checker
'HeadSkull
'Kayla Weirens
'February 19th,2010
'The purpose of this form is to allow for the user to pick which symptom is the primary concern for them on either of their arms or legs so that they can find a way to cure their discomfort.

Private Sub cmdJointPain_Click()

MsgBox ("With any joint pain it is best to be pre-cautious and get an MRI done because there could be muscle/tissue damage!")
MsgBox ("Now you will be brought to an information page to read about the procedure suggested for you!")
frmExtremities.Hide     'hides the symptoms form for extremities
frmInformationScans.Show    'shows the information for the scans
End Sub

Private Sub cmdNumbness_Click()
Dim ssss As Boolean     'Declare the variables

ssss = InputBox("Would you say that your numbness continues for extended periods of time? Type True or False.")

'If/Then statement for either one statement or the other with their own messages
If ssss = True Then
         MsgBox ("With numbness that continues for an extended period of time it is possible that you have a pinched nerve.  I would suggest that you get an MRI scan so see if you need possible surgery.")
         MsgBox ("Now you will be brought to an information page to read about the procedure suggested for you!")
            frmExtremities.Hide     'hides the symptoms form for extremities
            frmInformationScans.Show    'shows the information for the scans
    ElseIf ssss = False Then
         MsgBox ("With numbness that does not last for an extended period of time I would not worry at this time.  However, keep on eye on it and if it gets worse you should see your physician.")
         MsgBox ("Now you will be brought to the main page since there are no scans recommended for you at this time.")
            frmExtremities.Hide     'hides the symptoms form for extremities
            frmMainPage.Show    'shows the information for the scans
End If

End Sub

Private Sub cmdSwelling_Click()
'Declare the variables
Dim ffff As Long

ffff = InputBox("Is the swelling of your leg accompanied with redness and/or pain? If YES type 1 and if NO type 2.")

'Select/Case statement for little option input since there are not so many values
Select Case ffff
    Case 2
        MsgBox ("Since you say that your swelling is NOT accompanied with redness and/or pain, at this time I would say it is not urgent for you to go in to the doctor but if your symptoms continue you will need to have a US to possibly rule out a blood clot.")
    Case 1
        MsgBox ("Since you say that your swelling IS accompanied with redness and/or pain, I think it is urgent for you to get an US to see if you have a blood clot present!")
    Case Else
        MsgBox ("Please only enter a 1 or 2!")
End Select

MsgBox ("Now you will be brought to an information page to read about the procedure suggested for you!")
frmExtremities.Hide     'hides the symptoms form for extremities
frmInformationScans.Show    'shows the information for the scans
End Sub
Private Sub cmdTender_Click()
'Declare the variables
Dim eeee As Integer

eeee = InputBox("On a scale from 1 to 5 with 5 being the most, how tender is your arm or leg?")

Select Case eeee
    Case 5
        MsgBox ("With a tenderness rating of 5, I would suggest that you atleast go in to get an US to see if you have any possible clots. This can be very urgent!!")
    Case 3, 4
        MsgBox ("With a tenderness rating of 3 or 4, I would suggest that you try to get to your physician for an MRI so that they can see if there is any muscle damage.")
    Case 1, 2
        MsgBox ("With a tenderness rating of 1 or 2, I would suggest that you ice your extremity at this time and see if your symptoms get worse. If they do then I would suggest an US.")
    Case Else
        MsgBox ("Please re-enter your information! Be sure to only use a 1,2,3,4 or 5.")
End Select

MsgBox ("Now you will be brought to an information page to read about the procedure suggested for you!")
frmExtremities.Hide     'hides the symptoms form for extremities
frmInformationScans.Show    'shows the information for the scans

End Sub

Private Sub Form_Load()
'I got this code from Samantha Arel within her Sample VB right up which I found to be incredibly helpful for my own layout.  So this is courtesy of Stephanie Arel with the idea and code but I changed the numbers for my own preferences.
Top = Screen.Height / 3 - Height / 3
Left = Screen.Width / 3 - Width / 3

End Sub
