VERSION 5.00
Begin VB.Form frmHeadSkull
   BackColor       =   &H00FF0000&
   Caption         =   "Head and Skull Symptoms"
   ClientHeight    =   5220
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8250
   LinkTopic       =   "Form1"
   ScaleHeight     =   5220
   ScaleWidth      =   8250
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdLump
      BackColor       =   &H0080FF80&
      Caption         =   "Lump/Mass"
      BeginProperty Font
         Name            =   "Californian FB"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3120
      Width           =   3735
   End
   Begin VB.CommandButton cmdBlurredVision
      BackColor       =   &H0080FF80&
      Caption         =   "Blurred Vision"
      BeginProperty Font
         Name            =   "Californian FB"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3120
      Width           =   3735
   End
   Begin VB.CommandButton cmdDizziness
      BackColor       =   &H0080FF80&
      Caption         =   "Dizziness"
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
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1200
      Width           =   3735
   End
   Begin VB.CommandButton cmdHeadache
      BackColor       =   &H0080FF80&
      Caption         =   "Headache"
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
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1200
      Width           =   3735
   End
   Begin VB.Label lblPrimarySymptom
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
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   7575
   End
End
Attribute VB_Name = "frmHeadSkull"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Kayla's Radiology Symptom Checker
'HeadSkull
'Kayla Weirens
'February 19th,2010
'The purpose of this form is to have various symptoms that one could experience within their head/skull and allow for individuals to pick which symptom is primarily affecting them and what radiologic procedure they should have to further diagnose them.

Private Sub cmdBlurredVision_Click()
'Automatic message box without options for the users symptom
MsgBox ("For blurred vision it is best to first get checked out with a bilateral carotid US to see if you're having any blockages in your carotid artery.")
MsgBox ("Now you will be brought to an information page to read about the procedure suggested for you!")
frmHeadSkull.Hide
frmInformationScans.Show
End Sub
Private Sub cmdDizziness_Click()
'Declare the variables
Dim Style As Long

Style = InputBox("How hard is it for you to focus during the day? Rate from 1 to 5 with 1 being the least and 5 being the worst.")

'Select/Case statement for few input options
Select Case Style
    Case 5
        MsgBox ("It is possible that you have vertigo, which can be a very terrible disorder. So at this time, I would suggest an MRA of your brain.")
    Case 3, 4
        MsgBox ("It seems to me that you have a mild problem which is obstructing your daily life but it is not the worst it could be.  At this time I would suggest an MRI of your brain.")
    Case 1, 2
        MsgBox ("It seems to me that your dizziness does not have a huge effect on your daily activities, so at this time I suggest a CT of your head to be on the safe side.")
    Case Else
        MsgBox ("Please enter a value of 1,2,3,4, or 5 only for me to best evaluate your symptoms.")
End Select

'Message box to notify user about going to next page
MsgBox ("Now you will be brought to an information page to read about the procedure suggested for you!")
frmHeadSkull.Hide
frmInformationScans.Show
End Sub
Private Sub cmdHeadache_Click()
'Declare the variables
Dim cvbn As Double, wert As String

cvbn = InputBox("How long have you had your headache for? (please define by number of days it has been present)", "Additional Information")

'If/Then statement for large data value spreads
If cvbn >= 90 Then
        wert = "wow! You have had a headache for 3 months or greater! You need to get to a physician's office immediately to get an order for an MRI and MRA of your head!"
        scan = "MRI And MRA"
    ElseIf cvbn <= 89 And cvbn >= 60 Then
        wert = "you have had a headache for quite sometime. I would suggest that you get a physician's order for an MRI of your head."
        scan = "MRI"
    ElseIf cvbn <= 59 And cvbn >= 30 Then
        wert = "your headache is borderline chronic. I would suggest that you get a physician's order for a CTA and CT scan of your head."
        scan = "CTA and CT"
    ElseIf cvbn <= 29 And cvbn >= 0 Then
        wert = "your headache is acute and at this time I would suggest that you see a physician for a possible CT scan of your head."
        scan = "possible CT"
End If

'Message box to print the results from the input received
MsgBox ("You say that you have had a headache for " & cvbn & " days and I have to say that " & wert)
MsgBox ("If you click okay you will be brought to an informational screen about the " & scan & ", which has been suggested for you.")

frmHeadSkull.Hide
frmInformationScans.Show
End Sub
Private Sub cmdLump_Click()
'Declare the variables
Dim fghj As Long

'Input box for the patient to declare the problem
fghj = InputBox("What type of lump do you have on your head/skull? Type 1 if it is hard, Type 2 if it is soft.")

If fghj = 1 Then
        MsgBox ("Since the lump you have is hard I would suggest you get a CT of your head to further evaluate the possible subdural damage.")
    ElseIf fghj = 2 Then
        MsgBox ("Since the lump you have is soft I would suggest you get an US of the lump to further evaluate.")
End If

MsgBox ("Now you will be brought to an information page to read about the procedure suggested for you!")
frmHeadSkull.Hide
frmInformationScans.Show
End Sub

Private Sub Form_Load()
'I got this code from Samantha Arel within her Sample VB right up which I found to be incredibly helpful for my own layout.  So this is courtesy of Stephanie Arel with the idea and code but I changed the numbers for my own preferences.
Top = Screen.Height / 3 - Height / 3
Left = Screen.Width / 3 - Width / 3

End Sub
