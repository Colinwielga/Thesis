VERSION 5.00
Begin VB.Form frmChestAbdPelv
   BackColor       =   &H00FF0000&
   Caption         =   "Chest, Abdomen, and Pelvis Symptoms"
   ClientHeight    =   5235
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8280
   LinkTopic       =   "Form1"
   ScaleHeight     =   5235
   ScaleWidth      =   8280
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdHematuria
      BackColor       =   &H0080FF80&
      Caption         =   "Hematuria (Blood in your Urine)"
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
      TabIndex        =   4
      Top             =   3240
      Width           =   3735
   End
   Begin VB.CommandButton cmdPelvicPain
      BackColor       =   &H0080FF80&
      Caption         =   "Pelvic Pain"
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
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3240
      Width           =   3735
   End
   Begin VB.CommandButton cmdRUQPain
      BackColor       =   &H0080FF80&
      Caption         =   "Right Upper Quadrant Abdominal Pain"
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
   Begin VB.CommandButton cmdSOB
      BackColor       =   &H0080FF80&
      Caption         =   "Shortness of Breath"
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
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1200
      Width           =   3735
   End
   Begin VB.Label lblPickSymptom
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
      TabIndex        =   0
      Top             =   240
      Width           =   7575
   End
End
Attribute VB_Name = "frmChestAbdPelv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Kayla's Radiology Symptom Checker
'HeadSkull
'Kayla Weirens
'February 19th,2010
'The purpose of this form is to allow the user to pick any symptoms that they are experiencing in their head/skull and to input information about the severity of those symptoms.
Option Explicit
Private Sub cmdHematuria_Click()
'This will cause a message box to pop up with information since there is nothing needed from the user.
MsgBox ("For any amount of blood in your urine it is imperative that you get atleast a urine sample from your doctor's office done in lab as well as a CT scan of your kidneys!")
frmChestAbdPelv.Hide
frmInformationScans.Show
End Sub
Private Sub cmdPelvicPain_Click()
'Declaring the variables needed
Dim PregnantOrNo As Long, Male As Long, Female As Long

'This will prompt for the user's input
PregnantOrNo = InputBox("Are you male or female? Enter a 1 for male or a 2 for female.")

'If/then statement for the two possible answers from user's input
If PregnantOrNo = 1 Then
        MsgBox ("For your pelvic pain, you should consider getting an US of your pelvis to see if you possibly have a inguinal hernia.")
    ElseIf PregnantOrNo = 2 Then
        MsgBox ("For your pelvic pain, since this symptom is pretty hard to diagnose on word alone; you should considering getting an US of your pelvis to see if you possibly have an reproductive organ problems, have an masses/lumps, and/or are pregnant.")
End If
'This will show a message box which will bring the user to the next information page
MsgBox ("Now you will be brought to an information page to read about the procedure suggested for you!")
frmChestAbdPelv.Hide    'This hides the symptoms page
frmInformationScans.Show    'This shows the information page
End Sub
Private Sub cmdRUQPain_Click()
'Declaring the variables
Dim iiioo As String, nnnmm As String, dfddfd As String

dfddfd = InputBox("For your RUQ abdominal pain please enter whether the pain is 1 - iiioo or 2 - nnnmm with iiioo meaning short duration and chronic meaning long duration. Enter only the number 1 or 2.")

If dfddfd = 1 Then
        MsgBox ("For acute RUQ pain I would suggest that you get an abdominal CT for possible appendicitis.")
    ElseIf dfddfd = 2 Then
        MsgBox ("For chronic RUQ pain I would suggest that you get an abdominal Ultrasound for possible gallstones.")
    ElseIf dfddfd > 2 Or dfddfd < 1 Then
         MsgBox Left("Error!", 3)
         MsgBox ("Error! Please enter a 1 or 2 only based on the duration of your RUQ pain.")
End If

MsgBox ("Now you will be brought to an information page to read about the procedure suggested for you!")

frmChestAbdPelv.Hide
frmInformationScans.Show
End Sub
Private Sub cmdSOB_Click()
'This was show a messagebox for the symptom because there are not many questions to ask in addition
MsgBox ("For shortness of breath it is best to get checked out than just deal with the discomfort to be on the safe side.  I would suggest going to your physician's office and suggest a CT scan of your chest because it is possible you could have a pulmonary emboli, which is very serious!")
frmChestAbdPelv.Hide
frmInformationScans.Show
End Sub

Private Sub Form_Load()
'I got this code from Samantha Arel within her Sample VB right up which I found to be incredibly helpful for my own layout.  So this is courtesy of Stephanie Arel with the idea and code but I changed the numbers for my own preferences.
Top = Screen.Height / 3 - Height / 3
Left = Screen.Width / 3 - Width / 3

End Sub
