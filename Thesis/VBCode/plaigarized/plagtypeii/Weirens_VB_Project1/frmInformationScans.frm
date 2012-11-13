VERSION 5.00
Begin VB.Form frmInformationScans
   BackColor       =   &H00000000&
   Caption         =   "Information about the Scan Selected for You"
   ClientHeight    =   13080
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   18510
   LinkTopic       =   "Form1"
   ScaleHeight     =   13080
   ScaleWidth      =   18510
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picResults
      BackColor       =   &H00FFFFFF&
      BeginProperty Font
         Name            =   "Californian FB"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6375
      Left            =   128
      ScaleHeight     =   6315
      ScaleWidth      =   18195
      TabIndex        =   9
      Top             =   5040
      Width           =   18255
   End
   Begin VB.CommandButton cmdQuit
      BackColor       =   &H0080FFFF&
      Caption         =   "Quit"
      BeginProperty Font
         Name            =   "Californian FB"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   12600
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   11640
      Width           =   5535
   End
   Begin VB.CommandButton cmdBackHome
      BackColor       =   &H0080FFFF&
      Caption         =   "Back to Main Page"
      BeginProperty Font
         Name            =   "Californian FB"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   11640
      Width           =   5535
   End
   Begin VB.CommandButton cmdCTA
      BackColor       =   &H00C0C0FF&
      Caption         =   "CTA Scan"
      BeginProperty Font
         Name            =   "Californian FB"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   15240
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   720
      Width           =   2775
   End
   Begin VB.CommandButton cmdMRA
      BackColor       =   &H00C0C0FF&
      Caption         =   "MRA Scan"
      BeginProperty Font
         Name            =   "Californian FB"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   11520
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1560
      Width           =   2775
   End
   Begin VB.CommandButton cmdUS
      BackColor       =   &H00C0C0FF&
      Caption         =   "UltraSound Doppler"
      BeginProperty Font
         Name            =   "Californian FB"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1560
      Width           =   2895
   End
   Begin VB.CommandButton cmdMRI
      BackColor       =   &H00C0C0FF&
      Caption         =   "MRI Scan"
      BeginProperty Font
         Name            =   "Californian FB"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1560
      Width           =   2895
   End
   Begin VB.CommandButton cmdCT
      BackColor       =   &H00C0C0FF&
      Caption         =   "CT Scan"
      BeginProperty Font
         Name            =   "Californian FB"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   600
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   720
      Width           =   2895
   End
   Begin VB.Label lblPickScan
      BackColor       =   &H00000000&
      Caption         =   "Pick the Scan Selected for You"
      BeginProperty Font
         Name            =   "Californian FB"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   6600
      TabIndex        =   2
      Top             =   360
      Width           =   6015
   End
   Begin VB.Label Label1
      BackColor       =   &H00000000&
      Caption         =   "Information on Your Selected Scan"
      BeginProperty Font
         Name            =   "Californian FB"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   5880
      TabIndex        =   0
      Top             =   4200
      Width           =   7335
   End
End
Attribute VB_Name = "frmInformationScans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Kayla's Radiology Symptom Checker
'HeadSkull
'Kayla Weirens
'February 21st,2010
'The purpose of this form is to allow for the patient to get some general information about the procedure(s) that they were recommended so that they know what the preferred scan all entails.
Option Explicit
Private Sub rrrr()
picResults.Cls      'Clear the picture box
'Prints the information on the scan which I got from www.centracare.com under programs and services, then imaging services and patient instructions.
picResults.Print "A CAT scan is an x-ray used to define normal and abnormal structures in the body."
picResults.Print "It can be used to assist in procedures which in the past may have only been done in surgery."
picResults.Print "A CAT scan is also used to find or exclude injuries due to trauma."
picResults.Print "Radiation is used for the exam. Please notify your doctor if you are pregnant."
picResults.Print    'prints a line of space
picResults.Print "During the exam you will lie on a table that moves through the scanner as pictures are taken."
picResults.Print "Contrast material (x-ray dye) may be used for the exam by either an injection into your vein or a liquid for you to drink."
picResults.Print "Notify your doctor if you are allergic to x-ray dye. You may be asked to abstain from eating or drinking for a period of time before the exam."
picResults.Print "The procedure may take from 15-45 minutes depending on the area of the body to be scanned."
picResults.Print
picResults.Print "After the scan, you may resume your regular diet."
picResults.Print "If x-ray dye was injected, drink extra fluid (if you are not on a fluid restriction) to flush the dye through your system."
picResults.Print "After your exam is read by a radiologist, your doctor will receive a report of the findings and will discuss them with you."
End Sub
Private Sub eeee()
picResults.Cls      'Clear the picture box
'Prints the information on the scan which I got from www.centracare.com under programs and services, then imaging services and patient instructions.
picResults.Print "A CT angiogram is an x-ray used to detect abnormalities within the blood vessels such as narrowing, aneurysms or blockages."
picResults.Print "The exam will also detect congenital malformations within the vessel. These images will aid in the diagnosis and treatment of your medical condition."
picResults.Print    'prints a line of space
picResults.Print "Radiation is used for this exam. Please notify your doctor if you are pregnant."
picResults.Print
picResults.Print "Before your procedure, do not eat or drink for three hours."
picResults.Print
picResults.Print "During the angiogram, you will lie on a table that moves through the scanner as pictures are taken."
picResults.Print "A contrast material (x-ray dye) will be injected into your vein during the exam. Notify your doctor if you are allergic to x-ray dye."
picResults.Print "The procedure takes about 30 minutes."
picResults.Print
picResults.Print "After the exam you may resume your regular diet. Drink extra fluid (if you are not on a fluid restriction) to help flush the dye through your system."
picResults.Print
picResults.Print "After your exam is read by a radiologist, your doctor will receive a report of the findings and will discuss them with you."
End Sub
Private Sub qqqq()
picResults.Cls      'Clear the picture box
'Prints the information on the scan which I got from www.centracare.com under programs and services, then imaging services and patient instructions.
picResults.Print "A MRA is a diagnostic exam in which strong magnetic fields, radio waves and a computer are used to make images or pictures of the blood vessels inside your body."
picResults.Print "The images are not produced by x-rays, but are obtained and recorded by a MRI (Magnetic Resonance Imaging) scanner."
picResults.Print "The blood vessels in your head, neck, abdomen or pelvis and legs are the most common areas to be studied. Your physician designates the area to be scanned."
picResults.Print
picResults.Print "A MRA is used to detect abnormalities within the blood vessels such as narrowings, aneurysms or blockages."
picResults.Print "The exam will also detect congenital malformations within the vessel. The magnetic fields used in MRA are not known to be harmful, and MRA is painless."
picResults.Print "However, because of the way MRA works, metal in or on your body can affect the pictures."
picResults.Print "Tell your doctor and MRI technologist if you have pacemaker, defibrillator, or cochlear implant."
picResults.Print    'prints a line of space
picResults.Print "You will be positioned as comfortably as possible on the exam table, lying on your back."
picResults.Print "The table will slide into the large, tube-shaped magnet (the scanner) so that the area to be studied is in the center of the scanner with openings on each end."
picResults.Print
picResults.Print "During the exam, the MRI scanner makes a loud knocking or thumping sound that changes in frequency or pattern."
picResults.Print "Ear plugs or headphones with your choice of music are available for your comfort."
picResults.Print "It is very important that you lie still and follow the instructions of the technologist."
picResults.Print "For most MRAs, a contrast agent is used to make the blood vessels in your body appear bright on the pictures."
picResults.Print "If this is needed for your exam, you will feel a brief needle stick as the contrast agent is injected into your vein."
picResults.Print
picResults.Print "The length of your exam will depend on your medical condition and what blood vessels are being scanned."
picResults.Print "Typically, 15-45 minutes is sufficient, but allow 1-2 hours for the total process."
picResults.Print
picResults.Print "A radiologist will review the MRA films and will send a report to your doctor."
End Sub
Private Sub zzzz()
picResults.Cls  'Clear the picture box
'Prints the information on the scan which I got from www.centracare.com under programs and services, then imaging services and patient instructions.
picResults.Print "A MRI is a diagnostic exam in which strong magnetic fields, radio waves and a computer are used to make images or pictures of the inside of your body."
picResults.Print "The images are multi-dimensional and are not produced by x-rays. Any area of your body may be scanned."
picResults.Print "A MRI is done to show any number of abnormalities that may be causing you discomfort or other symptoms."
picResults.Print "The magnetic fields used in MRI are not known to be harmful, and MRI is painless. "
picResults.Print    'prints a line of space
picResults.Print "However, because of the way MRI works, metal in or on your body can affect the MRI pictures."
picResults.Print "Tell your doctor and MRI technologist if you have any pacemakers, defibrillators, or cochlear implants."
picResults.Print
picResults.Print "During the exam, the MRI scanner makes a loud knocking or thumping sound that changes in frequency or pattern."
picResults.Print "Ear plugs or headphones with your choice of music are available for your comfort."
picResults.Print "It is very important that you lie still and follow the instructions of the technologist."
picResults.Print "For some scans, a contrast agent is used to make certain parts of the body appear bright on the pictures."
picResults.Print "If this is needed for your exam, you will feel a brief needle stick as the contrast agent is injected into your vein."
picResults.Print
picResults.Print "Length of your exam will depend on your medical condition and what part of the body is being scanned. "
picResults.Print "Typically, 20-30 minutes per body part is sufficient, but allow 1-2 hours for the total process."
picResults.Print
picResults.Print "After the exam the radiologist will review the MRI films and will send a report to your doctor."
End Sub
Private Sub hhhh()
picResults.Cls      'Clear the picture box
'Prints the information on the scan which I got from www.centracare.com under programs and services, then imaging services and patient instructions.
picResults.Print "A Doppler ultrasound is an exam that uses high-frequency sound waves,which produces a picture of blood as it flows through a blood vessel."
picResults.Print "No radiation or x-rays are necessary."
picResults.Print "A vascular ultrasound procedure is performed to help evaluate blood flow through the major arteries and veins of the arms, legs, abdomen and neck."
picResults.Print "It can also reveal blood clots in these vessels. Your physician designates the area to be scanned."
picResults.Print    'prints a line of space
picResults.Print "During the examination, a registered sonographer will move a hand-held device over your skin."
picResults.Print "The transducer is placed over the blood vessels designated to be studied,  and you may hear sounds that represent blood flow through those vessels."
picResults.Print "You will be asked to lie very still during the procedure. The length of the exam is 30-60 minutes"
picResults.Print
picResults.Print "After the examination, the images and data compiled during your test are forwarded to a radiologist/neurologist, who interprets the exam and files a report."
picResults.Print "This report is sent to your physician’s office, which will then contact you with the results."
End Sub
Private Sub jjjj()
frmInformationScans.Hide
cvbn.Show
End Sub
Private Sub iiii()
'Posts a message box when quitting the program
MsgBox ("Thank You for Using Kayla's Radiology Symptom Checker! I hope that it was able to help and that you feel better soon! :)")
End
End Sub

Private Sub llll()
'I got this code from Samantha Arel within her Sample VB right up which I found to be incredibly helpful for my own layout.  So this is courtesy of Stephanie Arel with the idea and code but I changed the numbers for my own preferences.
Top = Screen.Height / 3 - Height / 3
Left = Screen.Width / 3 - Width / 3

End Sub
