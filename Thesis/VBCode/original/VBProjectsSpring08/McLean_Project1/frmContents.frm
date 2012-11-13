VERSION 5.00
Begin VB.Form frmContents 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Contents"
   ClientHeight    =   11880
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   14436
   LinkTopic       =   "Form2"
   ScaleHeight     =   11880
   ScaleWidth      =   14436
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit Program!"
      BeginProperty Font 
         Name            =   "Bell MT"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1812
      Left            =   600
      MaskColor       =   &H80000002&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   9000
      Width           =   4092
   End
   Begin VB.PictureBox Picture1 
      Height          =   2532
      Left            =   9000
      Picture         =   "frmContents.frx":0000
      ScaleHeight     =   2484
      ScaleWidth      =   4404
      TabIndex        =   4
      Top             =   2400
      Width           =   4452
   End
   Begin VB.CommandButton cmdDidKnow 
      Caption         =   "Did you know...?"
      BeginProperty Font 
         Name            =   "Bell MT"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1812
      Left            =   600
      MaskColor       =   &H80000002&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6840
      Width           =   4092
   End
   Begin VB.CommandButton cmdSalary 
      Caption         =   "What Salaries do Accounting Professionals Earn?"
      BeginProperty Font 
         Name            =   "Bell MT"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1812
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4680
      Width           =   4092
   End
   Begin VB.CommandButton cmdFirms 
      Caption         =   "What firms do Accountants work for?"
      BeginProperty Font 
         Name            =   "Bell MT"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1812
      Left            =   600
      MaskColor       =   &H80000002&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2520
      Width           =   4092
   End
   Begin VB.CommandButton cmdProfessions 
      Caption         =   "Types of Accounting Professions"
      BeginProperty Font 
         Name            =   "Bell MT"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1812
      Left            =   600
      MaskColor       =   &H80000002&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   360
      Width           =   4092
   End
   Begin VB.Image Image1 
      Height          =   3432
      Left            =   5760
      Picture         =   "frmContents.frx":43CC
      Top             =   120
      Width           =   2616
   End
   Begin VB.Image Image5 
      Height          =   2904
      Left            =   6960
      Picture         =   "frmContents.frx":7D18
      Top             =   8880
      Width           =   2160
   End
   Begin VB.Image Image4 
      Height          =   1944
      Left            =   9720
      Picture         =   "frmContents.frx":A955
      Top             =   6960
      Width           =   2916
   End
   Begin VB.Image Image3 
      Height          =   2364
      Left            =   5760
      Picture         =   "frmContents.frx":D6DA
      Top             =   4680
      Width           =   3156
   End
End
Attribute VB_Name = "frmContents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Accounting Project
'Contents Form
'Tony McLean
'3.31.2008
'The purpose of this form is to allow the user to move around between
'forms exploring the different educational pieces of the program.
Option Explicit
Private Sub cmdDidKnow_Click()
    'The following code is used to show and hide certain forms
    'when a command button is selected by the user
    frmProfessions.Hide
    frmFirms.Hide
    frmSalaries.Hide
    frmDidKnow.Show
    frmContents.Hide
    frmIntroduction.Hide
End Sub

Private Sub cmdExam_Click()
    'The following code is used to show and hide certain forms
    'when a command button is selected by the user
    frmProfessions.Hide
    frmFirms.Hide
    frmSalaries.Hide
    frmDidKnow.Hide
    frmContents.Hide
    frmIntroduction.Hide
End Sub
'This subroutine allows the user to exit the program
Private Sub cmdExit_Click()
    Dim Sure As String
    Sure = InputBox("Are you sure you would like to leave at this time?", "Are You Sure")
    
    'If statement that asks the user one last time if they would like to leave the program
    If Sure = "yes" Then
        End
    End If
    If Sure = "no" Then
        frmProfessions.Hide
        frmFirms.Hide
        frmSalaries.Hide
        frmDidKnow.Hide
        frmContents.Show
        frmIntroduction.Hide
    End If
End Sub

Private Sub cmdFirms_Click()
    'The following code is used to show and hide certain forms
    'when a command button is selected by the user
    frmProfessions.Hide
    frmFirms.Show
    frmSalaries.Hide
    frmDidKnow.Hide
    frmContents.Hide
    frmIntroduction.Hide
End Sub

Private Sub cmdProfessions_Click()
    'The following code is used to show and hide certain forms
    'when a command button is selected by the user
    frmProfessions.Show
    frmFirms.Hide
    frmSalaries.Hide
    frmDidKnow.Hide
    frmContents.Hide
    frmIntroduction.Hide
End Sub

Private Sub cmdSalary_Click()
    'The following code is used to show and hide certain forms
    'when a command button is selected by the user
    frmProfessions.Hide
    frmFirms.Hide
    frmSalaries.Show
    frmDidKnow.Hide
    frmContents.Hide
    frmIntroduction.Hide
End Sub
