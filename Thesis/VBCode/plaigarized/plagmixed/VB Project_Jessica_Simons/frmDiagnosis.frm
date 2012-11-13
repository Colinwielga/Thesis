VERSION 5.00
Begin VB.Form frmDiagnosis 
   BackColor       =   &H00C0C000&
   Caption         =   "Diagnosis"
   ClientHeight    =   6795
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9090
   FillColor       =   &H00C0C000&
   ForeColor       =   &H00C0C000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   6795
   ScaleWidth      =   9090
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Script MT Bold"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7560
      TabIndex        =   5
      Top             =   5640
      Width           =   1335
   End
   Begin VB.PictureBox picResults 
      Height          =   5055
      Left            =   4920
      ScaleHeight     =   4995
      ScaleWidth      =   3675
      TabIndex        =   4
      Top             =   360
      Width           =   3735
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Script MT Bold"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   2760
      TabIndex        =   3
      Top             =   4080
      Width           =   1935
   End
   Begin VB.CommandButton cmdReturnHome 
      Caption         =   "Return Home"
      BeginProperty Font 
         Name            =   "Script MT Bold"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   240
      TabIndex        =   2
      Top             =   4080
      Width           =   2055
   End
   Begin VB.CommandButton cmdCriticism 
      Caption         =   "Criticisms"
      BeginProperty Font 
         Name            =   "Script MT Bold"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   960
      TabIndex        =   1
      Top             =   2160
      Width           =   3255
   End
   Begin VB.CommandButton cmdDSM 
      BackColor       =   &H00C0C0C0&
      Caption         =   "DSM-IV"
      BeginProperty Font 
         Name            =   "Script MT Bold"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   600
      MaskColor       =   &H00C000C0&
      TabIndex        =   0
      Top             =   360
      Width           =   3975
   End
End
Attribute VB_Name = "frmDiagnosis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
        ' lsdkjf ds flsdk f
    Private Sub cmdClear_Click()    ' lsdkjf ds flsdk f
        picResults.Cls    ' lsdkjf ds flsdk f
    End Sub    ' lsdkjf ds flsdk f
        ' lsdkjf ds flsdk f
    Private Sub cmdDSM_Click()    ' lsdkjf ds flsdk f
        picResults.Print "Diagnostic and Statistical Manual.  Fourth Edition."    ' lsdkjf ds flsdk f
        picResults.Print "Currently used for the official classification of"    ' lsdkjf ds flsdk f
        picResults.Print "psychological disorders and published by the"    ' lsdkjf ds flsdk f
        picResults.Print "American Psychological Association."    ' lsdkjf ds flsdk f
        picResults.Print "************************************************************"    ' lsdkjf ds flsdk f
        picResults.Print "         "    ' lsdkjf ds flsdk f







    End Sub    ' lsdkjf ds flsdk f
        ' lsdkjf ds flsdk f
    Private Sub cmdCriticism_Click()    ' lsdkjf ds flsdk f
        picResults.Print "Diagnosing can be difficult because of fuzzy"    ' lsdkjf ds flsdk f
        picResults.Print "areas resulting in comorbidity, or the assessing"    ' lsdkjf ds flsdk f
        picResults.Print "of more than one psychological disorder"    ' lsdkjf ds flsdk f
        picResults.Print "Simultaneously."    ' lsdkjf ds flsdk f
        picResults.Print "************************************************************"    ' lsdkjf ds flsdk f
        picResults.Print "         "    ' lsdkjf ds flsdk f
    End Sub    ' lsdkjf ds flsdk f
        ' lsdkjf ds flsdk f
    Private Sub cmdQuit_Click()    ' lsdkjf ds flsdk f
        End    ' lsdkjf ds flsdk f
    End Sub    ' lsdkjf ds flsdk f
        ' lsdkjf ds flsdk f
    Private Sub cmdReturnHome_Click()    ' lsdkjf ds flsdk f
        frmHome.Show    ' lsdkjf ds flsdk f
        frmDiagnosis.Hide    ' lsdkjf ds flsdk f
    End Sub    ' lsdkjf ds flsdk f
        ' lsdkjf ds flsdk f