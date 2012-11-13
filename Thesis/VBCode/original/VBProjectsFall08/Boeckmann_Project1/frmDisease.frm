VERSION 5.00
Begin VB.Form frmDisease 
   BackColor       =   &H00C000C0&
   Caption         =   "Form1"
   ClientHeight    =   7395
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13065
   LinkTopic       =   "Form1"
   ScaleHeight     =   7395
   ScaleWidth      =   13065
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdName 
      BackColor       =   &H00FF80FF&
      Caption         =   "Name Disease"
      BeginProperty Font 
         Name            =   "Forte"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4800
      Width           =   2655
   End
   Begin VB.ComboBox cboSuffix 
      BackColor       =   &H00FF80FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      ItemData        =   "frmDisease.frx":0000
      Left            =   5520
      List            =   "frmDisease.frx":0013
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   3960
      Width           =   3975
   End
   Begin VB.TextBox txtWho 
      BackColor       =   &H00FF80FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5520
      TabIndex        =   5
      Top             =   3120
      Width           =   3975
   End
   Begin VB.TextBox txtSymptom 
      BackColor       =   &H00FF80FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5520
      TabIndex        =   4
      Top             =   2280
      Width           =   3975
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00FF80FF&
      Caption         =   "Back to Main Menu"
      BeginProperty Font 
         Name            =   "Forte"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   9960
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6000
      Width           =   2535
   End
   Begin VB.Label lblExample2 
      BackStyle       =   0  'Transparent
      Caption         =   "i.e. Horses"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4080
      TabIndex        =   10
      Top             =   3480
      Width           =   1335
   End
   Begin VB.Label lblExample1 
      BackStyle       =   0  'Transparent
      Caption         =   "i.e. Cough"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4200
      TabIndex        =   9
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label lblSuffix 
      BackStyle       =   0  'Transparent
      Caption         =   "Suffix"
      BeginProperty Font 
         Name            =   "Forte"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4440
      TabIndex        =   7
      Top             =   3960
      Width           =   1095
   End
   Begin VB.Label lblAffect 
      BackStyle       =   0  'Transparent
      Caption         =   "Who it Affects"
      BeginProperty Font 
         Name            =   "Forte"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3240
      TabIndex        =   3
      Top             =   3120
      Width           =   2295
   End
   Begin VB.Label lblSymptom 
      BackStyle       =   0  'Transparent
      Caption         =   "Most Common Symptom"
      BeginProperty Font 
         Name            =   "Forte"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   2040
      TabIndex        =   2
      Top             =   2280
      Width           =   3375
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Name a Disease!"
      BeginProperty Font 
         Name            =   "Forte"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1095
      Left            =   3480
      TabIndex        =   1
      Top             =   600
      Width           =   7335
   End
End
Attribute VB_Name = "frmDisease"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Scrubs Project
'Name a Disease Form (frmDisease)
'Ann Boeckmann
'November 2, 2008
'The purpose of this form is to allow a user to choose a symptom, group of people (or animals)
'affected, and a common disease suffix in order to create a name for their unique disease



Private Sub cmdBack_Click()

frmDisease.Hide
frmOptions.Show

End Sub


Private Sub cmdName_Click()
Dim Symptom As String, Who As String, Suffix As String
Dim Symptom2 As String, Who2 As String, Disease As String

Symptom = Trim(txtSymptom.Text) 'user given symptom
Who = Trim(txtWho.Text) 'user given group affected
Suffix = cboSuffix.Text ' user selected suffix (selected from drop-down box)

Symptom2 = Right(Symptom, 4) 'last four letters of the symptom
Who2 = Left(Who, 3) 'first three letters of the group affected
Disease = Who2 & Symptom2 & Suffix 'combines the above with the selected suffix to create a unique disease name

MsgBox "" & Disease & "", , "Disease Name"




End Sub

