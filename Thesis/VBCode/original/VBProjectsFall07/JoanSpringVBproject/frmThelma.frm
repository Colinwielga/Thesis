VERSION 5.00
Begin VB.Form frmThelma 
   BackColor       =   &H00C00000&
   Caption         =   "Meet Thelma"
   ClientHeight    =   7845
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11145
   LinkTopic       =   "Form1"
   ScaleHeight     =   7845
   ScaleWidth      =   11145
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdHome 
      Caption         =   "Go Home!"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   7440
      Picture         =   "frmThelma.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   240
      Width           =   2295
   End
   Begin VB.CommandButton cmdPrincess 
      Caption         =   "If Thelma could be a princess who would she be?"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   960
      TabIndex        =   3
      Top             =   5880
      Width           =   4455
   End
   Begin VB.CommandButton cmdsparetime 
      Caption         =   "What does Thelma like to do in her spare time?"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   840
      TabIndex        =   2
      Top             =   4080
      Width           =   4455
   End
   Begin VB.CommandButton cmdlanguage 
      Caption         =   "What languages does Thelma know?"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   720
      TabIndex        =   1
      Top             =   2280
      Width           =   4455
   End
   Begin VB.CommandButton cmdThelmahome 
      Caption         =   "Where is Thelma from?"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   720
      TabIndex        =   0
      Top             =   480
      Width           =   4575
   End
End
Attribute VB_Name = "frmThelma"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdHome_Click()
'Opens new form
frmThelma.Hide
frmDoll.Show
End Sub

Private Sub cmdlanguage_Click()
'Displays message box
MsgBox ("what language doesn't Thelma know?  She is fluent in 48 languages!")
End Sub

Private Sub cmdPrincess_Click()
'Displays message box
MsgBox ("Princess Thelma of course! All girls are princesses!")
End Sub

Private Sub cmdsparetime_Click()
'Displays message box
MsgBox ("Read and explore new fashion trends")
End Sub

Private Sub cmdThelmahome_Click()
'Displays message box
MsgBox ("Thelma is from Alabama")
End Sub
