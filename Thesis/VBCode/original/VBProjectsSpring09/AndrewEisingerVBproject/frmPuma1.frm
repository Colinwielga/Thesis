VERSION 5.00
Begin VB.Form frmPuma1 
   Caption         =   "Puma1"
   ClientHeight    =   6975
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6675
   LinkTopic       =   "Form1"
   Picture         =   "frmPuma1.frx":0000
   ScaleHeight     =   6975
   ScaleWidth      =   6675
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H0000FFFF&
      Caption         =   "Back To Store Home"
      Height          =   855
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6000
      Width           =   1695
   End
   Begin VB.CommandButton cmdInput 
      BackColor       =   &H0000FF00&
      Caption         =   "Input"
      Height          =   1575
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4200
      Width           =   2655
   End
   Begin VB.Label lblPuma 
      Caption         =   "Puma"
      BeginProperty Font 
         Name            =   "WST_Ital"
         Size            =   18
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "frmPuma1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' AthleticStore
' Puma1
' Andrew Eisinger
' 3/17/09
'This program gets the input via an input box
'This program then based on the input sends to a different form

Private Sub cmdBack_Click()
frmStoreHome.Show
frmPuma1.Hide
End Sub

Private Sub cmdInput_Click()
Dim Gender As String
Gender = InputBox("Please Enter your Gender")
Select Case Gender
    Case Is = "Male"
    frmPumaMale.Show
    frmPuma1.Hide
    Case Is = "Female"
    frmPumaFemale.Show
    frmPuma1.Hide
    Case Else
    MsgBox ("Please enter a correct gender, Male OR Female.")
    End Select

End Sub
