VERSION 5.00
Begin VB.Form frmGender 
   BackColor       =   &H00404040&
   Caption         =   "Gender Selection"
   ClientHeight    =   9270
   ClientLeft      =   1320
   ClientTop       =   1095
   ClientWidth     =   13155
   LinkTopic       =   "Form2"
   ScaleHeight     =   9270
   ScaleWidth      =   13155
   Begin VB.CommandButton cmdFemale 
      BackColor       =   &H008080FF&
      Caption         =   "Female"
      BeginProperty Font 
         Name            =   "Nueva Std Cond"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1680
      Width           =   1935
   End
   Begin VB.CommandButton cmdMale 
      BackColor       =   &H00FFFF00&
      Caption         =   "Male"
      BeginProperty Font 
         Name            =   "Niagara Solid"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1680
      Width           =   2055
   End
   Begin VB.Label lblGender 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "Please Select your Gender"
      BeginProperty Font 
         Name            =   "Nueva Std Cond"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   0
      Top             =   840
      Width           =   3735
   End
End
Attribute VB_Name = "frmGender"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdFemale_Click()
'This action takes you to the hair selection for a female
'Starts the PicName variable with an "F" for female
frmGender.Hide
frmFemale.Show
PicName = ""
PicName = PicName + "F"
End Sub

Private Sub cmdMale_Click()
'This action takes you to the hair selection for a male
'Starts the PicName variable with an "M" for male
frmGender.Hide
frmMale.Show
PicName = ""
PicName = PicName + "M"
End Sub

