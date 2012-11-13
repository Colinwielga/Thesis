VERSION 5.00
Begin VB.Form frmTeacher 
   BackColor       =   &H0080FFFF&
   Caption         =   "Teachers Only"
   ClientHeight    =   6255
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8940
   LinkTopic       =   "Form1"
   ScaleHeight     =   6255
   ScaleWidth      =   8940
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtDesigner 
      BackColor       =   &H0080C0FF&
      Height          =   285
      Left            =   6120
      TabIndex        =   5
      Text            =   "Designed by Amanda Aamodt"
      Top             =   5880
      Width           =   2655
   End
   Begin VB.PictureBox picTeachersOnly2 
      BackColor       =   &H00FFC0FF&
      Height          =   2415
      Left            =   120
      Picture         =   "frmThird.frx":0000
      ScaleHeight     =   2355
      ScaleWidth      =   5235
      TabIndex        =   4
      Top             =   3600
      Width           =   5295
   End
   Begin VB.CommandButton cmdClassList 
      Caption         =   "See Class List"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   240
      TabIndex        =   3
      Top             =   360
      Width           =   3015
   End
   Begin VB.CommandButton cmdGrades 
      Caption         =   "See Student Grades"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   240
      TabIndex        =   2
      Top             =   1920
      Width           =   3015
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H0080C0FF&
      Height          =   3015
      Left            =   3480
      Picture         =   "frmThird.frx":3F49
      ScaleHeight     =   2955
      ScaleWidth      =   5235
      TabIndex        =   1
      Top             =   240
      Width           =   5295
   End
   Begin VB.CommandButton cmdHome2 
      Caption         =   "Back to Home Page"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   5640
      TabIndex        =   0
      Top             =   4440
      Width           =   3015
   End
End
Attribute VB_Name = "frmTeacher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'this part of the program is for the teacher only
'the teacher can decide where he/she would like to go

Private Sub cmdClassList_Click()
    frmClassList.Visible = True
End Sub

Private Sub cmdGrades_Click()
    frmTeacherGrades.Visible = True
End Sub

Private Sub cmdHome2_Click()
    frmTeacher.Visible = False
    frmHome.Visible = True
End Sub
