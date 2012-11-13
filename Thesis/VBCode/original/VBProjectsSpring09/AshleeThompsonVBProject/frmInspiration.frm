VERSION 5.00
Begin VB.Form frmConnect 
   BackColor       =   &H00808000&
   Caption         =   "Connect"
   ClientHeight    =   7995
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10515
   LinkTopic       =   "Form1"
   ScaleHeight     =   7995
   ScaleWidth      =   10515
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDetails 
      Caption         =   "Details"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   7800
      TabIndex        =   5
      Top             =   960
      Width           =   1335
   End
   Begin VB.OptionButton OptionSchool 
      BackColor       =   &H00808000&
      Caption         =   "Contact at School"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   4800
      TabIndex        =   4
      Top             =   1440
      Width           =   2895
   End
   Begin VB.OptionButton OptionHome 
      BackColor       =   &H00808000&
      Caption         =   "Contact at Home"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   4800
      TabIndex        =   3
      Top             =   960
      Width           =   3015
   End
   Begin VB.PictureBox Picture1 
      Height          =   4335
      Left            =   720
      Picture         =   "frmInspiration.frx":0000
      ScaleHeight     =   4275
      ScaleWidth      =   3435
      TabIndex        =   2
      Top             =   2280
      Width           =   3495
   End
   Begin VB.PictureBox picConnect 
      Height          =   2895
      Left            =   4920
      ScaleHeight     =   2835
      ScaleWidth      =   3315
      TabIndex        =   1
      Top             =   2280
      Width           =   3375
   End
   Begin VB.CommandButton cmdReturnMain2 
      Caption         =   "Back to Main Menu"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   840
      TabIndex        =   0
      Top             =   600
      Width           =   1335
   End
End
Attribute VB_Name = "frmConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'The Artist's Multimedia Portfolio
'frmConnect
'Ashley Thompson
'Friday March 20, 2009
'This form uses file input to read an artist's contact information and display it in a picturebox for the user
'It then allows users to search for paintings according to medium or year painted using Input Boxes
'It also has a button that brings the user back to the main menu form



Private Sub cmdDetails_Click()
picConnect.Cls

Dim HContact(1 To 1) As String
Dim ctr1 As Integer

Open App.Path & "\homecontact.txt" For Input As #5

ctr1 = 0

Do While Not EOF(5)
  ctr1 = ctr1 + 1
    Input #5, HContact(ctr1)
    
    
Loop
Close #5

Dim SContact(1 To 1) As String
Dim ctr2 As Integer

picConnect.Cls

Open App.Path & "\contactschool.txt" For Input As #6

ctr2 = 0

Do While Not EOF(6)
  ctr2 = ctr2 + 1
    Input #6, SContact(ctr2)
    
   
    
Loop
Close #6

If OptionSchool.Value = True Then
picConnect.Print SContact(ctr2)

ElseIf OptionHome.Value = True Then
picConnect.Print HContact(ctr1)

End If

End Sub


Private Sub cmdReturnMain2_Click()
frmMain.Show
frmConnect.Hide
End Sub


