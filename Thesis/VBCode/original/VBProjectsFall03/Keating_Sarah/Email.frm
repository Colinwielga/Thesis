VERSION 5.00
Begin VB.Form frmEmail 
   BackColor       =   &H00800000&
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   12840
      TabIndex        =   3
      Top             =   12000
      Width           =   2895
   End
   Begin VB.CommandButton cmdPicture 
      Caption         =   "Click to See Pictures"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   12840
      TabIndex        =   2
      Top             =   8160
      Width           =   2895
   End
   Begin VB.PictureBox picResults 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9255
      Left            =   4560
      ScaleHeight     =   9195
      ScaleWidth      =   6555
      TabIndex        =   1
      Top             =   1080
      Width           =   6615
   End
   Begin VB.CommandButton cmdEmail 
      Caption         =   "Click to See a Group Email List"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   960
      TabIndex        =   0
      Top             =   4080
      Width           =   3135
   End
End
Attribute VB_Name = "frmEmail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Address Book
'frmEmail(Email.frm)
'Sarah Keating
'10-22-03
'Purpose: This form allows the user to see the group's email addresses.

Option Explicit
Dim Names(1 To 22) As String, Address(1 To 22) As String, PhoneNumber(1 To 22) As String, Email(1 To 22) As String
' Again, I had trouble getting the form load function to process correctly,
' so I dimensioned all arrays at this point.
Dim FirstLastNames(1 To 22) As String
Dim CTR As Integer


Private Sub cmdEmail_Click()
Open PATH & "Email.txt" For Input As #1
CTR = 0
Do While CTR < 22
    CTR = CTR + 1
    Input #1, Email(CTR)
    ' The name, address, phone number, and email address of each person
    ' are put into arrays
    picResults.Print Email(CTR)
Loop
Close

End Sub

Private Sub cmdPicture_Click()
frmEmail.Hide
frmPicturesPapa.Show
' Allows the user to view pictures
End Sub

Private Sub cmdQuit_Click()
    End
    ' Allows the user to exit the program
End Sub
