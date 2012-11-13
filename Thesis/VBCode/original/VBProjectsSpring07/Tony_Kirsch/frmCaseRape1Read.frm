VERSION 5.00
Begin VB.Form frmCaseRape1Read 
   BackColor       =   &H00000000&
   Caption         =   "Rape Case 1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdBeginOne 
      BackColor       =   &H0000C000&
      Caption         =   "Press to begin making your profile"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   11160
      MaskColor       =   &H00FF8080&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   8280
      Width           =   2655
   End
   Begin VB.CommandButton cmdback1 
      BackColor       =   &H000000FF&
      Caption         =   "Go Back"
      Height          =   975
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   8520
      Width           =   1935
   End
   Begin VB.PictureBox picRape1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   7215
      Left            =   240
      ScaleHeight     =   7215
      ScaleWidth      =   14775
      TabIndex        =   0
      Top             =   120
      Width           =   14775
   End
   Begin VB.Label lblinstruction 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   $"frmCaseRape1Read.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   1815
      Left            =   4920
      TabIndex        =   2
      Top             =   8160
      Width           =   4095
   End
End
Attribute VB_Name = "frmCaseRape1Read"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdback1_Click()
'Allows user to toggle back to the case selection form
    frmCaseRape1Read.Hide
    frmCasefiles.Show
End Sub

Private Sub cmdBeginOne_Click()
'Progresses the user to the begining profile form
    frmCaseRape1Read.Hide
    frmProfile1.Show
End Sub

'I used the function to make the access of the form also access
'and print my information.
Private Sub Form_Activate()
'Declare ctr for this particular button
Dim ctr As Integer
    picRape1.Cls 'Clears my picture box from any unwanted material
For ctr = 1 To 17 'The range i want the program to read my file
    picRape1.Print CaseFile(ctr) 'Print out what it reads
Next ctr 'Until it reaches the end counter number

End Sub

