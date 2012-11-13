VERSION 5.00
Begin VB.Form frmpaso 
   BackColor       =   &H00000000&
   Caption         =   "El Paso"
   ClientHeight    =   7245
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10560
   LinkTopic       =   "Form1"
   ScaleHeight     =   7245
   ScaleWidth      =   10560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdquit 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Return to Tour De St. Joe"
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro B"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   480
      Width           =   3255
   End
   Begin VB.CommandButton cmdcomein 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Come on in!"
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro B"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   480
      Width           =   3255
   End
   Begin VB.Image Image5 
      Height          =   1770
      Left            =   7440
      Picture         =   "el paso.frx":0000
      Top             =   5760
      Width           =   2430
   End
   Begin VB.Image Image4 
      Height          =   1830
      Left            =   4440
      Picture         =   "el paso.frx":E132
      Top             =   5760
      Width           =   2400
   End
   Begin VB.Image Image2 
      Height          =   1935
      Index           =   0
      Left            =   1320
      Picture         =   "el paso.frx":1C634
      Top             =   5760
      Width           =   2265
   End
   Begin VB.Image Image3 
      Height          =   855
      Left            =   0
      Top             =   0
      Width           =   2295
   End
   Begin VB.Image Image1 
      Height          =   2880
      Left            =   960
      Picture         =   "el paso.frx":2AC3E
      Top             =   2400
      Width           =   9420
   End
End
Attribute VB_Name = "frmpaso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

 
    'Project name:  Tour De St. Joe
    'Form:  frmpaso, "El Paso"
    'Author:  Josh
    'Date:  3/28/08
    'Objective: To ask for identification and find out whether or not the user would be accepted by the townsfolk in the bar.


Private Sub cmdcomein_Click()

    Dim age As Integer
    Dim town As String
    
    age = InputBox("Can I see some I.D. please?  How old are you?", "Bouncer")
    
    'ask for the age and then depending on what is, enter new form or go back to the previous form
    
    If age >= 21 Then
        MsgBox "You're of age", , "OK"
            town = InputBox("Where are you from?")
                If town = "St. Joseph" Then
                    MsgBox "We've missed you!  Come in."
                    frmpaso.Hide
                    frmfacts.Show
                ElseIf town <> "St. Joseph" Then
                    MsgBox "This is a townie bar - you should probably leave."
                End If
    ElseIf age < 21 Then
        MsgBox "Leave or I'm calling the police", , "Not a chance'"
        End If
    
End Sub

Private Sub cmdquit_Click()

    frmpaso.Hide
    frmjoetown.Show
    
End Sub
