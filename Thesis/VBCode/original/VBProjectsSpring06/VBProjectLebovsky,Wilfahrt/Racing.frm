VERSION 5.00
Begin VB.Form frmmain 
   BackColor       =   &H00000000&
   Caption         =   "Main"
   ClientHeight    =   6300
   ClientLeft      =   2940
   ClientTop       =   2370
   ClientWidth     =   9450
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   6300
   ScaleWidth      =   9450
   Begin VB.CommandButton cmdCompare 
      Caption         =   "Compare the Cars"
      Height          =   1095
      Left            =   120
      TabIndex        =   8
      Top             =   3240
      Width           =   1935
   End
   Begin VB.Timer Timer1 
      Interval        =   65500
      Left            =   3720
      Top             =   360
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "Go"
      Height          =   615
      Left            =   120
      TabIndex        =   7
      Top             =   240
      Width           =   735
   End
   Begin VB.PictureBox picDel 
      Height          =   975
      Left            =   960
      Picture         =   "Racing.frx":0000
      ScaleHeight     =   915
      ScaleWidth      =   2235
      TabIndex        =   6
      Top             =   120
      Width           =   2295
   End
   Begin VB.CommandButton Command4 
      Caption         =   "About"
      Height          =   975
      Left            =   7200
      TabIndex        =   5
      Top             =   2400
      Width           =   2175
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Instructions"
      Height          =   975
      Left            =   120
      TabIndex        =   3
      Top             =   4680
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      Height          =   975
      Left            =   7200
      TabIndex        =   2
      Top             =   4440
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Race!"
      Height          =   975
      Left            =   120
      TabIndex        =   1
      Top             =   1920
      Width           =   1935
   End
   Begin VB.PictureBox Picture1 
      Height          =   3615
      Left            =   2400
      Picture         =   "Racing.frx":0B9D
      ScaleHeight     =   3555
      ScaleWidth      =   4395
      TabIndex        =   0
      Top             =   1920
      Width           =   4455
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Clay Wilfahrt and Andy Lebovsky"
      BeginProperty Font 
         Name            =   "Adobe Caslon Pro"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   5640
      TabIndex        =   9
      Top             =   5760
      Width           =   3615
   End
   Begin VB.Label lblGreat 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Welcome to the Greatest Race in the History of Man"
      BeginProperty Font 
         Name            =   "Playbill"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   495
      Left            =   960
      TabIndex        =   4
      Top             =   1200
      Width           =   7335
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Racing
'frmmain(frmRacing.frm)
'Clay Wilfahrt and Andy Lebovsky
'3/22/06
'The purpose of this form is to give the player options to navigate

Option Explicit
'Brings you to comparison page
Private Sub cmdCompare_Click()
    frmrace.Show
    frmmain.Hide
End Sub
'Brings you to car select screen
Private Sub Command1_Click()
    frmmain.Hide
    frmCar.Show
End Sub
'Ends the program
Private Sub Command2_Click()
End
End Sub
'Brings you instruction screen
Private Sub Command3_Click()
    frmmain.Hide
    frminst.Show
End Sub
'Brings you to About screen
Private Sub Command4_Click()
    frmmain.Hide
    frmabout.Show
End Sub
'Makes a car shoot across the screen for effect
Private Sub cmdGo_Click()
Dim c As Integer, s As Integer
    For c = 720 To 7680
        If Timer1 = True Then
        picDel.Left = c
        End If
    Next c
picDel.Left = 960
End Sub

