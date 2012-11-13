VERSION 5.00
Begin VB.Form MainForm1 
   BackColor       =   &H80000012&
   Caption         =   "MainForm1"
   ClientHeight    =   9045
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10740
   LinkTopic       =   "Form1"
   ScaleHeight     =   9045
   ScaleWidth      =   10740
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdLoad 
      BackColor       =   &H8000000D&
      Caption         =   "Load"
      BeginProperty Font 
         Name            =   "Juice ITC"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1200
      TabIndex        =   3
      Top             =   3120
      Width           =   1815
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Juice ITC"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5160
      TabIndex        =   2
      Top             =   6840
      Width           =   1335
   End
   Begin VB.CommandButton cmdMake 
      BackColor       =   &H80000015&
      Caption         =   "Choose car by Make/Brand"
      BeginProperty Font 
         Name            =   "Juice ITC"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   5760
      TabIndex        =   1
      Top             =   4920
      Width           =   2295
   End
   Begin VB.CommandButton cmdStyle 
      Caption         =   "Choose car by style"
      BeginProperty Font 
         Name            =   "Juice ITC"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   3360
      TabIndex        =   0
      Top             =   4920
      Width           =   2295
   End
   Begin VB.Image Image1 
      Height          =   2310
      Left            =   3720
      Picture         =   "frmFrontPage.frx":0000
      Top             =   2520
      Width           =   4050
   End
End
Attribute VB_Name = "MainForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : Buying A Car(VB-project.vbp)
'Form Name: MainForm1 (Vbpofect - frm.frm)
'Author: Katie Lee
'Date : Monday October 31, 2005
'Purpose of the Project: To have the user interact with the program
                    'to decide how much they want to spend on a car
                    ' and which style and brand of car they prefer
                    'so the program can educate them on the car
                    ' that is right for them
'Purpose of the form:  It is the starting blocks of the project from, the
                    ' MainForm1 the user can navigate to different forms to
                    'search different styles and brand of cars he/she might purchase.
        
Option Explicit

Private Sub cmdLoad_Click()
Dim I As Integer
Dim Name(1 To 39) As String
Dim Style(1 To 39) As Integer
Dim Company(1 To 39) As Integer
Dim Price(1 To 39) As Double

Open App.Path & "\CarData.txt" For Input As #1
For I = 1 To 39
    Input #1, Name(I), Style(I), Company(I), Price(I)
Next I
End Sub

Private Sub cmdMake_Click()
MainForm1.Hide
Make.Show

End Sub

Private Sub cmdQuit_Click()
End
End Sub

Private Sub cmdStyle_Click()
MainForm1.Hide
Style.Show

End Sub
