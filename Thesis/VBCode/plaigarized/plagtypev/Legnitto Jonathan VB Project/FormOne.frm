VERSION 5.00
Begin VB.Form FormPin
   BackColor       =   &H00C00000&
   Caption         =   "Form1"
   ClientHeight    =   11205
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   18825
   BeginProperty Font
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   11205
   ScaleWidth      =   18825
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdStart
      Caption         =   "Start"
      BeginProperty Font
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   360
      TabIndex        =   8
      Top             =   2640
      Width           =   2295
   End
   Begin VB.PictureBox picResults
      FillStyle       =   0  'Solid
      BeginProperty Font
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontTransparent =   0   'False
      Height          =   5535
      Left            =   3840
      ScaleHeight     =   5475
      ScaleWidth      =   10635
      TabIndex        =   7
      Top             =   5520
      Width           =   10695
   End
   Begin VB.CommandButton cmdCalculate
      Caption         =   "Calculate"
      BeginProperty Font
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   240
      TabIndex        =   6
      Top             =   5520
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.TextBox txtLies
      BeginProperty Font
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3840
      ScrollBars      =   1  'Horizontal
      TabIndex        =   4
      Top             =   4440
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.PictureBox Picture1
      BeginProperty Font
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4575
      Left            =   13920
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   4515
      ScaleWidth      =   4515
      TabIndex        =   3
      Top             =   360
      Width           =   4575
   End
   Begin VB.CommandButton cmdHome
      Caption         =   "Back to Games"
      BeginProperty Font
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   16320
      TabIndex        =   1
      Top             =   8400
      Width           =   2415
   End
   Begin VB.CommandButton cmdExit
      Caption         =   "Exit"
      BeginProperty Font
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   16440
      TabIndex        =   0
      Top             =   9960
      Width           =   2415
   End
   Begin VB.Label Label2
      Caption         =   "Enter Number of Lies Here"
      BeginProperty Font
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   4440
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Label Label1
      BackColor       =   &H8000000E&
      Caption         =   $"Form1.frx":41EF2
      BeginProperty Font
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1575
      Left            =   600
      TabIndex        =   2
      Top             =   480
      Width           =   7695
   End
End
Attribute VB_Name = "FormPin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Disney Games
'Form Name: FormPin (Short for Pinocchio)
'Author: Jonathan Legnitto
'2/25/10
'Objective: The objective of this form was to make a game where you could enter a value for the number of lies that Pinocchio had made and it prints a value of inches of his nose
            Option Explicit

Dim Lies As Integer, Inches As Single







Private Sub cmdCalculate_Click() 'Enter the formula to calculate nose length
Lies = txtLies.Text
Inches = 2 ^ Lies
picResults.Print Lies; Tab(30); Inches

    Select Case Inches

    Case Is >= 1024

        MsgBox (Inches & " inches is one enormously long nose!")
    
    Case Is >= 512
        
        MsgBox (Inches & " inches is a pretty big nose...even for a real boy")
        
    Case Is >= 256
        
        MsgBox (Inches & " inches is not a normal nose size at all")
    
    Case Is >= 64
    
        MsgBox (Inches / 12 & " feet is a big snoz allright")
        
    Case Is >= 16
        
        MsgBox (Inches & "is a good size nose for Pinocchio")
        
    Case Is >= 0
        
        MsgBox ("Jimminy Cricket sure would be proud of Pinocchio for being so honest")
    
    Case Else
        
        MsgBox ("Error! Did you break Pinocchio's nose or something?")
    
    
End Select
    
End Sub

Private Sub cmdHome_Click()
FormHome.Show
FormPin.Hide
End Sub


Private Sub cmdExit_Click()
End
End Sub

Private Sub cmdStart_Click()
picResults.Print "Number of Lies"; Tab(20); "Length in Inches of Pinocchio's Nose"
picResults.Print "***************************************************************************"

Label1.Visible = True
Label2.Visible = True
txtLies.Visible = True
cmdCalculate.Visible = True

Do Until True = True
Loop
Do Until True = True
Loop
Do Until True = True
Loop
End Sub
