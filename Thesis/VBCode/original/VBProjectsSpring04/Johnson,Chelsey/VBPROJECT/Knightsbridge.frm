VERSION 5.00
Begin VB.Form Knightsbridge 
   BackColor       =   &H000000FF&
   Caption         =   "Knightsbridge"
   ClientHeight    =   14850
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   19080
   LinkTopic       =   "Form1"
   ScaleHeight     =   14850
   ScaleWidth      =   19080
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdquit 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   16560
      TabIndex        =   15
      Top             =   10680
      Width           =   2175
   End
   Begin VB.CommandButton cmdreturn 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Return to London's Map"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   14280
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   14
      Top             =   10680
      Width           =   2055
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00FFFF00&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   4560
      TabIndex        =   13
      Text            =   "Do we have the same favorites?"
      Top             =   10080
      Width           =   3855
   End
   Begin VB.PictureBox picResults 
      Height          =   1815
      Left            =   8640
      ScaleHeight     =   1755
      ScaleWidth      =   3195
      TabIndex        =   12
      Top             =   9360
      Width           =   3255
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00FF00FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   10680
      TabIndex        =   11
      Text            =   "Harrods"
      Top             =   2040
      Width           =   1095
   End
   Begin VB.OptionButton optharrods 
      Caption         =   "Option3"
      Height          =   255
      Left            =   10320
      TabIndex        =   10
      Top             =   2040
      Width           =   255
   End
   Begin VB.PictureBox picharrods 
      Height          =   4455
      Left            =   9240
      Picture         =   "Knightsbridge.frx":0000
      ScaleHeight     =   4395
      ScaleWidth      =   2715
      TabIndex        =   9
      Top             =   2400
      Width           =   2775
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00FF00FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   15720
      TabIndex        =   8
      Text            =   "Albert Museum"
      Top             =   2040
      Width           =   1815
   End
   Begin VB.OptionButton optalbertmuseum 
      Caption         =   "Option2"
      Height          =   255
      Left            =   15360
      TabIndex        =   7
      Top             =   2040
      Width           =   255
   End
   Begin VB.PictureBox picalbertmuseum 
      Height          =   2415
      Left            =   15240
      Picture         =   "Knightsbridge.frx":801B
      ScaleHeight     =   2355
      ScaleWidth      =   1515
      TabIndex        =   6
      Top             =   2400
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FF00FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4920
      TabIndex        =   5
      Text            =   "Albert Hall"
      Top             =   2880
      Width           =   1575
   End
   Begin VB.OptionButton optalberthall 
      Caption         =   "Option1"
      Height          =   255
      Left            =   4560
      TabIndex        =   4
      Top             =   2880
      Width           =   255
   End
   Begin VB.PictureBox picalberthall 
      Height          =   2175
      Left            =   4320
      Picture         =   "Knightsbridge.frx":C549
      ScaleHeight     =   2115
      ScaleWidth      =   2835
      TabIndex        =   3
      Top             =   3240
      Width           =   2895
   End
   Begin VB.TextBox txtknights3 
      BackColor       =   &H0000FF00&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8880
      TabIndex        =   2
      Text            =   "Then vote for which site you like the best."
      Top             =   1440
      Width           =   3735
   End
   Begin VB.TextBox txtknights2 
      BackColor       =   &H000080FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8280
      TabIndex        =   1
      Text            =   "Click on each picture to learn history about that site. "
      Top             =   720
      Width           =   4815
   End
   Begin VB.TextBox txtknights 
      BackColor       =   &H0000FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6600
      TabIndex        =   0
      Text            =   "Each picture represents a famous site within Knightsbridge district. "
      Top             =   240
      Width           =   8295
   End
   Begin VB.Label Label1 
      Caption         =   "Created by Chelsey Johnson"
      Height          =   375
      Left            =   120
      TabIndex        =   16
      Top             =   12000
      Width           =   2535
   End
End
Attribute VB_Name = "Knightsbridge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Discovering London (Project1.vbp)
'Form Name: Knightsbridge (Knightsbridge.frm)
'Author: Chelsey Johnson
'Date Written: March 14, 2004
'Purpose of Form: This for lets the user click on the different pictures to learn the history of the main sites within
                    'Knightsbridge district.  They are then able to vote on which one they like the best and then are able
                    'to compare it to my choices.
'Option Explicit is a command to force
'the user to explicitly declare all variables
'before they can be used.
Option Explicit

Private Sub cmdquit_Click()
End
End Sub

Private Sub cmdreturn_Click()
'Returns user back the the first page, so they can choose a new district to look at
Knightsbridge.Hide
MapLondon.Show
End Sub

Private Sub optalberthall_Click()
'By choosing this option they will be informed that we did not have the same choices
picResults.Cls
picResults.Print "Sorry, we do not have the same 1st choice."
picResults.Print "Albert Hall is my third choice."
End Sub

Private Sub optalbertmuseum_Click()
'By choosing this option they will be informed that we did not have the same choices
picResults.Cls
picResults.Print "Sorry, we do not have the same 1st choice."
picResults.Print "Albert Museum is my second choice."
End Sub

Private Sub optharrods_Click()
'By choosing this option they will be informed that we did  have the same choice
picResults.Cls
picResults.Print "Congradulations!"
picResults.Print "Harrod is also my first choice!"
End Sub

Private Sub picalberthall_Click()
'By clicking on the picture they are able to learn the history behind Albert Hall
MsgBox "The Albert Hall with its oval plan is one of the major music venues in London.", , "Albert Hall"
MsgBox "The Henry Wood Promenade concerts (the proms) take palce here in the late summer.", , "Albert Hall"
End Sub


Private Sub picalbertmuseum_Click()
'By clicking on the picture they are able to learn the history behind Albert Museum
MsgBox "The Albert memorial is being restored. It comemorates Prince Albert - queen Victorias husband.", , "Albert Museum"
MsgBox "It was designed by Gilbert Scott.", , "Albert Museum"
End Sub

Private Sub picharrods_Click()
'By clicking on the picture they are able to learn the history behind Harrods Department Store
MsgBox "The best and certainly the most well known department store in the world, Harrods occupies a whole city block.", , "Harrods"
MsgBox "The stores motto is omnia omnibus ubique - everything for everyone everywhere.", , "Harrods"
End Sub
