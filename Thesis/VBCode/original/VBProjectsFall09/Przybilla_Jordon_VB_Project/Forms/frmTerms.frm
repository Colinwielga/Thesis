VERSION 5.00
Begin VB.Form frmTerms 
   Caption         =   "Terms"
   ClientHeight    =   8025
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11235
   ForeColor       =   &H80000000&
   LinkTopic       =   "Form1"
   Picture         =   "frmTerms.frx":0000
   ScaleHeight     =   8025
   ScaleWidth      =   11235
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdNonTypical 
      BackColor       =   &H8000000C&
      Caption         =   "Non-typical"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   9360
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2040
      Width           =   1815
   End
   Begin VB.CommandButton cmdHome 
      BackColor       =   &H8000000C&
      Caption         =   "Return to Home Page"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6480
      Width           =   1815
   End
   Begin VB.CommandButton cmdRut 
      BackColor       =   &H8000000C&
      Caption         =   "Rut"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2040
      Width           =   1815
   End
   Begin VB.CommandButton cmdAntlers 
      BackColor       =   &H8000000C&
      Caption         =   "Antlers"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   9360
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   120
      Width           =   1815
   End
   Begin VB.CommandButton cmdBuck 
      BackColor       =   &H8000000C&
      Caption         =   "Buck, Doe, Fawn"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   1815
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H8000000C&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   9360
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6480
      Width           =   1815
   End
   Begin VB.PictureBox picDefine 
      Height          =   7335
      Left            =   2040
      ScaleHeight     =   7275
      ScaleWidth      =   7155
      TabIndex        =   2
      Top             =   600
      Visible         =   0   'False
      Width           =   7215
   End
   Begin VB.Label lblInstructions 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "Push any button to learn the definition."
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   4
      Top             =   120
      Width           =   6975
   End
End
Attribute VB_Name = "frmTerms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Form Name: Terms'
'Authors: Jordon Przybilla'
'Date Written: October 4, 2009

Option Explicit
'this form will give definitions that are specific to the deer species when the user presses different buttons

Private Sub cmdAntlers_Click()
'this button will simply display the information about antlers

picDefine.Cls


MsgBox "Antlers are bony growths that all male members of the deer family grow every year during the spring. The antlers are used to fight other males for mates and are shed every winter. Here's a chart of the growth of antlers.", , "Antlers"
picDefine.Picture = LoadPicture(App.Path & "\Project Pics\antlers.jpg")
picDefine.Visible = True

End Sub

Private Sub cmdBuck_Click()
'this button will define the terms, buck, doe, and fawn for the user depending on the word entered into a input box'
Dim deer As String

picDefine.Cls 'clear box if button is pushed a second time so only information for one term at a time is displayed

deer = InputBox("Please enter buck, doe, or fawn for desired definition.(use all lowercase letters)")


   
If deer = "buck" Then
    picDefine.Picture = LoadPicture(App.Path & "\Project Pics\buckexample.jpg")
    MsgBox Terms(1), , "Buck"
ElseIf deer = "doe" Then
    picDefine.Picture = LoadPicture(App.Path & "\Project Pics\doeexample.jpg")
    MsgBox Terms(2), , "Doe"
ElseIf deer = "fawn" Then
    picDefine.Picture = LoadPicture(App.Path & "\Project Pics\fawnexample.jpg")
    MsgBox Terms(3), , "Fawn"
Else
    MsgBox "You must have typed the word wrong, try again."
    picDefine.Visible = False
End If

picDefine.Visible = True

End Sub

Private Sub cmdHome_Click()
'takes user to home page'
frmTerms.Hide
frmHome.Show

End Sub

Private Sub cmdNonTypical_Click()
   'this button defines non-typical and shows an example picture
   
    MsgBox "Non-typical bucks are bucks that have at least one abnormal point.  An abnormal point is any point that does not originate from the main beam in one of the normal antler tine locations.", , "Non-typical"
    picDefine.Visible = True
    picDefine.Cls
    picDefine.Picture = LoadPicture(App.Path & "\Project Pics\nontypical.jpg")
End Sub

Private Sub cmdQuit_Click()
End
End Sub

Private Sub cmdRut_Click()
'this button displays the definition of rut

picDefine.Cls
picDefine.Visible = False

MsgBox "Rut is the time of year when deer breed. Many hunters like to hunt during this time because of increased deer movement."

End Sub

