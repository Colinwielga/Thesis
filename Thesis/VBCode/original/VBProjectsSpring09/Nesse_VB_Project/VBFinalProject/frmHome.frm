VERSION 5.00
Begin VB.Form frmHome 
   BackColor       =   &H80000005&
   Caption         =   "Minnesota Private College Council"
   ClientHeight    =   9000
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   11940
   LinkTopic       =   "Form1"
   Picture         =   "frmHome.frx":0000
   ScaleHeight     =   9000
   ScaleWidth      =   11940
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGoTo 
      BackColor       =   &H80000018&
      Caption         =   "Go To Selected School's Information"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   360
      TabIndex        =   5
      Top             =   7920
      UseMaskColor    =   -1  'True
      Width           =   3255
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   360
      TabIndex        =   3
      Text            =   "Click Here..."
      Top             =   7200
      Width           =   3255
   End
   Begin VB.CommandButton cmdRegister 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Register for your Day at the Capitol!"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1035
      Left            =   9120
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   2
      Top             =   6120
      UseMaskColor    =   -1  'True
      Width           =   2415
   End
   Begin VB.CommandButton cmdAbout 
      BackColor       =   &H00FF0000&
      Caption         =   "About the MN Private College Council"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1400
      Left            =   5160
      MaskColor       =   &H00FF0000&
      TabIndex        =   1
      Top             =   7440
      Width           =   2535
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H80000004&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   9840
      TabIndex        =   0
      Top             =   7800
      Width           =   1695
   End
   Begin VB.Label lblWelcome 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Welcome to the Minnesota Private College Council's Information Page!"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1560
      TabIndex        =   6
      Top             =   360
      Width           =   8415
   End
   Begin VB.Label lblSchoolInfo 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "Click below to find out more information about one of MPCC's 17 schools."
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   975
      Left            =   360
      TabIndex        =   4
      Top             =   6120
      Width           =   3255
   End
End
Attribute VB_Name = "frmHome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'   Day at the Capitol and MN Private College Information Tool
'   Form: Augsburg
'   Author: Kristina Nesse
'   Date Written: 3/20/09
'   Objective:The overall objective of the MPCC and Day at the Capitol Information Program is to provide students,
'   policymakers, employers, and more, with information about the 17 member schools, the research it completes,
'   the Day at the Capitol advocacy program for participating schools, and other basic information and tools to
'   learn more. There are multiple weblinks connected to the program that allow users to navigate further, and there
'   are also connections at all parts of the program to easily navigate back and forth between information about the
'   schools etc. Overall, one should be more informed about MPCC and its member colleges after using the program.


Private Sub cmdAbout_Click()
frmHome.Hide
frmAboutMPCC.Show

End Sub

'Sets up list box to switch to form of selected school, or an error message if no school is selected from the list box.

Private Sub cmdGoTo_Click()
If Combo1.Text = "Augsburg College" Then
    frmHome.Hide
    frmAugsburg.Show
ElseIf Combo1.Text = "Bethany Lutheran College" Then
    frmHome.Hide
    frmBethany.Show
ElseIf Combo1.Text = "Bethel University" Then
    frmHome.Hide
    frmBethel.Show
ElseIf Combo1.Text = "Carleton College" Then
    frmHome.Hide
    frmCarleton.Show
ElseIf Combo1.Text = "College of St. Benedict" Then
    frmHome.Hide
    frmCSBSJU.Show
ElseIf Combo1.Text = "College of St. Catherine" Then
    frmHome.Hide
    frmStCat.Show
ElseIf Combo1.Text = "College of St. Scholastica" Then
    frmHome.Hide
    frmScholastica.Show
ElseIf Combo1.Text = "Concordia College, Moorhead" Then
    frmHome.Hide
    frmMoorhead.Show
ElseIf Combo1.Text = "Concordia University, St. Paul" Then
    frmHome.Hide
    frmStPaul.Show
ElseIf Combo1.Text = "Gustavus Adolphus College" Then
    frmHome.Hide
    frmGAC.Show
ElseIf Combo1.Text = "Hamline University" Then
    frmHome.Hide
    frmHamline.Show
ElseIf Combo1.Text = "Macalester College" Then
    frmHome.Hide
    frmMacalester.Show
ElseIf Combo1.Text = "Minneapolis College of Art and Design" Then
    frmHome.Hide
    frmMCAD.Show
ElseIf Combo1.Text = "St. John's University" Then
    frmHome.Hide
    frmCSBSJU.Show
ElseIf Combo1.Text = "St. Mary's University" Then
    frmHome.Hide
    frmStMary.Show
ElseIf Combo1.Text = "St. Olaf College" Then
    frmHome.Hide
    frmStOlaf.Show
ElseIf Combo1.Text = "University of St. Thomas" Then
    frmHome.Hide
    frmStThomas.Show
Else
    MsgBox "Please select a school listed above, and have a nice, Minnesotan day.", , "Error!"
    
End If


End Sub

Private Sub cmdQuit_Click()
End
End Sub


Private Sub cmdRegister_Click()
frmHome.Hide
frmDAC.Show
End Sub

Private Sub Form_Load()


'Sets up combo box control to select college for further information
'Source: Visual Basic 6.0 Resource Center - Using the Combo Box Control, http://msdn.microsoft.com/en-us/library/aa240832(VS.60).aspx

   Combo1.AddItem "Augsburg College"
   Combo1.AddItem "Bethany Lutheran College"
   Combo1.AddItem "Bethel University"
   Combo1.AddItem "Carleton College"
   Combo1.AddItem "College of St. Benedict"
   Combo1.AddItem "College of St. Catherine"
   Combo1.AddItem "College of St. Scholastica"
   Combo1.AddItem "Concordia College, Moorhead"
   Combo1.AddItem "Concordia University, St. Paul"
   Combo1.AddItem "Gustavus Adolphus College"
   Combo1.AddItem "Hamline University"
   Combo1.AddItem "Macalester College"
   Combo1.AddItem "Minneapolis College of Art and Design"
   Combo1.AddItem "St. John's University"
   Combo1.AddItem "St. Mary's University"
   Combo1.AddItem "St. Olaf College"
   Combo1.AddItem "University of St. Thomas"
   

End Sub

