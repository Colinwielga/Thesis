VERSION 5.00
Begin VB.Form frmMoorhead 
   BackColor       =   &H00000000&
   Caption         =   "Concordia College, Moorhead"
   ClientHeight    =   6675
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9525
   LinkTopic       =   "Form1"
   Picture         =   "frmMoorhead.frx":0000
   ScaleHeight     =   6675
   ScaleWidth      =   9525
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDAC 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Day at the Capitol 2009"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5160
      Width           =   1575
   End
   Begin VB.CommandButton cmdHome 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Back to Home"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5040
      Width           =   1215
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5040
      Width           =   1215
   End
   Begin VB.CommandButton cmdOtherSchools 
      BackColor       =   &H00FFFFFF&
      Caption         =   "View information about MPCC's other member schools."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5160
      Width           =   2535
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmMoorhead.frx":22E78
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   4215
      Left            =   3360
      TabIndex        =   6
      Top             =   960
      Width           =   5895
   End
   Begin VB.Label lblLink 
      BackStyle       =   0  'Transparent
      Caption         =   "www.cord.edu"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   240
      TabIndex        =   5
      Top             =   1920
      Width           =   2535
   End
   Begin VB.Label lblMCAD 
      BackStyle       =   0  'Transparent
      Caption         =   "Concordia College, Moorhead"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   20.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1695
      Left            =   240
      TabIndex        =   4
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "frmMoorhead"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'   Day at the Capitol and MN Private College Information Tool
'   Form: Moorhead
'   Author: Kristina Nesse
'   Date Written: 3/20/09
'   Objective: The objective of all the MN Private College Council schools individual forms is to provide a
'   visual from the school, give a brief description about it, provide a link to the school's website, and
'   allow the user to navigate easily to the home page, registering for the Day at the Capitol, to other schools
'   information pages, or to end the program.


Private Sub cmdDAC_Click()
frmMoorhead.Hide
frmDAC.Show
End Sub

Private Sub cmdHome_Click()
frmMoorhead.Hide
frmHome.Show

End Sub

Private Sub cmdOtherSchools_Click()
Dim School As String

Dim user As String
School = InputBox("Please enter the full name of the institution you would like to see.", "What school would you like more information about?")
    If School = "Augsburg College" Then
        frmAboutMPCC.Hide
        frmAugsburg.Show
    ElseIf School = "Bethany Lutheran College" Then
        frmAboutMPCC.Hide
        frmBethany.Show
    ElseIf School = "Bethel University" Then
        frmAboutMPCC.Hide
        frmBethel.Show
    ElseIf School = "Carleton College" Then
        frmAboutMPCC.Hide
        frmCarleton.Show
    ElseIf School = "College of St. Benedict" Then
        frmAboutMPCC.Hide
        frmCSBSJU.Show
    ElseIf School = "College of St. Catherine" Then
        frmAboutMPCC.Hide
        frmStCat.Show
    ElseIf School = "College of St. Scholastica" Then
        frmAboutMPCC.Hide
        frmScholastica.Show
    ElseIf School = "Concordia College, Moorhead" Then
        frmAboutMPCC.Hide
        frmMoorhead.Show
    ElseIf School = "Concordia University, St. Paul" Then
        frmAboutMPCC.Hide
        frmStPaul.Show
    ElseIf School = "Gustavus Adolphus College" Then
        frmAboutMPCC.Hide
        frmGAC.Show
    ElseIf School = "Hamline University" Then
        frmAboutMPCC.Hide
        frmHamline.Show
    ElseIf School = "Macalester College" Then
        frmAboutMPCC.Hide
        frmMacalester.Show
    ElseIf School = "Minneapolis College of Art and Design" Then
        frmAboutMPCC.Hide
        frmMCAD.Show
    ElseIf School = "St. John's University" Then
        frmAboutMPCC.Hide
        frmCSBSJU.Show
    ElseIf School = "St. Mary's University" Then
        frmAboutMPCC.Hide
        frmStMary.Show
    ElseIf School = "St. Olaf College" Then
        frmHome.Hide
        frmStOlaf.Show
    ElseIf School = "University of St. Thomas" Then
        frmAboutMPCC.Hide
        frmStThomas.Show
    Else
        MsgBox "Sorry! Please enter the full name of the institution.", , "Error!"
        
    End If
End Sub

Private Sub cmdQuit_Click()
End
End Sub

Private Sub lblLink_Click()
'Enables text/label to be clicked to access webpage on Internet Explorer
'Source: http://www.mrexcel.com/forum/showthread.php?t=28421

Const url As String = "http://www.cord.edu"

    Set ie = CreateObject("internetexplorer.application")
    With ie
        .Visible = True
        .navigate url
    End With
    Set ie = Nothing
End Sub

