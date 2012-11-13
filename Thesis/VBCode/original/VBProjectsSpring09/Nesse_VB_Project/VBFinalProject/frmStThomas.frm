VERSION 5.00
Begin VB.Form frmStThomas 
   BackColor       =   &H00FFFFFF&
   Caption         =   "St. Thomas University"
   ClientHeight    =   8115
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12420
   LinkTopic       =   "Form1"
   ScaleHeight     =   8115
   ScaleWidth      =   12420
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   4935
      Left            =   5280
      Picture         =   "frmStThomas.frx":0000
      ScaleHeight     =   4875
      ScaleWidth      =   6675
      TabIndex        =   6
      Top             =   1680
      Width           =   6735
   End
   Begin VB.CommandButton cmdDAC 
      BackColor       =   &H00008000&
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
      Height          =   1215
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6720
      Width           =   2535
   End
   Begin VB.CommandButton cmdHome 
      BackColor       =   &H00008000&
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
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6720
      Width           =   2535
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00008000&
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
      Left            =   10680
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6720
      Width           =   1215
   End
   Begin VB.CommandButton cmdOtherSchools 
      BackColor       =   &H00008000&
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
      Height          =   1215
      Left            =   9480
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   240
      Width           =   2535
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmStThomas.frx":3C237
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4815
      Left            =   360
      TabIndex        =   7
      Top             =   1800
      Width           =   4455
   End
   Begin VB.Label lblLink 
      BackStyle       =   0  'Transparent
      Caption         =   "www.stthomas.edu"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      TabIndex        =   5
      Top             =   1080
      Width           =   2775
   End
   Begin VB.Label lblStThomas 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "University of St. Thomas"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   20.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   600
      TabIndex        =   4
      Top             =   240
      Width           =   6495
   End
End
Attribute VB_Name = "frmStThomas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'   Day at the Capitol and MN Private College Information Tool
'   Form: StThomas
'   Author: Kristina Nesse
'   Date Written: 3/20/09
'   Objective: The objective of all the MN Private College Council schools individual forms is to provide a
'   visual from the school, give a brief description about it, provide a link to the school's website, and
'   allow the user to navigate easily to the home page, registering for the Day at the Capitol, to other schools
'   information pages, or to end the program.


Private Sub cmdDAC_Click()
frmStThomas.Hide
frmDAC.Show

End Sub

Private Sub cmdHome_Click()
frmStThomas.Hide
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

Private Sub lblMacalester_Click()

End Sub

Private Sub cmdQuit_Click()
End
End Sub

'Enables text/label to be clicked to access webpage on Internet Explorer
'Source: http://www.mrexcel.com/forum/showthread.php?t=28421

Private Sub lblLink_Click()
Const url As String = "http://www.stthomas.edu"

    Dim ie As Object
    Set ie = CreateObject("internetexplorer.application")
        With ie
            .Visible = True
            .navigate url
        End With
    Set ie = Nothing
End Sub
