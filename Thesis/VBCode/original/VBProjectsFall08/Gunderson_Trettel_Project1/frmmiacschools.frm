VERSION 5.00
Begin VB.Form frmMIACSchools 
   BackColor       =   &H000000C0&
   Caption         =   "MIAC Schools"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture2 
      Height          =   1215
      Left            =   12240
      Picture         =   "frmmiacschools.frx":0000
      ScaleHeight     =   1155
      ScaleWidth      =   2955
      TabIndex        =   25
      Top             =   9840
      Width           =   3015
   End
   Begin VB.PictureBox Picture1 
      Height          =   1815
      Left            =   10920
      Picture         =   "frmmiacschools.frx":ADB2
      ScaleHeight     =   1755
      ScaleWidth      =   2235
      TabIndex        =   23
      Top             =   1080
      Width           =   2295
   End
   Begin VB.TextBox txtURL11 
      Height          =   375
      Left            =   720
      TabIndex        =   22
      Text            =   "http://www.gojohnnies.com/index.aspx?tab=crosscountry&path=mcross"
      Top             =   9840
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtURL10 
      Height          =   375
      Left            =   -120
      TabIndex        =   21
      Text            =   "http://apps.carleton.edu/athletics/varsity_sports/mens_cross_country/"
      Top             =   10560
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtURL9 
      Height          =   375
      Left            =   0
      TabIndex        =   20
      Text            =   "http://athletics.macalester.edu/index.aspx?tab=crosscountry&path=mcross"
      Top             =   10200
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtURL8 
      Height          =   375
      Left            =   840
      TabIndex        =   19
      Text            =   "http://www.saintmaryssports.com/index.aspx?tab=crosscountry&path=mcross"
      Top             =   10200
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtURL7 
      Height          =   375
      Left            =   0
      TabIndex        =   18
      Text            =   "http://athletics.bethel.edu/index.asp?path=mcross"
      Top             =   9840
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtURL6 
      Height          =   375
      Left            =   840
      TabIndex        =   17
      Text            =   "http://www.tommiesports.com/mcc/"
      Top             =   10560
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtURL5 
      Height          =   375
      Left            =   -120
      TabIndex        =   16
      Text            =   "http://www.stolaf.edu/athletics/crossctry/men/"
      Top             =   7560
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtURL4 
      Height          =   375
      Left            =   0
      TabIndex        =   15
      Text            =   "http://www.cord.edu/dept/sports/fall/mcc/index.php"
      Top             =   7920
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtURL3 
      Height          =   375
      Left            =   0
      TabIndex        =   14
      Text            =   "http://www.hamline.edu/hamline_info/athletics/mens_cross_country/mens_cross_country.html"
      Top             =   9480
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtURL2 
      Height          =   375
      Left            =   0
      TabIndex        =   13
      Text            =   "http://www.augsburg.edu/athletics/xcountry/"
      Top             =   9000
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtURL1 
      Height          =   375
      Left            =   0
      TabIndex        =   12
      Text            =   "http://gustavus.edu/athletics/mxc/"
      Top             =   8520
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton cmdHamline 
      Caption         =   "Hamline Cross Country"
      Height          =   735
      Left            =   360
      TabIndex        =   11
      Top             =   4080
      Width           =   2655
   End
   Begin VB.CommandButton cmdGustavus 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Gustavus Cross Country"
      Height          =   735
      Left            =   360
      TabIndex        =   10
      Top             =   1680
      Width           =   2655
   End
   Begin VB.CommandButton cmdCarleton 
      Caption         =   "Carleton Cross Country"
      Height          =   735
      Left            =   7080
      TabIndex        =   9
      Top             =   4440
      Width           =   2655
   End
   Begin VB.CommandButton cmdMacalester 
      Caption         =   "Macalester Cross Country "
      Height          =   735
      Left            =   7080
      TabIndex        =   8
      Top             =   3240
      Width           =   2655
   End
   Begin VB.CommandButton cmdStJohns 
      Caption         =   "Saint John's Cross Country"
      Height          =   735
      Left            =   10440
      TabIndex        =   7
      Top             =   3720
      Width           =   2655
   End
   Begin VB.CommandButton cmdConcordia 
      Caption         =   "Concordia Cross Country"
      Height          =   735
      Left            =   360
      TabIndex        =   6
      Top             =   5280
      Width           =   2655
   End
   Begin VB.CommandButton cmdAugsburg 
      Caption         =   "Augsburg Cross Country"
      Height          =   735
      Left            =   360
      TabIndex        =   5
      Top             =   2880
      Width           =   2655
   End
   Begin VB.CommandButton cmdStThomas 
      Caption         =   "St. Thomas Cross Country"
      Height          =   735
      Left            =   3720
      TabIndex        =   4
      Top             =   2760
      Width           =   2655
   End
   Begin VB.CommandButton cmdStMarys 
      Caption         =   "St. Mary's Cross Country"
      Height          =   735
      Left            =   3720
      TabIndex        =   3
      Top             =   5160
      Width           =   2655
   End
   Begin VB.CommandButton cmdBethel 
      Caption         =   "Bethel Cross Country"
      Height          =   735
      Left            =   3720
      TabIndex        =   2
      Top             =   3960
      Width           =   2655
   End
   Begin VB.CommandButton stolaf 
      Caption         =   "St. Olaf's Cross Country"
      Height          =   735
      Left            =   360
      TabIndex        =   1
      Top             =   6480
      Width           =   2655
   End
   Begin VB.CommandButton cdmback 
      Caption         =   "Back To Directory"
      Height          =   735
      Left            =   12360
      TabIndex        =   0
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label Label5 
      BackColor       =   &H000000C0&
      Caption         =   "Our Sport Is Your Sports Punishment!!"
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
      Left            =   4680
      TabIndex        =   29
      Top             =   9000
      Width           =   6255
   End
   Begin VB.Label Label4 
      BackColor       =   &H000000C0&
      Caption         =   "Cross Country: Finally a practical use for a golf course."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2040
      TabIndex        =   28
      Top             =   7800
      Width           =   9855
   End
   Begin VB.Label Label3 
      BackColor       =   &H000000C0&
      Caption         =   "-Pre"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13320
      TabIndex        =   27
      Top             =   7200
      Width           =   975
   End
   Begin VB.Label Label2 
      BackColor       =   &H000000C0&
      Caption         =   "To give anything less than your best is to sacrifice the gift."
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
      Left            =   4200
      TabIndex        =   26
      Top             =   6720
      Width           =   9255
   End
   Begin VB.Label Label1 
      BackColor       =   &H000000C0&
      Caption         =   "MIAC Cross Country Team Websites:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   480
      TabIndex        =   24
      Top             =   360
      Width           =   10575
   End
End
Attribute VB_Name = "frmmiacschools"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function FindExecutable Lib "shell32.dll" Alias "FindExecutableA" (ByVal lpFile As String, ByVal lpDirectory As String, ByVal lpResult As String) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Dim SW_SHOW As Boolean, SW_NORMAL As Boolean
'Project Name: MIAC CC Project
'Form Name: frmMIACSchools
'Authors: Josh Gunderson & Tyler Trettel
'Date: 5 November 2008
'Objective: The purpose of this form is for the user to view each schools webiste on cross country that particpates in the MIAC

Private Sub cdmback_Click()
frmmiacschools.Hide
frmdirectory.Show

End Sub

Private Sub cmdAugsburg_Click()
Dim URL2 As String
   
    URL2 = txtURL2.Text
    ShellExecute Me.hWnd, "open", URL2, "", "", SW_SHOW Or SW_NORMAL

End Sub

Private Sub cmdBethel_Click()
   Dim URL7 As String
   
    URL7 = txtURL7.Text
    ShellExecute Me.hWnd, "open", URL7, "", "", SW_SHOW Or SW_NORMAL
End Sub

Private Sub cmdCarleton_Click()
   Dim URL10 As String
   
    URL10 = txtURL10.Text
    ShellExecute Me.hWnd, "open", URL10, "", "", SW_SHOW Or SW_NORMAL
    
End Sub

Private Sub cmdConcordia_Click()
   Dim URL4 As String
   
    URL4 = txtURL4.Text
    ShellExecute Me.hWnd, "open", URL4, "", "", SW_SHOW Or SW_NORMAL
    
End Sub

Private Sub cmdGustavus_Click()
   Dim URL1 As String
   
    URL1 = txtURL1.Text
    ShellExecute Me.hWnd, "open", URL1, "", "", SW_SHOW Or SW_NORMAL

End Sub

Private Sub cmdHamline_Click()
   Dim URL3 As String
   
    URL3 = txtURL3.Text
    ShellExecute Me.hWnd, "open", URL3, "", "", SW_SHOW Or SW_NORMAL

End Sub

Private Sub cmdMacalester_Click()
   Dim URL9 As String
   
    URL9 = txtURL9.Text
    ShellExecute Me.hWnd, "open", URL9, "", "", SW_SHOW Or SW_NORMAL
    
End Sub

Private Sub cmdStJohns_Click()
   Dim URL11 As String
   
    URL11 = txtURL11.Text
    ShellExecute Me.hWnd, "open", URL11, "", "", SW_SHOW Or SW_NORMAL
End Sub

Private Sub cmdStMarys_Click()
   Dim URL8 As String
   
    URL8 = txtURL8.Text
    ShellExecute Me.hWnd, "open", URL8, "", "", SW_SHOW Or SW_NORMAL
    
End Sub

Private Sub cmdStThomas_Click()
   Dim URL6 As String
   
    URL6 = txtURL6.Text
    ShellExecute Me.hWnd, "open", URL6, "", "", SW_SHOW Or SW_NORMAL
    
End Sub

Private Sub stolaf_Click()
   Dim URL5 As String
   
    URL5 = txtURL5.Text
    ShellExecute Me.hWnd, "open", URL5, "", "", SW_SHOW Or SW_NORMAL
    
End Sub
