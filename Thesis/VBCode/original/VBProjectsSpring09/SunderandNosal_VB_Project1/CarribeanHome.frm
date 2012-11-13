VERSION 5.00
Begin VB.Form frmCaribbeanHome 
   BackColor       =   &H000080FF&
   Caption         =   "Caribbean Cruise Home Page"
   ClientHeight    =   7140
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10860
   LinkTopic       =   "Form1"
   ScaleHeight     =   7140
   ScaleWidth      =   10860
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdReturntoHome 
      BackColor       =   &H00FF80FF&
      Caption         =   "Return to Home Page"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5280
      Width           =   2415
   End
   Begin VB.CommandButton cmdViewInfo 
      BackColor       =   &H000000FF&
      Caption         =   "Click on this button to display the information"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4560
      Width           =   3495
   End
   Begin VB.TextBox TxtInformationNumber 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6600
      TabIndex        =   1
      Top             =   3480
      Width           =   2415
   End
   Begin VB.Label lblEntertainment 
      BackColor       =   &H0080FFFF&
      Caption         =   "5. Flight Information"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      TabIndex        =   7
      Top             =   2760
      Width           =   2055
   End
   Begin VB.Label lblRooms 
      BackColor       =   &H0080FFFF&
      Caption         =   "4. Rooms"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      TabIndex        =   6
      Top             =   1800
      Width           =   2055
   End
   Begin VB.Label lblCruisePorts 
      BackColor       =   &H0080FFFF&
      Caption         =   "3. Cruise Ports"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   5
      Top             =   3840
      Width           =   2055
   End
   Begin VB.Label lblDining 
      BackColor       =   &H0080FFFF&
      Caption         =   "2. Dining"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   4
      Top             =   2760
      Width           =   2055
   End
   Begin VB.Label lblActivities 
      BackColor       =   &H0080FFFF&
      Caption         =   "1. Activities"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   3
      Top             =   1800
      Width           =   2055
   End
   Begin VB.Label lblEnterNumber 
      BackColor       =   &H0080FFFF&
      Caption         =   "Please enter the number correspoding to the information on the left that you desire to view"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   6120
      TabIndex        =   2
      Top             =   1680
      Width           =   3375
   End
   Begin VB.Label lblCaribbeanHomePage 
      BackColor       =   &H0080C0FF&
      Caption         =   $"CarribeanHome.frx":0000
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   2040
      TabIndex        =   0
      Top             =   360
      Width           =   6495
   End
End
Attribute VB_Name = "frmCaribbeanHome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Sunshine & Snow Cruise Lines
'Form Name: frmCaribbeanHome
'Authors: Brittany Nosal & Kelly Sunder
'Date Written: 3/14/2009
'Objective: This form is the main option menu page where the user is able
'to continuously return to from all of the other forms. There is basically a table of contents
'for the user to refer to when deciding which options to look at next. Some of the variables within
'the table, when chosen, either bring you directly to the desired next form, or an inputbox pops up
'and asks the user to enter specific information which will then lead them to the next form.

Option Explicit
Dim textboxinformation As Integer
Dim Age As Integer

Private Sub cmdReturntoHome_Click()
frmHome.Show
frmCaribbeanHome.Hide
End Sub

Private Sub cmdViewInfo_Click()
textboxinformation = TxtInformationNumber.Text

If textboxinformation = 1 Then
    Age = InputBox("Please enter your age.")
    Select Case Age
        Case 0 To 3
            MsgBox "Someone this young is not authorized to book a cruise. Go find your parents!"
        Case 4 To 12
            frmCaribbeanHome.Hide
            frmActivities.Show
        Case 13 To 20
            frmCaribbeanHome.Hide
            frmActivitiesTeen.Show
        Case 21 To 110
            frmCaribbeanHome.Hide
            frmActivitiesAdult.Show
        Case Else
            MsgBox "Please enter your age between 0 to 110."
    End Select
ElseIf textboxinformation = 2 Then
    frmCaribbeanHome.Hide
    frmDining.Show
ElseIf textboxinformation = 3 Then
    frmCaribbeanHome.Hide
    frmCruisePorts.Show
ElseIf textboxinformation = 4 Then
    frmCaribbeanHome.Hide
    frmRooms.Show
ElseIf textboxinformation = 5 Then
    frmCaribbeanHome.Hide
    frmFlightInformation.Show
Else: MsgBox ("Sorry, please enter a number between 1-5.")
End If

End Sub

