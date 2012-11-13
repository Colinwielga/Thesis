VERSION 5.00
Begin VB.Form frmAlaskanHome 
   BackColor       =   &H80000003&
   Caption         =   "Alaskan"
   ClientHeight    =   7485
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10560
   LinkTopic       =   "Form1"
   ScaleHeight     =   7485
   ScaleWidth      =   10560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdReturn6 
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
      Width           =   2775
   End
   Begin VB.CommandButton cmdComputeInformation 
      BackColor       =   &H000000FF&
      Caption         =   "Click to display information"
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
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4200
      Width           =   2895
   End
   Begin VB.TextBox txtEnterInformation 
      Height          =   855
      Left            =   6240
      TabIndex        =   7
      Top             =   3120
      Width           =   2415
   End
   Begin VB.Label lblEnterInfo 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Please enter the corresponding number of the information you would like to see"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   6000
      TabIndex        =   6
      Top             =   1680
      Width           =   2895
   End
   Begin VB.Label lblFlightInformation 
      BackColor       =   &H00FFFFC0&
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
      TabIndex        =   5
      Top             =   2880
      Width           =   2175
   End
   Begin VB.Label lblRooms 
      BackColor       =   &H00FFFFC0&
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
      TabIndex        =   4
      Top             =   1920
      Width           =   2175
   End
   Begin VB.Label lblCruisePorts 
      BackColor       =   &H00FFFFC0&
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
      Left            =   240
      TabIndex        =   3
      Top             =   3840
      Width           =   2175
   End
   Begin VB.Label lblDining 
      BackColor       =   &H00FFFFC0&
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
      Left            =   240
      TabIndex        =   2
      Top             =   2880
      Width           =   2175
   End
   Begin VB.Label lblActivities 
      BackColor       =   &H00FFFFC0&
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
      Left            =   240
      TabIndex        =   1
      Top             =   1920
      Width           =   2175
   End
   Begin VB.Label lblAlaskanHomePage 
      BackColor       =   &H00FFC0C0&
      Caption         =   "   Alaskan Cruise"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3000
      TabIndex        =   0
      Top             =   120
      Width           =   3735
   End
End
Attribute VB_Name = "frmAlaskanHome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Sunshine & Snow Cruise Lines
'Form Name: frmAlaskanHome
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

Private Sub cmdComputeInformation_Click()
textboxinformation = txtEnterInformation.Text

If textboxinformation = 1 Then
    Age = InputBox("Please enter your age.")
    Select Case Age
        Case 0 To 3
            MsgBox "Someone this young is not authorized to book a cruise. Go find your parents!"
        Case 4 To 12
            frmAlaskanHome.Hide
            frmActivities2.Show
        Case 13 To 20
            frmAlaskanHome.Hide
            frmActivitiesTeen2.Show
        Case 21 To 110
            frmAlaskanHome.Hide
            frmActivitiesAdult2.Show
        Case Else
            MsgBox "Please enter your age between 0 to 110."
    End Select
ElseIf textboxinformation = 2 Then
    frmAlaskanHome.Hide
    frmDining2.Show
ElseIf textboxinformation = 3 Then
    frmAlaskanHome.Hide
    frmCruisePorts2.Show
ElseIf textboxinformation = 4 Then
    frmAlaskanHome.Hide
    frmRooms2.Show
ElseIf textboxinformation = 5 Then
    frmAlaskanHome.Hide
    frmFlightInformation2.Show
Else: MsgBox ("Please enter a number 1-5")
End If

End Sub

Private Sub cmdReturn6_Click()
frmHome.Show
frmAlaskanHome.Hide
End Sub
