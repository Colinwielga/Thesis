VERSION 5.00
Begin VB.Form Application 
   BackColor       =   &H0000C0C0&
   Caption         =   "Form1"
   ClientHeight    =   7965
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10305
   LinkTopic       =   "Form1"
   Picture         =   "Application.frx":0000
   ScaleHeight     =   7965
   ScaleWidth      =   10305
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdView 
      Caption         =   "View Form"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   0
      Left            =   5520
      Picture         =   "Application.frx":1B69C6
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4200
      Width           =   1215
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00FFFFFF&
      Height          =   4335
      Left            =   5520
      ScaleHeight     =   4275
      ScaleWidth      =   4755
      TabIndex        =   0
      Top             =   960
      Width           =   4815
      Begin VB.CommandButton cmdReturn 
         Caption         =   "Return to MN"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   3480
         Picture         =   "Application.frx":1BAF2C
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   3240
         Width           =   1335
      End
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackColor       =   &H00C00000&
      Caption         =   "Schwan's Employment Application"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   855
      Left            =   5520
      TabIndex        =   1
      Top             =   120
      Width           =   4815
   End
End
Attribute VB_Name = "Application"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Project Name: Minnesota
'Form Name: Application
'Author: Tony Blum and Danielle johnson
'Date Written: March 26th 2008
'The purpose of this form is to allow the user to view the information that they put in on the employment form
Private Sub cmdReturn_Click()
'This hides the current form and returns you to the Minnesota home page
'Brings up message box to show you that you applied for an application
MsgBox "You have successfully applied for a job at Schwan's Food Company", , "Thank You"
Application.Hide
Minnesota.Show
End Sub

Private Sub cmdView_Click(Index As Integer)
'This prints all of the data that was entered into the employment form and displays it for the user to see
picResults.Print "Name:", Tab(40); YourName
picResults.Print "Date of Birth:", Tab(40); Month; " /"; Day; "/"; Year
picResults.Print "E-Mail Address:", Tab(40); EMail
picResults.Print "Country:"; Tab(40); Country
picResults.Print "State:"; Tab(40); State
picResults.Print "City:", Tab(40); City
picResults.Print "Address:", Tab(40); Address1
picResults.Print "Address:", Tab(40); Address2
picResults.Print "Phone Number:", Tab(40); Number1; "-"; Number2; "-"; Number3; ""
picResults.Print "Personal Description:", Tab(40); Description
End Sub

