VERSION 5.00
Begin VB.Form frmUserInfo 
   BackColor       =   &H00FFC0C0&
   Caption         =   "User Information (Megan Fitzgerald)"
   ClientHeight    =   7320
   ClientLeft      =   3420
   ClientTop       =   2370
   ClientWidth     =   10020
   BeginProperty Font 
      Name            =   "Century Schoolbook"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   7320
   ScaleWidth      =   10020
   Begin VB.TextBox txtAddress 
      Height          =   855
      Left            =   1920
      TabIndex        =   4
      Top             =   6000
      Width           =   5055
   End
   Begin VB.TextBox txtEmail 
      Height          =   855
      Left            =   1920
      TabIndex        =   3
      Top             =   5040
      Width           =   5055
   End
   Begin VB.CommandButton cmdWrite 
      BackColor       =   &H00FF8080&
      Caption         =   "Send Information"
      Height          =   855
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5040
      Width           =   2055
   End
   Begin VB.TextBox txtName 
      Height          =   855
      Left            =   1920
      TabIndex        =   1
      Top             =   4080
      Width           =   5055
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FF8080&
      Caption         =   "Return to Mission Trips"
      Height          =   615
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6480
      UseMaskColor    =   -1  'True
      Width           =   2775
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Contact Amigos for Christ"
      BeginProperty Font 
         Name            =   "Baskerville Old Face"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2280
      TabIndex        =   11
      Top             =   120
      Width           =   5175
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFC0C0&
      Caption         =   $"UserInfo.frx":0000
      Height          =   735
      Left            =   720
      TabIndex        =   10
      Top             =   3360
      Width           =   8295
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFC0C0&
      Caption         =   $"UserInfo.frx":00AF
      BeginProperty Font 
         Name            =   "Baskerville Old Face"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1575
      Left            =   600
      TabIndex        =   9
      Top             =   2160
      Width           =   8895
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFC0C0&
      Caption         =   $"UserInfo.frx":01F3
      BeginProperty Font 
         Name            =   "Baskerville Old Face"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1335
      Left            =   600
      TabIndex        =   8
      Top             =   720
      Width           =   8655
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Enter your Home Address"
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   240
      TabIndex        =   7
      Top             =   5880
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Enter your Email Address"
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   240
      TabIndex        =   6
      Top             =   4920
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Enter your Name"
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   240
      TabIndex        =   5
      Top             =   4200
      Width           =   1695
   End
End
Attribute VB_Name = "frmUserInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : ProjectAmigosForChrist (Megan Fitzgerald's Visual Basic Project.vbp)
'Form Name : frmUserInfo (UserInfo.frm)
'Author: Megan Fitzgerald
'Date Written: October 28, 2003
'Purpose of This Form: To allow the user to input contact information
                        'so that they can receive information from
                        'Amigos for Christ regarding the mission trip.
                        'This information will be sent to a database that
                        'automatically saves the information and keeps track
                        'of all of the users that have entered information.

Option Explicit

Dim PATH As String, I As Integer
Dim Name1() As String, Email() As String, Address() As String

Private Sub cmdWrite_Click()

Dim NumUsers As Integer


'This will open the database for user information at this file address and
'redimension the array so that a new line of information can be added to the database.
Open PATH & "UserInformation.txt" For Input As #1
Input #1, NumUsers
ReDim Name1(NumUsers + 1)
ReDim Email(NumUsers + 1)
ReDim Address(NumUsers + 1)


For I = 1 To NumUsers
    Input #1, Name1(I), Email(I), Address(I)
Next I
Close #1

'This allows each user to input their Name, Email address, and home address
'into 3 separate text boxes.
NumUsers = NumUsers + 1
Name1(NumUsers) = txtName.Text
Email(NumUsers) = txtEmail.Text
Address(NumUsers) = txtAddress.Text

'This information will then be sent (or written) to the database at this file address.
Open PATH & "UserInformation.txt" For Output As #1
Write #1, NumUsers
For I = 1 To NumUsers
    Write #1, Name1(I), Email(I), Address(I)
Next I
Close #1

'This message will pop up after the user has clicked the "send information" button.
MsgBox "Thank you! We will be in contact with you soon!", , "Amigos for Christ"

End Sub

Private Sub Command1_Click()

'Take the user back to Homepage "Amigos for Christ".
frmUserInfo.Hide
frmMissionTrips.Show

End Sub


Private Sub Form_Load()
PATH = "N:\CS130\handin\Fitzgerald_Megan\"
End Sub

