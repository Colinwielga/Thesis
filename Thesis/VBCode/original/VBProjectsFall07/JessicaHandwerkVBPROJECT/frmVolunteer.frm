VERSION 5.00
Begin VB.Form frmVolunteer 
   BackColor       =   &H80000012&
   Caption         =   "Volunteer"
   ClientHeight    =   7335
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9630
   LinkTopic       =   "Form1"
   ScaleHeight     =   7335
   ScaleWidth      =   9630
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picSearch 
      Height          =   2895
      Left            =   240
      ScaleHeight     =   2835
      ScaleWidth      =   3795
      TabIndex        =   21
      Top             =   4200
      Width           =   3855
   End
   Begin VB.CommandButton cmdSearchVolunteer 
      Caption         =   "Search Volunteer Information"
      Height          =   615
      Left            =   4440
      TabIndex        =   20
      Top             =   5400
      Width           =   2295
   End
   Begin VB.CommandButton cmdSubmit 
      Caption         =   "Submit Information!"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Rockwell Extra Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   4440
      TabIndex        =   17
      Top             =   4080
      Width           =   4695
   End
   Begin VB.TextBox txtEmail 
      Height          =   285
      Left            =   4680
      TabIndex        =   10
      Top             =   3450
      Width           =   3615
   End
   Begin VB.TextBox txtZip 
      Height          =   375
      Left            =   1080
      TabIndex        =   9
      Top             =   2760
      Width           =   1095
   End
   Begin VB.TextBox txtPhone 
      Height          =   285
      Left            =   960
      TabIndex        =   8
      Top             =   3450
      Width           =   2055
   End
   Begin VB.TextBox txtAge 
      Height          =   495
      Left            =   4560
      TabIndex        =   7
      Top             =   840
      Width           =   615
   End
   Begin VB.TextBox txtState 
      Height          =   375
      Left            =   3600
      TabIndex        =   6
      Top             =   2160
      Width           =   615
   End
   Begin VB.TextBox txtCity 
      Height          =   375
      Left            =   1080
      TabIndex        =   5
      Top             =   2160
      Width           =   1455
   End
   Begin VB.TextBox txtAddress 
      Height          =   375
      Left            =   1080
      TabIndex        =   4
      Top             =   1560
      Width           =   3615
   End
   Begin VB.TextBox txtName 
      Height          =   495
      Left            =   1080
      TabIndex        =   3
      Top             =   840
      Width           =   2535
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8520
      TabIndex        =   1
      Top             =   6480
      Width           =   1095
   End
   Begin VB.CommandButton cmdMenu 
      Caption         =   "Main Menu"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7440
      TabIndex        =   0
      Top             =   6480
      Width           =   1095
   End
   Begin VB.Label lblVolunteer 
      BackColor       =   &H80000012&
      Caption         =   "Volunteer"
      BeginProperty Font 
         Name            =   "Rockwell Extra Bold"
         Size            =   18
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   3360
      TabIndex        =   19
      Top             =   120
      Width           =   4215
   End
   Begin VB.Image Image3 
      Height          =   2295
      Left            =   5640
      Picture         =   "frmVolunteer.frx":0000
      Stretch         =   -1  'True
      Top             =   720
      Width           =   3330
   End
   Begin VB.Image Image2 
      Height          =   1695
      Left            =   6120
      Top             =   480
      Width           =   2295
   End
   Begin VB.Image Image1 
      Height          =   2655
      Left            =   5520
      Top             =   360
      Width           =   3735
   End
   Begin VB.Label lblEmail 
      BackColor       =   &H80000012&
      Caption         =   "Email"
      BeginProperty Font 
         Name            =   "Rockwell Extra Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   3840
      TabIndex        =   18
      Top             =   3480
      Width           =   855
   End
   Begin VB.Label lblPhone 
      BackColor       =   &H80000012&
      Caption         =   "Phone"
      BeginProperty Font 
         Name            =   "Rockwell Extra Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   120
      TabIndex        =   16
      Top             =   3480
      Width           =   1095
   End
   Begin VB.Label lblZip 
      BackColor       =   &H80000012&
      Caption         =   "Zip Code"
      BeginProperty Font 
         Name            =   "Rockwell Extra Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   120
      TabIndex        =   15
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label lblState 
      BackColor       =   &H80000012&
      Caption         =   "State"
      BeginProperty Font 
         Name            =   "Rockwell Extra Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   735
      Left            =   3000
      TabIndex        =   14
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Label lblCity 
      BackColor       =   &H80000012&
      Caption         =   "City"
      BeginProperty Font 
         Name            =   "Rockwell Extra Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label lblAge 
      BackColor       =   &H80000012&
      Caption         =   "Age"
      BeginProperty Font 
         Name            =   "Rockwell Extra Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   4080
      TabIndex        =   12
      Top             =   960
      Width           =   975
   End
   Begin VB.Label lblAddress 
      BackColor       =   &H80000012&
      Caption         =   "Address"
      BeginProperty Font 
         Name            =   "Rockwell Extra Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label lblName 
      BackColor       =   &H80000012&
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "Rockwell Extra Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   1455
   End
End
Attribute VB_Name = "frmVolunteer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_Click()
End
End Sub

Private Sub cmdMenu_Click()
    frmVolunteer.Hide
    frmMenu.Show
End Sub

Private Sub cmdSearchVolunteer_Click()
picSearch.Cls

Dim Age(1 To 100) As String
Dim Name(1 To 100) As String
Dim Address(1 To 100) As String
Dim City(1 To 100) As String
Dim State(1 To 100) As String
Dim Phone(1 To 100) As String
Dim Zip(1 To 100) As String
Dim Email(1 To 100) As String
Dim ctr As Integer
Dim L As String
Dim Found As Boolean
Dim K As Integer

ctr = 0

Open App.Path & "\Volunteers.txt" For Input As #1
Do Until EOF(1)
    ctr = ctr + 1
    Input #1, Name(ctr), Age(ctr), Address(ctr), City(ctr), State(ctr), Phone(ctr), Zip(ctr), Email(ctr)
    Loop
Close #1
L = InputBox("What name would you like to search for?", "Search")
'searches the arrays to find volunteer name
K = 0
Found = False
Do While ((Not Found) And (K < ctr))
K = K + 1
If L = Name(K) Then Found = True
Loop
If (Found) Then
    picSearch.Print "Name: ", Name(K)
    picSearch.Print "Age: ", Age(K)
    picSearch.Print "Address: ", Address(K)
    picSearch.Print "City: ", City(K)
    picSearch.Print "State: ", State(K)
    picSearch.Print "Zip: ", Zip(K)
    picSearch.Print "Phone: ", Phone(K)
    picSearch.Print "Email: ", Email(K)
Else
    Beep
    MsgBox ("No Volunteer Found")
End If

End Sub

Private Sub cmdSubmit_Click()
Dim Age As String
Dim Name As String
Dim Address As String
Dim City As String
Dim State As String
Dim Phone As String
Dim Zip As String
Dim Email As String
Dim J As Integer
J = MsgBox("Thank You For Volunteering With Habitat!", vbExclamation)

Name = txtName.Text
Age = txtAge.Text
Address = txtAddress.Text
City = txtCity.Text
State = txtState.Text
Phone = txtPhone.Text
Zip = txtZip.Text
Email = txtEmail.Text

Open App.Path & "\Volunteers.txt" For Append As #1
'puts the new names in a the volunteers.txt file

Print #1, Name & "," & Age & "," & Address & "," & City & "," & State & ", " & Zip & ", " & Phone & ", " & Email
Close #1

End Sub

Private Sub txtEmail_Change()
cmdSubmit.Enabled = True
End Sub
