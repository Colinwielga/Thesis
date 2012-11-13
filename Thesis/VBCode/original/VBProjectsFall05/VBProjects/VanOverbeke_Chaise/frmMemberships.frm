VERSION 5.00
Begin VB.Form frmMemberships 
   Caption         =   "Sign up for a membership"
   ClientHeight    =   7425
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9780
   LinkTopic       =   "Form1"
   ScaleHeight     =   7425
   ScaleWidth      =   9780
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdLogin 
      Caption         =   "Login"
      Height          =   975
      Left            =   840
      TabIndex        =   10
      Top             =   5400
      Width           =   2175
   End
   Begin VB.TextBox txtBoxPassword 
      Height          =   375
      Left            =   600
      TabIndex        =   6
      Top             =   4560
      Width           =   2775
   End
   Begin VB.TextBox txtBoxName 
      Height          =   375
      Left            =   600
      TabIndex        =   5
      Top             =   3720
      Width           =   2775
   End
   Begin VB.CommandButton cmdSignUp 
      Caption         =   "Click here to sign up for a membership"
      Height          =   1095
      Left            =   600
      TabIndex        =   4
      Top             =   1320
      Width           =   2295
   End
   Begin VB.CommandButton cmdBenefits 
      Caption         =   "Why should I become a member of this website?"
      Height          =   975
      Left            =   7320
      TabIndex        =   2
      Top             =   720
      Width           =   2055
   End
   Begin VB.CommandButton cmdMainMenu3 
      Caption         =   "Return to Main Menu"
      Height          =   975
      Left            =   7440
      TabIndex        =   0
      Top             =   6360
      Width           =   2175
   End
   Begin VB.Image Image1 
      Height          =   2235
      Left            =   5160
      Picture         =   "frmMemberships.frx":0000
      Top             =   2640
      Width           =   2850
   End
   Begin VB.Label lblMembername 
      Height          =   375
      Left            =   3720
      TabIndex        =   11
      Top             =   6840
      Width           =   2775
   End
   Begin VB.Label lblPassword 
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   9
      Top             =   4320
      Width           =   1935
   End
   Begin VB.Label lblexist 
      Caption         =   "If you have already signed up for a membership, login here:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   8
      Top             =   2880
      Width           =   3135
   End
   Begin VB.Label lblsignname 
      Caption         =   "Username"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   7
      Top             =   3480
      Width           =   2895
   End
   Begin VB.Label lblName 
      Caption         =   "By: Chaise VanOverbeke"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   7080
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Please fill out the following information:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   720
      Width           =   3975
   End
End
Attribute VB_Name = "frmMemberships"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : Airline Option(Project1.vbp)
'Form Name : frmMemberships(frmMemberships.frm)
'Author: Chaise VanOverbeke
'Date : Friday October 28, 2005
'Purpose of the form:  This form lets the user sign up for a membership and then
                     'login to the program, which shows then shows the user's name
                     'on each of the forms.  This form also allows the user to see
                     'why he/she should become a member and the reasons are displayed
                     'in Message Boxes.


Option Explicit
Dim Members(1 To 100) As String
Dim Passwords(1 To 100) As String
Dim CTR As Integer

Private Sub cmdBenefits_Click() 'prints three consecutive message boxes with information about benefits.
        MsgBox "Members of this website are rewarded for their loyalty by receiving e-mails of new promos and deals", , "Benefits"
        MsgBox "By becoming a member, this website will keep track of all frequent flyer miles that can be credited towards discounts or free flights", , "Another Benefit"
        MsgBox "Members will automatically be entered into random drawings for special deals, including free flights!", , "Another Benefit"
End Sub

Private Sub cmdLogin_Click()
    Dim notFound As Boolean
    Dim I As Integer
    Dim S As String
    Dim F As String
    S = txtBoxName.Text   'sets s = to txtBoxName, so when "S" is used in the code it is actually what the user typed in the text box.
    F = txtBoxPassword.Text    'sets F = to txtBoxPassword, so when "F" is used in the code it is actually what the user typed in the text box.
    I = 0
    notFound = True   'sets notFound equal to True
    Do While notFound And I < CTR
        I = I + 1
        If S = Members(I) And Len(S) > 0 And F = Passwords(I) And Len(F) > 0 Then   'Checks to see if the user typed in a username and password that are longer than 0.
            notFound = False
        End If
    Loop
    If notFound Then    'notFound = true, message box is displayed, because the user is not an authorized member
        MsgBox "You are not an authorized member.", , "Error"
    Else
        MsgBox "Congratulations, you have successfully logged in!", , "Found User"
    membername = S  'sets all the forms and their individual labels = to membername which is found in the module, so everytime the user logs into the progam from this form and then goes to any other form, his/her name will appear on that particular form.
    lblMembername.Caption = membername
    frmMainForm1.lblMembername.Caption = membername
    frmDestinations.lblMembername.Caption = membername
    frmPrices.lblMembername.Caption = membername
    frmInformation.lblMembername.Caption = membername
    End If
    
End Sub

Private Sub cmdMainMenu3_Click()
    frmMemberships.Hide
    frmMainForm1.Show
    
End Sub

Private Sub cmdSignUp_Click()
    Dim Member As String
    Dim Password As String
    CTR = CTR + 1
    
    Member = InputBox("Please enter a username that you would like to use: ")   'an input box pops up asking the user to type in a username
    If Len(Member) <> 0 Then    'checks to see if the length of the user name is more than 0 characters.
        Members(CTR) = Member   'the counter keeps track of that particular user and their username.
        Password = InputBox("Please enter a password ") 'an input box pops up asking th euser to pick a password
        If Len(Password) <> 0 Then  'checks to see if the length of the password is more than 1 character.
            Passwords(CTR) = Password   'the counter keeps track of that particular user and their password.
            MsgBox "Congratulations you have successfully logged in", , "New Member"
        Else    'the requirements for a password were not met
            MsgBox "Must type a valid password", , "Error"
        End If
    Else    'the requirements for a username were not met
        MsgBox "Must type a valid username", , "Error"
    End If
    

    
    
End Sub

Private Sub Label2_Click()

End Sub

