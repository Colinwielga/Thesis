VERSION 5.00
Begin VB.Form frmRegister 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Please Register With Games.Or.Cutt Today!"
   ClientHeight    =   6495
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11370
   LinkTopic       =   "Form1"
   ScaleHeight     =   6495
   ScaleWidth      =   11370
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtPassword 
      Height          =   285
      Left            =   2640
      TabIndex        =   17
      Top             =   3360
      Width           =   3015
   End
   Begin VB.CommandButton cmdEnter 
      Caption         =   "Enter Now!"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   14
      Top             =   5160
      Width           =   3255
   End
   Begin VB.CommandButton cmdSubscribe 
      Caption         =   "Subscribe Now!"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5280
      TabIndex        =   13
      Top             =   5160
      Width           =   3255
   End
   Begin VB.TextBox txtEmail 
      Height          =   285
      Left            =   2280
      TabIndex        =   11
      Top             =   2880
      Width           =   3615
   End
   Begin VB.ComboBox ComboCountry 
      Height          =   315
      ItemData        =   "frmRegister.frx":0000
      Left            =   4680
      List            =   "frmRegister.frx":000D
      TabIndex        =   10
      Text            =   "Please Choose a Country"
      Top             =   2400
      Width           =   2535
   End
   Begin VB.TextBox txtPostal 
      Height          =   285
      Left            =   2040
      TabIndex        =   9
      Top             =   2400
      Width           =   1935
   End
   Begin VB.ComboBox ComboState 
      Height          =   315
      ItemData        =   "frmRegister.frx":0030
      Left            =   4560
      List            =   "frmRegister.frx":00CD
      TabIndex        =   8
      Text            =   "Please Select a State"
      Top             =   1920
      Width           =   2535
   End
   Begin VB.TextBox txtCity 
      Height          =   285
      Left            =   1080
      TabIndex        =   7
      Top             =   1920
      Width           =   2655
   End
   Begin VB.TextBox txtAddress 
      Height          =   285
      Left            =   1440
      TabIndex        =   5
      Top             =   1440
      Width           =   3015
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   1320
      TabIndex        =   4
      Top             =   960
      Width           =   3015
   End
   Begin VB.Label lblPassword 
      BackColor       =   &H000000FF&
      Caption         =   "Choose Password:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   16
      Top             =   3360
      Width           =   2295
   End
   Begin VB.Label lblOr 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "OR"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4080
      TabIndex        =   15
      Top             =   5040
      Width           =   855
   End
   Begin VB.Label lblDetails 
      BackColor       =   &H000000FF&
      Caption         =   "Please provide details below:"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   12
      Top             =   240
      Width           =   5775
   End
   Begin VB.Label lblCity 
      BackColor       =   &H000000FF&
      Caption         =   "City:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label lblName 
      BackColor       =   &H000000FF&
      Caption         =   "Name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   960
      Width           =   855
   End
   Begin VB.Label lblAddress 
      BackColor       =   &H000000FF&
      Caption         =   "Address:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label lblPostal 
      BackColor       =   &H000000FF&
      Caption         =   "Postal Code:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Label lblEmail 
      BackColor       =   &H000000FF&
      Caption         =   "Email Address:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   2880
      Width           =   1815
   End
   Begin VB.Image Image1 
      Height          =   6705
      Left            =   0
      Picture         =   "frmRegister.frx":030D
      Top             =   0
      Width           =   8745
   End
End
Attribute VB_Name = "frmRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Chris Orcutt
'frmRegister
'26 March 2007

Option Explicit
'This form asks users to enter details such as name, address, city and country
'before asking users to enter a valid email and unique user name to gain access
'access to the rest of the program.
Private Sub cmdEnter_Click()
    Name1 = txtName.Text
    Password = txtPassword.Text
    Address = txtAddress.Text
    City = txtCity.Text
    Postal = txtPostal.Text
    Email = txtEmail.Text
    Country = ComboCountry
    State = ComboState
    Email1 = txtEmail.Text
    
    frmRegister.Hide
    MsgBox "Thank You for Registering!", , "Thank You"
    Email = InputBox("Please enter email address", "Verify Email")
    If Email = Email1 Then
        Password1 = InputBox("Please enter user password", "Verify Account")
    Else
        MsgBox "Sorry, You Entered an Incorrect User Name", , "Error Login"
        frmMain.Show
        
    End If
    If Password = Password1 Then
            MsgBox "Congratulations!", , "Login Successful"
            frmSelectWant.Show
    End If
End Sub
Private Sub cmdSubscribe_Click()
    frmRegister.Hide                    'Hides Register form
    frmSubscribe.Show                   'Shows Subscribe form
    MsgBox "Please Fill in Each Form!!" 'Alerts user to fill in each form carefully
End Sub

