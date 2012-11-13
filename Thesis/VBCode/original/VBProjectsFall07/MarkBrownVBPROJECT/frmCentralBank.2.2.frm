VERSION 5.00
Begin VB.Form frmCustomerIdentity 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Please Identify Customer"
   ClientHeight    =   8145
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9870
   LinkTopic       =   "Form2"
   ScaleHeight     =   8145
   ScaleWidth      =   9870
   Begin VB.CommandButton Command1 
      BackColor       =   &H000000FF&
      Caption         =   "Leave Bank"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4200
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7560
      Width           =   1935
   End
   Begin VB.CommandButton cmdReturn 
      BackColor       =   &H00008000&
      Caption         =   "Return to Member selection page"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6240
      Width           =   2295
   End
   Begin VB.CommandButton cmdViewAccount 
      BackColor       =   &H00008000&
      Caption         =   "View Your Account"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6240
      Width           =   2295
   End
   Begin VB.TextBox txtInputpw 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   120
      MaxLength       =   15
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   3000
      Width           =   2295
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H000000FF&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   18720
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   12120
      Width           =   2055
   End
   Begin VB.PictureBox picIdentity 
      BackColor       =   &H00FFFFFF&
      Height          =   4815
      Left            =   3360
      ScaleHeight     =   4755
      ScaleWidth      =   3675
      TabIndex        =   1
      Top             =   1200
      Width           =   3735
   End
   Begin VB.Label lblmember 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   1560
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Please enter your verification password:"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   735
      Left            =   120
      TabIndex        =   7
      Top             =   2040
      Width           =   2655
   End
   Begin VB.Label lblIdentity 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Verify Identity"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   495
      Left            =   4080
      TabIndex        =   0
      Top             =   360
      Width           =   2295
   End
End
Attribute VB_Name = "frmCustomerIdentity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This bank system was designed and created by Mark Brown and David Bernardy

Option Explicit

Private Sub cmdQuit_Click()
End                                                     'Exits the bank

End Sub

Private Sub cmdReturn_Click()

picIdentity.Picture = Nothing                           'Makes the picture box empty

frmCustomerIdentity.Hide
frmCustomerChoose.Show

End Sub


Private Sub cmdViewAccount_Click()

Dim password1 As String
Dim found As Boolean

password1 = txtInputpw.Text                             'Has the user enter their password
found = False

'Searches the array for the password and askes if it equals what the user entered as their password
If password1 = password(position + 1) Then
    found = True
End If

If Not found Then
    MsgBox "I'm sorry, that was the incorrect password.", , "Incorrect Password"
    Else
    frmCustomerIdentity.Hide
    frmWithdrawlsandDeposits.Show
End If
  
txtInputpw.Text = ""                                    'Sets the password input box to empty so the user can easily enter a new password if they typed it wrong
  
End Sub

Private Sub Command1_Click()
End                                                     'Exits the bank
End Sub

Private Sub Form_Load()

lblmember.Caption = firstname(position + 1) & " " & lastname(position + 1)      'Displays the member's name
picIdentity.Picture = LoadPicture(App.Path & "\" & id(position + 1))            'Displays the member's picture id for easy confirmation

End Sub

