VERSION 5.00
Begin VB.Form frmMainForm1 
   BackColor       =   &H80000002&
   Caption         =   "Airline Industry"
   ClientHeight    =   6855
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9885
   LinkTopic       =   "Form1"
   ScaleHeight     =   6855
   ScaleWidth      =   9885
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   975
      Left            =   7920
      TabIndex        =   10
      Top             =   3240
      Width           =   1815
   End
   Begin VB.PictureBox Picture1 
      Height          =   3495
      Left            =   2040
      Picture         =   "MainForm1.frx":0000
      ScaleHeight     =   3435
      ScaleWidth      =   4995
      TabIndex        =   7
      Top             =   1920
      Width           =   5055
   End
   Begin VB.CommandButton cmdInfo 
      Caption         =   "Click here for further information about the Chaise Van Air flight company"
      Height          =   975
      Left            =   7680
      TabIndex        =   3
      Top             =   5520
      Width           =   1935
   End
   Begin VB.CommandButton cmdMember 
      Caption         =   "Click here to sign up for a membership and enjoy first rate deals"
      Height          =   975
      Left            =   5280
      TabIndex        =   2
      Top             =   5520
      Width           =   2055
   End
   Begin VB.CommandButton cmdPrices 
      Caption         =   "If you want to see our good deals and compare prices, Click Here!"
      Height          =   975
      Left            =   2880
      TabIndex        =   1
      Top             =   5520
      Width           =   1935
   End
   Begin VB.CommandButton cmdDestinations 
      Caption         =   "Click here to see our available flight destinations"
      Height          =   975
      Left            =   480
      TabIndex        =   0
      Top             =   5520
      Width           =   1935
   End
   Begin VB.Label lblMembername 
      BackColor       =   &H80000002&
      Height          =   255
      Left            =   3600
      TabIndex        =   9
      Top             =   6600
      Width           =   3255
   End
   Begin VB.Label lblSlang 
      BackColor       =   &H80000002&
      Caption         =   "Don't worry this is Happy Airlines, not Chaise Van Air!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   480
      TabIndex        =   8
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Label lblName 
      BackColor       =   &H80000002&
      Caption         =   "By: Chaise VanOverbeke"
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
      Left            =   0
      TabIndex        =   6
      Top             =   6600
      Width           =   2415
   End
   Begin VB.Label lblSlogan 
      BackColor       =   &H80000002&
      Caption         =   $"MainForm1.frx":90042
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   480
      TabIndex        =   5
      Top             =   960
      Width           =   8895
   End
   Begin VB.Label lblTitle 
      BackColor       =   &H80000002&
      Caption         =   "   Chaise Van Air"
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2160
      TabIndex        =   4
      Top             =   120
      Width           =   5895
   End
End
Attribute VB_Name = "frmMainForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : Airline Option(Project1.vbp)
'Form Name : frmMainForm1(MainForm1.frm)
'Author: Chaise VanOverbeke
'Date : Friday October 25, 2005
'Purpose of the Project: To have the user interact with the program
                    'in such a way so they can cycle through the different
                    'forms and see what states and cities Chaise Van Air
                    'flies to, their various prices, to login and become a
                    'member of the program, as well as viewing further
                    'information about the company and its founder.
'purpose of the form:  The frmMainForm1 is the central form that serves as
                    'a directory that leads to the other various forms the
                    'user wishes to look at.  It is an introductory form that
                    'offers the title of the airline company and initially
                    'shows the user what this program really entails.


Private Sub cmdDestinations_Click()
    frmMainForm1.Hide
    frmDestinations.Show    'allows the user to visit the Destinations form to see what states and cities Chaise Van Air flies to.
End Sub

Private Sub cmdInfo_Click()
    frmMainForm1.Hide
    frmInformation.Show    'allows the user to visit the Information form to read furhter information about the company and its founder
End Sub

Private Sub cmdMember_Click()
    frmMainForm1.Hide
    frmMemberships.Show     'allows the user to visit the Memberships form to sign up and become a member of Chaise Van Air.
End Sub

Private Sub cmdPrices_Click()
    frmMainForm1.Hide
    frmPrices.Show     'allows the user to visit the Prices form to compare the prices of flights and deals Chaise Van Air offers
End Sub

Private Sub cmdQuit_Click()
    End     'allows the user to exit the program
End Sub

Private Sub Form_Load()
lblMembername.Caption = membername  'Everytime someone is logged in to the program, this allows the user's name to show up on this form.
End Sub
