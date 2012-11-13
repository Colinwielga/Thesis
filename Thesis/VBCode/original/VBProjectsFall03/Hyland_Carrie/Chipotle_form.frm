VERSION 5.00
Begin VB.Form frmMainMenu 
   BackColor       =   &H00FF8080&
   Caption         =   "Main Menu"
   ClientHeight    =   6690
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   9480
   FillColor       =   &H0000C000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   6690
   ScaleWidth      =   9480
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00FFFF80&
      Caption         =   "QUIT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3480
      MaskColor       =   &H00FF00FF&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5400
      Width           =   1455
   End
   Begin VB.CommandButton cmdFixings2 
      BackColor       =   &H00FF80FF&
      Caption         =   "Want to know about all the fixings?  Click here!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   4800
      MaskColor       =   &H00FF80FF&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3240
      Width           =   2775
   End
   Begin VB.CommandButton cmdCheapest 
      BackColor       =   &H00FF80FF&
      Caption         =   "Low on Cash?  To arrange choices from least to most expensive."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1200
      Width           =   2775
   End
   Begin VB.CommandButton cmdStores 
      BackColor       =   &H00FF80FF&
      Caption         =   "To See if There is a Chipotle in Your Minnesota City!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3240
      Width           =   2775
   End
   Begin VB.CommandButton cmdOrder 
      BackColor       =   &H00FF80FF&
      Caption         =   "To Order from Chipotle"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1200
      Width           =   2775
   End
   Begin VB.Label lblName 
      BackColor       =   &H00FF8080&
      Caption         =   "Designed by Carrie Hyland"
      Height          =   495
      Left            =   5640
      TabIndex        =   6
      Top             =   5400
      Width           =   2535
   End
   Begin VB.Label lblMainMenu 
      BackColor       =   &H00FF8080&
      Caption         =   "Welcome to CHIPOTLE!  Please click a button below according to your desired action!  Enjoy!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   855
      Left            =   1080
      TabIndex        =   4
      Top             =   120
      Width           =   6135
   End
End
Attribute VB_Name = "frmMainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: ProjChipotleOrder (Carrie Hyland's VB Project.vbp)
'Form Name: frmMainMenu (MainMenu_form.frm)
'Author: Carrie Hyland
'Date Written: October 19, 2003
'Purpose of Project: To give the user many information about the Chipotle restaurant.
                    'To allow the user a fun interaction between several different
                    'options: order, arrange in order from least to greatest, see if there
                    'is a store in your area and see all the fixings that can go
                    'on the Chipotle products.
'Purpose of Form: To give the user several options as to what
                 ' they desire to do by directing them to another form.
                 ' It serves as a starting place for the user.
                 
                 
'Option Explicit is a command to force
'the user to explicitly declare all variables
'before they can be used.
Option Explicit


Private Sub cmdCheapest_Click()
'Hides the frmMainMenu and shows the frmCheap (switches from
'the main menu form to the cheap form).
frmMainMenu.Hide
frmCheap.Show
End Sub

Private Sub cmdFixings2_Click()
'Hides the frmMainMenu and shows the frmFixings (switches from
'the main menu form to the fixings form).
frmMainMenu.Hide
frmFixings.Show
End Sub

Private Sub cmdOrder_Click()
'Hides the frmMainMenu and shows the frmOrder (switches from
'the main menu form to the Order form).
frmMainMenu.Hide
frmOrder.Show
End Sub

Private Sub cmdQuit_Click()
'Ends the program
End
End Sub

Private Sub cmdStores_Click()
'Hides the frmMainMenu and shows the frmStores (switches from
'the main menu form to the stores form).
frmMainMenu.Hide
frmStores.Show
End Sub
