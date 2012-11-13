VERSION 5.00
Begin VB.Form Battersea 
   BackColor       =   &H0000FFFF&
   Caption         =   "Battersea"
   ClientHeight    =   14850
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   19080
   LinkTopic       =   "Form1"
   ScaleHeight     =   14850
   ScaleWidth      =   19080
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdquit 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9360
      TabIndex        =   11
      Top             =   10560
      Width           =   2055
   End
   Begin VB.CommandButton cmdreturn 
      Caption         =   "Return to Map of London"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   7320
      TabIndex        =   10
      Top             =   10560
      Width           =   1575
   End
   Begin VB.CommandButton cmdright 
      Caption         =   "Click Here to find if you were right ."
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
      Left            =   11760
      TabIndex        =   9
      Top             =   9000
      Width           =   2295
   End
   Begin VB.TextBox txtexpensive 
      Height          =   855
      Left            =   8880
      TabIndex        =   6
      Top             =   9000
      Width           =   2415
   End
   Begin VB.CommandButton cmdcompute 
      Caption         =   "Click Here to Find out how much Chelsea Bridge cost to be built."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   6360
      TabIndex        =   5
      Top             =   6600
      Width           =   1815
   End
   Begin VB.PictureBox picResults 
      Height          =   1335
      Left            =   8760
      ScaleHeight     =   1275
      ScaleWidth      =   6795
      TabIndex        =   4
      Top             =   6720
      Width           =   6855
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H008080FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   9240
      TabIndex        =   3
      Text            =   "This is Chelsea Bridge"
      Top             =   1200
      Width           =   3135
   End
   Begin VB.PictureBox picchelseabridge 
      Height          =   4815
      Left            =   7200
      Picture         =   "Battersea.frx":0000
      ScaleHeight     =   4755
      ScaleWidth      =   7275
      TabIndex        =   2
      Top             =   1560
      Width           =   7335
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00FF00FF&
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
      Left            =   8400
      TabIndex        =   1
      Text            =   "Click on the picture of the Bridge to learn more of it's history."
      Top             =   600
      Width           =   5415
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFF00&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7920
      TabIndex        =   0
      Text            =   "Battersea's main attraction is the Chelsea Bridge."
      Top             =   120
      Width           =   6255
   End
   Begin VB.Label Label3 
      Caption         =   "Created by Chelsey Johnson"
      Height          =   375
      Left            =   4800
      TabIndex        =   12
      Top             =   11880
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000FFFF&
      Caption         =   "Type ""Yes"" or ""No""."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7200
      TabIndex        =   8
      Top             =   9120
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FFFF&
      Caption         =   "Do you think that 88,000 pounds was a lot of money to build a bridge when Chelsea Bridge was built?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3960
      TabIndex        =   7
      Top             =   9000
      Width           =   2775
   End
End
Attribute VB_Name = "Battersea"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Dicovering London (Project1.vbp)
'Form Name: Battersea (Battersea.frm)
'Author: Chelsey Johnson
'Date Written: March,14, 2004
'Purpose of Form: The purpose of this form is to let the user come familiar with the famous cites within Battersea.
                    'It lets the user learn more about the history of Chelsea Bridge, learn how much it cost to be
                    'built, and decide if they think that was an expensive price.
'Option Explicit is a command to force
'the user to explicitly declare all variables
'before they can be used.
Option Explicit

Private Sub cmdcompute_Click()
'Prints out the cost to build the bridge so the user can then decide if it was expensive
picResults.Cls
picResults.Print "Chelsea Bridge cost 88,000 pounds to build."
End Sub

Private Sub cmdquit_Click()
End
End Sub

Private Sub cmdreturn_Click()
'Takes the user back to the Map of London page, so they may look at a new district
Battersea.Hide
MapLondon.Show
End Sub

Private Sub cmdright_Click()
Dim Expensive As String
Expensive = txtexpensive.Text 'Taking the users opinion
picResults.Cls
If Expensive = "Yes" Then  'Comparing the users opinion to my opinion
    picResults.Print "You are right!"
    picResults.Print "88,000 pounds was incrediably expensive price for the building of the bridge."
End If
If Expensive = "No" Then 'Comparing the users opinion to my opinion
    picResults.Print "Sorry, that is incorrect."
    picResults.Print "88,000 pounds was incrediably expensive price for the building of the bridge."
End If
End Sub

Private Sub picchelseabridge_Click()
'Presenting history to the user in the form of a message box, so they have to read the history before moving on
MsgBox "In 1851-8 the Chelsea Suspension Bridge was designed and built by Thomas Page, going from the base of Grosvenor Canal on the north side of the river to the edge of Battersea Park on the south side.", , "Chelsea Bridge"
MsgBox "The crossing opened in 1858 and charged a toll until 1879.", , "Chelsea Bridge"
End Sub
