VERSION 5.00
Begin VB.Form frmHomepage 
   BackColor       =   &H000000C0&
   Caption         =   "HomePage"
   ClientHeight    =   8175
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10830
   LinkTopic       =   "Form1"
   ScaleHeight     =   8175
   ScaleWidth      =   10830
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGear 
      BackColor       =   &H0000FFFF&
      Caption         =   "Shop for Baseball Gear "
      Height          =   1215
      Left            =   3600
      TabIndex        =   5
      Top             =   4320
      Width           =   3735
   End
   Begin VB.CommandButton cmdDrawing 
      Caption         =   "Click to enter the drawing!"
      Height          =   255
      Left            =   3600
      TabIndex        =   2
      Top             =   6960
      Width           =   2655
   End
   Begin VB.CommandButton cmdMerchandise 
      BackColor       =   &H0000FFFF&
      Caption         =   "Shop for 2006 Cardinals Championship Merchandise"
      Height          =   1215
      Left            =   3600
      TabIndex        =   0
      Top             =   3000
      Width           =   3735
   End
   Begin VB.Label Label3 
      BackColor       =   &H000000C0&
      Caption         =   "We sell Cardinals 2006 World Series Championship gear as well as various MLB Baseball gear."
      BeginProperty Font 
         Name            =   "Myriad Condensed Web"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   1575
      Left            =   4560
      TabIndex        =   9
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackColor       =   &H000000C0&
      Caption         =   "Welcome to MLB Online.  "
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3600
      TabIndex        =   8
      Top             =   240
      Width           =   4815
   End
   Begin VB.Label Label1 
      Caption         =   "Click on Albert Pujols to Rate Him!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   0
      TabIndex        =   7
      Top             =   6600
      Width           =   3015
   End
   Begin VB.Label Label5 
      BackColor       =   &H000000C0&
      Caption         =   "Click on Cart to View Purchasing Total"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8400
      TabIndex        =   6
      Top             =   2040
      Width           =   2295
   End
   Begin VB.Label Label4 
      BackColor       =   &H000000C0&
      Caption         =   "My Cart"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9240
      TabIndex        =   4
      Top             =   240
      Width           =   975
   End
   Begin VB.Image cmdCart 
      Height          =   1395
      Left            =   9120
      Picture         =   "frmHomepage.frx":0000
      Top             =   600
      Width           =   1170
   End
   Begin VB.Label lblMagazine 
      BackColor       =   &H00FFFF00&
      Caption         =   "Subscribe to Sports Illustrated Magazine"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   6720
      TabIndex        =   3
      Top             =   5880
      Width           =   1935
   End
   Begin VB.Image imgMag 
      Height          =   2985
      Left            =   7800
      Picture         =   "frmHomepage.frx":55FE
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   2295
   End
   Begin VB.Label lblDrawing 
      BackColor       =   &H00FFFFFF&
      Caption         =   $"frmHomepage.frx":224A8
      Height          =   1215
      Left            =   3600
      TabIndex        =   1
      Top             =   5760
      Width           =   2655
   End
   Begin VB.Image imgAlbert 
      Height          =   2595
      Left            =   120
      Picture         =   "frmHomepage.frx":2257F
      Stretch         =   -1  'True
      Top             =   3960
      Width           =   2760
   End
   Begin VB.Image imgMLBlogo 
      Height          =   3240
      Left            =   240
      Picture         =   "frmHomepage.frx":276C5
      Top             =   120
      Width           =   3240
   End
End
Attribute VB_Name = "frmHomepage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Buy Major League Gear Online (MLBonline.vbp)
'frmHomePage (frmHomePage.frm)
'Chris Van Guilder and Pete Steele, 11/2/2006
'This project allows the user to buy Major League gear and Cardinals Championship Merchandise online.
'The user can also subscribe for magazines and enter into drawings for free tickets.
'This program allows the user to purchase all Cardinals Merchandise and other Baseball gear quickly over the internet.
'The Module is used to save the totals for each sub-section (Merchandise, Gear)
'so they can be put in a grand total in the cart form.
Option Explicit
Private Sub imgAlbert_Click()
    Dim X As Integer
    X = InputBox("On a Scale of 1 - 10 (Ten being the best and One being the worst) How do you rate Albert Pujols?", "Rate Albert")  'asks for value 1 - 10, use Case Select statement to output different results for input
    Select Case X
        Case 1 To 3
            MsgBox "You think he is overrated!?", , "No Way!"
        Case 4 To 7
            MsgBox "Consider this: He did bring the Cardinals to a World Series Championship.", , "So-So"
        Case 7 To 10
            MsgBox "You must know your Baseball very well.", , "Excellent!"
        Case Else
            MsgBox "Invalid Entry", , "Invalid"
    End Select
End Sub

Private Sub cmdDrawing_Click()
    Dim X As String
    Dim X1 As String
    X = InputBox("please enter your receipt number", "Purchase receipt number")    'asks for a receipt number
    X1 = InputBox("please enter your last name followed by telephone number (e.g. Johnson, (363)555-5555)", "Contestant information)") 'gathers info about user
    MsgBox "Thank you for your purchase and good luck in the drawing", , "You've been entered!"
End Sub


Private Sub imgMag_Click()
    Dim Found As Boolean
    Dim Y As String
    Dim X As String
    Y = InputBox("Please enter your full name", "Name")
    X = InputBox("Would you like to receive 3 free issues of Sports Illustrated Free?", "Accept?") 'asks user for Yes/No input
    If X = "Yes" Then
        Found = True
            If X = "No" Then
                Found = False
            End If
    End If
    If Found = True Then          'if user types Yes
        X = InputBox("Please enter mailing address", "Delivery Adress")
        MsgBox "You will receive your first issue shortly", , "Subscription Validated"
    End If
    If Found = False Then         'if user types No
         MsgBox "That's Unfortunate. The offer still stands " & Y, , "Subscription Denied"
    End If
End Sub

Private Sub cmdCart_Click()
    frmHomepage.Visible = False
    frmCart.Visible = True

End Sub
Private Sub cmdMerchandise_Click()
    frmHomepage.Visible = False
    frmMerchandise.Visible = True
End Sub

Private Sub cmdGear_Click()
    frmHomepage.Visible = False
    frmGear.Visible = True
End Sub



Private Sub imgMLBlogo_Click()
    MsgBox "Please don't touch me.", , "Ouch!"
End Sub
