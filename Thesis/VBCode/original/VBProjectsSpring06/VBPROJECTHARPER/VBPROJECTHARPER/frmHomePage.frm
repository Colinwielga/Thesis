VERSION 5.00
Begin VB.Form frmHomePage 
   BackColor       =   &H000080FF&
   Caption         =   "Buy Instruments Online"
   ClientHeight    =   10470
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14040
   LinkTopic       =   "Form1"
   ScaleHeight     =   10470
   ScaleMode       =   0  'User
   ScaleWidth      =   140
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmddrawing 
      Caption         =   "Enter Me in the drawing>>>"
      Height          =   255
      Left            =   4080
      TabIndex        =   9
      Top             =   6840
      Width           =   2175
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Yamaha Electric"
      Height          =   255
      Left            =   9000
      TabIndex        =   7
      Top             =   10560
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Pulse X Kit"
      Height          =   255
      Left            =   9000
      TabIndex        =   6
      Top             =   8400
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ludwig Ultra Series"
      Height          =   255
      Left            =   8880
      TabIndex        =   5
      Top             =   6480
      Width           =   1935
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   8895
      Left            =   3720
      TabIndex        =   4
      Top             =   2160
      Width           =   255
   End
   Begin VB.ComboBox menuopt 
      Height          =   315
      ItemData        =   "frmHomePage.frx":0000
      Left            =   4320
      List            =   "frmHomePage.frx":000D
      TabIndex        =   3
      Text            =   "Go to..."
      Top             =   2160
      Width           =   1935
   End
   Begin VB.CommandButton cmdAcc 
      BackColor       =   &H00FF00FF&
      Caption         =   "Shop Accessories"
      Height          =   1215
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5640
      Width           =   3615
   End
   Begin VB.CommandButton cmdCymbals 
      BackColor       =   &H000080FF&
      Caption         =   "Shop Cymbals"
      Height          =   1215
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4080
      Width           =   3615
   End
   Begin VB.CommandButton cmdDrums 
      BackColor       =   &H0000FFFF&
      Caption         =   "Shop Drums"
      Height          =   1215
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2520
      Width           =   3615
   End
   Begin VB.Label Label5 
      BackColor       =   &H000000C0&
      Caption         =   "By: Ben Harper"
      Height          =   255
      Left            =   9600
      TabIndex        =   13
      Top             =   0
      Width           =   1215
   End
   Begin VB.Label DMB 
      BackColor       =   &H0000FFFF&
      Caption         =   "Is Carter Beauford the Messiah?"
      BeginProperty Font 
         Name            =   "Myriad Pro Cond"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   12
      Top             =   9960
      Width           =   2655
   End
   Begin VB.Image Carter 
      Height          =   3390
      Left            =   240
      Picture         =   "frmHomePage.frx":002E
      Top             =   6840
      Width           =   3240
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000006&
      FillStyle       =   0  'Solid
      Height          =   8295
      Left            =   0
      Top             =   2160
      Width           =   3735
   End
   Begin VB.Label Label2 
      BackColor       =   &H000080FF&
      Caption         =   "TOP SELLERS"
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
      Left            =   9000
      TabIndex        =   11
      Top             =   4320
      Width           =   1815
   End
   Begin VB.Image Image20 
      Height          =   780
      Left            =   9000
      Picture         =   "frmHomePage.frx":23C80
      Top             =   1320
      Width           =   1800
   End
   Begin VB.Image Cart 
      Height          =   1395
      Left            =   9360
      Picture         =   "frmHomePage.frx":285E2
      Top             =   2760
      Width           =   1170
   End
   Begin VB.Label Label1 
      BackColor       =   &H000080FF&
      Caption         =   "My Cart"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9360
      TabIndex        =   10
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Image Mag 
      Height          =   2250
      Left            =   4920
      Picture         =   "frmHomePage.frx":2DBE0
      Stretch         =   -1  'True
      Top             =   8400
      Width           =   2505
   End
   Begin VB.Image Image4 
      Height          =   2295
      Left            =   4080
      Picture         =   "frmHomePage.frx":4B682
      Stretch         =   -1  'True
      Top             =   7080
      Width           =   2670
   End
   Begin VB.Label lblCase 
      BackColor       =   &H0080FF80&
      Caption         =   $"frmHomePage.frx":79524
      Height          =   1575
      Left            =   4080
      TabIndex        =   8
      Top             =   5280
      Width           =   2175
   End
   Begin VB.Image Image2 
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   1800
      Left            =   6240
      Picture         =   "frmHomePage.frx":795FB
      Stretch         =   -1  'True
      Top             =   5280
      Width           =   1800
   End
   Begin VB.Image Image15 
      Height          =   3300
      Left            =   3960
      Picture         =   "frmHomePage.frx":811FD
      Top             =   2520
      Width           =   4815
   End
   Begin VB.Image Image16 
      Height          =   480
      Left            =   8400
      Picture         =   "frmHomePage.frx":B4EAF
      Top             =   600
      Width           =   1800
   End
   Begin VB.Image Image14 
      Height          =   1800
      Left            =   2640
      Picture         =   "frmHomePage.frx":B7BF1
      Top             =   360
      Width           =   1800
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000FF&
      X1              =   0
      X2              =   107.692
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Image Image13 
      Height          =   1800
      Left            =   7080
      Picture         =   "frmHomePage.frx":C24F3
      Top             =   120
      Width           =   1380
   End
   Begin VB.Image Image12 
      Height          =   780
      Left            =   480
      Picture         =   "frmHomePage.frx":CA695
      Top             =   120
      Width           =   1800
   End
   Begin VB.Image Image11 
      Height          =   270
      Left            =   2880
      Picture         =   "frmHomePage.frx":CEFF7
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image10 
      Height          =   555
      Left            =   5040
      Picture         =   "frmHomePage.frx":D0989
      Top             =   600
      Width           =   1800
   End
   Begin VB.Image Image9 
      Height          =   285
      Left            =   4800
      Picture         =   "frmHomePage.frx":D3DD3
      Top             =   120
      Width           =   1800
   End
   Begin VB.Image Image8 
      Height          =   585
      Left            =   4680
      Picture         =   "frmHomePage.frx":D58CD
      Top             =   1320
      Width           =   1800
   End
   Begin VB.Image Image7 
      Height          =   1800
      Left            =   9120
      Picture         =   "frmHomePage.frx":D8FE7
      Top             =   8760
      Visible         =   0   'False
      Width           =   1635
   End
   Begin VB.Image Image6 
      Height          =   1575
      Left            =   9000
      Picture         =   "frmHomePage.frx":E29E9
      Top             =   6840
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.Image Image5 
      Height          =   1455
      Left            =   9000
      Picture         =   "frmHomePage.frx":EBDD3
      Top             =   5040
      Width           =   1800
   End
   Begin VB.Line Line1 
      X1              =   37.094
      X2              =   37.094
      Y1              =   2160
      Y2              =   8640
   End
   Begin VB.Image Image1 
      Height          =   1245
      Left            =   120
      Picture         =   "frmHomePage.frx":F467D
      Top             =   840
      Width           =   1800
   End
End
Attribute VB_Name = "frmHomePage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Buy Drums Online (OnlineDrums.vbp)
'frmHomePage (frmHomPage)
'Ben Harper
'3/23/06
'This project allows the user to buy drums and all drum accessories online.
'The user can also subscribe for magazines and ebter into drawings.
'This program allows the user to purchase all his drums quickly over the internet.
'The Module is used to save the totals for each drum sub-section (cymbals, drums, and accessories)
'so they can be put in a grand total inthe cart form.







Private Sub Carter_Click()
      X = InputBox("On a Scale of 1 - 10, How do you rate Carter?", "Rate Carter!")  'asks for value 1 - 10, use If-Then statement to output different results for input
If X <= 3 Then
    MsgBox "You should listen to some real music you dirty coward!", , "FOOL!"
Else
    If X <= 7 Then
        MsgBox "Listen to the drum solo on Holloween, then come back", , "So-So"
    Else
        If X <= 10 Then
            MsgBox "You are worthy of Ben Harper's presence... proceed with life", , "Sweet"
        Else
            MsgBox "Invalid Entry", "Invalid"
        End If
    End If
End If
End Sub

Private Sub cmdAcc_Click()
frmHomePage.Visible = False
frmAccessories.Visible = True
End Sub

Private Sub cmdCymbals_Click()
frmHomePage.Visible = False
frmCymbals.Visible = True
End Sub

Private Sub cmddrawing_Click()
    X = InputBox("please enter your receipt number", "receipt number")    'asks for a receipt number
    X1 = InputBox("please enter your last name followed by telephone number (e.g. Larson, (651)983-6313)", "Contestant information)") 'gathers info about user
    MsgBox "Thank you for your purchase and Good Luck in next months drawing", , "You've been entered!"
End Sub



Private Sub cmdDrums_Click()
frmHomePage.Visible = False
frmDrums.Visible = True
End Sub


Private Sub Command1_Click()
Image5.Visible = True
Image6.Visible = False
Image7.Visible = False
End Sub

Private Sub Command2_Click()
Image5.Visible = False
Image6.Visible = True
Image7.Visible = False
End Sub

Private Sub Command3_Click()
Image5.Visible = False
Image6.Visible = False
Image7.Visible = True
End Sub


Private Sub Mag_Click()
    Dim Found As Boolean
    Y = InputBox("Please enter your full name", "Name")
    X = InputBox("Would you like to receive 3 issues of Drum Magazine Free?", "Accept?") 'asks user for Yes/No input
    If X = "Yes" Then
        Found = True
            If X = "No" Then
                Found = False
            End If
    End If
    If Found = True Then          'if user types Yes
        X = InputBox("Please enter mailing address", "Delivery Adress")
        MsgBox "You will receive your first issue shortly", , "thank you"
    End If
    If Found = False Then         'if user types no
         MsgBox "you are a damn fool " & Y, , "you idiot!"
    End If
End Sub

Private Sub Cart_Click()
frmHomePage.Visible = False
frmCart.Visible = True

End Sub


Private Sub Magazine_Click()    'same as above
 Dim Found As Boolean
    Y = InputBox("Please enter your full name", "Name")
    X = InputBox("Would you like to receive 3 issues of Drum Magazine Free?", "Accept?")
    If X = "Yes" Then
        Found = True
            If X = "No" Then
                Found = False
            End If
    End If
    If Found = True Then
        X = InputBox("Please enter mailing address", "Delivery Adress")
        MsgBox "You will receive your first issue shortly", , "thank you"
    End If
    If Found = False Then
         MsgBox "you are a damn fool " & Y, , "you idiot!"
    End If
End Sub


