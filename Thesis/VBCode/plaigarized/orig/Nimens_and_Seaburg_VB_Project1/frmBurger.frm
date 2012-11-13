VERSION 5.00
Begin VB.Form frmBurger 
   BackColor       =   &H00FFFFC0&
   Caption         =   "Form1"
   ClientHeight    =   11895
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   14715
   LinkTopic       =   "Form1"
   ScaleHeight     =   11895
   ScaleWidth      =   14715
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClub 
      BackColor       =   &H00FF8080&
      Caption         =   "5-8 Club Menu"
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
      Left            =   11040
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4680
      Width           =   2055
   End
   Begin VB.CommandButton cmdMatt 
      BackColor       =   &H0080C0FF&
      Caption         =   "Matt's Bar Menu"
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
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4680
      Width           =   2175
   End
   Begin VB.CommandButton cmdOrder 
      BackColor       =   &H008080FF&
      Caption         =   "Alphabetical Order"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   8040
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3240
      Width           =   1695
   End
   Begin VB.CommandButton cmdPrice 
      BackColor       =   &H0000FF00&
      Caption         =   "Highest Price First"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3240
      Width           =   1815
   End
   Begin VB.PictureBox picResults 
      Height          =   6735
      Left            =   6000
      ScaleHeight     =   6675
      ScaleWidth      =   3675
      TabIndex        =   2
      Top             =   4560
      Width           =   3735
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   10800
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   10200
      Width           =   1335
   End
   Begin VB.CommandButton cmdGoBack 
      BackColor       =   &H0000FFFF&
      Caption         =   "Go Back"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3120
      Picture         =   "frmBurger.frx":0000
      TabIndex        =   0
      Top             =   10080
      Width           =   1335
   End
   Begin VB.Image imgClub 
      BorderStyle     =   1  'Fixed Single
      Height          =   4005
      Left            =   9960
      Picture         =   "frmBurger.frx":1EB44
      Stretch         =   -1  'True
      Top             =   5520
      Width           =   4440
   End
   Begin VB.Image imgMatts 
      BorderStyle     =   1  'Fixed Single
      Height          =   3885
      Left            =   1320
      Picture         =   "frmBurger.frx":3D688
      Stretch         =   -1  'True
      Top             =   5640
      Width           =   4560
   End
   Begin VB.Label lblClub 
      BackColor       =   &H00FF8080&
      Caption         =   "5-8 Club "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   10320
      TabIndex        =   9
      Top             =   3360
      Width           =   3375
   End
   Begin VB.Label lblMatt 
      BackColor       =   &H0080C0FF&
      Caption         =   "Matt's Bar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   2160
      TabIndex        =   8
      Top             =   3360
      Width           =   2775
   End
   Begin VB.Label lblBurger 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Birthplace of the Juicy Lucy in Minneapolis Minnesota"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   3600
      TabIndex        =   5
      Top             =   1200
      Width           =   11055
   End
End
Attribute VB_Name = "frmBurger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim pass As Integer, comp As Integer
'Man vs. Food
'frmBurger
'Ty Nimens and Josh Seaburg
'February 2010
'information about two resturants and have buttons that display their menu's and have two other buttons that shows their prices from ascending order and alphabetical order


Private Sub cmdClub_Click()
' This code fills the array and prints the menu items and prices

    X = 0
    picResults.Cls
    picResults.Print "5-8 Club Menu"
    picResults.Print "------------------------------------"
    picResults.Print "Item"; Tab(25); "Cost"
    picResults.Print "===================================="
    Open App.Path & "\5-8club.txt" For Input As #1
    Do While Not EOF(1)
    X = X + 1
    Input #1, ClubItem(X), ClubCost(X)
    picResults.Print ClubItem(X); Tab(25); FormatCurrency(ClubCost(X))
    
    Loop
    
    Close #1
    
   ' this code makes the 5-8 clubs picture show and matts picture not shown
    cmdPrice.Enabled = True
    cmdOrder.Enabled = True
    

    imgMatts.Visible = False
    imgClub.Visible = True
    
End Sub

Private Sub cmdGoback_Click()
    frmBurger.Hide
    frmMap.Show
End Sub

Private Sub cmdMatt_Click()
'This Code fills the array for matts bars menu and also prints the items and prices on the page
    I = 0
    picResults.Cls
    picResults.Print "Matt's Menu"
    picResults.Print "------------------------------------"
    picResults.Print "Item"; Tab(25); "Cost"
    picResults.Print "===================================="
    Open App.Path & "\matts.txt" For Input As #1
    Do While Not EOF(1)
    I = I + 1
    Input #1, MattItem(I), MattCost(I)
    picResults.Print MattItem(I); Tab(25); FormatCurrency(MattCost(I))
    
    Loop
    
    Close #1
' This code makes pictures of the juicy lucy show and not shown when dealing with ther respective menu.

    cmdPrice.Enabled = True
    cmdOrder.Enabled = True
    
    imgClub.Visible = False
    imgMatts.Visible = True
End Sub
   
Private Sub cmdOrder_Click()
          
    imgClub.Visible = True
    imgMatts.Visible = True
    
    picResults.Cls
'Matt's menu sorted by name
    picResults.Print "Matt's Menu"
    picResults.Print "------------------------------------"
    picResults.Print "Item"; Tab(25); "Cost"
    picResults.Print "===================================="
    
    For pass = 1 To I - 1
        For comp = 1 To I - pass
    
    If MattItem(comp) > MattItem(comp + 1) Then
            PosItem = MattItem(comp)
            MattItem(comp) = MattItem(comp + 1)
            MattItem(comp + 1) = PosItem
            
            PosCost = MattCost(comp)
            MattCost(comp) = MattCost(comp + 1)
            MattCost(comp + 1) = PosCost
            
    End If
    Next comp
Next pass
For CTR = 1 To I
    picResults.Print MattItem(CTR); Tab(25); FormatCurrency(MattCost(CTR))
Next CTR

'5-8 menu sorted by Name
    picResults.Print "***************************************************************************"
    
    picResults.Print "5-8 Club Menu"
    picResults.Print "------------------------------------"
    picResults.Print "Item"; Tab(25); "Cost"
    picResults.Print "===================================="
    
    For pass = 1 To I - 1
        For comp = 1 To I - pass
    
    
    If ClubItem(comp) > ClubItem(comp + 1) Then
            CItem = ClubItem(comp)
            ClubItem(comp) = ClubItem(comp + 1)
            ClubItem(comp + 1) = CItem
            
            CCost = ClubCost(comp)
            ClubCost(comp) = ClubCost(comp + 1)
            ClubCost(comp + 1) = CCost
            
    End If
    Next comp
Next pass
For CTR = 1 To I
    picResults.Print ClubItem(CTR); Tab(25); FormatCurrency(ClubCost(CTR))
Next CTR
End Sub

Private Sub cmdPrice_Click()
   
    imgClub.Visible = True
    imgMatts.Visible = True
    
    picResults.Cls
'Matt's menu sorted by price
    picResults.Print "Matt's Menu"
    picResults.Print "------------------------------------"
    picResults.Print "Item"; Tab(25); "Cost"
    picResults.Print "===================================="
    
    For pass = 1 To I - 1
        For comp = 1 To I - pass
    
    
    If MattCost(comp) < MattCost(comp + 1) Then
            PosCost = MattCost(comp)
            MattCost(comp) = MattCost(comp + 1)
            MattCost(comp + 1) = PosCost
            
            PosItem = MattItem(comp)
            MattItem(comp) = MattItem(comp + 1)
            MattItem(comp + 1) = PosItem
            
    End If
    Next comp
Next pass
For CTR = 1 To I
    picResults.Print MattItem(CTR); Tab(25); FormatCurrency(MattCost(CTR))
Next CTR

'5-8 menu sorted by price
    picResults.Print "-------------------------------------------------------------------------------"
    
    picResults.Print "5-8 Club Menu"
    picResults.Print "------------------------------------"
    picResults.Print "Item"; Tab(25); "Cost"
    picResults.Print "===================================="
    
    For pass = 1 To I - 1
        For comp = 1 To I - pass
    
    
    If ClubCost(comp) < ClubCost(comp + 1) Then
            CCost = ClubCost(comp)
            ClubCost(comp) = ClubCost(comp + 1)
            ClubCost(comp + 1) = CCost
            
            CItem = ClubItem(comp)
            ClubItem(comp) = ClubItem(comp + 1)
            ClubItem(comp + 1) = CItem
            
    End If
    Next comp
Next pass
For CTR = 1 To I
    picResults.Print ClubItem(CTR); Tab(25); FormatCurrency(ClubCost(CTR))
Next CTR
    
    
End Sub

Private Sub cmdQuit_Click()
    End
End Sub
