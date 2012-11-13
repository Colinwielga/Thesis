VERSION 5.00
Begin VB.Form frmAvailability 
   BackColor       =   &H80000007&
   Caption         =   "All Movies Available for Sale"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   Picture         =   "frmAvailability.frx":0000
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCheckOut 
      Caption         =   "Check Out"
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
      Left            =   6600
      TabIndex        =   23
      Top             =   8520
      Width           =   2535
   End
   Begin VB.CommandButton cmdTotal 
      Caption         =   "Total"
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
      Left            =   6600
      TabIndex        =   22
      Top             =   7320
      Width           =   2535
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
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
      Left            =   6600
      TabIndex        =   21
      Top             =   6120
      Width           =   2535
   End
   Begin VB.CommandButton cmdLethal 
      Caption         =   "Lethal Weapon 4"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2400
      TabIndex        =   20
      Top             =   3600
      Width           =   1695
   End
   Begin VB.CommandButton cmdBodyguard 
      Caption         =   "The Bodyguard from Beijing"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4440
      TabIndex        =   19
      Top             =   9000
      Width           =   1695
   End
   Begin VB.CommandButton cmdOnce 
      Caption         =   "Once Upon a Time in China VI"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2400
      TabIndex        =   18
      Top             =   5760
      Width           =   1695
   End
   Begin VB.CommandButton cmdTheOne 
      Caption         =   "The One"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      TabIndex        =   17
      Top             =   7920
      Width           =   1695
   End
   Begin VB.CommandButton cmdAdventure 
      Caption         =   "Adventure King"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2400
      TabIndex        =   16
      Top             =   7920
      Width           =   1695
   End
   Begin VB.CommandButton cmdRomeo 
      Caption         =   "Romeo Must Die"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      TabIndex        =   15
      Top             =   9000
      Width           =   1695
   End
   Begin VB.CommandButton cmdKing 
      Caption         =   "King of Assassins"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2400
      TabIndex        =   14
      Top             =   4680
      Width           =   1695
   End
   Begin VB.CommandButton cmdKiss 
      Caption         =   "Kiss of the Dragon"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      TabIndex        =   13
      Top             =   6840
      Width           =   1695
   End
   Begin VB.CommandButton cmdFist 
      Caption         =   "Fist of Legend"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4440
      TabIndex        =   12
      Top             =   6840
      Width           =   1695
   End
   Begin VB.CommandButton cmdTheEnforcer 
      Caption         =   "The Enforcer"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2400
      TabIndex        =   11
      Top             =   9000
      Width           =   1695
   End
   Begin VB.CommandButton cmdBlack 
      Caption         =   "Black Mask"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2400
      TabIndex        =   10
      Top             =   6840
      Width           =   1695
   End
   Begin VB.CommandButton cmdCradle 
      Caption         =   "Cradle 2 the Grave"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      TabIndex        =   9
      Top             =   5760
      Width           =   1695
   End
   Begin VB.CommandButton cmdShaolin 
      Caption         =   "Shaolin Kung Fu"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4440
      TabIndex        =   8
      Top             =   7920
      Width           =   1695
   End
   Begin VB.CommandButton cmdHigh 
      Caption         =   "High Risk"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4440
      TabIndex        =   7
      Top             =   4680
      Width           =   1695
   End
   Begin VB.CommandButton cmdLegend 
      Caption         =   "Legend of the Future Shaolin"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4440
      TabIndex        =   6
      Top             =   5760
      Width           =   1695
   End
   Begin VB.CommandButton cmdHero 
      Caption         =   "Hero"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      TabIndex        =   5
      Top             =   4680
      Width           =   1695
   End
   Begin VB.CommandButton cmdUnleashed 
      Caption         =   "Unleashed"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      TabIndex        =   4
      Top             =   3600
      Width           =   1695
   End
   Begin VB.CommandButton cmdFearless 
      Caption         =   "Fearless"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      TabIndex        =   3
      Top             =   2520
      Width           =   1695
   End
   Begin VB.PictureBox picResults 
      Height          =   10095
      Left            =   9600
      ScaleHeight     =   10035
      ScaleWidth      =   5115
      TabIndex        =   2
      Top             =   240
      Width           =   5175
   End
   Begin VB.CommandButton cmdMain 
      Caption         =   "Main Page"
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
      Left            =   1680
      TabIndex        =   1
      Top             =   240
      Width           =   1095
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
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
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label lblNote 
      Caption         =   "  Note: There is a 7% sales tax and 3% shipping and handling fee."
      Height          =   255
      Left            =   360
      TabIndex        =   25
      Top             =   1440
      Width           =   4695
   End
   Begin VB.Label lblSelect 
      Caption         =   " Click on any movie to add into Selection:"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   24
      Top             =   840
      Width           =   5055
   End
End
Attribute VB_Name = "frmAvailability"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Planet of Jet Li
'Form Name: frmAvailability
'Author: Chakong Thao
'Date Written: Wednesday, Nov. 1st
'Form Objective: This form contains all the Jet Li movies that are
                'available for sale on this program.  It finally
                'separates just listing and searching for titles
                'and actually computes the prices of the selections
                'made by the user and finds the subtotal, tax, shipping
                'fee, and grand total.
                
Option Explicit
Dim Subtotal As Single, Counter As Single

Private Sub cmdAdventure_Click()    'Chooses this movie and displays in picture box
    Dim Adventure As Single
    Adventure = 8.99
    Subtotal = Adventure + Subtotal
    picResults.Print "Adventure King"; Tab(50); FormatCurrency(Adventure)
End Sub

Private Sub cmdBack_Click() 'Brings user back to the movie search page
    frmAvailability.Hide
    frmMovieSale.Show
End Sub

Private Sub cmdBlack_Click()    'Chooses this movie and displays in picture box
    Dim Black As Single
    Black = 8.99
    Subtotal = Black + Subtotal
    picResults.Print "Black Mask"; Tab(50); FormatCurrency(Black)
End Sub

Private Sub cmdBodyguard_Click()    'Chooses this movie and displays in picture box
    Dim Bodyguard As Single
    Bodyguard = 15.99
    Subtotal = Bodyguard + Subtotal
    picResults.Print "The Bodyguard From Beijing"; Tab(50); FormatCurrency(Bodyguard)
End Sub

Private Sub cmdCheckOut_Click() 'An input box will pop up and ask for an address after selections have been made
    Dim A As String
    A = InputBox("Great! Now please enter a billing and shipping address.", "Shipping Information")
    MsgBox "Thank you for your purchase.  Your item(s) will arrive within the next 2 weeks.", , "Purchase Confirmed"
End Sub

Private Sub cmdClear_Click()    'This deletes everything already in the picture box
    Dim Clear As Single
    picResults.Cls
    Subtotal = 0
End Sub

Private Sub cmdCradle_Click()   'Chooses this movie and displays in picture box
    Dim Cradle As Single
    Cradle = 15.99
    Subtotal = Cradle + Subtotal
    picResults.Print "Cradle 2 the Grave"; Tab(50); FormatCurrency(Cradle)
End Sub

Private Sub cmdFearless_Click() 'Chooses this movie and displays in picture box
    Dim Fearless As Single
    Fearless = 19.99
    Subtotal = Fearless + Subtotal
    picResults.Print "Fearless"; Tab(50); FormatCurrency(Fearless)
End Sub

Private Sub cmdFist_Click() 'Chooses this movie and displays in picture box
    Dim Fist As Single
    Fist = 5.99
    Subtotal = Fist + Subtotal
    picResults.Print "Fist of Legend"; Tab(50); FormatCurrency(Fist)
End Sub

Private Sub cmdHero_Click() 'Chooses this movie and displays in picture box
    Dim Hero As Single
    Hero = 15.99
    Subtotal = Hero + Subtotal
    picResults.Print "Hero"; Tab(50); FormatCurrency(Hero)
End Sub

Private Sub cmdHigh_Click() 'Chooses this movie and displays in picture box
    Dim High As Single
    High = 6.99
    Subtotal = High + Subtotal
    picResults.Print "High Risk"; Tab(50); FormatCurrency(High)
End Sub

Private Sub cmdKing_Click() 'Chooses this movie and displays in picture box
    Dim King As String
    King = 10.99
    Subtotal = King + Subtotal
    picResults.Print "King of Assassins"; Tab(50); FormatCurrency(King)
End Sub

Private Sub cmdKiss_Click() 'Chooses this movie and displays in picture box
    Dim Kiss As Single
    Kiss = 15.99
    Subtotal = Kiss + Subtotal
    picResults.Print "Kiss of the Dragon"; Tab(50); FormatCurrency(Kiss)
End Sub

Private Sub cmdLegend_Click()   'Chooses this movie and displays in picture box
    Dim Legend As Single
    Legend = 6.99
    Subtotal = Legend + Subtotal
    picResults.Print "Legend of the Future Shaolin"; Tab(50); FormatCurrency(Legend)
End Sub

Private Sub cmdLethal_Click()   'Chooses this movie and displays in picture box
    Dim Lethal As String
    Lethal = 10.99
    Subtotal = Lethal + Subtotal
    picResults.Print "Lethal Weapon 4"; Tab(50); FormatCurrency(Lethal)
End Sub

Private Sub cmdMain_Click() 'Brings user back to beginning page
    frmAvailability.Hide
    frmJetLi.Show
End Sub

Private Sub cmdOnce_Click() 'Chooses this movie and displays in picture box
    Dim Once As Single
    Once = 10.99
    Subtotal = Once + Subtotal
    picResults.Print "Once Upon a Time in China VI"; Tab(50); FormatCurrency(Once)
End Sub

Private Sub cmdRomeo_Click()    'Chooses this movie and displays in picture box
    Dim Romeo As String
    Romeo = 12.99
    Subtotal = Romeo + Subtotal
    picResults.Print "Romeo Must Die"; Tab(50); FormatCurrency(Romeo)
End Sub

Private Sub cmdShaolin_Click()  'Chooses this movie and displays in picture box
    Dim Shaolin As Single
    Shaolin = 5.99
    Subtotal = Shaolin + Subtotal
    picResults.Print "Shaolin Kung Fu"; Tab(50); FormatCurrency(Shaolin)
End Sub

Private Sub cmdTheEnforcer_Click()  'Chooses this movie and displays in picture box
    Dim Enforcer As Single
    Enforcer = 8.99
    Subtotal = Enforcer + Subtotal
    picResults.Print "The Enforcer"; Tab(50); FormatCurrency(Enforcer)
End Sub

Private Sub cmdTheOne_Click()   'Chooses this movie and displays in picture box
    Dim TheOne As Single
    TheOne = 12.99
    Subtotal = TheOne + Subtotal
    picResults.Print "The One"; Tab(50); FormatCurrency(TheOne)
End Sub

Private Sub cmdTotal_Click()    'This button computes and displays the subtotal, tax, shipping fee, and grand total for all merchandise selected by user
    Dim Total As Single, Tax As Single, Shipping As Single
    picResults.Print "------------------------------------------"
    Tax = Subtotal * 0.07
    Shipping = Subtotal * 0.03
    Total = Subtotal + Subtotal * 0.07 + Subtotal * 0.03
    picResults.Print "Subtotal"; Tab(50); FormatCurrency(Subtotal)
    picResults.Print "Tax"; Tab(50); FormatCurrency(Tax)
    picResults.Print "Shipping & Handling"; Tab(50); FormatCurrency(Shipping)
    picResults.Print "Total"; Tab(50); FormatCurrency(Total)
End Sub

Private Sub cmdUnleashed_Click()    'Chooses this movie and displays in picture box
    Dim Unleashed As Single
    Unleashed = 19.99
    Subtotal = Unleashed + Subtotal
    picResults.Print "Unleashed"; Tab(50); FormatCurrency(Unleashed)
End Sub
