VERSION 5.00
Begin VB.Form Bid 
   BackColor       =   &H00C0C000&
   Caption         =   "Form1"
   ClientHeight    =   6870
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9840
   BeginProperty Font 
      Name            =   "MT Extra"
      Size            =   8.25
      Charset         =   2
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H0000C000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   6870
   ScaleWidth      =   9840
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CommandButton cmdquit 
      BackColor       =   &H00C0C000&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Imprint MT Shadow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3720
      Width           =   1095
   End
   Begin VB.CommandButton cmdnext 
      BackColor       =   &H0000FF00&
      Caption         =   "Click Here to continue to the next game!!"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   6600
      Picture         =   "bid.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1560
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.CommandButton cmdwin 
      BackColor       =   &H00008000&
      Caption         =   "Click Here if you win to see what you have won!!"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   7200
      MaskColor       =   &H00C0FFFF&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3360
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.PictureBox picinstruct 
      BackColor       =   &H00008000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      ScaleHeight     =   315
      ScaleWidth      =   6675
      TabIndex        =   3
      Top             =   240
      Width           =   6735
   End
   Begin VB.PictureBox picbid 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   1560
      Picture         =   "bid.frx":12C8
      ScaleHeight     =   3555
      ScaleWidth      =   5235
      TabIndex        =   2
      Top             =   3000
      Width           =   5295
      Begin VB.Shape shape2 
         BackColor       =   &H00004000&
         BorderColor     =   &H000000FF&
         Height          =   3495
         Left            =   0
         Shape           =   3  'Circle
         Top             =   0
         Visible         =   0   'False
         Width           =   5175
      End
   End
   Begin VB.CommandButton cmdbid 
      BackColor       =   &H00800000&
      Caption         =   "Click Here to Place your Bid on the item below!!!"
      DisabledPicture =   "bid.frx":60FE
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   2880
      Picture         =   "bid.frx":45AEE
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1560
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.CommandButton cdminstruct 
      BackColor       =   &H00FFFF00&
      Caption         =   "Click Here to view directions on the Game!"
      BeginProperty Font 
         Name            =   "Ravie"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   0
      MaskColor       =   &H0080C0FF&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      Width           =   2895
   End
   Begin VB.Shape Shape1 
      Height          =   615
      Left            =   3240
      Shape           =   2  'Oval
      Top             =   3720
      Width           =   1095
   End
End
Attribute VB_Name = "Bid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cdminstruct_Click()
'This button displays the directions for the game in a text box
    Dim Instructions As String
    Dim Instructions2 As String
    Instructions = "To play: place a bid and if your bid is closer"
    Instructions2 = " , but not over, the bid of the item, you will win!!"
    picinstruct.Print (Instructions); (Instructions2)
'This button also enables the bid button to become visible
    cmdbid.Visible = True
End Sub

Private Sub cmdbid_Click()
'This button enables the user to enter a bid in an
'input box and then decides with Case statements whether the
'player is a winner or loser
'Formats Name to display only fist name
    Dim Bid As Single
    Dim Oponentbid As Single
    Dim Item As Single
    Oponentbid = 500
    Item = 700
    Bid = InputBox("please enter a bid!", "Bid")
    
    
    Select Case Bid
        Case 501 To 699
            MsgBox (WholeName) & " , you have WON, the price of the item is $700.00 and the oponent guessed $500.00!!", , "Winner"
            cmdwin.Visible = True
            Runningtotal = Runningtotal
            TV = True
        Case 0 To 499
            MsgBox (WholeName) & " you have LOSSED, the price of the item is $700.00 and the openent was closer with a guess of $500.00!", , "Loser"
            Runningtotal = Runningtotal
        Case Is < 0
            MsgBox (WholeName) & " ,Error!! Please try again, you must enter a positive value!", , "Error!!"
        Case 700
            MsgBox (WholeName) & " you are AMAZING!! you have guessed the price exactly right!! You have won an extra $500.00!!", , "Winner"
            cmdwin.Visible = True
            Runningtotal = Runningtotal + 500
            TV = True
        Case Is > 700
            MsgBox (WholeName) & " you have LOSSED, the price of the item is $700.00 and you have gone over!!!", , "Loser"
            Runningtotal = Runningtotal
        Case 500
            MsgBox (WholeName) & " You have tied the oponents bid!! Please try again!", , "You Tied"
        End Select
        cmdnext.Visible = True
End Sub

Private Sub cmdnext_Click()
'This buttons enables the user to move to the next form and continue playing the game
Bid.Hide
wheel.Show

End Sub

Private Sub cmdquit_Click()
End
End Sub

Private Sub cmdwin_Click()
 'This button enables the winner to view what the winnings are
 'and circles the prize to further clarify what the prize is
 Dim Winnings As String
 shape2.Visible = True
 Winnings = MsgBox((WholeName) & " You have one the item shown to your left!!")
End Sub
