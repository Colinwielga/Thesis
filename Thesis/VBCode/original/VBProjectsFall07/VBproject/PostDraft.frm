VERSION 5.00
Begin VB.Form PostDraft 
   BackColor       =   &H00008000&
   Caption         =   "Post Draft"
   ClientHeight    =   4260
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   4260
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdReset 
      Caption         =   "Reset Post Draft Form"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   18
      Top             =   8760
      Width           =   3495
   End
   Begin VB.CommandButton cmdPlayerResponse2 
      Caption         =   " Second Agent/Player Response"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4200
      TabIndex        =   17
      Top             =   7560
      Width           =   3495
   End
   Begin VB.CommandButton cmdOffer2 
      Caption         =   "Make Second Offer To Player"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   480
      TabIndex        =   16
      Top             =   7560
      Width           =   3495
   End
   Begin VB.CommandButton cmdPlayerResponse 
      Caption         =   "Agent/Player Response"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4200
      TabIndex        =   15
      Top             =   6240
      Width           =   3495
   End
   Begin VB.CommandButton cmdMakeOffer 
      Caption         =   "Make Offer To Player"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   480
      TabIndex        =   14
      Top             =   6240
      Width           =   3495
   End
   Begin VB.PictureBox picJerry 
      Height          =   4575
      Left            =   7920
      ScaleHeight     =   4515
      ScaleWidth      =   6915
      TabIndex        =   11
      Top             =   2400
      Width           =   6975
   End
   Begin VB.CommandButton cmdContract2 
      Caption         =   "Get Contract Advice From Agent"
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
      Left            =   4200
      TabIndex        =   10
      Top             =   4800
      Width           =   3495
   End
   Begin VB.TextBox txtPickNumber2 
      BackColor       =   &H000080FF&
      Height          =   615
      Left            =   4200
      TabIndex        =   9
      Top             =   3840
      Width           =   3495
   End
   Begin VB.TextBox txtPlayer2 
      BackColor       =   &H000080FF&
      Height          =   615
      Left            =   4200
      TabIndex        =   7
      Top             =   2160
      Width           =   3495
   End
   Begin VB.CommandButton cmdContract 
      Caption         =   "Contract Negotiation"
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
      Left            =   480
      TabIndex        =   3
      Top             =   4800
      Width           =   3495
   End
   Begin VB.TextBox txtPickNumber 
      BackColor       =   &H000080FF&
      Height          =   615
      Left            =   480
      TabIndex        =   2
      Top             =   3840
      Width           =   3495
   End
   Begin VB.TextBox txtPlayer 
      BackColor       =   &H000080FF&
      Height          =   615
      Left            =   480
      TabIndex        =   1
      Top             =   2160
      Width           =   3495
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "End Draft Season"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4200
      TabIndex        =   0
      Top             =   8760
      Width           =   3495
   End
   Begin VB.Label lblFootballTeam 
      Caption         =   "Owner/GM Contract Negotiation"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   480
      TabIndex        =   13
      Top             =   360
      Width           =   3495
   End
   Begin VB.Label lblProfessional 
      Caption         =   "Agent/Player Contract Negotiation"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4200
      TabIndex        =   12
      Top             =   360
      Width           =   3495
   End
   Begin VB.Label lblPickNumber2 
      Caption         =   "Please Enter What Number Overall You Were Drafted:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4200
      TabIndex        =   8
      Top             =   3120
      Width           =   3495
   End
   Begin VB.Label lblPlayer2 
      Caption         =   "Please Input Your Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4200
      TabIndex        =   6
      Top             =   1440
      Width           =   3495
   End
   Begin VB.Label lblDraftNumber 
      Caption         =   "Please Enter The Overall Pick Number Of Your Team:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   5
      Top             =   3120
      Width           =   3495
   End
   Begin VB.Label lblPlayer 
      Caption         =   "Please Input the Name of the Player You Selected:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   4
      Top             =   1440
      Width           =   3495
   End
End
Attribute VB_Name = "PostDraft"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'NFL Draft by Justin Buysse and Pete Larson. (NFLPOSTDRAFT.vbp)
'November 6th, 2007
'Form Objective: This form allows for post draft one year contract negotiations
    'depending on when a certain player was drafted will grant them a certain
    'amount of money.  This form ultimately has the Agent/Player winning the
    'negotiation with added commentary by well known agent Jerry Maguire.
Option Explicit
Dim Draftee As String
Dim PickNumber As Integer
Dim Draftee2 As String
Dim PickNumber2 As Integer
Private Sub cmdContract_Click()
'When this button is selected, the Owner/GM will say how much the drafted player should be
'paid according to when they were drafted.
cmdContract.Enabled = False
cmdContract2.Enabled = True
cmdMakeOffer.Enabled = False
cmdOffer2.Enabled = False
cmdPlayerResponse.Enabled = False
cmdPlayerResponse2.Enabled = False
cmdQuit.Enabled = True
'Draftee allows the Owner/GM to input the name of the player that they selected
Draftee = txtPlayer.Text
'PickNumber allows the Owner/GM to input the draft pick number of the player that they selected
PickNumber = txtPickNumber.Text
Select Case PickNumber
    Case Is = 1
        MsgBox "Opening Offer should be 1 year " & FormatCurrency(4000000, 0), , "Contract"
    Case 2 To 10
        MsgBox "Opening offer should be 1 year " & FormatCurrency(3500000, 0), , "Contract"
    Case 11 To 19
        MsgBox "Opening offer should be 1 year " & FormatCurrency(3000000, 0), , "Contract"
    Case 20 To 32
        MsgBox "Opening offer should be 1 year " & FormatCurrency(2500000, 0), , "Contract"
    Case Else
        MsgBox "Not a first round pick, invalid entry", , "Contract"
End Select
End Sub
Private Sub cmdContract2_Click()
'When this button is selected, the Agent will say how much the drafted player should be
'paid according to when the player was drafted.
cmdContract2.Enabled = False
cmdMakeOffer.Enabled = True
cmdOffer2.Enabled = False
cmdPlayerResponse.Enabled = False
cmdPlayerResponse2.Enabled = False
'Draftee2 allows the player to input their name
Draftee2 = txtPlayer2.Text
'PickNumber2 allows the player to input the draft pick number that they were selected at
PickNumber2 = txtPickNumber2.Text
picJerry.Picture = LoadPicture(App.Path & "\jerry maguire.jpg")
Select Case PickNumber2
    Case Is = 1
        MsgBox "Jerry Maguire: You should get paid at least " & FormatCurrency(4500000, 0) & " for 1 year", , "Contract"
    Case 2 To 10
        MsgBox "Jerry Maguire: You should get paid at least " & FormatCurrency(4000000, 0) & " for 1 year", , "Contract"
    Case 11 To 19
        MsgBox "Jerry Maguire: You should get paid at least " & FormatCurrency(3500000, 0) & " for 1 year", , "Contract"
    Case 20 To 32
        MsgBox "Jerry Maguire: You should get paid at least " & FormatCurrency(3000000, 0) & " for 1 year", , "Contract"
    Case Else
        MsgBox "Not a first round pick, invalid entry", , "Contract"
End Select
End Sub
Private Sub cmdMakeOffer_Click()
'This button will display the original offer from the Owner/GM to the player/agent
cmdMakeOffer.Enabled = True
cmdOffer2.Enabled = False
cmdPlayerResponse.Enabled = True
cmdPlayerResponse2.Enabled = False
Select Case PickNumber
    Case Is = 1
        MsgBox "Owner/GM: We are offering a 1 year contract worth " & FormatCurrency(4000000, 0), , "Contract"
    Case 2 To 10
        MsgBox "Owner/GM: We are offering a 1 year contract worth " & FormatCurrency(3500000, 0), , "Contract"
    Case 11 To 19
        MsgBox "Owner/GM: We are offering a 1 year contract worth " & FormatCurrency(3000000, 0), , "Contract"
    Case 20 To 32
        MsgBox "Owner/GM: We are offering a 1 year contract worth " & FormatCurrency(2500000, 0), , "Contract"
    Case Else
        MsgBox "Not a first round pick, invalid entry", , "Contract"
End Select
End Sub
Private Sub cmdOffer2_Click()
'This button will display the second offer from the Owner/GM to the player/agent.
'To avoid a holdout, this offer will match the agent and player's original contract expectations.
'This assumption is made because the majority of the time, player's get their way.
cmdMakeOffer.Enabled = False
cmdOffer2.Enabled = False
cmdPlayerResponse.Enabled = False
cmdPlayerResponse2.Enabled = True
Select Case PickNumber
    Case Is = 1
        MsgBox "Owner/GM: We are offering a 1 year contract worth " & FormatCurrency(4500000, 0), , "Contract"
    Case 2 To 10
        MsgBox "Owner/GM: We are offering a 1 year contract worth " & FormatCurrency(4000000, 0), , "Contract"
    Case 11 To 19
        MsgBox "Owner/GM: We are offering a 1 year contract worth " & FormatCurrency(3500000, 0), , "Contract"
    Case 20 To 32
        MsgBox "Owner/GM: We are offering a 1 year contract worth " & FormatCurrency(3000000, 0), , "Contract"
    Case Else
        MsgBox "Not a first round pick, invalid entry", , "Contract"
End Select
End Sub

Private Sub cmdPlayerResponse_Click()
'This button will display the response of the agent/player to the original offer of the Owner/GM
'when the contract is not greater than or equal to the original contract expectation of the agent/player.
'The agent will give an appropriate response showing their disapproval of the contract offer.
cmdMakeOffer.Enabled = False
cmdOffer2.Enabled = True
cmdPlayerResponse.Enabled = False
cmdPlayerResponse2.Enabled = False
Select Case PickNumber2
    Case Is = 1
        MsgBox "Jerry Maguire: SHOW ME THE MONEY", , "Contract"
    Case 2 To 10
        MsgBox "Jerry Maguire: SHOW ME THE MONEY", , "Contract"
    Case 11 To 19
        MsgBox "Jerry Maguire: SHOW ME THE MONEY", , "Contract"
    Case 20 To 32
        MsgBox "Jerry Maguire: SHOW ME THE MONEY", , "Contract"
    Case Else
        MsgBox "Not a first round pick, invalid entry", , "Contract"
End Select
End Sub

Private Sub cmdPlayerResponse2_Click()
'This button will display the response of the agent/player to the second offfer of the Owner/GM
'The offer will meet the original expectation of the agent/player and a deal will be made.
'The agent will display the appropriate response showing their approval of the contract offer.
cmdPlayerResponse2.Enabled = True
Select Case PickNumber
    Case Is = 1
        MsgBox "Jerry Maguire: It's a deal", , "Contract"
    Case 2 To 10
        MsgBox "Jerry Maguire: It's a deal ", , "Contract"
    Case 11 To 19
        MsgBox "Jerry Maguire: It's a deal ", , "Contract"
    Case 20 To 32
        MsgBox "Jerry Maguire: It's a deal ", , "Contract"
    Case Else
        MsgBox "Not a first round pick, invalid entry", , "Contract"
End Select
End Sub

Private Sub cmdQuit_Click()
'This button allows the user to end the draft or quit the program at any time.
    End
End Sub

Private Sub cmdReset_Click()
'This button will reset the Post-Draft form and allow changes to be made to either the player's names
'or their overall pick number in the corresponding text boxes.
cmdContract.Enabled = True
cmdContract2.Enabled = True
cmdMakeOffer.Enabled = True
cmdOffer2.Enabled = True
cmdPlayerResponse.Enabled = True
cmdPlayerResponse2.Enabled = True
cmdQuit.Enabled = True
End Sub
