VERSION 5.00
Begin VB.Form frmRSInformation 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Republic of Srpska"
   ClientHeight    =   9855
   ClientLeft      =   270
   ClientTop       =   855
   ClientWidth     =   14355
   LinkTopic       =   "Form1"
   ScaleHeight     =   9855
   ScaleWidth      =   14355
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00FF8080&
      Caption         =   "Back to main page"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   8520
      Width           =   2295
   End
   Begin VB.CommandButton cmdCities 
      BackColor       =   &H00FF8080&
      Caption         =   "Get more information about large cities in Republic of Srpska"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   1320
      Width           =   3735
   End
   Begin VB.PictureBox picBanjaLuka 
      Height          =   1215
      Left            =   3480
      Picture         =   "frmRSInformation.frx":0000
      ScaleHeight     =   1155
      ScaleWidth      =   1275
      TabIndex        =   13
      Top             =   3480
      Width           =   1335
   End
   Begin VB.PictureBox picDoboj 
      Height          =   1215
      Left            =   6120
      Picture         =   "frmRSInformation.frx":465C
      ScaleHeight     =   1155
      ScaleWidth      =   1275
      TabIndex        =   12
      Top             =   3480
      Width           =   1335
   End
   Begin VB.PictureBox picPrijedor 
      Height          =   1215
      Left            =   840
      Picture         =   "frmRSInformation.frx":5B26
      ScaleHeight     =   1155
      ScaleWidth      =   1275
      TabIndex        =   11
      Top             =   6000
      Width           =   1335
   End
   Begin VB.PictureBox picTrebinje 
      Height          =   1215
      Left            =   3480
      Picture         =   "frmRSInformation.frx":7135
      ScaleHeight     =   1155
      ScaleWidth      =   1275
      TabIndex        =   10
      Top             =   6000
      Width           =   1335
   End
   Begin VB.PictureBox picSarajevo 
      Height          =   1215
      Left            =   6120
      Picture         =   "frmRSInformation.frx":8544
      ScaleHeight     =   1155
      ScaleWidth      =   1275
      TabIndex        =   9
      Top             =   6000
      Width           =   1335
   End
   Begin VB.PictureBox picBijeljina 
      Height          =   1215
      Left            =   840
      Picture         =   "frmRSInformation.frx":9176
      ScaleHeight     =   1155
      ScaleWidth      =   1275
      TabIndex        =   8
      Top             =   3480
      Width           =   1335
   End
   Begin VB.CommandButton cmdLoad 
      BackColor       =   &H00FF8080&
      Caption         =   "Elementar informations about Republic of Srpska"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1320
      Width           =   3735
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7575
      Left            =   8160
      ScaleHeight     =   7515
      ScaleWidth      =   6075
      TabIndex        =   0
      Top             =   1440
      Width           =   6135
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   "Welcome to Republic of Srpska"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1080
      TabIndex        =   16
      Top             =   240
      Width           =   11295
   End
   Begin VB.Label lblPrijedor 
      BackStyle       =   0  'Transparent
      Caption         =   "4 PRIJEDOR"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   840
      TabIndex        =   7
      Top             =   7320
      Width           =   1215
   End
   Begin VB.Label lblTrebinje 
      BackStyle       =   0  'Transparent
      Caption         =   "5 TREBINjE"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   3600
      TabIndex        =   6
      Top             =   7320
      Width           =   1215
   End
   Begin VB.Label lblBijeljina 
      BackStyle       =   0  'Transparent
      Caption         =   "6 EAST SARAJEVO"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   5880
      TabIndex        =   5
      Top             =   7320
      Width           =   1815
   End
   Begin VB.Label lblDoboj 
      BackStyle       =   0  'Transparent
      Caption         =   "3 Doboj"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   6360
      TabIndex        =   4
      Top             =   4800
      Width           =   975
   End
   Begin VB.Label lblBanjaLuka 
      BackStyle       =   0  'Transparent
      Caption         =   "2 BanJa Luka"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   3360
      TabIndex        =   3
      Top             =   4800
      Width           =   1575
   End
   Begin VB.Shape shpBanjaLuka 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000000&
      Height          =   1935
      Index           =   6
      Left            =   3120
      Top             =   3240
      Width           =   2055
   End
   Begin VB.Shape shpDoboj 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000000&
      Height          =   1935
      Index           =   4
      Left            =   5760
      Top             =   3240
      Width           =   2055
   End
   Begin VB.Shape shpPrijedor 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000000&
      Height          =   1935
      Index           =   3
      Left            =   480
      Top             =   5760
      Width           =   2055
   End
   Begin VB.Shape shpEastSarajevo 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000000&
      Height          =   1935
      Index           =   2
      Left            =   5760
      Top             =   5760
      Width           =   2055
   End
   Begin VB.Shape shpTrebinje 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000000&
      Height          =   1935
      Index           =   1
      Left            =   3120
      Top             =   5760
      Width           =   2055
   End
   Begin VB.Label lblBijeljina 
      BackStyle       =   0  'Transparent
      Caption         =   "1 BIJELJINA"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   960
      TabIndex        =   2
      Top             =   4800
      Width           =   1215
   End
   Begin VB.Shape shpBijeljina 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000000&
      Height          =   1935
      Index           =   0
      Left            =   480
      Top             =   3240
      Width           =   2055
   End
End
Attribute VB_Name = "frmRSInformation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'This code defineds Global Values
Dim History(1 To 100) As String

Private Sub cmdBack_Click()

'Cod Option that connect two forms
'In this case this form and Main Page
frmMainPage.Show
frmRSInformation.Hide

End Sub

Private Sub cmdCities_Click()
'This code defineds Privat Values
Dim Ctr As Integer, EnterNumber As Integer

'Those codes option Organized and Clean output Box
picResults.Cls
picResults.Print
picResults.Print

'This code option Open Data files as input
Open App.Path & "\Bijeljina.txt" For Input As #1
Open App.Path & "\BanjaLuka.txt" For Input As #2
Open App.Path & "\Doboj.txt" For Input As #3
Open App.Path & "\Prijedor.txt" For Input As #4
Open App.Path & "\Trebinje.txt" For Input As #5
Open App.Path & "\EastSarajevo.txt" For Input As #6

EnterNumber = InputBox("Please write number of city about want to know more", "City")



'This code option organized array and read Data files
'In same time force program that take information from guest
'connect information whit information in data and
'search for valid condition.
Ctr = 0
Select Case EnterNumber
Case 1 To 6
    Do While Not EOF(EnterNumber)
    Ctr = Ctr + 1
    Input #EnterNumber, History(Ctr)
    picResults.Print History(Ctr)
    Loop
Case Else
    MsgBox "Sorry your number is incorrect, please try again"
End Select

'Close Outputs
Close #1
Close #2
Close #3
Close #4
Close #5
Close #6

End Sub

Private Sub cmdLoad_Click()

Dim Ctr As Integer


'Those codes option Organized and Clean output Box
picResults.Cls
picResults.Print
picResults.Print
picResults.Print Tab(5); "Information about Republic of Srpska"
picResults.Print

'This code option Open Data file as input
Open App.Path & "\RS-informations.txt" For Input As #1

'This code option organized array and read Data files
'Then print information from data files in output
Ctr = 0
Do While Not EOF(1)
Ctr = Ctr + 1
Input #1, History(Ctr)
picResults.Print Tab(5); History(Ctr) ' print information in output
Loop

Close #1 'close output

End Sub


