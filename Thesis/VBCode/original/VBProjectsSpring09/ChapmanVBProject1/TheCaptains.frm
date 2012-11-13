VERSION 5.00
Begin VB.Form TheCaptainsForm 
   BackColor       =   &H00000000&
   Caption         =   "Form1"
   ClientHeight    =   9675
   ClientLeft      =   1545
   ClientTop       =   1125
   ClientWidth     =   12225
   LinkTopic       =   "Form1"
   MousePointer    =   6  'Size NE SW
   ScaleHeight     =   9675
   ScaleWidth      =   12225
   Begin VB.TextBox txtClarence 
      Height          =   375
      Left            =   9600
      TabIndex        =   6
      Text            =   "Clarence Manuel"
      Top             =   6120
      Width           =   1335
   End
   Begin VB.TextBox txtBobby 
      Height          =   375
      Left            =   3240
      TabIndex        =   5
      Text            =   "Bobby Chapman"
      Top             =   5760
      Width           =   1335
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00FF0000&
      Caption         =   "Back to Team Info"
      BeginProperty Font 
         Name            =   "Ravie"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1100
      Left            =   9240
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7200
      Width           =   2415
   End
   Begin VB.PictureBox picResults3 
      BackColor       =   &H000000C0&
      BeginProperty Font 
         Name            =   "JazzText"
         Size            =   14.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   4320
      ScaleHeight     =   1035
      ScaleWidth      =   4635
      TabIndex        =   3
      Top             =   7200
      Width           =   4695
   End
   Begin VB.CommandButton cmdNicknames 
      BackColor       =   &H000080FF&
      Caption         =   "Our Nicknames"
      BeginProperty Font 
         Name            =   "Ravie"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1100
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7200
      Width           =   2415
   End
   Begin VB.PictureBox picResults 
      Height          =   4455
      Left            =   720
      Picture         =   "TheCaptains.frx":0000
      ScaleHeight     =   4395
      ScaleWidth      =   6675
      TabIndex        =   1
      Top             =   1080
      Width           =   6735
   End
   Begin VB.PictureBox picResults2 
      Height          =   5775
      Left            =   9000
      Picture         =   "TheCaptains.frx":54FC
      ScaleHeight     =   5715
      ScaleWidth      =   2355
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "TheCaptainsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'WaterPoloProject
'TheCaptainsForm
'Bobby Chapman
'Written 3/16/2009
'Objective-tell the user who the captains are, provide pictures, and give
'the user their nicknames
Option Explicit

Private Sub cmdNicknames_Click()
'declare local variables
Dim BobNick As String, ClarNick As String, BFirst As String, CFirst As String
Dim BobbyWhole As String, ClarWhole As String, N As Integer

'declares the text in the textbox under Bobby as BobbyWhole
BobbyWhole = txtBobby.Text

'gets the number of characteres in the textbox
N = InStr(BobbyWhole, " ")

'finds the first name in the textbox
BFirst = Left(BobbyWhole, N - 1)

'declares that BobNick is the first letter in the textbox
BobNick = Left(BFirst, 1)

'declares the text in the textbox under Clarence as ClarWhole
ClarWhole = txtClarence.Text

'gets the number of characters in the textbox
N = InStr(ClarWhole, " ")

'finds the first name in the textbox
CFirst = Left(ClarWhole, N - 1)

'declares ClarNick as the first letter in the textbox
ClarNick = Left(CFirst, 1)

'prints Bobby and Clarence's nicknames
picResults3.Print "Bobby's nickname is "; BobNick; " and Clarence's is "; ClarNick; "."

End Sub

Private Sub cmdBack_Click()
'goes back to the TeamForm
TheCaptainsForm.Hide
TeamForm.Show
End Sub

