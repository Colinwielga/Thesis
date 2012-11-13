VERSION 5.00
Begin VB.Form frmSteamboatresorts 
   BackColor       =   &H00000000&
   Caption         =   "Steamboat Resorts"
   ClientHeight    =   9240
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14010
   LinkTopic       =   "Form1"
   ScaleHeight     =   9240
   ScaleWidth      =   14010
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmbBack 
      Caption         =   "Back to Resorts"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   6
      Top             =   9480
      Width           =   1335
   End
   Begin VB.PictureBox picResults 
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7215
      Left            =   5520
      ScaleHeight     =   7155
      ScaleWidth      =   7515
      TabIndex        =   5
      Top             =   840
      Width           =   7575
   End
   Begin VB.CommandButton cmd5 
      Height          =   1695
      Left            =   2880
      Picture         =   "frmSteamboatresorts.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   840
      Width           =   2055
   End
   Begin VB.CommandButton cmd4 
      Height          =   1695
      Left            =   240
      Picture         =   "frmSteamboatresorts.frx":3DCC
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6360
      Width           =   2055
   End
   Begin VB.CommandButton cmd3 
      Height          =   1695
      Left            =   240
      Picture         =   "frmSteamboatresorts.frx":7950
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4440
      Width           =   2055
   End
   Begin VB.CommandButton cmd2 
      Height          =   1695
      Left            =   240
      Picture         =   "frmSteamboatresorts.frx":D5DB
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2640
      Width           =   2055
   End
   Begin VB.CommandButton cmd1 
      Height          =   1695
      Left            =   240
      Picture         =   "frmSteamboatresorts.frx":10F13
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   840
      Width           =   2055
   End
   Begin VB.Label lblname 
      Caption         =   "By: Levi Glines and John Krebsbach"
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   10680
      Width           =   2775
   End
End
Attribute VB_Name = "frmSteamboatresorts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : Colorado Spring Break(Final.vbp)
'Form Name : frmSteamboatresorts(frmSteamboatresorts.frm)
'Author: Levi Glines and John Krebsbach
'Date : Thursday March 23, 2006
'Purpose of this form:  this form allows the user to get a greater sense of what each
'resort is about. it explains each resort in greater detail than just the price.
Private Sub cmbBack_Click()
    frmSteamboatresorts.Hide 'hides this form
    frmSteamboatlodge.Show 'brings you back to the steamboatlodge form
End Sub

Private Sub cmd1_Click()
    picResults.Cls 'clears out any info in the pic box
    picResults.Print "Sheraton Steamboat Hotel"
    picResults.Print 'prints a blank line
    picResults.Print "Located just steps away from the famous Champagne Powder, Sheraton"
    picResults.Print "Steamboat Springs Resort & Conference Center is Steamboat's"
    picResults.Print "slopeside conference resort. Hailed as one of the best ski resorts"
    picResults.Print "by Condé Nast Traveler's Top Ski Resorts readers' survey, one of"
    picResults.Print "Twelve Perfect Places in North America"
End Sub

Private Sub cmd2_Click()
    picResults.Cls 'clears out any info in the pic box
    picResults.Print "Steamboat Grand Hotel  "
    picResults.Print 'prints a blank line
    picResults.Print "As soon as you arrive, you feel the warmth and beauty of the mountain"
    picResults.Print "architecture ... from the Grand Lobby entrance and the 327 luxury"
    picResults.Print "guestrooms and suites ... to our Priest Creek Ballroom and"
    picResults.Print "17,000 sq. ft. of meeting space. Whether it's business, pleasure"
    picResults.Print "or both, the Steamboat Grand has it all!"
End Sub

Private Sub cmd3_Click()
    picResults.Cls 'clears out any info in the pic box
    picResults.Print "Best Western Ptarmigan Inn Hotel"
    picResults.Print 'prints a blank line
    picResults.Print "The Best Western Ptarmigan Inn is located at the base of the Mount"
    picResults.Print "Werner/Steamboat Ski area, a full service hotel, offering"
    picResults.Print "Sportstalker Ski Shop, complimentary valet ski storage, outdoor"
    picResults.Print "heated pool, hot tub and sauna."
End Sub

Private Sub cmd4_Click()
    picResults.Cls 'clears out any info in the pic box
    picResults.Print "Fairfield Inn & Suites"
    picResults.Print 'prints a blank line
    picResults.Print "Steamboat is indeed the real thing, with it's natural hot springs,"
    picResults.Print "friendly local people and a great old western town full of character"
    picResults.Print "and history, and of course, famous champagne powder skiing. There is"
    picResults.Print "something for everyone - festivals, ballooning, golf,"
    picResults.Print "and fishing, a winter driving."
End Sub

Private Sub cmd5_Click()
    picResults.Cls 'clears out any info in the pic box
    picResults.Print "Thunderhead Hotel"
    picResults.Print 'prints a blank line
    picResults.Print "Knowledgeable skiers select lodging by location. Those same skiers"
    picResults.Print "are loyal and frequent guests at Thunderhead Lodge. Thunderhead Lodge"
    picResults.Print "offers true slope side access, a Ski Time Square location, a variety"
    picResults.Print "of on-site amenities and convenient access to popular mountain"
    picResults.Print "village shops and restaraunts."
End Sub

