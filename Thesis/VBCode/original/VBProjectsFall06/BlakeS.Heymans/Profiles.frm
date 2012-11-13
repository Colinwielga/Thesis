VERSION 5.00
Begin VB.Form frmProfiles 
   BackColor       =   &H80000012&
   Caption         =   "Player Profiles"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   Picture         =   "Profiles.frx":0000
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picResults 
      BackColor       =   &H00008000&
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   4680
      ScaleHeight     =   1035
      ScaleWidth      =   5715
      TabIndex        =   13
      Top             =   6000
      Visible         =   0   'False
      Width           =   5775
   End
   Begin VB.CommandButton cmdFindprofile 
      Caption         =   "Find Player's Profile"
      BeginProperty Font 
         Name            =   "Rockwell Condensed"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      TabIndex        =   12
      Top             =   2880
      Width           =   1575
   End
   Begin VB.TextBox txtProfile 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   720
      TabIndex        =   11
      Text            =   "Enter # Here"
      Top             =   2400
      Width           =   1095
   End
   Begin VB.CommandButton cmdBackmenu3 
      Caption         =   "Back to Main Menu"
      BeginProperty Font 
         Name            =   "Rockwell Condensed"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   12960
      TabIndex        =   0
      Top             =   9360
      Width           =   2055
   End
   Begin VB.Image ImageD 
      Height          =   5550
      Left            =   5760
      Picture         =   "Profiles.frx":18430
      Top             =   240
      Visible         =   0   'False
      Width           =   3705
   End
   Begin VB.Image ImageP 
      Height          =   5550
      Left            =   5760
      Picture         =   "Profiles.frx":1C93F
      Top             =   240
      Visible         =   0   'False
      Width           =   3705
   End
   Begin VB.Image ImageE 
      Height          =   5550
      Left            =   5760
      Picture         =   "Profiles.frx":20E85
      Top             =   240
      Visible         =   0   'False
      Width           =   3705
   End
   Begin VB.Image ImageT 
      Height          =   5550
      Left            =   5760
      Picture         =   "Profiles.frx":24B19
      Top             =   240
      Visible         =   0   'False
      Width           =   3705
   End
   Begin VB.Image ImageC 
      Height          =   5550
      Left            =   5760
      Picture         =   "Profiles.frx":28410
      Top             =   240
      Visible         =   0   'False
      Width           =   3705
   End
   Begin VB.Image ImageTr 
      Height          =   5550
      Left            =   5760
      Picture         =   "Profiles.frx":2C4F9
      Top             =   240
      Visible         =   0   'False
      Width           =   3705
   End
   Begin VB.Image ImageS 
      Height          =   5550
      Left            =   5760
      Picture         =   "Profiles.frx":2FFD1
      Top             =   240
      Visible         =   0   'False
      Width           =   3705
   End
   Begin VB.Image ImageA 
      Height          =   5550
      Left            =   5760
      Picture         =   "Profiles.frx":33E7B
      Top             =   240
      Visible         =   0   'False
      Width           =   3705
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "To See a Profile Enter the Corresponding Number of the Player Below then Press the Button"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   2055
      Left            =   240
      TabIndex        =   10
      Top             =   240
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "SAINT JOHN'S ROSTER 2006"
      BeginProperty Font 
         Name            =   "Rockwell Condensed"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   0
      TabIndex        =   9
      Top             =   3600
      Width           =   3135
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "1.) Steve Tacl"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   4080
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "2.) Trevor Beach"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   4440
      Width           =   2055
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "3.) Adam Putschoegl"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   4800
      Width           =   2535
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "4.) Curtis Horton"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   5160
      Width           =   2055
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "5.) Ted Lauer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   5520
      Width           =   1935
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "6.) Eric Holmgren"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   5880
      Width           =   2175
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "7.) Pat Boerner"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   6240
      Width           =   2055
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "8.) Dan Ruehl"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   6600
      Width           =   1935
   End
   Begin VB.Image Image7 
      Height          =   3240
      Left            =   2400
      Picture         =   "Profiles.frx":37F87
      Top             =   240
      Width           =   2250
   End
   Begin VB.Image Image6 
      Height          =   3780
      Left            =   6240
      Picture         =   "Profiles.frx":3D428
      Top             =   6480
      Width           =   2760
   End
   Begin VB.Image Image5 
      Height          =   2160
      Left            =   7920
      Picture         =   "Profiles.frx":403AB
      Top             =   3720
      Width           =   2880
   End
   Begin VB.Image Image4 
      Height          =   1500
      Left            =   5880
      Picture         =   "Profiles.frx":416C2
      Top             =   3960
      Width           =   1425
   End
   Begin VB.Image Image3 
      Height          =   2550
      Left            =   5880
      Picture         =   "Profiles.frx":432F1
      Top             =   600
      Width           =   4980
   End
   Begin VB.Image Image2 
      Height          =   2925
      Left            =   3000
      Picture         =   "Profiles.frx":4766C
      Top             =   3960
      Width           =   2295
   End
   Begin VB.Image Image1 
      Height          =   2250
      Left            =   2160
      Picture         =   "Profiles.frx":4A59C
      Top             =   7440
      Width           =   3015
   End
End
Attribute VB_Name = "frmProfiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
'2006 MIAC Tennis Tournament Distribution
'Opponent Search Form
'Blake Heymans
'10/30/06
'Opponent Search Form Objective
    'This form is intended to show the user the individual player profile of
    'each Saint John's player on the roster. First it promts the user to enter
    'a number to the corresponding name on the roster. This number is then read into
    'a file containing the profile of an individual player. That number is then
    'compared to a the player profile and if it is found the data is printed in a
    'picture box and the player's picture will appear. If the player number is not
    'found in the file then the user is promted that they must enter a different
    'player number. This allows for a user to learn some information about each player.
'Pictures were taken from Google image search as well as the Saint John's Univerity web site.
    
Private Sub cmdBackmenu3_Click()
    'Brings the user back to Title Form
    frmTitle.Show
    frmProfiles.Hide
End Sub

Private Sub cmdFindprofile_Click()
Dim Pname As String, Found As Boolean, I As Integer, Ctr As Integer
    
    'makes the picture box visible
    picResults.Visible = True
    picResults.Cls

    'clears images from the screen
    Image1.Visible = False
    Image2.Visible = False
    Image3.Visible = False
    Image4.Visible = False
    Image5.Visible = False
    Image6.Visible = False
    Image7.Visible = False
    
    Ctr = 0
    
    Pname = txtProfile.Text
    
    'For Piture to appear all other pictures must be hidden except the input
    If Pname = 1 Then
        ImageS.Visible = True
        ImageA.Visible = False
        ImageTr.Visible = False
        ImageC.Visible = False
        ImageT.Visible = False
        ImageE.Visible = False
        ImageP.Visible = False
        ImageD.Visible = False
    End If
    If Pname = 2 Then
        ImageTr.Visible = True
        ImageS.Visible = False
        ImageA.Visible = False
        ImageC.Visible = False
        ImageT.Visible = False
        ImageE.Visible = False
        ImageP.Visible = False
        ImageD.Visible = False
    End If
    If Pname = 3 Then
        ImageA.Visible = True
        ImageS.Visible = False
        ImageTr.Visible = False
        ImageC.Visible = False
        ImageT.Visible = False
        ImageE.Visible = False
        ImageP.Visible = False
        ImageD.Visible = False
    End If
    If Pname = 4 Then
        ImageC.Visible = True
        ImageS.Visible = False
        ImageA.Visible = False
        ImageTr.Visible = False
        ImageT.Visible = False
        ImageE.Visible = False
        ImageP.Visible = False
        ImageD.Visible = False
    End If
    If Pname = 5 Then
        ImageT.Visible = True
        ImageS.Visible = False
        ImageA.Visible = False
        ImageTr.Visible = False
        ImageC.Visible = False
        ImageE.Visible = False
        ImageP.Visible = False
        ImageD.Visible = False
    End If
    If Pname = 6 Then
        ImageE.Visible = True
        ImageS.Visible = False
        ImageA.Visible = False
        ImageTr.Visible = False
        ImageC.Visible = False
        ImageT.Visible = False
        ImageP.Visible = False
        ImageD.Visible = False
    End If
    If Pname = 7 Then
        ImageP.Visible = True
        ImageS.Visible = False
        ImageA.Visible = False
        ImageTr.Visible = False
        ImageC.Visible = False
        ImageT.Visible = False
        ImageE.Visible = False
        ImageD.Visible = False
    End If
    If Pname = 8 Then
        ImageD.Visible = True
        ImageS.Visible = False
        ImageA.Visible = False
        ImageTr.Visible = False
        ImageC.Visible = False
        ImageT.Visible = False
        ImageE.Visible = False
        ImageP.Visible = False
    End If
    
    'Opens Profile file into arrays
    Open App.Path & "\Profile.txt" For Input As #1
    
    Do While Not EOF(1)
        Ctr = Ctr + 1
        Input #1, Playername(Ctr), YearS(Ctr), Hometown(Ctr)
    Loop
    
    Close #1
    
    'Boolean search to find player profile from file 1
    I = 0
    Found = False
    
    Do While ((Not Found) And (I < Ctr))
        I = I + 1
        If Pname = I Then Found = True
    Loop
    
    picResults.Print "Player's Name"; Tab(30); "Year In School"; Tab(60); "Hometown"
    picResults.Print "______________"; Tab(30); "______________"; Tab(60); "______________"
    picResults.Print
    
    If (Not Found) Then
            MsgBox "That Player # Is Not On The 2006 Roster. Please Enter A Different Player #.", , "Error"
        Else
            picResults.Print Playername(I); Tab(30); YearS(I), Tab(60); Hometown(I)
    End If
    
End Sub

Private Sub txtProfile_click()
    'clears text box when user clicks on it
    txtProfile.Text = " "
End Sub
