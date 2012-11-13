VERSION 5.00
Begin VB.Form frmDisco 
   BackColor       =   &H00000000&
   Caption         =   "Discography"
   ClientHeight    =   9165
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10710
   LinkTopic       =   "Form1"
   ScaleHeight     =   9165
   ScaleWidth      =   10710
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Exit This Rad Program"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6120
      TabIndex        =   19
      Top             =   8280
      Width           =   1095
   End
   Begin VB.CommandButton cmdSongs 
      Caption         =   "Go to Songs Page"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4800
      TabIndex        =   18
      Top             =   8280
      Width           =   1215
   End
   Begin VB.CommandButton cmdMain 
      Caption         =   "Go to Main Page"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3360
      TabIndex        =   17
      Top             =   8280
      Width           =   1335
   End
   Begin VB.CommandButton cmdCheck 
      BackColor       =   &H00FFFFFF&
      Caption         =   "See Track Listing"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   16
      Top             =   1920
      Width           =   3255
   End
   Begin VB.TextBox txtAlbum 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1095
      Left            =   5160
      TabIndex        =   9
      Top             =   720
      Width           =   2055
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   4095
      Left            =   3240
      ScaleHeight     =   4035
      ScaleWidth      =   4035
      TabIndex        =   6
      Top             =   4080
      Width           =   4095
   End
   Begin VB.PictureBox picMB 
      Height          =   2895
      Left            =   7440
      Picture         =   "frmDisco.frx":0000
      ScaleHeight     =   2835
      ScaleWidth      =   2955
      TabIndex        =   5
      Top             =   3120
      Width           =   3015
   End
   Begin VB.PictureBox picRed 
      Height          =   2895
      Left            =   7440
      Picture         =   "frmDisco.frx":4528
      ScaleHeight     =   2835
      ScaleWidth      =   2955
      TabIndex        =   4
      Top             =   6120
      Width           =   3015
   End
   Begin VB.PictureBox cmdPink 
      Height          =   2895
      Left            =   120
      Picture         =   "frmDisco.frx":21FA6
      ScaleHeight     =   2835
      ScaleWidth      =   2955
      TabIndex        =   3
      Top             =   3120
      Width           =   3015
   End
   Begin VB.PictureBox picGreen 
      Height          =   2895
      Left            =   120
      Picture         =   "frmDisco.frx":3FA24
      ScaleHeight     =   2835
      ScaleWidth      =   2955
      TabIndex        =   2
      Top             =   6120
      Width           =   3015
   End
   Begin VB.PictureBox picMala 
      Height          =   2895
      Left            =   7440
      Picture         =   "frmDisco.frx":41DD2
      ScaleHeight     =   2835
      ScaleWidth      =   2955
      TabIndex        =   1
      Top             =   120
      Width           =   3015
   End
   Begin VB.PictureBox picBlue 
      Height          =   2895
      Left            =   120
      Picture         =   "frmDisco.frx":457BC
      ScaleHeight     =   2835
      ScaleWidth      =   2955
      TabIndex        =   0
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label lblMala 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "4. Maladroit"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   5400
      TabIndex        =   15
      Top             =   2400
      Width           =   1935
   End
   Begin VB.Label lblPink 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "2. Pinkerton"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   3240
      TabIndex        =   14
      Top             =   3000
      Width           =   1935
   End
   Begin VB.Label lblMB 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "5. Make Believe"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   5400
      TabIndex        =   13
      Top             =   3000
      Width           =   1935
   End
   Begin VB.Label lblGreen 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "3. Weezer (Green Album)"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   3240
      TabIndex        =   12
      Top             =   3480
      Width           =   1935
   End
   Begin VB.Label lblRed 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "6. Weezer (Red Album)"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   5400
      TabIndex        =   11
      Top             =   3480
      Width           =   1935
   End
   Begin VB.Label lblBlue 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "1. Weezer (Blue Album)"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   3240
      TabIndex        =   10
      Top             =   2400
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Type the album name here exactly as is below to  see its track listing:"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   3360
      TabIndex        =   8
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label lblDisco 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "DISCOGRAPHY"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   3360
      TabIndex        =   7
      Top             =   120
      Width           =   3855
   End
End
Attribute VB_Name = "frmDisco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Weezer
'Form Name: frmDisco.frm
'Author: Emily Balamut
'Date Written: 11/4/08
'Objective: This form allows the user to enter in an album they wish to see the
'track listing and release date for.
Option Explicit
Private Sub cmdCheck_Click()
    Dim AlbumInfo As String
    
    AlbumInfo = txtAlbum.Text
    picResults.Cls
    
    If AlbumInfo = "Weezer (Blue Album)" Then
        picResults.Print "Year Released: May 10th, 1994"
        picResults.Print "1.) My Name Is Jonas"
        picResults.Print "2.) No One Else"
        picResults.Print "3.) The World Has Turned and Left Me Here"
        picResults.Print "4.) Buddy Holly"
        picResults.Print "5.) Undone (The Sweater Song)"
        picResults.Print "6.) Surf Wax America"
        picResults.Print "7.) Say It Ain't So"
        picResults.Print "8.) In the Garage"
        picResults.Print "9.) Holiday"
        picResults.Print "10.) Only In Dreams"
    End If
    If AlbumInfo = "Pinkerton" Then
        picResults.Print "Date Released: September 24th, 1996"
        picResults.Print "1.) Tired of Sex"
        picResults.Print "2.) Getchoo"
        picResults.Print "3.) No Other One"
        picResults.Print "4.) Why Bother"
        picResults.Print "5.) Across the Sea"
        picResults.Print "6.) The Good Life"
        picResults.Print "7.) El Scorcho"
        picResults.Print "8.) Pink Triangle"
        picResults.Print "9.) Falling For You"
        picResults.Print "10.) Butterfly"
    End If
    If AlbumInfo = "Weezer (Green Album)" Then
        picResults.Print "Date Released: May 15th, 2001"
        picResults.Print "1.) Don't Let Go"
        picResults.Print "2.) Photograph"
        picResults.Print "3.) Hash Pipe"
        picResults.Print "4.) Island in the Sun"
        picResults.Print "5.) Crab"
        picResults.Print "6.) Knock-Down Drag-Out"
        picResults.Print "7.) Smile"
        picResults.Print "8.) Simple Pages"
        picResults.Print "9.) Glorious Days"
        picResults.Print "10.) O Girlfriend"
     End If
    If AlbumInfo = "Maladroit" Then
        picResults.Print "Date Released: May 14th, 2002"
        picResults.Print "1.) American Gigolo"
        picResults.Print "2.) Dope Nose"
        picResults.Print "3.) Keep Fishin'"
        picResults.Print "4.) Take Control"
        picResults.Print "5.) Death and Destruction"
        picResults.Print "6.) Slob"
        picResults.Print "7.) Burndt Jamb"
        picResults.Print "8.) Space Rock"
        picResults.Print "9.) Slave"
        picResults.Print "10.) Fall Together"
        picResults.Print "11.) Possibilites"
        picResults.Print "12.) Love Explosion"
        picResults.Print "13.) December"
    End If
    If AlbumInfo = "Make Believe" Then
        picResults.Print "Date Released: May 10th, 2005"
        picResults.Print "1.) Beverly Hills"
        picResults.Print "2.) Perfect Situation"
        picResults.Print "3.) This Is Such a Pity"
        picResults.Print "4.) Hold Me"
        picResults.Print "5.) Peace"
        picResults.Print "6.) We Are All On Drugs"
        picResults.Print "7.) The Damage In Your Heart"
        picResults.Print "8.) Pardon Me"
        picResults.Print "9.) My Best Friend"
        picResults.Print "10.) The Other Way"
        picResults.Print "11.) Freak Me Out"
        picResults.Print "12.) Haunt You Everyday"
    End If
    If AlbumInfo = "Weezer (Red Album)" Then
        picResults.Print "Date Released: June 17th, 2008"
        picResults.Print "1.) Troublemaker"
        picResults.Print "2.) The Greatest Man That Ever Lived"
        picResults.Print "3.) Pork and Beans"
        picResults.Print "4.) Heart Songs"
        picResults.Print "5.) Everybody Get Dangerous"
        picResults.Print "6.) Dreamin'"
        picResults.Print "7.) Thought I Knew"
        picResults.Print "8.) Cold Dark World"
        picResults.Print "9.) Automatic"
        picResults.Print "10.) The Angel and the One"
    End If

End Sub

Private Sub cmdMain_Click()
    frmDisco.Hide
    frmBeginning.Show
End Sub

Private Sub cmdQuit_Click()
MsgBox "Thanks for rocking out with Weezer, " & UserName & "! See you later!", , "Bye!"
End
End Sub

Private Sub cmdSongs_Click()
    frmDisco.Hide
    frmSongs.Show
End Sub
