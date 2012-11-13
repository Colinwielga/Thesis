VERSION 5.00
Begin VB.Form frmWarsaw 
   BackColor       =   &H80000008&
   Caption         =   "Form2"
   ClientHeight    =   10875
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   15240
   LinkTopic       =   "Form2"
   Picture         =   "Form2.frx":0000
   ScaleHeight     =   10875
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CommandButton cmdNato 
      Caption         =   "NATO"
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   12600
      Picture         =   "Form2.frx":21EC8
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   7320
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Surrender"
      BeginProperty Font 
         Name            =   "Felix Titling"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   12600
      Picture         =   "Form2.frx":9E43A
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   8760
      Width           =   2655
   End
   Begin VB.PictureBox picPicture 
      BackColor       =   &H80000009&
      Height          =   5895
      Left            =   0
      ScaleHeight     =   5835
      ScaleWidth      =   9435
      TabIndex        =   7
      Top             =   4920
      Visible         =   0   'False
      Width           =   9495
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H80000009&
      Height          =   2775
      Left            =   600
      ScaleHeight     =   2715
      ScaleWidth      =   5835
      TabIndex        =   6
      Top             =   1440
      Visible         =   0   'False
      Width           =   5895
   End
   Begin VB.CommandButton cmd127x108 
      BeginProperty Font 
         Name            =   "Gill Sans Ultra Bold Condensed"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   10920
      MaskColor       =   &H00FFFFFF&
      Picture         =   "Form2.frx":A348B
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2640
      Width           =   3735
   End
   Begin VB.CommandButton cmd762t 
      BackColor       =   &H80000009&
      Caption         =   "7.62x25mm"
      BeginProperty Font 
         Name            =   "Gill Sans Ultra Bold Condensed"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   11280
      Picture         =   "Form2.frx":A86A4
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5520
      Width           =   3135
   End
   Begin VB.CommandButton cmd762r 
      BackColor       =   &H80000009&
      Caption         =   "7.62x54r"
      BeginProperty Font 
         Name            =   "Gill Sans Ultra Bold Condensed"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   10320
      Picture         =   "Form2.frx":A93FB
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   0
      Width           =   4815
   End
   Begin VB.CommandButton cmd545 
      BackColor       =   &H80000009&
      Caption         =   "5.45x39.5mm"
      BeginProperty Font 
         Name            =   "Gill Sans Ultra Bold Condensed"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   10320
      Picture         =   "Form2.frx":C02E1
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1320
      Width           =   4815
   End
   Begin VB.Label lblWeapons 
      Caption         =   "Type of Weapon that fires this ammunition"
      BeginProperty Font 
         Name            =   "Myriad Web Pro Condensed"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   9
      Top             =   4560
      Visible         =   0   'False
      Width           =   5415
   End
   Begin VB.Label lblAmmo 
      Caption         =   "AMMUNITION INFO"
      BeginProperty Font 
         Name            =   "Myriad Web Pro"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   8
      Top             =   960
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000009&
      Caption         =   "12.7x108mm"
      BeginProperty Font 
         Name            =   "Gill Sans Ultra Bold Condensed"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   12000
      TabIndex        =   5
      Top             =   5160
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000009&
      Caption         =   "WARSAW Pact Munitions"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5055
   End
End
Attribute VB_Name = "frmWarsaw"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'We started off dimming every variable at the top and opening all the files when the program started

Dim ctr As Integer, fivefourgr(1 To 10) As Single, fivefourfps(1 To 10) As Single, fivefourfpe(1 To 10) As Single
Dim fidyfourgr(1 To 10) As Single, fidyfourfps(1 To 10) As Single, fidyfourfpe(1 To 10) As Single
Dim tofifgr(1 To 10) As Single, tofiffps(1 To 10) As Single, tofiffpe(1 To 10) As Single






Private Sub cmd127x108_Click()
'This was the button that had no array with it
ctr = 1 'This sets the counter as one
    lblAmmo.Visible = True 'this makes the ammo label visible
    lblWeapons.Visible = True 'this makes the weapons label visible
     picResults.Visible = True 'this makes the results picbox visible
    picPicture.Visible = True 'this makes the weapon picbox visible
    picPicture.Cls 'this clears the weapon picbox
     picResults.Cls 'this clears the results picbox
  picResults.Print "Bullet Size"; Tab(25); "Feet Per Second"; Tab(45); "Foot Pounds of Energy"
  picResults.Print "********************************************************************************"
 picResults.Print "802 grains"; Tab(25); "2800"; Tab(45); "14,180"
 'above puts an organized file that is uniform to the other arrays in from the other files
 'listed already are the weight, speed, and foot pounds of energy
 picPicture.Picture = LoadPicture(App.Path & "\mg12_7mm-2.jpg")
 'This loads a picture of a "Dushka" heavy machine gun for viewing pleasure
 
End Sub

Private Sub cmd545_Click()
ctr = 2 'There are two rows of info from this file
    lblAmmo.Visible = True 'this makes the ammo label visible
    lblWeapons.Visible = True 'this makes the weapons label visible
     picResults.Visible = True 'this makes the results picbox visible
    picPicture.Visible = True 'this makes the weapon picbox visible
    picPicture.Cls 'this clears the weapon picbox
     picResults.Cls 'this clears the results picbox
   picResults.Print "Bullet Size"; Tab(25); "Feet Per Second"; Tab(45); "Foot Pounds of Energy"
   picResults.Print "********************************************************************************"
   'the above two lines create an organized table that is displayed in the picResults picture box
Dim J As Integer
  
  For J = 1 To ctr
   picResults.Print fivefourgr(J); "grains"; Tab(25); fivefourfps(J); Tab(45); FormatNumber((fivefourfpe(J)), 0)
  Next J
'This for next loop takes all of the information in the file and organizes it into a list, first it is the bullet's weight, then its speed, then its hit power
picPicture.Picture = LoadPicture(App.Path & "\AK-74.jpg")
'this file loads a picture of an AK-74 that shoots this type of round





End Sub

Private Sub cmd762r_Click()
ctr = 3
     lblAmmo.Visible = True 'this makes the ammo label visible
     lblWeapons.Visible = True 'this makes the weapons label visible
    picResults.Visible = True 'this makes the results picbox visible
     picPicture.Visible = True 'this makes the weapon picbox visible
     picPicture.Cls 'this clears the weapon picbox
    picResults.Cls 'this clears the results picbox

    picResults.Print "Bullet Size"; Tab(25); "Feet Per Second"; Tab(45); "Foot Pounds of Energy"
    picResults.Print "********************************************************************************"
    'the above two lines create an organized table that is displayed in the picResults picture box
Dim J As Integer
 For J = 1 To ctr
  picResults.Print fidyfourgr(J); "grains"; Tab(25); fidyfourfps(J); Tab(45); FormatNumber((fidyfourfpe(J)), 0)
 Next J
 'This for next loop takes all of the information in the file and organizes it into a list, first it is the bullet's weight, then its speed, then its hit power
picPicture.Picture = LoadPicture(App.Path & "\pkms.jpg")
 'This loads a picture of a PKMS into the picPicture picture box



End Sub

Private Sub cmd762t_Click()
ctr = 4
      lblAmmo.Visible = True 'this makes the ammo label visible
      lblWeapons.Visible = True 'this makes the weapons label visible
     picResults.Visible = True 'this makes the results picbox visible
    picPicture.Visible = True 'this makes the weapon picbox visible
    picPicture.Cls 'this clears the weapon picbox
     picResults.Cls 'this clears the results picbox

    picResults.Print "Bullet Size"; Tab(25); "Feet Per Second"; Tab(45); "Foot Pounds of Energy"
    picResults.Print "********************************************************************************"
    'the above two lines create an organized table that is displayed in the picResults picture box
Dim J As Integer
 For J = 1 To ctr
  picResults.Print tofifgr(J); "grains"; Tab(25); tofiffps(J); Tab(45); FormatNumber((tofiffpe(J)), 0)
Next J
 'This for next loop takes all of the information in the file and organizes it into a list, first it is the bullet's weight, then its speed, then its hit power

picPicture.Picture = LoadPicture(App.Path & "\type79smg.jpg")
'This loads a picture of a PKMS into the picPicture picture box

End Sub

Private Sub cmdNato_Click() 'This button allows the user to switch out to the NATO form to view their bullet information
 frmWarsaw.Hide  'this hides the current form
 frmNATO.Show   'this righteously displays the nato form

End Sub

Private Sub Command1_Click()
 MsgBox "What??? YOU COULDN'T HANDLE IT?"
 End 'This button ends the program.  Why would the user leave the program? Probably personal reasons. Who knows?
 
End Sub

Private Sub Form_Load()
'We decided to open all the arrays at the beginning of the form, just like on the nato form
'all the variables are dimmed at the top and are assembled into arrays below at the start of the program
'The fourth ammunition (the 12.7x108mm) unfortunately only had one weight type we could find, so we opted to just show the information when the button was clicked and not waste time on a fourth .txt file

Open App.Path & "\54539.txt" For Input As #1
ctr = 0 'having had previous problems with loading lots of files and having lots of arrays,
'we decided to reset the counter after every file was loaded and arrayed
Do While Not EOF(1)
    ctr = ctr + 1
    Input #1, fivefourgr(ctr), fivefourfps(ctr), fivefourfpe(ctr)
Loop
Close #1
'This is the end of the first file
Open App.Path & "\76225.txt" For Input As #2
ctr = 0
Do While Not EOF(2)
    ctr = ctr + 1
    Input #2, tofifgr(ctr), tofiffps(ctr), tofiffpe(ctr)
Loop
Close #2
'This is the end of the second file
Open App.Path & "\76254r.txt" For Input As #3
ctr = 0
Do While Not EOF(3)
    ctr = ctr + 1
    Input #3, fidyfourgr(ctr), fidyfourfps(ctr), fidyfourfpe(ctr)
Loop
Close #3
'You guessed it! This is the end of the third file

MsgBox "WARSAW PACT IS AMMO IS READY TO ROCK!!!"
'Mother Russia and her allies are ready to kick some Capitalist asses, and the arrays are also successfully loaded


End Sub


