VERSION 5.00
Begin VB.Form MainForm 
   BackColor       =   &H00004000&
   Caption         =   "Hunting Information  for Minnesota Seasons"
   ClientHeight    =   10320
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   18240
   LinkTopic       =   "Form1"
   ScaleHeight     =   10320
   ScaleWidth      =   18240
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text5 
      BackColor       =   &H000080FF&
      Height          =   855
      Left            =   14760
      TabIndex        =   23
      Top             =   2640
      Width           =   3375
   End
   Begin VB.CommandButton cmdChoose 
      BackColor       =   &H000080FF&
      Caption         =   "Click to see what your desired animal."
      Height          =   855
      Left            =   13320
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   2640
      Width           =   1335
   End
   Begin VB.CommandButton cmdSeasonOrder 
      BackColor       =   &H0000C000&
      Caption         =   "Order of hunting seasons in Minnesota"
      Height          =   1095
      Left            =   16680
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   0
      Width           =   1335
   End
   Begin VB.CommandButton cmdType 
      BackColor       =   &H000080FF&
      Caption         =   "Available Game according to weapon used."
      Height          =   975
      Left            =   13320
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   4920
      Width           =   1335
   End
   Begin VB.TextBox txtGame 
      BackColor       =   &H000080FF&
      Height          =   855
      Left            =   14760
      TabIndex        =   17
      Top             =   3960
      Width           =   3375
   End
   Begin VB.TextBox txtSeason 
      BackColor       =   &H000080FF&
      Height          =   735
      Left            =   14760
      TabIndex        =   13
      Top             =   6240
      Width           =   3375
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H000080FF&
      Height          =   735
      Left            =   14760
      TabIndex        =   12
      Top             =   7320
      Width           =   3375
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H000080FF&
      Height          =   615
      Left            =   14760
      TabIndex        =   11
      Top             =   8400
      Width           =   3375
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H000080FF&
      Height          =   765
      Left            =   14760
      TabIndex        =   10
      Top             =   5160
      Width           =   3375
   End
   Begin VB.CommandButton cmdGame 
      BackColor       =   &H000080FF&
      Caption         =   "How Much Game Can be Taken "
      Height          =   1215
      Left            =   13320
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3600
      Width           =   1335
   End
   Begin VB.CommandButton cmdSeason 
      BackColor       =   &H000080FF&
      Caption         =   "Length of Desired Season"
      Height          =   975
      Left            =   13320
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6000
      Width           =   1335
   End
   Begin VB.CommandButton cmdSky 
      BackColor       =   &H000080FF&
      Caption         =   "Current Sky Conditions"
      Height          =   855
      Left            =   13320
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   8160
      Width           =   1335
   End
   Begin VB.CommandButton cmdTemperature 
      BackColor       =   &H000080FF&
      Caption         =   "Temperature outlooks for supplies needed"
      Height          =   975
      Left            =   13320
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7080
      Width           =   1335
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H000000FF&
      Caption         =   "Quit"
      Height          =   975
      Left            =   16680
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1200
      Width           =   1335
   End
   Begin VB.CommandButton cmdGrouse 
      BackColor       =   &H0000C000&
      Caption         =   "Grouse Hunting Information"
      Height          =   975
      Left            =   15000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1200
      Width           =   1335
   End
   Begin VB.CommandButton cmdBear 
      BackColor       =   &H0000C000&
      Caption         =   "Bear Hunting Information"
      Height          =   1095
      Left            =   15000
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   0
      Width           =   1335
   End
   Begin VB.CommandButton cmdWaterfowl 
      BackColor       =   &H0000C000&
      Caption         =   "WaterFowl Hunting Information"
      Height          =   975
      Left            =   13320
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1200
      Width           =   1335
   End
   Begin VB.CommandButton cmdDeer 
      BackColor       =   &H0000C000&
      Caption         =   "Deer Hunting Information"
      Height          =   1095
      Left            =   13320
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   0
      Width           =   1335
   End
   Begin VB.PictureBox Results 
      BackColor       =   &H000080FF&
      Height          =   9975
      Left            =   0
      ScaleHeight     =   9915
      ScaleWidth      =   13035
      TabIndex        =   0
      Top             =   0
      Width           =   13095
   End
   Begin VB.Label Label2 
      BackColor       =   &H000080FF&
      Caption         =   "Choose  # 1-4: 1-Deer, 2-Bear, 3-Waterfowl, 4-Grouse"
      Height          =   255
      Left            =   13320
      TabIndex        =   24
      Top             =   2400
      Width           =   4815
   End
   Begin VB.Label Label1 
      BackColor       =   &H000080FF&
      Caption         =   "What type of Weapon will you be using?"
      Height          =   255
      Left            =   14760
      TabIndex        =   21
      Top             =   4920
      Width           =   3375
   End
   Begin VB.Label Label7 
      BackColor       =   &H000080FF&
      Caption         =   "What type of game do you wish to harvest? Be sure to Capitalize the name of the animal."
      Height          =   375
      Left            =   14760
      TabIndex        =   18
      Top             =   3600
      Width           =   3375
   End
   Begin VB.Label Label6 
      BackColor       =   &H000080FF&
      Caption         =   "What are the sky conditions today:"
      Height          =   255
      Left            =   14760
      TabIndex        =   16
      Top             =   8160
      Width           =   3375
   End
   Begin VB.Label Label5 
      BackColor       =   &H000080FF&
      Caption         =   "What is the todays Temperature (in Fahrenheit) :"
      Height          =   255
      Left            =   14760
      TabIndex        =   15
      Top             =   7080
      Width           =   3375
   End
   Begin VB.Label Label3 
      BackColor       =   &H000080FF&
      Caption         =   "What Species would you like to hunt:"
      Height          =   255
      Left            =   14760
      TabIndex        =   14
      Top             =   6000
      Width           =   3375
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Game I Enjoy Hunting in Minnesota
'MainForm
'Mark Hines
'3-23-09
'This form provides the information and pathways to get to the other huntable game.  you can access weather data and seasons that are open as well as some of the other
'regulations that apply. It also includes information for novice hunters that might prove to be useful.
'All commands clear screen when run to avoid overlaping images with text.
Option Explicit
Dim weapon As String, rifle As String, shotgun As String, bow As String, found As Boolean, ctr3 As Integer, deer As String, bear As String, Grouse As String, phesant As String, trukey As String, waterfowl As String, smallgame As String, game As String, season(1 To 5) As String
'Shows the form for the Bear information
Private Sub cmdBear_Click()
FormBear.Show
End Sub

Private Sub cmdChoose_Click()
Dim animal As String
'take user input to produce a picture based on the their response.
animal = Text5.Text
    If animal = "1" Then
        Results.Cls
        Results.Picture = LoadPicture("")
        Results.Picture = LoadPicture(App.Path & "\Marcheldeer_200.jpg")
    ElseIf animal = "2" Then
        Results.Cls
        Results.Picture = LoadPicture("")
        Results.Picture = LoadPicture(App.Path & "\blackbear.jpg")
    ElseIf animal = "3" Then
        Results.Cls
        Results.Picture = LoadPicture("")
        Results.Picture = LoadPicture(App.Path & "\2450-2290.jpg")
    ElseIf animal = "4" Then
        Results.Cls
        Results.Picture = LoadPicture("")
        Results.Picture = LoadPicture(App.Path & "\Marchelgrouse_200.jpg")
    Else
        MsgBox "You have enter an invalid entry.  Please try again.", , "Error"
    
    End If
End Sub



'show the deer form where information is all on deer hunting
Private Sub cmdDeer_Click()
formDeer.Show
End Sub

Private Sub cmdGame_Click()
Dim I As Integer, found As Boolean, animal As String, baglimit(1 To 6) As Integer, species(1 To 6) As String, tempLimit As Integer, ctr As Integer
Results.Cls
Results.Picture = LoadPicture("")
'open array and load it wtih variable names
Open App.Path & "\game.txt" For Input As #1
ctr = 0
Do While Not EOF(1)
    ctr = ctr + 1
    Input #1, species(ctr), baglimit(ctr)
Loop
Close #1
'user input information changed to variable for application
animal = txtGame.Text
I = 0
found = False
'search array for the information concerning the animal input by the user.
For I = 1 To ctr
If animal = species(I) Then
    tempLimit = baglimit(I)
    found = True
  End If
Next I
'once the animal has been found it is printed or user is sent a message box message saying error
    If Not found Then
        MsgBox "There arent any animals under that name in the directory.", , "Baglimit"
    Else
        Results.Print animal; " have a baglimit of "; tempLimit
        
    End If

End Sub
'brings up the Grouse form where more information is held
Private Sub cmdGrouse_Click()
formGrouse.Show
End Sub
'ends the program
Private Sub cmdQuit_Click()
End
End Sub

Private Sub cmdSeason_Click()
Dim I As Integer
Results.Cls
Results.Picture = LoadPicture("")
'user enters the type of game they wish to hunt and see the season for.
game = txtSeason.Text
Results.Print "**************************************************************************************************************************************"
        'if the input matches one of the names listed then the according message will be printed in the picture box.
        If game = "deer" Then
          Results.Print "For Deer: "
          Results.Print "9-19-09 through 12--13-09 is Minnesotas archery season"
          Results.Print "11-7-09 through 11-21-09 is Minnesotas rifle season"
          Results.Print "11-28-09 through 12-13-09 is Minnesotas muzzleloader season."
        ElseIf game = "phesant" Then
        Results.Print "Legal season starts October 10th and ends January 3rd."
        ElseIf game = "Turkey" Then
        Results.Print "10/14/09 - 10/18/09 1st turkey season and 10/21/09 - 10/25/09 is the 2nd turkey season in the fall."
        ElseIf game = "grouse" Then
        Results.Print "Legal season starts September 19th through January 3rd."
        ElseIf game = "bear" Then
        Results.Print "Legal season starts September 1st through October 18th."
        ElseIf game = "waterfowl" Then
        Results.Print "Legal season starts October 3rd and ends November 28th."
        ElseIf game = "smallgame" Then
        Results.Print "Legal season starts on September 19th and ends February 28th."
        'result if there arent any corresponding species.
        Else
            MsgBox "There isnt a season on this particular animal or it is not a game animal. Possibly is not a native species of Minnesota.", , "Error"
            
    End If
    Results.Print "**************************************************************************************************************************************"


End Sub
Private Sub cmdSeasonOrder_Click()
Dim ctr As Single, pass As Integer, pos As Integer, day(1 To 5) As Single, temp As Integer, tempsea As String, tempdate As Single, J As Integer
Results.Cls
Results.Picture = LoadPicture("")
'opens the array and loads it and attatches variables
Open App.Path & "\Seasons.txt" For Input As #1
ctr = 0
Do While Not EOF(1)
    ctr = ctr + 1
    Input #1, season(ctr), day(ctr)
Loop

' the function then takes the seasons of that array and sorts them according to the begining date of that season
Results.Print season(ctr)
    For pass = 1 To ctr - 1
        For pos = 1 To ctr - pass
            If day(pos) > day(pos + 1) Then
            tempsea = season(pos)
            season(pos) = season(pos + 1)
            season(pos + 1) = tempsea
            tempdate = day(pos)
            day(pos) = day(pos + 1)
            day(pos + 1) = tempdate
        End If
    Next pos
Next pass
  'the results are then printed with this header above it
    Results.Print "Game.", "Start Date"
    Results.Print "**********************************************************************************************"
 'this is where the dates are set with the animal to be hunted and printed out
For J = 1 To ctr
    Results.Print season(J), day(J)
Next J
   
End Sub

Private Sub cmdSky_Click()
Dim condition As String
Results.Cls
Results.Picture = LoadPicture("")
'take user input and match it with the corresponding answer to get your information to pop up in a message box.
condition = Text2.Text
If condition = "cloudy" Then
    MsgBox "better get out there the ducks will be moving today."
ElseIf condition = "sunny" Then
    MsgBox "most wildlife is rather lazy when it is sunny but it is always nice to soak up the rays."
ElseIf condition = "raining" Then
    MsgBox "best duck hunting weather known to man.  Get up and get out there!"
ElseIf condition = "foggy" Then
    MsgBox "what is the point if you cant see what you are shooting at."
ElseIf condition = "snowing" Then
    MsgBox "make things easier on yourself go out and get a deer. Tracking will be easy as it gets."
Else
    MsgBox "I dont know what to tell you. You might have to make the decision yourself."
    End If
End Sub

Private Sub cmdTemperature_Click()
Dim temp As Single
Results.Cls
Results.Picture = LoadPicture("")
'obtain information about weather from user.
temp = Text3.Text
Results.Cls
' then take user information and apply. if answer falls within one of the ranges then that answer will be displayed.
Select Case temp
    Case Is >= 100
        Results.Print "Holy crap you are gonna sweat and be rather miserable."
    Case Is >= 90
        Results.Print "Damn it is hot. You better dress in cool clothing or suffer the consequences."
    Case Is >= 80
        Results.Print "I dont even know why you are considering hunting right now!'"
    Case Is >= 70
        Results.Print "I would wear as little Camo as possible but watchout for the mosquitos."
    Case Is >= 60
        Results.Print "Starting to cool off finally and better get out and start scouting."
    Case Is >= 50
        Results.Print "Time to strap up get preped and get some sleep, ducks are flying and its time to focus on your passion."
    Case Is >= 40
        Results.Print "Time to put some layers on and hold your ground Ducks will be flying and lots of wildlife movement."
    Case Is >= 30
        Results.Print "Get on your cold weather gear and keep at it. Bound to pull something in with this cold front."
    Case Is >= 20
        Results.Print "Cold! Damn i know you dont want to get up but some of the best hunting occurs during cold snaps."
    Case Is >= 10
        Results.Print "Cover up and bundle up and dont let old man winter beat you."
    Case Is <= 9
        Results.Print "Holy crap most people would say forget it but you are heading out."
        Results.Print "You are a trooper and can truely call yourself dedicated."
    End Select

End Sub

Private Sub cmdType_Click()
weapon = Text1.Text
Results.Cls
Results.Picture = LoadPicture("")
'take user input and match with a possible outcome.  all are printed in the picture box
If weapon = "bow" Then
Results.Print "There are four types of game that can be hunted with a bow. Turkey, deer, bear and small game."
ElseIf weapon = "rifle" Then
Results.Print "There are two types of game that can be hunted with a rifle. Bear and Deer"
ElseIf weapon = "shotgun" Then
Results.Print " there are six different types of game that can be hunted with a shotgun.  Turkey, deer, waterfowl, grouse, pheasant, and small game."
Else
    MsgBox ("Could not read input please check spelling or weapon type.")
End If

End Sub
'go to the waterfowl form where more data can be obtained
Private Sub cmdWaterfowl_Click()
formWaterfowl.Show
End Sub


Private Sub Label4_Click()

End Sub

