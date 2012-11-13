VERSION 5.00
Begin VB.Form formWaterfowl 
   BackColor       =   &H00004000&
   Caption         =   "Minnesota Waterfowl Hunting Information"
   ClientHeight    =   7800
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14730
   LinkTopic       =   "Form1"
   ScaleHeight     =   7800
   ScaleWidth      =   14730
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdTotal 
      BackColor       =   &H0000C000&
      Caption         =   "Total # of Birds Bagged for the Year and how the season is going for you."
      Height          =   855
      Left            =   11760
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4440
      Width           =   2895
   End
   Begin VB.CommandButton cmdIllegal 
      BackColor       =   &H0000C000&
      Caption         =   "Illegal Ducks l in Minnesota (what you cant shoot)."
      Height          =   975
      Left            =   13320
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3360
      Width           =   1335
   End
   Begin VB.CommandButton cmdBagG 
      BackColor       =   &H0000C000&
      Caption         =   "Bag Limit for Geese in Minnesota"
      Height          =   975
      Left            =   11760
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3360
      Width           =   1455
   End
   Begin VB.CommandButton cmdBag 
      BackColor       =   &H0000C000&
      Caption         =   "Bag Limits for Ducks in Minnesota"
      Height          =   975
      Left            =   13320
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2280
      Width           =   1335
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H000000FF&
      Caption         =   "Quit"
      Height          =   975
      Left            =   12360
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5520
      Width           =   1935
   End
   Begin VB.CommandButton cmdReturn 
      BackColor       =   &H0000C000&
      Caption         =   "Return to the Main"
      Height          =   975
      Left            =   12360
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6600
      Width           =   1935
   End
   Begin VB.CommandButton cmdPicture 
      BackColor       =   &H0000C000&
      Caption         =   "Pictures of Geese and Ducks"
      Height          =   975
      Left            =   11760
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2280
      Width           =   1455
   End
   Begin VB.CommandButton cmdSort 
      BackColor       =   &H0000C000&
      Caption         =   "Sort the Waterfowl Seasons in Order."
      Height          =   975
      Left            =   13320
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1200
      Width           =   1335
   End
   Begin VB.CommandButton cmdAreas 
      BackColor       =   &H0000C000&
      Caption         =   "Suggested areas for Waterfowl Hunting in Minnesota"
      Height          =   975
      Left            =   13320
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton cmdAmmunition 
      BackColor       =   &H0000C000&
      Caption         =   "Ammunition qualifications for waterfowl hunting"
      Height          =   975
      Left            =   11760
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1200
      Width           =   1455
   End
   Begin VB.CommandButton cmdSeason 
      BackColor       =   &H0000C000&
      Caption         =   "Waterfowl Seasons for Minnesota"
      Height          =   975
      Left            =   11760
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
   Begin VB.PictureBox results 
      BackColor       =   &H000080FF&
      Height          =   7575
      Left            =   120
      ScaleHeight     =   7515
      ScaleWidth      =   11475
      TabIndex        =   0
      Top             =   120
      Width           =   11535
   End
End
Attribute VB_Name = "formWaterfowl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Game I Enjoy Hunting in Minnesota
'formWaterfowl
'Mark Hines
'3-23-09
'This is a sub form that allows the user to specifically target bear information from the GUI.
Option Explicit

Private Sub cmdAmmunition_Click()
Dim ammunition As String, steel As String
results.Cls
results.Picture = LoadPicture("")
'obtains information as to what type of ammunition the user would like to use.
ammunition = InputBox("what type of ammunition do you think you should use?")
'response of user is checked and if the answer was correct then message printed in picture box.
If ammunition = "steel" Then
    results.Print "Steel is the only legal ammunuition to use due to lead causing lead poisioning."
Else
    results.Print " NO! NO! NO! you have to use steel or you will kill the wildlife with lead poisoning."
End If
'more information displayed regardless of answer.
results.Print "It is suggested you use a heavier shot in your ammunition, using #2,4,BB,T are all good "
results.Print "when hunting waterfowl. This reduces the likelyhood of a crippled bird."
End Sub

'This bit of code is a click and see type design where the user clicks on the command button and the
'information is printed in the picture box for the user to read.

Private Sub cmdAreas_Click()
results.Cls
results.Picture = LoadPicture("")
results.Print "**************************************************************************************************************************************"
results.Print "Wildlife Mangement Areas (WMAs), National wildlife refuges, Wildlife Production Areas (WPAs)."
results.Print "Some of the best areas to hunt are overlooked public lands."
results.Print "If you do hunt private make sure you have permission from the land owner."
results.Print "More information on possible hunting land for Waterfowl, as well as, other forms of game can be found on Minnesota's DNR website"
results.Print "http://www.dnr.state.mn.us/hunting/tips/locations.html."
results.Print "**************************************************************************************************************************************"
End Sub

Private Sub cmdBag_Click()
Dim baglimit(1 To 5) As Integer, game(1 To 5) As String, tempLimit As Integer, I As Integer, ctr As Integer
Dim found As Boolean
'clears all contents in the picture box
results.Cls
results.Picture = LoadPicture("")
'opens the array and applies variables to the names in array
Open App.Path & "\game.txt" For Input As #1
    ctr = 0
        Do While Not EOF(1)
        ctr = ctr + 1
        Input #1, game(ctr), baglimit(ctr)
    Loop
Close #1
'finds specific name in the array
I = 0
found = False
For I = 1 To ctr
If "waterfowl" = game(I) Then
    tempLimit = baglimit(I)
    found = True
    
End If
Next I
'displays the information if the word was not found in the array. displayed in message box.
    If Not found Then
        MsgBox "There arent any animals under that name in the directory.", , "Baglimit"
'displays the bag limit and message if word was found. displayed in message box.
    Else
        MsgBox "waterfowl have a baglimit of " & tempLimit & "this is for most waterfowl please see your state regulations for more information on the correct number per species of waterfowl.", , "Waterfowl bag limit"
        
    End If

End Sub

Private Sub cmdBagG_Click()
Dim baglimit(1 To 5) As Integer, game(1 To 5) As String, tempLimit As Integer, I As Integer, ctr As Integer
Dim found As Boolean
'clear contents of picture box
results.Cls
results.Picture = LoadPicture("")
'opens and loads the array while assigning the variables to each.
Open App.Path & "\game.txt" For Input As #1
    ctr = 0
        Do While Not EOF(1)
        ctr = ctr + 1
        Input #1, game(ctr), baglimit(ctr)
    Loop
Close #1
'search for particular species in the list
I = 0
found = False
For I = 1 To ctr
If "waterfowl" = game(I) Then
    tempLimit = baglimit(I)
    found = True
    
End If
Next I
'print the message if not found in the array as a message box.
    If Not found Then
        MsgBox "There arent any animals under that name in the directory.", , "Baglimit"
'prints the results if the particular word was found in the array in a message box.
    Else
        MsgBox "Geese have a baglimit of " & tempLimit & "this is for most waterfowl please see your state regulations for more information on the correct number per species of waterfowl.", , "Waterfowl bag limit"
        
    End If

End Sub

'function will simply print out the picture of the illegal waterfowl of the duck species.

Private Sub cmdIllegal_Click()
results.Cls
results.Picture = LoadPicture("")
results.Picture = LoadPicture(App.Path & "\canvas_back.jpg")
Close

End Sub

'this will print a large picture of many species of ducks and geese for the user once clicked.

Private Sub cmdPicture_Click()
results.Cls
results.Picture = LoadPicture("")
results.Picture = LoadPicture(App.Path & "\2450-2290.jpg")
Close
End Sub

'ends program alltogether

Private Sub cmdQuit_Click()
End
End Sub

'This will return you to the main screen of the GUI

Private Sub cmdReturn_Click()
formWaterfowl.Hide
MainForm.Show

End Sub

'This prints out the results of the waterfowl season.  this prints out in the picture box.

Private Sub cmdSeason_Click()
results.Cls
results.Picture = LoadPicture("")
results.Print "Goose - Spring Light Goose 03/01/09 - 04/30/09;"
results.Print "Early Canada Goose (Tentative) 09/05/09 - 09/22/09;"
results.Print "Waterfowl season opener 10/03/09 - 11-28-09."

End Sub

Private Sub cmdSort_Click()
Dim ctr As Single, pass As Integer, pos As Integer, day(1 To 5) As Single, tempsea As String, tempdate As Single, J As Integer
Dim season(1 To 3) As String
'Clear all picture box contents
results.Cls
results.Picture = LoadPicture("")
'open up the array and load it.
Open App.Path & "\WaterFowlSeason.txt" For Input As #1
ctr = 0
Do While Not EOF(1)
    ctr = ctr + 1
    Input #1, season(ctr), day(ctr)
Loop
'once the array is open it is sorted according to start date of the season.
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

    results.Print "Game."; Tab(20); "Start Date"
    results.Print "**********************************************************************************************"
'results for all items in the array are printed in the picture box.
For J = 1 To ctr
    results.Print season(J); Tab(20); day(J)
Next J
   
End Sub

Private Sub cmdTotal_Click()
Dim duckTotal As Single, gooseTotal As Single, avg As Single, ctr As Integer, bird As Single
Dim duck As Single, goose As Single
'Clear all information in the picture box
results.Cls
results.Picture = LoadPicture("")
'input what kind of bird you are choosing to talk about by entering in a specific number for one or the other
bird = InputBox("What kind of waterfowl did you shoot? enter a 1 for duck and 2 for a goose.")

If bird = "1" Then
'Then asks for information about how many
    duck = InputBox("How many ducks have you shot so far this year?")
        duckTotal = duckTotal + duck
   'If the number falls within one of the ranges the appropriate message appears.
   Select Case duckTotal
    Case Is >= 100
        results.Print "You are a slaughtering king. Better start eating them or the freezer will over flow."
    Case Is >= 80
        results.Print "You are doing well and havent reached your full potential yet."
    Case Is >= 60
        results.Print "Having a good year but still plenty of time left in the season."
    Case Is >= 40
        results.Print "Keep it up, must still be early season you should have more than that if the weather is cold."
    Case Is >= 20
        results.Print "Forget the slough, you need to go shoot some trap to get your shots down."
    Case Is > 0
        results.Print "try scouting and maybe try learning to call cause that is pathetic."
    Case Is < 0
        results.Print "dont even try and call yourself a hunter."
    End Select
End If
'another option for geese in this case
If bird = "2" Then
    goose = InputBox("How many Geese have you shot so far this year?")
        gooseTotal = gooseTotal + goose
    'after more information is input then the values is checked to see if it falls into a range. if so, the message is printed.
Select Case gooseTotal
    Case Is >= 100
        results.Print "You are a slaughtering king. Better start eating them or the freezer will over flow."
    Case Is >= 80
        results.Print "You are doing well and havent reached your full potential yet."
    Case Is >= 60
        results.Print "Having a good year but still plenty of time left in the season."
    Case Is >= 40
        results.Print "Keep it up, must still be early season you should have more than that if the weather is cold."
    Case Is >= 20
        results.Print "Forget the slough, you need to go shoot some trap to get your shots down."
    Case Is > 0
        results.Print "try scouting and maybe try learning to call cause that is pathetic."
    Case Is < 0
        results.Print "dont even try and call yourself a hunter."
    End Select
End If
End Sub
