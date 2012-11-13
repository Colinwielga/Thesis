VERSION 5.00
Begin VB.Form formGrouse 
   BackColor       =   &H00004000&
   Caption         =   "Minnesota Grouse Hunting Information"
   ClientHeight    =   7620
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14385
   LinkTopic       =   "Form1"
   ScaleHeight     =   7620
   ScaleWidth      =   14385
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGrouse 
      BackColor       =   &H0000C000&
      Caption         =   "Pictures of Grouse"
      Height          =   735
      Left            =   11280
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3960
      Width           =   3015
   End
   Begin VB.CommandButton cmdArea 
      BackColor       =   &H0000C000&
      Caption         =   "Suggested Areas to Hunt Grouse (Capitalise name of particular bird)"
      Height          =   735
      Left            =   11280
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3120
      Width           =   3015
   End
   Begin VB.CommandButton cmdDogs 
      BackColor       =   &H0000C000&
      Caption         =   "Types of Dogs that are great for Grouse Hunting"
      Height          =   735
      Left            =   11280
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2280
      Width           =   3015
   End
   Begin VB.CommandButton cmdReturn 
      BackColor       =   &H0000C000&
      Caption         =   "Return to Main"
      Height          =   735
      Left            =   12840
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4800
      Width           =   1455
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H000000FF&
      Caption         =   "Quit"
      Height          =   735
      Left            =   11280
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4800
      Width           =   1455
   End
   Begin VB.CommandButton cmdBag 
      BackColor       =   &H0000C000&
      Caption         =   "Bag Limits for Grouse"
      Height          =   735
      Left            =   12840
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1440
      Width           =   1455
   End
   Begin VB.CommandButton cmdAmmo 
      BackColor       =   &H0000C000&
      Caption         =   "Suggested Amunition"
      Height          =   735
      Left            =   11280
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1440
      Width           =   1455
   End
   Begin VB.TextBox txtType 
      Height          =   375
      Left            =   11280
      TabIndex        =   2
      Top             =   360
      Width           =   3135
   End
   Begin VB.CommandButton cmdType 
      BackColor       =   &H0000C000&
      Caption         =   "Click here to find out more about Sharptail or Ruffed Grouse (Capitalize the species name)"
      Height          =   615
      Left            =   11400
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   720
      Width           =   2895
   End
   Begin VB.PictureBox results 
      BackColor       =   &H000080FF&
      Height          =   7335
      Left            =   120
      ScaleHeight     =   7275
      ScaleWidth      =   10995
      TabIndex        =   0
      Top             =   120
      Width           =   11055
   End
   Begin VB.Label Label1 
      Caption         =   "What type of Grouse do you wish to hunt?"
      Height          =   255
      Left            =   11280
      TabIndex        =   3
      Top             =   120
      Width           =   3135
   End
End
Attribute VB_Name = "formGrouse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Game I Enjoy Hunting in Minnesota
'FormGrouse
'Mark Hines
'3-23-09
'This form provides the information necessary to obtain information about grouse hunting and its regulations in Minnesota.
'It also includes information for novice hunters that might prove to be useful.  Some regulations are also displayed.
'All commands clear screen when run to avoid overlaping images with text.
Option Explicit
Private Sub cmdAmmo_Click()
'Clear picture box and displays message in message box
results.Cls
results.Picture = LoadPicture("")
MsgBox "Grouse are a relatively small bird with very few feathes.  Therefore, i would suggest 2 3/4in. #8s.  There is more than enough kick in a .410 to kill a grouse so it is up to the shooter.  Gauges often range from .410 to 12 gauge.", , "Ammunition"
End Sub

Private Sub cmdArea_Click()
Dim Grouse As String
results.Cls
results.Picture = LoadPicture("")
'obtain information from user about what type of grouse.
Grouse = InputBox("What type of grouse are you looking to hunt?")
'if the input satisfies requirement then message printed
If Grouse = "Sharptail" Then
    MsgBox "These birds are native to southern Minnesota as well as western.  they tend to prefer prarie grasses over wooded areas.  You may want to check Minnesotas DNR website for possible public hunting lands within those areas of Minnesota.", , "Sharptail Grouse"
ElseIf Grouse = "Ruffed" Then
    MsgBox "These birds are native to northern Minnesota.  Little Falls and Northward often hold vast quanitities of ruffed grouse.  They tend to prefer poplar groves and other wooded areas with a lot of cover.  You may want to check Minnesotas DNR website for possible public hunting lands within those areas of Minnesota.", , "Ruffed Grouse"
'THis is what you would get if you entered invalid data
Else
    MsgBox "Your input was not valid for the species of Grouse.", , "Error"
End If
End Sub

Private Sub cmdBag_Click()
Dim game(1 To 5) As String, baglimit(1 To 5) As Integer, I As Integer, found As Boolean, tempLimit As Single, ctr As Integer
results.Cls
results.Picture = LoadPicture("")
'opens the array of animals and the daily bag limits
Open App.Path & "\game.txt" For Input As #1
    ctr = 0
Do While Not EOF(1)
    ctr = ctr + 1
    Input #1, game(ctr), baglimit(ctr)
  Loop
Close #1
'searches for specific species in the array
I = 0
found = False
For I = 1 To ctr
If "Grouse" = game(I) Then
    tempLimit = baglimit(I)
    found = True
    
End If
Next I
'if the species was found in the array then this message would pop up.
    If Not found Then
        MsgBox "There arent any animals under that name in the directory.", , "Baglimit"
'other wise if the species was found this would be printed in the window.
    Else
        results.Print "Grouse have a baglimit of "; tempLimit; " This bag limit can be "
        results.Print "associated with either species of grouse."
        
    End If
End Sub

Private Sub cmdDogs_Click()
Dim Dog(1 To 7) As String, Specialty(1 To 7) As String, I As Integer, found As Boolean, tempLimit As Single, whoctr As Integer, ctr As Integer
results.Cls
results.Picture = LoadPicture("")
'opens the array of dogs and their specific expertise
Open App.Path & "\Dogs.txt" For Input As #1
    ctr = 0
Do While Not EOF(1)
ctr = ctr + 1
Input #1, Dog(ctr), Specialty(ctr)
Loop
Close #1
'print header in picture box
    results.Print "Dog"; Tab(20); "Speciality"
        results.Print "********************************************************************************"
'obtains and prints specified information
I = 0
For I = 1 To ctr
  If "Upland" = Specialty(I) Then
        results.Print Dog(I); Tab(20); "Upland"
  End If
Next I
results.Print "********************************************************************************"
End Sub

Private Sub cmdGrouse_Click()
'prints the specific picture from the file
results.Cls
results.Picture = LoadPicture("")
results.Picture = LoadPicture(App.Path & "\Marchelgrouse_200.jpg")
End Sub
'ends the program entirely
Private Sub cmdQuit_Click()
End
End Sub
'switches the forms to the main GUI
Private Sub cmdReturn_Click()
formGrouse.Hide
MainForm.Show
End Sub

Private Sub cmdType_Click()
Dim Grouse As String
results.Cls
results.Picture = LoadPicture("")
'tells what the text should be represented by after being entered by the user
Grouse = txtType.Text
'find if the imput text matches any of the possible outcomes and if so print the results in a picture box.
  If Grouse = "Sharptail" Then
    results.Print "Sharptail season is the same as Ruffed but the habitat is very different between the two."
    results.Print "Sharptail grouse tend to live in prarie grass type habitats and can be hunted in much "
    results.Print "the same way as Phesant.  The Sharptail season runs from Sept. 19th through Oct. 30th."
  ElseIf Grouse = "Ruffed" Then
    results.Print "Ruffed grouse are common in the northern areas of Minnesota, and are most commonly found in "
    results.Print "poplar goves.  They are very fast fliers and have thunderous wing beats when taking off.  "
    results.Print "The ruffed grouse season start on Sept.19th through jan. 3rd."
'result if the information input is incorrect.
Else
    MsgBox "You have entered and invalid entry. Please try again.", , "ERROR"
End If
End Sub
