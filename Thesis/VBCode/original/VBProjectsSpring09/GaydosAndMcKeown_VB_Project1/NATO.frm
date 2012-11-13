VERSION 5.00
Begin VB.Form frmNATO 
   Caption         =   "Form1"
   ClientHeight    =   11010
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   14565
   LinkTopic       =   "Form1"
   Picture         =   "NATO.frx":0000
   ScaleHeight     =   11010
   ScaleWidth      =   14565
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdWarsaw 
      Caption         =   "Warsaw Pact"
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   6120
      Picture         =   "NATO.frx":153CE
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   9720
      Width           =   4575
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "SURRENDER!"
      BeginProperty Font 
         Name            =   "Myriad Web Pro Condensed"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   2760
      Picture         =   "NATO.frx":17D05
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   9600
      Width           =   2535
   End
   Begin VB.PictureBox picGun 
      Height          =   3615
      Left            =   6000
      ScaleHeight     =   3555
      ScaleWidth      =   8235
      TabIndex        =   6
      Top             =   5760
      Visible         =   0   'False
      Width           =   8295
   End
   Begin VB.PictureBox picStats 
      Height          =   3615
      Left            =   240
      ScaleHeight     =   3555
      ScaleWidth      =   5235
      TabIndex        =   5
      Top             =   5760
      Visible         =   0   'False
      Width           =   5295
   End
   Begin VB.CommandButton cmd762x51 
      Caption         =   "7.62 X 51MM NATO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4920
      Picture         =   "NATO.frx":3B967
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3960
      Width           =   5415
   End
   Begin VB.CommandButton cmd50cal 
      Caption         =   ".50 Caliber"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   3240
      Picture         =   "NATO.frx":5284D
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1440
      Width           =   8655
   End
   Begin VB.CommandButton cmd45ACP 
      Caption         =   ".45 ACP"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6360
      Picture         =   "NATO.frx":56285
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4920
      Width           =   2775
   End
   Begin VB.CommandButton cmd556x45 
      Caption         =   "5.56 X 45mm"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   4440
      Picture         =   "NATO.frx":5F327
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2760
      Width           =   6375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF0000&
      Caption         =   "North Atlantic Treaty Organization                        Munitions"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2640
      TabIndex        =   0
      Top             =   240
      Width           =   9615
   End
End
Attribute VB_Name = "frmNATO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'We have four arrays that contain the different weights, feet per second, and foot pounds of energy
'We needed to dim three variables per each array then
Dim fortyfive(1 To 10) As Single, fiddycal(1 To 10) As Single, Nato(1 To 10) As Single, fyfySix(1 To 10) As Single, ctr As Integer
Dim FVSize(1 To 10) As Single, FVspeed(1 To 10) As Single, FVstrength(1 To 10) As Single
Dim MSize(1 To 10) As Single, Mspeed(1 To 10) As Single, Mstrength(1 To 10) As Single
Dim fyfySize(1 To 10) As Single, fyfyspeed(1 To 10) As Single, fyfystrength(1 To 10) As Single
Dim fiddySize(1 To 10) As Single, fiddyspeed(1 To 10) As Single, fiddystrength(1 To 10) As Single
Dim GunPic(1 To 10) As Single

Private Sub cmd50Cal_Click()

ctr = 5 'There were five different types of rounds for this ammunition


    picStats.Visible = True 'We started the pictures off as invisible
    picStats.Cls  'Everytime a button was clicked, the statistics picture cleared
    picStats.Print "Size", Tab(25), "ft/sec", Tab(45), "lbs/ft"
    picStats.Print "------------------------------------------------------------------------------------------------------------------------------------------------"
'Above is how we began the table for the ammunition
  Dim x As Integer
   For x = 1 To ctr 'Below we matched the size, speed, and strength up with the labels above them in the table, we used the tab function to get them spaced out correctly
    picStats.Print fiddySize(x); " grains", Tab(25), fiddyspeed(x), Tab(45), FormatNumber((fiddystrength(x)), 0)
   Next x
'We used a For Next loop in order to print off all ammunition information
    picGun.Visible = True  'On top of already showing of the ammunition on the buttons, we decided that this project would miss the point if we did not add the weapons that use the different types of ammunition
    picGun.Cls 'We first made the gun picture slot visible, then we added the pictures!
    picGun.Picture = LoadPicture(App.Path & "\50_Cal.jpg") 'We finally managed to find out exactly how to load a picture without the need for an array
End Sub

Private Sub cmd45ACP_Click()

ctr = 4 'There are four different types of ammunition
    picStats.Visible = True 'If this were to be the first button clicked, it would make the pictures visible
    picStats.Cls 'this clears all previous ammunition statistics from the table
    picStats.Print "Size", Tab(25), "ft/sec", Tab(45), "lbs/ft"
    picStats.Print "------------------------------------------------------------------------------------------------------------------------------------------------"
'Above is how we made the table uniform, we used Tab(25) and Tab(45) to get the overall lengths
 Dim x As Integer
  For x = 1 To ctr
   picStats.Print FVSize(x); " grains", Tab(25), FVspeed(x), Tab(45), FormatNumber((FVstrength(x)), 0)
  Next x
'Again we used a For Next loop that found and arranged all the variables in the file

    picGun.Visible = True
    picGun.Cls
    picGun.Picture = LoadPicture(App.Path & "\45ACPSpringfield.jpg")
'Here we loaded a picture of a weapon that fires this type of ammunition
End Sub

Private Sub cmd556x45_Click()

ctr = 3 'There are three lines of information in this text
    picStats.Visible = True 'this makes the stats picture visible if it is the first ammunition selected
    picStats.Cls 'this clears the stats picture of any previous ammunition info before it
    picStats.Print "Size"; , Tab(25), "ft/sec", Tab(45), "lbs/ft"
    picStats.Print "------------------------------------------------------------------------------------------------------------------------------------------------"
'this gives the picStats an organized looking table
  Dim x As Integer
   For x = 1 To ctr
    picStats.Print fyfySize(x); " grains", Tab(25), fyfyspeed(x), Tab(45), FormatNumber((fyfystrength(x)), 0)
   Next x
'Again we used a For Next Loop in order to load all the files uniformily
    picGun.Visible = True 'if this is the first ammunition selected, it will make the picGun visible to the user
    picGun.Cls 'clears the picture box of any previous pictures
    picGun.Picture = LoadPicture(App.Path & "\5.56x45.jpg")
'Here we loaded a picture of a weapon that fires this type of ammunition
End Sub

Private Sub cmd762x51_Click()
ctr = 2 'There are two lines of info in this txt document
    picStats.Visible = True 'This makes the picbox visible
    picStats.Cls 'this clears the pic box
    picStats.Print "Size", Tab(25), "ft/sec", Tab(45), "lbs/ft"
    picStats.Print "------------------------------------------------------------------------------------------------------------------------------------------------"
'the above two lines is the header for the table
Dim x As Integer
  For x = 1 To ctr
   picStats.Print MSize(x); " grains", Tab(25), Mspeed(x), Tab(45), FormatNumber((Mstrength(x)), 0)
Next x
'the for next loop displays all of the information in the array
    picGun.Visible = True 'this makes the picbox visible
    picGun.Cls 'this clears the picture box
    picGun.Picture = LoadPicture(App.Path & "\762x51.jpg")
'above fills the picture box with weapony goodness
End Sub


Private Sub cmdReturn_Click()
MsgBox "What??? YOU COULDN'T HANDLE IT?"
End  'ENDS PROGRAM FOREVER.... until the next time that the user starts the program
End Sub

Private Sub cmdWarsaw_Click() 'This allows the user to switch to the Warsaw Pact form to see the warsaw pact ammunitions
frmWarsaw.Show 'this makes the Warsaw pact form visible
frmNATO.Hide 'this doesn't make the nato form visible, in fact it hides it

End Sub

Private Sub Form_Load()

'Below we opened up four different .txt documents in the same way
'we wanted to dim all the variables right away, so each form that has loaded has their variables with them
'there are twelve variables since each has three variables, those are the weight of the bullet, the speed of the bullet, and the knock down "foot pounds of energy" of each bullet
'there are two or more weights for each round which affect the FPS and FPE so we decided it would be better to put them in arrays
Open App.Path & "\45ACP.txt" For Input As #1 'This opens the array

Do While Not EOF(1) 'this Do While loop reads through the file and puts it into three arrays, the weight, speed, and hit power
    ctr = ctr + 1
    Input #1, FVSize(ctr), FVspeed(ctr), FVstrength(ctr)
Loop
Close #1
ctr = 0 'This puts the counter to zero, we had a problem loading the other arrays before, so we found that putting the counter to zero made everything run very smoothly
Open App.Path & "\50Cal.txt" For Input As #2

Do While Not EOF(2)
    ctr = ctr + 1
    Input #2, fiddySize(ctr), fiddyspeed(ctr), fiddystrength(ctr)
Loop
Close #2
'This is the second file, it is also put into three arrays just like the first one
ctr = 0
Open App.Path & "\556x45mm.txt" For Input As #3

Do While Not EOF(3)
    ctr = ctr + 1
    Input #3, fyfySize(ctr), fyfyspeed(ctr), fyfystrength(ctr)
Loop
Close #3
'This is the third file
ctr = 0
Open App.Path & "\762X51mm.txt" For Input As #4

Do While Not EOF(4)
    ctr = ctr + 1
    Input #4, MSize(ctr), Mspeed(ctr), Mstrength(ctr)
Loop
Close #4
'this is the fourth file


MsgBox "NATO AMMO IS LOCKED AND LOADED!!!!"
'This message box lets the user know that their ammo info is ready to roll, when it seconds as also letting the user know that the arrays were successfully loaded


End Sub


