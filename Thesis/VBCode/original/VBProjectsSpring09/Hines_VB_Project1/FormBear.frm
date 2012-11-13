VERSION 5.00
Begin VB.Form FormBear 
   BackColor       =   &H00004000&
   Caption         =   "Minnesota Bear Hunting Information"
   ClientHeight    =   7620
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13005
   LinkTopic       =   "Form1"
   ScaleHeight     =   7620
   ScaleWidth      =   13005
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBag 
      BackColor       =   &H0000C000&
      Caption         =   "Bag Limits for Black Bears in Minnesota"
      Height          =   975
      Left            =   9840
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3360
      Width           =   1455
   End
   Begin VB.CommandButton cmdTotal 
      BackColor       =   &H0000C000&
      Caption         =   "Total Number of Bears Put Down and your average per year."
      Height          =   975
      Left            =   11520
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2280
      Width           =   1455
   End
   Begin VB.CommandButton cmdRegulations 
      BackColor       =   &H0000C000&
      Caption         =   "Minnesota Big Game Regulations"
      Height          =   975
      Left            =   9840
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H000000FF&
      Caption         =   "Quit"
      Height          =   855
      Left            =   9840
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5520
      Width           =   2775
   End
   Begin VB.CommandButton cmdReturn 
      BackColor       =   &H0000C000&
      Caption         =   "Return to Main "
      Height          =   855
      Left            =   9840
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6480
      Width           =   2775
   End
   Begin VB.CommandButton cmdLottery 
      BackColor       =   &H0000C000&
      Caption         =   "Lottery Information for  Bear Hunting"
      Height          =   975
      Left            =   11520
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1200
      Width           =   1455
   End
   Begin VB.CommandButton cmdSeason 
      BackColor       =   &H0000C000&
      Caption         =   "BlackBear hunting season in Minnesota"
      Height          =   975
      Left            =   9840
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1200
      Width           =   1455
   End
   Begin VB.CommandButton cmdhunting 
      BackColor       =   &H0000C000&
      Caption         =   "Suggested areas to hunt Black Bear"
      Height          =   975
      Left            =   11520
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton cmdPicBear 
      BackColor       =   &H0000C000&
      Caption         =   "Picture of a Black Bear"
      Height          =   975
      Left            =   9840
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2280
      Width           =   1455
   End
   Begin VB.PictureBox Results 
      BackColor       =   &H000080FF&
      Height          =   7335
      Left            =   120
      ScaleHeight     =   7275
      ScaleWidth      =   9555
      TabIndex        =   0
      Top             =   120
      Width           =   9615
   End
End
Attribute VB_Name = "FormBear"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Game I Enjoy Hunting in Minnesota
'FormBear
'Mark Hines
'3-23-09
'This form provides the information necessary to obtain information about bear hunting regulations in Minnesota.
'It also includes information for novice hunters that might prove to be useful.  Some regulations are also displayed.
'All commands clear screen when run to avoid overlaping images with text.
Option Explicit
Dim bears As Integer, sum As Integer, total As Integer, avg As Single, ctr As Integer

Private Sub cmdBag_Click()
'clear main picture box and tells you the exact limit of bear in minnesotas hunting season
results.Cls
results.Picture = LoadPicture("")
results.Print "There is a limit of one bear per season in Minnesota."
End Sub

Private Sub cmdhunting_Click()
'clear picture box and print data with the click of the button
results.Cls
results.Picture = LoadPicture("")
results.Print "**************************************************************************************************************************************"
results.Print "Wildlife Mangement Areas (WMAs), National wildlife refuges, Wildlife Production Areas (WPAs), National forests, Industrial forest land"
results.Print "Shooting preserves, County land, and State forests. You may want to focus on the Northern areas of Minnesota."
results.Print "**************************************************************************************************************************************"
results.Print "More information on possible hunting land for bear, as well as, other forms of game can be found on Minnesota's DNR website"
results.Print "http://www.dnr.state.mn.us/index.html."
results.Print "**************************************************************************************************************************************"
End Sub

Private Sub cmdLottery_Click()
'clear picture box and displays information as seen through a message box
results.Cls
results.Picture = LoadPicture("")
MsgBox "Applications available late March. Deadline first Friday in May. Lottery results available end of May. This information is posted on the MN DNR website."
End Sub

Private Sub cmdPicBear_Click()
'clear screen and displays picture of a black bear
results.Cls
results.Picture = LoadPicture("")
results.Picture = LoadPicture(App.Path & "\blackbear.jpg")
End Sub
'end the whole program
Private Sub cmdQuit_Click()
End
End Sub
'specific regulations can be seen at the website displayed in a message box.
Private Sub cmdRegulations_Click()
results.Cls
results.Picture = LoadPicture("")
MsgBox ("The Regulations can be found on the DNR website at--> http://files.dnr.state.mn.us/rlp/regulations/hunting/2008/full_regs.pdf#page=95")
End Sub
'change back to the main form.
Private Sub cmdReturn_Click()
FormBear.Hide
MainForm.Show
End Sub

Private Sub cmdSeason_Click()
'displays the season for bear in a message box
MsgBox ("The Black Bear season starts on September 1st and continues through October 12th. You may start baiting bear as of August 15th.")
End Sub

Private Sub cmdTotal_Click()
results.Cls
results.Picture = LoadPicture("")
'input information from user concerning the amount of bears killed in the past 4 years and then averages them and prints a rounded result for total and avg.
bears = InputBox("How many bears did you shoot in 2006?")
    ctr = bears + ctr
bears = InputBox("How many bears did you shoot in 2007?")
    ctr = bears + ctr
bears = InputBox("How many bears did you shoot in 2008?")
    ctr = bears + ctr
bears = InputBox("How many bears did you shoot in 2009?")
    total = bears + ctr
    avg = total / 4
MsgBox "You have shot " & total & " bear in the past 4 years and are roughly averaging " & Round(avg) & " bears per year."
End Sub
