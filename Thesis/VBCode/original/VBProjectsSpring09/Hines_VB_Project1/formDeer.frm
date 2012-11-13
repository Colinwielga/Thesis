VERSION 5.00
Begin VB.Form formDeer 
   BackColor       =   &H00004000&
   Caption         =   "Deer Information for Minnesota"
   ClientHeight    =   7125
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13515
   LinkTopic       =   "Form2"
   ScaleHeight     =   7125
   ScaleWidth      =   13515
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdTotalAvg 
      BackColor       =   &H0000C000&
      Caption         =   "Total the # of Deer Harvested in the past 3 year and find your aveage per year."
      Height          =   855
      Left            =   9960
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3000
      Width           =   3495
   End
   Begin VB.CommandButton cmdBag 
      BackColor       =   &H0000C000&
      Caption         =   "Bag limits for Whitetail Deer in Minnesota"
      Height          =   855
      Left            =   11760
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2040
      Width           =   1695
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H000000FF&
      Caption         =   "Quit"
      Height          =   855
      Left            =   10560
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3960
      Width           =   2295
   End
   Begin VB.CommandButton cmdDeerPic 
      BackColor       =   &H0000C000&
      Caption         =   "Picture of Whitetail Deer"
      Height          =   855
      Left            =   9960
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2040
      Width           =   1695
   End
   Begin VB.PictureBox Results 
      BackColor       =   &H000080FF&
      Height          =   6855
      Left            =   120
      ScaleHeight     =   6795
      ScaleWidth      =   9675
      TabIndex        =   5
      Top             =   120
      Width           =   9735
   End
   Begin VB.CommandButton cmdReturn 
      BackColor       =   &H0000C000&
      Caption         =   "Return to main screen"
      Height          =   855
      Left            =   10560
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4920
      Width           =   2295
   End
   Begin VB.CommandButton cmdLocation 
      BackColor       =   &H0000C000&
      Caption         =   "Possible Hunting Land locations"
      Height          =   855
      Left            =   9960
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1080
      Width           =   1695
   End
   Begin VB.CommandButton cmdBullet 
      BackColor       =   &H0000C000&
      Caption         =   "Legal Calibers For Deer Hunting "
      Height          =   855
      Left            =   11760
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1080
      Width           =   1695
   End
   Begin VB.CommandButton cmdZoning 
      BackColor       =   &H0000C000&
      Caption         =   "Map For Deer Hunting Zones"
      Height          =   855
      Left            =   11760
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton cmdSeason 
      BackColor       =   &H0000C000&
      Caption         =   "Various Different Seasons For Deer Hunting"
      Height          =   855
      Left            =   9960
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "formDeer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Game I Enjoy Hunting in Minnesota
'FormDeer
'Mark Hines
'3-23-09
'This form provides the information necessary to obtain information about deer hunting regulations in Minnesota.
'It also includes information for novice hunters that might prove to be useful.  Some regulations are also displayed.
'All commands clear screen when run to avoid overlaping images with text.
Option Explicit

'This command allows the user to find out about the bag limit for deer in minnesota and how many they may legally take.
'This command button is a simple click and print where user presses the button and the information appears on the picture box.

Private Sub cmdBag_Click()
results.Cls
results.Picture = LoadPicture("")
results.Print "Lottery deer areas: The bag limit is one deer total per year, regardless of license type. "
results.Print "Bonus permits are not valid in lottery deer areas. Managed deer areas: The bag limit for "
results.Print "managed deer areas is two deer and hunters can use any combination of valid licenses or "
results.Print "permits to tag both deer. Intensive deer areas: Using any combination of licenses and "
results.Print "permits, the bag limit for intensive deer areas is five deer. Early antlerless areas: "
results.Print "Up to two deer can be taken in addition to the statewide limit of five."
End Sub

Private Sub cmdBullet_Click()
Dim caliber As Single
'Clear picture box
results.Cls
results.Picture = LoadPicture("")
'Obtain information from user about what caliber they would like to use for deer hunting
caliber = InputBox("what caliber rifle are you planning on using for deer hunting? (enter your information with just bullet size in .30 style.)")
'check to see if caliber was above the legal minimum and if it was then would print the message.
If caliber >= 0.223 Then
    results.Print "This is a legal caliber for deer hunting in Minnesota."
'this is what would be printed in the picture box if the answer was below the minimum caliber.
    Else
    results.Print "This is a non-legal caliber, You need to use a larger caliber bullet for large game."
    results.Print "otherwise the caliber was not input in correct form."
  End If
End Sub

'this particular function allows the user to view a picture of a common whitetailed deer.

Private Sub cmdDeerPic_Click()
results.Cls
results.Picture = LoadPicture(App.Path & "\Marcheldeer_200.jpg")
End Sub

'This command will display the information as written in the picResults screen for the user to read.

Private Sub cmdLocation_Click()
results.Cls
results.Picture = LoadPicture("")
results.Print "**************************************************************************************************************************************"
results.Print "Wildlife Mangement Areas (WMAs), National wildlife refuges, Wildlife Production Areas (WPAs), National forests, Industrial forest land"
results.Print "Shooting preserves, County land, and State forests"
results.Print "**************************************************************************************************************************************"
results.Print "More information on possible hunting land for Deer, as well as, other forms of game can be found on Minnesota's DNR website"
results.Print "http://www.dnr.state.mn.us/index.html."
results.Print "**************************************************************************************************************************************"
End Sub

'Quits program

Private Sub cmdQuit_Click()
End
End Sub

'Returns user to main from

Private Sub cmdReturn_Click()
formDeer.Hide
MainForm.Show
End Sub

'provides the user with season information (start and end dates) for deer hunting through a msg box.

Private Sub cmdSeason_Click()
results.Cls
results.Picture = LoadPicture("")
MsgBox "For Deer: 9-19-09 through 12--13-09 is Minnesotas archery season, 11-7-09 through 11-21-09 is Minnesotas rifle season, and 11-28-09 through 12-13-09 is Minnesotas muzzleloader season.", , "Deer Hunting Season"
End Sub



Private Sub cmdTotalAvg_Click()
Dim deer As Integer, total As Integer, avg As Single, ctr As Integer
results.Cls
results.Picture = LoadPicture("")
'user is asked to input specific information to them and is then totaled and averaged and rounded up to make them feel better.
deer = InputBox("How many deer did you shoot in 2007? (press enter when number has been entered)")
    ctr = deer + ctr
deer = InputBox("How many deer did you shoot in 2008? (press enter when number has been entered)")
    ctr = deer + ctr
deer = InputBox("How many deer did you shoot in 2009? (press enter when number has been entered)")
    total = deer + ctr
    avg = total / 3
MsgBox "You have shot " & total & " deer in the past 3 years and are roughly averaging " & Round(avg) & " deer per year."
End Sub

'This function displays the statewide map of all the zones for deer hunting in Minnesota.

Private Sub cmdZoning_Click()
results.Cls
results.Picture = LoadPicture("")
results.Picture = LoadPicture(App.Path & "\zonemap.bmp")
End Sub


