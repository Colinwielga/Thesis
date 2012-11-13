VERSION 5.00
Begin VB.Form frmMonarchyGovernment 
   BackColor       =   &H00400000&
   Caption         =   "Monarchy and Government Attractions"
   ClientHeight    =   11025
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12150
   BeginProperty Font 
      Name            =   "Lucida Handwriting"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   11025
   ScaleWidth      =   12150
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdTowerBridge 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Tower Bridge"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Light"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5760
      Width           =   2415
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00808080&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Light"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   9480
      Width           =   2415
   End
   Begin VB.CommandButton cmdGoToPopularAttractions 
      BackColor       =   &H00808080&
      Caption         =   "Return to Popular Attractions Page"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Light"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   7800
      Width           =   2415
   End
   Begin VB.CommandButton cmdGoToHome 
      BackColor       =   &H00808080&
      Caption         =   "Return to Home Page"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Light"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   8640
      Width           =   2415
   End
   Begin VB.PictureBox picTubeResults 
      BackColor       =   &H00400000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   3360
      ScaleHeight     =   495
      ScaleWidth      =   4215
      TabIndex        =   6
      Top             =   9600
      Width           =   4215
   End
   Begin VB.TextBox txtResults 
      BackColor       =   &H00400000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2895
      Left            =   3480
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   7320
      Width           =   8175
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00400000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   6615
      Left            =   3480
      ScaleHeight     =   6615
      ScaleWidth      =   8175
      TabIndex        =   4
      Top             =   480
      Width           =   8175
   End
   Begin VB.CommandButton cmdWestminsterAbbey 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Westminster Abbey"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Light"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4440
      Width           =   2415
   End
   Begin VB.CommandButton cmdBuckinghamPalace 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Buckingham Palace"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Light"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1800
      Width           =   2415
   End
   Begin VB.CommandButton cmdTowerOfLondon 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Tower of London"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Light"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3120
      Width           =   2415
   End
   Begin VB.CommandButton cmdBigBen 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Big Ben and Parliament"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Light"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   480
      Width           =   2415
   End
End
Attribute VB_Name = "frmMonarchyGovernment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Project Name: London Attractions
'Form Name: Monarchy and Government
'Author: Heather Arnhalt
'Date Written: October 18, 2009
'Objective: The user can click on the attraction they would like to learn more about and a picture of that attraction is displayed
'along with information about that attraction.

Private Sub cmdBigBen_Click()

    'clear both picture boxes
    picResults.Cls
    picTubeResults.Cls
    
    'Display a picture of the attraction
    picResults.Picture = LoadPicture(App.Path & "\London_Pictures\Parliament.jpg")
    
    'Print information about the attraction in the text results box
    txtResults.Text = "The Big Ben clock tower is attached to the Houses of Parliament and has become a familiar and much loved landmark, with its great bell chiming on the hour (and every quarter of an hour too) keeping time with Greenwich meantime. The name Big Ben was initially given to the Great Bell which was created at the Whitechapel Bell Foundry and first struck in 1859. During the 14th century what is now the Houses of Parliament housed the courts of law as well as shops and stalls selling legal equipment - wigs and pens. Following a fire in 1512, King Henry VIII abandoned the Palace and it has been home to the two seats of Parliament - the Commons and the Lords - ever since."

    'Print the tube station the attraction is nearest to in the PicTubeResults box
    picTubeResults.Print "Tube Station: Westminster"
    
End Sub


Private Sub cmdBuckinghamPalace_Click()

    'clear both picture boxes
    picResults.Cls
    picTubeResults.Cls
    
    'Display a picture of the attraction
    picResults.Picture = LoadPicture(App.Path & "\London_Pictures\BuckinghamPalace.jpg")
    
    'Print information about the attraction in the text results box
    txtResults.Text = "England's most famous royal palace, and the official residence of Queen Elizabeth II, Buckingham Palace opens the doors of its State Rooms to the public every summer. Originally acquired by King George III for his wife Queen Charlotte, Buckingham House was increasingly known as the 'Queen's House' and 14 of George III's children were born there. On his accession to the throne, George IV decided to convert the house into a palace and employed John Nash to help him extend the building. Queen Victoria was the first sovereign to live here (from 1837). The State Rooms are now still used by the Royal Family to receive and entertain guests on State and ceremonial occasions."

    'Print the tube station the attraction is nearest to in the PicTubeResults box
    picTubeResults.Print "Tube Station: Victoria"
    
End Sub

Private Sub cmdGoToHome_Click()
    'Hide the Monarchy and Government Attractions form and show the Home Page form
    frmMonarchyGovernment.Hide
    frmHomePage.Show
End Sub

Private Sub cmdGoToPopularAttractions_Click()
    'Hide the Monarchy and Government Attractions form and show the Popular Attractions form
    frmMonarchyGovernment.Hide
    frmPopularAttractions.Show
End Sub

Private Sub cmdQuit_Click()
    'End the program
    End
End Sub

Private Sub cmdTowerBridge_Click()

    'clear both picture boxes
    picResults.Cls
    picTubeResults.Cls
    
    'Display a picture of the attraction
    picResults.Picture = LoadPicture(App.Path & "\London_Pictures\TowerBridge.bmp")
    
    'Print information about the attraction in the text results box
    txtResults.Text = "Tower Bridge offers one of the best vantage points in the city from its spectacular high walkways, elevated 140ft above the Thames. The East Walkway boasts fantastic views of the Docklands and the elegant Canary Wharf, while from the West Walkway you can compare the mixed architectural styles of City Hall, the Tower Of London, St Paul's Cathedral, the City, Big Ben and the London Eye. Now galleried, these walkways were originally built to transport pedestrians across the Thames when the bridge was being lifted to let tall ships sail past."

    'Print the tube station the attraction is nearest to in the PicTubeResults box
    picTubeResults.Print "Tube Station: Tower Hill"


End Sub

Private Sub cmdTowerOfLondon_Click()
    
    'clear both picture boxes
    picResults.Cls
    picTubeResults.Cls
    
    'Display a picture of the attraction
    picResults.Picture = LoadPicture(App.Path & "\London_Pictures\toweroflondon.jpg")
    
    'Print information about the attraction in the text results box
    txtResults.Text = "At the Tower of London, admire Henry VIII's armour, weaponry and torture instruments in the White Tower before being dazzled by the array of royal jewels, crowns and diamonds encased in the Jewel and Martin Towers. With its stunning riverside backdrop, the tower has been used as a prison, palace and place of execution, arsenal, mint and menagerie, since its construction following the Norman Conquest of 1066. After King Henry VIII's break with the Catholic Church it housed religious prisoners including two of Henry VIII's six wives (Anne Boleyn and Catherine Howard), both of whom were beheaded on the scaffolds at Tower Green. It is now one of the most famous structures in the world and hosts a range of exhibitions and re-enactments which celebrate and represent some of the most spectacular aspects of its gory and glorious past."

    'Print the tube station the attraction is nearest to in the PicTubeResults box
    picTubeResults.Print "Tube Station: Tower Hill"

End Sub

Private Sub cmdWestminsterAbbey_Click()

    'clear both picture boxes
    picResults.Cls
    picTubeResults.Cls
    
    'Display a picture of the attraction
    picResults.Picture = LoadPicture(App.Path & "\London_Pictures\WestminsterAbbey.jpg")
    
    'Print information about the attraction in the text results box
    txtResults.Text = "Westminster Abbey is the setting for almost every coronation since 1066 - 38 to be precise - the abbey has also formed the burial ground for statesmen, scientists, musicians and poets. Approximately, 3,300 people are said to have been buried in the church, including Chaucer, Sir Isaac Newton, Laurence Olivier and Charles Dickens. Stunning Gothic architecture, the fascinating literary history represented by Poets' Corner, the artistic talent that went into the statues, murals, paintings and tombs and the fantastic stained glass combine to make Westminster Abbey the most enduringly stunning of London's churches and a treasure trove of royal history."

    'Print the tube station the attraction is nearest to in the PicTubeResults box
    picTubeResults.Print "Tube Station: Westminster"
    
End Sub
